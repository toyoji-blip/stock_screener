"""
app.py
======
株式財務スクリーニングツール - Web版
CLI（input()を順番に打つ）の代わりに、ブラウザのウィザード画面から
Step1〜Step4を進められるようにした Flask アプリ。

起動方法:
    pip install -r requirements.txt
    python app.py
    → ブラウザで http://127.0.0.1:5000 を開く

※ 個人利用のローカルツールのため、進行中の状態はサーバー側の
   メモリ（STATE）に1件だけ保持する単純な作りにしています。
   同時に複数人・複数タブで別銘柄を並行して進める用途には未対応です。
"""

import base64
import io
import os
import traceback
import uuid
from datetime import datetime

from flask import Flask, jsonify, render_template, request, send_from_directory
import matplotlib.pyplot as plt

import screening_core as core

app = Flask(__name__)

# 進行中の1銘柄分の状態を保持する簡易ストア（個人利用ツール向け）
STATE = {}


@app.after_request
def add_no_cache_headers(response):
    """
    開発中にコードを修正しても、ブラウザにHTML/JSがキャッシュされたままだと
    変更が反映されずに古い挙動のまま動き続けてしまうことがある。
    個人利用のローカルツールでキャッシュの恩恵はほぼ無いため、
    常に最新のHTML/JSを取得させるようにする。
    """
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    return response


def fig_to_base64(fig):
    if fig is None:
        return None
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=100, bbox_inches="tight")
    buf.seek(0)
    return "data:image/png;base64," + base64.b64encode(buf.read()).decode("ascii")


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/step1", methods=["POST"])
def api_step1():
    """銘柄コードを受け取り、基本情報（名称・株価・決算月推定）を返す"""
    data = request.get_json(force=True) or {}
    code = (data.get("code") or "").strip()
    if not code:
        return jsonify({"error": "銘柄コードを入力してください。"}), 400

    try:
        basic = core.fetch_basic(code)
    except Exception as e:
        return jsonify({"error": f"取得に失敗しました: {e}"}), 400

    STATE.clear()
    STATE["code"] = code
    STATE["ticker"] = basic["ticker"]
    STATE["name"] = basic["name"]
    STATE["price"] = basic["price"]

    return jsonify(basic)


@app.route("/api/step2", methods=["POST"])
def api_step2():
    """決算月・予想DPSを確定し、予想利回りの計算 + 過去3年財務データを取得して返す"""
    if "ticker" not in STATE:
        return jsonify({"error": "先に銘柄コードを入力してください（Step1からやり直してください）。"}), 400

    data = request.get_json(force=True) or {}
    try:
        fiscal_month = int(data.get("fiscal_month"))
        if not (1 <= fiscal_month <= 12):
            raise ValueError
    except (TypeError, ValueError):
        return jsonify({"error": "決算月は1〜12の数値で入力してください。"}), 400

    forecast_dps = data.get("forecast_dps")
    try:
        forecast_dps = float(forecast_dps)
        if forecast_dps < 0:
            raise ValueError
    except (TypeError, ValueError):
        return jsonify({"error": "予想DPSは0以上の数値で入力してください。"}), 400

    price = STATE.get("price")
    fore_yield = (forecast_dps / price * 100) if price else 0

    try:
        fys, res = core.fetch_financials(STATE["ticker"], fiscal_month)
    except Exception as e:
        return jsonify({"error": f"財務データ取得に失敗しました: {e}"}), 400

    STATE["fiscal_month"] = fiscal_month
    STATE["forecast_dps"] = forecast_dps
    STATE["fore_yield"] = fore_yield
    STATE["fys"] = fys
    STATE["res"] = res

    rows = core.build_metric_rows(fys, res)

    return jsonify({
        "fys": fys,
        "fore_yield": fore_yield,
        "rows": rows,
    })


@app.route("/api/step3", methods=["POST"])
def api_step3():
    """短信からの手入力データ（現預金・配当支払総額）を受け取り、グラフとExcelを生成する"""
    required_keys = ["ticker", "name", "price", "forecast_dps", "fore_yield",
                      "fiscal_month", "fys", "res"]
    if not all(k in STATE for k in required_keys):
        return jsonify({"error": "手順が正しくありません。Step1からやり直してください。"}), 400

    data = request.get_json(force=True) or {}
    cash_vals = data.get("cash_vals") or []
    div_pay_vals = data.get("div_pay_vals") or []

    def to_float_list(vals, n):
        out = []
        for v in vals[:n]:
            if v in (None, ""):
                out.append(None)
                continue
            try:
                out.append(float(v))
            except (TypeError, ValueError):
                out.append(None)
        while len(out) < n:
            out.append(None)
        return out

    fys = STATE["fys"]
    cash_vals = to_float_list(cash_vals, len(fys))
    div_pay_vals = to_float_list(div_pay_vals, len(fys))

    res = STATE["res"]
    name = STATE["name"]
    code = STATE["code"]
    price = STATE["price"]
    forecast_dps = STATE["forecast_dps"]
    fore_yield = STATE["fore_yield"]
    fiscal_month = STATE["fiscal_month"]
    ticker = STATE["ticker"]
    years_label = [f"FY{fy}" for fy in fys]

    fcf_vals = [res[fy].get("FCF") or 0 for fy in fys]
    debt_vals = [res[fy].get("有利子負債") or 0 for fy in fys]

    payout_vals = []
    for i, fy in enumerate(fys):
        fcf = res[fy].get("FCF")
        dv = div_pay_vals[i]
        if fcf and dv and fcf != 0:
            payout_vals.append(round(dv / fcf * 100, 1))
        else:
            payout_vals.append(None)

    div_pay_safe = [v or 0 for v in div_pay_vals]
    cash_safe = [v or 0 for v in cash_vals]

    fig_a = fig_b = fig_c = None
    try:
        fig_a = core.make_fig_a(name, years_label, fcf_vals, div_pay_safe, payout_vals)
        fig_b = core.make_fig_b(name, years_label, cash_safe, debt_vals, div_pay_safe)
        fig_c = core.make_fig_c(ticker, name, fiscal_month, forecast_dps, price, fys, res)

        images = {
            "fig_a": fig_to_base64(fig_a),
            "fig_b": fig_to_base64(fig_b),
            "fig_c": fig_to_base64(fig_c),
        }

        out_path = core.write_excel(
            name, code, price, forecast_dps, fore_yield,
            fys, res, years_label,
            cash_vals, div_pay_vals,
            fig_a, fig_b, fig_c,
        )
        filename = os.path.basename(out_path)

        # 履歴（過去の分析資料）として保存
        entry_id = uuid.uuid4().hex[:12]
        archive_images = core.save_archive_images(entry_id, fig_a, fig_b, fig_c)
        history_entry = {
            "id": entry_id,
            "code": code,
            "ticker": ticker,
            "name": name,
            "created_at": datetime.now(core.TZ_JP).isoformat(timespec="seconds"),
            "price": price,
            "forecast_dps": forecast_dps,
            "fore_yield": fore_yield,
            "fiscal_month": fiscal_month,
            "fys": fys,
            "res": {str(fy): v for fy, v in res.items()},
            "cash_vals": cash_vals,
            "div_pay_vals": div_pay_vals,
            "filename": filename,
            "images": archive_images,
        }
        core.append_history_entry(history_entry)
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"グラフ・Excel作成に失敗しました: {e}"}), 500
    finally:
        for f in [fig_a, fig_b, fig_c]:
            if f is not None:
                plt.close(f)

    return jsonify({
        "images": images,
        "filename": filename,
        "download_url": f"/download/{filename}",
        "history_id": entry_id,
        "fys": fys,
        "rows": core.build_metric_rows(fys, res, cash_vals, div_pay_vals),
    })


@app.route("/download/<path:filename>")
def download(filename):
    return send_from_directory(core.OUTPUT_DIR, filename, as_attachment=True)


@app.route("/archive/<path:relpath>")
def archive_file(relpath):
    return send_from_directory(core.ARCHIVE_DIR, relpath)


@app.route("/api/history")
def api_history_list():
    """過去の分析資料を検索する。qが空なら全件を新しい順で返す。"""
    q = request.args.get("q", "")
    entries = core.search_history(q)
    results = [{
        "id": e.get("id"),
        "code": e.get("code"),
        "ticker": e.get("ticker"),
        "name": e.get("name"),
        "created_at": e.get("created_at"),
        "price": e.get("price"),
        "fore_yield": e.get("fore_yield"),
        "filename": e.get("filename"),
    } for e in entries]
    return jsonify({"results": results, "count": len(results)})


@app.route("/api/history/<entry_id>")
def api_history_detail(entry_id):
    """過去の分析資料1件分の詳細（財務テーブル・グラフ・ダウンロードURL）を返す。"""
    entry = core.get_history_entry(entry_id)
    if not entry:
        return jsonify({"error": "指定された履歴データが見つかりませんでした。"}), 404

    fys = entry.get("fys") or []
    res_raw = entry.get("res") or {}
    # 履歴JSONではFYがJSON上の都合で文字列キーになっているためintに戻す
    res = {int(k): v for k, v in res_raw.items()}
    cash_vals = entry.get("cash_vals")
    div_pay_vals = entry.get("div_pay_vals")
    rows = core.build_metric_rows(fys, res, cash_vals, div_pay_vals)

    images = {k: f"/archive/{v}" for k, v in (entry.get("images") or {}).items()}

    return jsonify({
        "id": entry.get("id"),
        "code": entry.get("code"),
        "ticker": entry.get("ticker"),
        "name": entry.get("name"),
        "created_at": entry.get("created_at"),
        "price": entry.get("price"),
        "forecast_dps": entry.get("forecast_dps"),
        "fore_yield": entry.get("fore_yield"),
        "fiscal_month": entry.get("fiscal_month"),
        "fys": fys,
        "rows": rows,
        "cash_vals": cash_vals,
        "div_pay_vals": div_pay_vals,
        "images": images,
        "filename": entry.get("filename"),
        "download_url": f"/download/{entry.get('filename')}" if entry.get("filename") else None,
    })


@app.route("/api/history/<entry_id>", methods=["DELETE"])
def api_history_delete(entry_id):
    """過去の分析資料を1件削除する（Excel・アーカイブ画像も併せて削除）。"""
    ok = core.delete_history_entry(entry_id)
    if not ok:
        return jsonify({"error": "指定された履歴データが見つかりませんでした。"}), 404
    return jsonify({"status": "deleted", "id": entry_id})


@app.route("/api/batch", methods=["POST"])
def api_batch():
    """
    バッチ処理: Excelファイルをアップロードして全銘柄を一括スクリーニング。
    multipart/form-data で file フィールドに .xlsx を添付する。
    処理は同期実行（全銘柄完了後にレスポンスを返す）。
    """
    if "file" not in request.files:
        return jsonify({"error": "fileフィールドが見つかりません。"}), 400

    f = request.files["file"]
    if not f.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify({"error": "Excelファイル（.xlsx）をアップロードしてください。"}), 400

    try:
        file_bytes = f.read()
        result = core.run_batch(file_bytes)
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"バッチ処理に失敗しました: {e}"}), 500

    return jsonify({
        "succeeded":       result["succeeded"],
        "failed":          result["failed"],
        "total":           len(result["succeeded"]) + len(result["failed"]),
        "succeeded_count": len(result["succeeded"]),
        "failed_count":    len(result["failed"]),
    })


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)
