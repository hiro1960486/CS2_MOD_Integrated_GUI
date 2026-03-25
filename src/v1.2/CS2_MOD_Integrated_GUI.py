# ==========================================================
# CS2 MOD Cache Integrated Tool
# Version : v1.2
# ==========================================================

# -*- coding: utf-8 -*-
"""
CS2 PdxModsCache 統合ツール
- PdxModsCache.json のコピー
- PdxModsCache.json のExcel変換
- すべてGUIで選択式
- 前回指定したファイル/フォルダを次回起動時に復元
- 日本語 / English 切替対応
"""

from __future__ import annotations

import json
import os
import shutil
import sys
import threading
from pathlib import Path
from typing import Any, Dict, List, Optional, Set

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Excelセル文字数上限
EXCEL_CELL_LIMIT = 32767

# 設定保存ファイル名
SETTINGS_FILE_NAME = "CS2_MOD_Integrated_GUI_settings.json"


# =========================================
# 多言語表示
# =========================================
UI_TEXT = {
    "ja": {
        "window_title": "CS2 MOD Cache 統合ツール",
        "language_label": "言語 / Language",
        "language_ja": "日本語",
        "language_en": "English",
        "input_json": "入力JSON:",
        "copy_dir": "コピー先フォルダ:",
        "output_excel": "出力Excel:",
        "browse": "参照...",
        "do_copy": "JSONをコピーする",
        "do_excel": "Excelへ変換する",
        "run": "実行",
        "open_output_folder": "出力フォルダを開く",
        "log": "ログ:",
        "pick_json_title": "入力JSONを選択",
        "pick_copy_dir_title": "コピー先フォルダを選択",
        "pick_xlsx_title": "出力Excelを指定",
        "error_title": "エラー",
        "done_title": "完了",
        "err_no_action": "実行内容が未選択です。",
        "err_no_json": "入力JSONが未指定です。",
        "err_no_copy_dir": "コピー先フォルダが未指定です。",
        "err_no_xlsx": "出力Excelが未指定です。",
        "start_log": "[INFO] 開始",
        "done_msg": "処理が完了しました。",
        "settings_path_log": "[INFO] 設定保存先: {path}",
    },
    "en": {
        "window_title": "CS2 MOD Cache Integrated Tool",
        "language_label": "言語 / Language",
        "language_ja": "日本語",
        "language_en": "English",
        "input_json": "Input JSON:",
        "copy_dir": "Copy Destination Folder:",
        "output_excel": "Output Excel:",
        "browse": "Browse...",
        "do_copy": "Copy JSON",
        "do_excel": "Export to Excel",
        "run": "Run",
        "open_output_folder": "Open Output Folder",
        "log": "Log:",
        "pick_json_title": "Select Input JSON",
        "pick_copy_dir_title": "Select Destination Folder",
        "pick_xlsx_title": "Specify Output Excel",
        "error_title": "Error",
        "done_title": "Completed",
        "err_no_action": "No action is selected.",
        "err_no_json": "Input JSON is not specified.",
        "err_no_copy_dir": "Destination folder is not specified.",
        "err_no_xlsx": "Output Excel is not specified.",
        "start_log": "[INFO] Start",
        "done_msg": "The process has completed.",
        "settings_path_log": "[INFO] Settings file: {path}",
    },
}


# =========================================
# 共通処理
# =========================================

def safe_text(x: Any) -> Optional[str]:
    """Excelに書き込める形へ整形する。"""
    if x is None:
        return None
    if isinstance(x, (dict, list)):
        s = json.dumps(x, ensure_ascii=False)
    else:
        s = str(x)

    if len(s) > EXCEL_CELL_LIMIT:
        return s[: EXCEL_CELL_LIMIT - 20] + " …(truncated)"
    return s


def build_df(rows: List[Dict[str, Any]], key_union: Set[str]) -> pd.DataFrame:
    """列が増減する辞書配列を、列をそろえてDataFrame化する。"""
    cols = ["Key"] + sorted([c for c in key_union if c != "Key"])
    if rows:
        df = pd.DataFrame(rows)
        for c in cols:
            if c not in df.columns:
                df[c] = None
        df = df[cols]
    else:
        df = pd.DataFrame(columns=cols)
    return df


def load_root(json_path: Path) -> Dict[str, Any]:
    """JSONを読み込む。UTF-8 BOMあり/なし対応。"""
    with json_path.open("r", encoding="utf-8-sig") as f:
        root = json.load(f)

    if isinstance(root, list):
        return {str(i): v for i, v in enumerate(root)}
    if isinstance(root, dict):
        return root
    return {}


def get_app_base_dir() -> Path:
    """
    設定ファイルの保存先となる基準フォルダを返す。
    EXE化した場合はEXEの場所、
    py実行時はこのスクリプトの場所を返す。
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_settings_path() -> Path:
    """設定ファイルのフルパスを返す。"""
    return get_app_base_dir() / SETTINGS_FILE_NAME


# =========================================
# コピー処理
# =========================================

def copy_json_file(source_json: Path, target_dir: Path, log_fn=print) -> Path:
    """
    JSONファイルを指定フォルダへ PdxModsCache.json としてコピーする。
    bat の copy /Y 相当。
    """
    if not source_json.exists():
        raise FileNotFoundError(f"コピー元ファイルが見つかりません:\n{source_json}")

    if not target_dir.exists():
        log_fn(f"[INFO] コピー先フォルダを作成: {target_dir}")
        target_dir.mkdir(parents=True, exist_ok=True)

    target_file = target_dir / "PdxModsCache.json"

    log_fn(f"[INFO] コピー開始: {source_json}")
    log_fn(f"[INFO] コピー先   : {target_file}")

    # 上書きコピー
    shutil.copy2(source_json, target_file)

    log_fn("[DONE] JSONコピー完了")
    return target_file


# =========================================
# Excel変換処理
# =========================================

def export_json_to_excel(json_path: Path, out_xlsx: Path, log_fn=print) -> None:
    """PdxModsCache.json を正規化してExcel出力する。"""
    if not json_path.exists():
        raise FileNotFoundError(f"入力JSONが見つかりません:\n{json_path}")

    out_dir = out_xlsx.parent
    if str(out_dir) != "" and not out_dir.exists():
        out_dir.mkdir(parents=True, exist_ok=True)

    log_fn(f"[INFO] JSON読み込み: {json_path}")
    root = load_root(json_path)

    scalar_fields = [
        "Id", "Name", "AuthorId", "ShortDescription", "Description",
        "ServerSize", "ForumLink", "VersionName", "Version",
        "Subscribers", "VoteCount", "IsCollection", "IsRemoved",
        "IsIncompatible", "IsBanned", "IsInvalid",
        "SuggestedGameVersion", "LatestVersion", "ThumbnailUrl",
        "ServerTime", "Timestamp", "Guid",
    ]

    mods_rows: List[Dict[str, Any]] = []
    tags_rows: List[Dict[str, Any]] = []
    req_rows: List[Dict[str, Any]] = []
    dlc_rows: List[Dict[str, Any]] = []
    img_rows: List[Dict[str, Any]] = []
    link_rows: List[Dict[str, Any]] = []

    req_key_union: Set[str] = set()
    dlc_key_union: Set[str] = set()
    img_key_union: Set[str] = set()
    link_key_union: Set[str] = set()

    log_fn(f"[INFO] レコード数: {len(root)}")

    for key, rec in root.items():
        if rec is None:
            rec = {}
        if not isinstance(rec, dict):
            rec = {}

        # Mods シート用
        m: Dict[str, Any] = {"Key": key}
        for fkey in scalar_fields:
            m[fkey] = safe_text(rec.get(fkey))
        mods_rows.append(m)

        # Tags シート用
        tags_obj = rec.get("Tags")
        if isinstance(tags_obj, dict):
            for t_k, t_v in tags_obj.items():
                tag = t_k if t_k is not None else t_v
                if tag is not None:
                    tags_rows.append({"Key": key, "Tag": safe_text(tag)})
        elif isinstance(tags_obj, list):
            for t in tags_obj:
                if t is not None:
                    tags_rows.append({"Key": key, "Tag": safe_text(t)})

        # Requirements シート用
        req_list = rec.get("Requirements") or []
        if isinstance(req_list, list):
            for r in req_list:
                if not isinstance(r, dict):
                    continue
                req_key_union.update(r.keys())
                rr: Dict[str, Any] = {"Key": key}
                for k2, v2 in r.items():
                    rr[k2] = safe_text(v2)
                req_rows.append(rr)

        # DlcRequirements シート用
        dlc_list = rec.get("DlcRequirements") or []
        if isinstance(dlc_list, list):
            for d in dlc_list:
                if not isinstance(d, dict):
                    continue
                dlc_key_union.update(d.keys())
                dr: Dict[str, Any] = {"Key": key}
                for k2, v2 in d.items():
                    dr[k2] = safe_text(v2)
                dlc_rows.append(dr)

        # Images シート用
        imgs = rec.get("Images") or []
        if isinstance(imgs, list):
            for im in imgs:
                if not isinstance(im, dict):
                    continue
                img_key_union.update(im.keys())
                ir: Dict[str, Any] = {"Key": key}
                for k2, v2 in im.items():
                    ir[k2] = safe_text(v2)
                img_rows.append(ir)

        # Links シート用
        links = rec.get("Links") or []
        if isinstance(links, list):
            for lk in links:
                if not isinstance(lk, dict):
                    continue
                link_key_union.update(lk.keys())
                lr: Dict[str, Any] = {"Key": key}
                for k2, v2 in lk.items():
                    lr[k2] = safe_text(v2)
                link_rows.append(lr)

    mods_df = pd.DataFrame(mods_rows)
    tags_df = pd.DataFrame(tags_rows) if tags_rows else pd.DataFrame(columns=["Key", "Tag"])
    req_df = build_df(req_rows, req_key_union)
    dlc_df = build_df(dlc_rows, dlc_key_union)
    img_df = build_df(img_rows, img_key_union)
    link_df = build_df(link_rows, link_key_union)

    log_fn(f"[INFO] Excel出力: {out_xlsx}")

    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        mods_df.to_excel(writer, sheet_name="Mods", index=False)
        tags_df.to_excel(writer, sheet_name="Tags", index=False)
        req_df.to_excel(writer, sheet_name="Requirements", index=False)
        dlc_df.to_excel(writer, sheet_name="DlcRequirements", index=False)
        img_df.to_excel(writer, sheet_name="Images", index=False)
        link_df.to_excel(writer, sheet_name="Links", index=False)

    log_fn("[DONE] Excel出力完了")


# =========================================
# GUI
# =========================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        # 設定ファイルパス
        self.settings_path = get_settings_path()

        # 表示言語
        self.lang = "ja"
        self.language_var = tk.StringVar(value=UI_TEXT["ja"]["language_ja"])

        # 入力値
        self.json_var = tk.StringVar()
        self.copy_dir_var = tk.StringVar()
        self.xlsx_var = tk.StringVar()

        # 実行オプション
        self.do_copy_var = tk.BooleanVar(value=True)
        self.do_excel_var = tk.BooleanVar(value=True)

        # 初期値
        default_json = Path(
            r"C:\Users\penpe\AppData\LocalLow\Colossal Order\Cities Skylines II\ModsData\Skyve\PdxModsCache.json"
        )
        default_copy_dir = Path(
            r"F:\05.game_管理\Cities Skylines II\PdxModsCache\Skyve_PdxModsCache"
        )
        default_xlsx = default_copy_dir / "PdxModsCache_normalized.xlsx"

        self.json_var.set(str(default_json))
        self.copy_dir_var.set(str(default_copy_dir))
        self.xlsx_var.set(str(default_xlsx))

        # 保存済み設定を読み込む
        self.load_settings()

        self.geometry("900x560")
        self._build_ui()
        self._apply_language()

        # 入力変更時に自動保存
        self.json_var.trace_add("write", self.on_setting_changed)
        self.copy_dir_var.trace_add("write", self.on_setting_changed)
        self.xlsx_var.trace_add("write", self.on_setting_changed)

        # ウィンドウ終了時にも保存
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.log_write(self.tr("settings_path_log", path=self.settings_path))

    def tr(self, key: str, **kwargs) -> str:
        text = UI_TEXT[self.lang].get(key, key)
        if kwargs:
            return text.format(**kwargs)
        return text

    def _build_ui(self):
        # 上部タイトル + 言語切替
        top = tk.Frame(self)
        top.pack(fill="x", padx=10, pady=(10, 4))

        lang_area = tk.Frame(top)
        lang_area.pack(side="right")

        self.language_label = tk.Label(lang_area, text="")
        self.language_label.pack(anchor="e")

        self.language_combo = ttk.Combobox(
            lang_area,
            textvariable=self.language_var,
            values=[UI_TEXT["ja"]["language_ja"], UI_TEXT["ja"]["language_en"]],
            state="readonly",
            width=12,
        )
        self.language_combo.pack(anchor="e", pady=(3, 0))
        self.language_combo.bind("<<ComboboxSelected>>", self.set_language)

        # 入力JSON
        frm1 = tk.Frame(self)
        frm1.pack(fill="x", padx=10, pady=8)

        self.label_json = tk.Label(frm1, text="", width=18, anchor="w")
        self.label_json.pack(side="left")
        tk.Entry(frm1, textvariable=self.json_var).pack(side="left", fill="x", expand=True, padx=6)
        self.btn_json = tk.Button(frm1, text="", command=self.pick_json)
        self.btn_json.pack(side="left")

        # コピー先フォルダ
        frm2 = tk.Frame(self)
        frm2.pack(fill="x", padx=10, pady=8)

        self.label_copy_dir = tk.Label(frm2, text="", width=18, anchor="w")
        self.label_copy_dir.pack(side="left")
        tk.Entry(frm2, textvariable=self.copy_dir_var).pack(side="left", fill="x", expand=True, padx=6)
        self.btn_copy_dir = tk.Button(frm2, text="", command=self.pick_copy_dir)
        self.btn_copy_dir.pack(side="left")

        # 出力Excel
        frm3 = tk.Frame(self)
        frm3.pack(fill="x", padx=10, pady=8)

        self.label_xlsx = tk.Label(frm3, text="", width=18, anchor="w")
        self.label_xlsx.pack(side="left")
        tk.Entry(frm3, textvariable=self.xlsx_var).pack(side="left", fill="x", expand=True, padx=6)
        self.btn_xlsx = tk.Button(frm3, text="", command=self.pick_xlsx)
        self.btn_xlsx.pack(side="left")

        # 実行オプション
        frm4 = tk.Frame(self)
        frm4.pack(fill="x", padx=10, pady=8)

        self.chk_copy = tk.Checkbutton(
            frm4,
            text="",
            variable=self.do_copy_var,
            command=self.save_settings
        )
        self.chk_copy.pack(side="left", padx=8)

        self.chk_excel = tk.Checkbutton(
            frm4,
            text="",
            variable=self.do_excel_var,
            command=self.save_settings
        )
        self.chk_excel.pack(side="left", padx=8)

        # ボタン
        frm5 = tk.Frame(self)
        frm5.pack(fill="x", padx=10, pady=8)

        self.run_btn = tk.Button(frm5, text="", command=self.run_clicked, height=2, width=16)
        self.run_btn.pack(side="left")

        self.open_dir_btn = tk.Button(frm5, text="", command=self.open_out_dir)
        self.open_dir_btn.pack(side="left", padx=10)

        # ログ
        frm6 = tk.Frame(self)
        frm6.pack(fill="both", expand=True, padx=10, pady=8)

        self.label_log = tk.Label(frm6, text="")
        self.label_log.pack(anchor="w")
        self.log = tk.Text(frm6, height=18)
        self.log.pack(fill="both", expand=True)

    def _apply_language(self):
        self.title(self.tr("window_title"))
        self.language_label.config(text=self.tr("language_label"))
        self.language_combo.configure(values=[self.tr("language_ja"), self.tr("language_en")])
        self.language_var.set(self.tr("language_en") if self.lang == "en" else self.tr("language_ja"))

        self.label_json.config(text=self.tr("input_json"))
        self.label_copy_dir.config(text=self.tr("copy_dir"))
        self.label_xlsx.config(text=self.tr("output_excel"))

        self.btn_json.config(text=self.tr("browse"))
        self.btn_copy_dir.config(text=self.tr("browse"))
        self.btn_xlsx.config(text=self.tr("browse"))

        self.chk_copy.config(text=self.tr("do_copy"))
        self.chk_excel.config(text=self.tr("do_excel"))
        self.run_btn.config(text=self.tr("run"))
        self.open_dir_btn.config(text=self.tr("open_output_folder"))
        self.label_log.config(text=self.tr("log"))

    def set_language(self, event=None):
        selected = self.language_var.get()
        self.lang = "en" if selected == UI_TEXT["en"]["language_en"] else "ja"
        self._apply_language()

    def log_write(self, msg: str):
        """ログ欄へ書き込む。"""
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.update_idletasks()

    def on_setting_changed(self, *args):
        """入力欄が変わったら設定を保存する。"""
        self.save_settings()

    def load_settings(self):
        """前回保存した設定を読み込む。"""
        try:
            if not self.settings_path.exists():
                return

            with self.settings_path.open("r", encoding="utf-8") as f:
                data = json.load(f)

            self.json_var.set(data.get("json_path", self.json_var.get()))
            self.copy_dir_var.set(data.get("copy_dir", self.copy_dir_var.get()))
            self.xlsx_var.set(data.get("xlsx_path", self.xlsx_var.get()))
            self.do_copy_var.set(bool(data.get("do_copy", self.do_copy_var.get())))
            self.do_excel_var.set(bool(data.get("do_excel", self.do_excel_var.get())))
        except Exception:
            # 設定ファイルが壊れていても起動は継続
            pass

    def save_settings(self):
        """現在の設定を保存する。"""
        try:
            data = {
                "json_path": self.json_var.get(),
                "copy_dir": self.copy_dir_var.get(),
                "xlsx_path": self.xlsx_var.get(),
                "do_copy": self.do_copy_var.get(),
                "do_excel": self.do_excel_var.get(),
            }
            with self.settings_path.open("w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            # 保存失敗でもアプリ自体は止めない
            pass

    def pick_json(self):
        """入力JSONを選ぶ。"""
        p = filedialog.askopenfilename(
            title=self.tr("pick_json_title"),
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if p:
            self.json_var.set(p)
            self.save_settings()

    def pick_copy_dir(self):
        """コピー先フォルダを選ぶ。"""
        p = filedialog.askdirectory(title=self.tr("pick_copy_dir_title"))
        if p:
            self.copy_dir_var.set(p)
            self.save_settings()

    def pick_xlsx(self):
        """出力Excelを選ぶ。"""
        p = filedialog.asksaveasfilename(
            title=self.tr("pick_xlsx_title"),
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if p:
            self.xlsx_var.set(p)
            self.save_settings()

    def open_out_dir(self):
        """出力Excelのフォルダを開く。"""
        out_path = Path(self.xlsx_var.get())
        if out_path.parent.exists():
            os.startfile(out_path.parent)

    def run_clicked(self):
        """実行ボタン押下時。"""
        json_path = Path(self.json_var.get())
        copy_dir = Path(self.copy_dir_var.get())
        out_xlsx = Path(self.xlsx_var.get())

        if not self.do_copy_var.get() and not self.do_excel_var.get():
            messagebox.showerror(self.tr("error_title"), self.tr("err_no_action"))
            return

        if not self.json_var.get().strip():
            messagebox.showerror(self.tr("error_title"), self.tr("err_no_json"))
            return

        if self.do_copy_var.get() and not self.copy_dir_var.get().strip():
            messagebox.showerror(self.tr("error_title"), self.tr("err_no_copy_dir"))
            return

        if self.do_excel_var.get() and not self.xlsx_var.get().strip():
            messagebox.showerror(self.tr("error_title"), self.tr("err_no_xlsx"))
            return

        # 実行前にも保存
        self.save_settings()

        self.run_btn.config(state="disabled")
        self.log_write(self.tr("start_log"))

        def worker():
            try:
                working_json = json_path

                # 1. JSONコピー
                if self.do_copy_var.get():
                    copied_json = copy_json_file(json_path, copy_dir, log_fn=self.log_write)
                    working_json = copied_json

                # 2. Excel変換
                if self.do_excel_var.get():
                    export_json_to_excel(working_json, out_xlsx, log_fn=self.log_write)

                # 実行後も保存
                self.save_settings()
                messagebox.showinfo(self.tr("done_title"), self.tr("done_msg"))

            except Exception as e:
                self.log_write(f"[ERROR] {e}")
                messagebox.showerror(self.tr("error_title"), str(e))
            finally:
                self.run_btn.config(state="normal")

        threading.Thread(target=worker, daemon=True).start()

    def on_close(self):
        """終了時に設定保存して閉じる。"""
        self.save_settings()
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
