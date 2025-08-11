import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import fitz  # PyMuPDF
import cv2
import numpy as np
import pytesseract
from PIL import Image, ImageTk
import os
import shutil
from datetime import datetime
import re
import io
import csv

class PDFRenamerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Renamer")
        self.root.state('zoomed')  # Start maximized
        
        # Configuration
        self.config = self.load_config()
        
        # Variables - 新仕様対応
        self.pdf_files = []
        self.current_pdf_index = 0
        self.current_pdf_doc = None
        self.selecting_area = False  # 'center' or 'right' or False
        self.start_x = 0
        self.start_y = 0
        self.rect_id = None
        self.current_selection_color = 'red'  # デフォルトの枠線色
        # Rendering state for accurate coordinate transforms (cover mode)
        self._render_scale = 1.0
        self._crop_left = 0
        self._crop_top = 0
        
        # Create folders
        self.create_folders()

        # Prepare per-run CSV log file name (YYYYMMDD_hhmmss.csv)
        try:
            ts_name = datetime.now().strftime("%Y%m%d_%H%M%S") + ".csv"
            log_dir = self.config.get('log_output_folder') or 'log_output'
            os.makedirs(log_dir, exist_ok=True)
            self.current_csv_path = os.path.join(log_dir, ts_name)
            # 実行ごとに新規CSVファイルを作成（空ファイル）
            try:
                with open(self.current_csv_path, 'w', encoding='utf-8', newline='') as f:
                    pass
            except Exception as e:
                # UI未初期化のためprintに留める
                print(f"CSV初期作成エラー: {e}")
        except Exception as e:
            # フォールバック: 既定log_output直下にタイムスタンプ名で作成を試行
            try:
                ts_name = datetime.now().strftime("%Y%m%d_%H%M%S") + ".csv"
                log_dir = 'log_output'
                os.makedirs(log_dir, exist_ok=True)
                self.current_csv_path = os.path.join(log_dir, ts_name)
                with open(self.current_csv_path, 'w', encoding='utf-8', newline='') as f:
                    pass
                # UI未初期化のためprintに留める
                print(f"CSVログフォールバック作成: {self.current_csv_path}")
            except Exception as e2:
                # 最終手段: メモリ上のパス（起動後の保存時に例外で通知）
                self.current_csv_path = os.path.join(os.getcwd(), ts_name)
                print(f"CSV作成に失敗しました: {e2}")

        # Setup UI
        self.setup_ui()
        # UI準備後にログ出力
        try:
            self.log_message(f"CSVログファイル: {self.current_csv_path}")
        except Exception:
            pass
        
        # Setup Tesseract path
        self.setup_tesseract()

        # Load PDF files on startup
        self.load_pdf_files()
    
    def load_config(self):
        """Load configuration from config.txt"""
        config = {}
        try:
            with open('config.txt', 'r', encoding='utf-8') as f:
                for line in f:
                    if line.strip() and not line.startswith('#'):
                        key, value = line.strip().split('=', 1)
                        # 数値項目の定義を新仕様に対応
                        if key in ['red_frame_x', 'red_frame_y', 'red_frame_width', 'red_frame_height',
                                  'blue_frame_x', 'blue_frame_y', 'blue_frame_width', 'blue_frame_height']:
                            config[key] = int(value)
                        else:
                            config[key] = value
        except FileNotFoundError:
            # Default configuration - 新仕様対応
            config = {
                'pdf_input_folder': 'pdf_input',
                'pdf_output_folder': 'pdf_output',
                'log_output_folder': 'log_output',
                'ocr_image_folder': 'ocr_get_image',
                # 赤枠（中央表示用）の座標
                'red_frame_x': 600,
                'red_frame_y': 250,
                'red_frame_width': 300,
                'red_frame_height': 200,
                # 青枠（右側表示用）の座標
                'blue_frame_x': 400,
                'blue_frame_y': 350,
                'blue_frame_width': 250,
                'blue_frame_height': 150
            }
        return config
    
    def save_config(self):
        """Save configuration to config.txt"""
        with open('config.txt', 'w', encoding='utf-8') as f:
            f.write("# PDF Renamer Configuration File\n")
            f.write(f"pdf_input_folder={self.config['pdf_input_folder']}\n")
            f.write(f"pdf_output_folder={self.config['pdf_output_folder']}\n")
            f.write(f"log_output_folder={self.config['log_output_folder']}\n")
            f.write(f"ocr_image_folder={self.config['ocr_image_folder']}\n")
            f.write("\n# Red Frame Coordinates (Center Display Area)\n")
            f.write(f"red_frame_x={self.config.get('red_frame_x', 600)}\n")
            f.write(f"red_frame_y={self.config.get('red_frame_y', 250)}\n")
            f.write(f"red_frame_width={self.config.get('red_frame_width', 300)}\n")
            f.write(f"red_frame_height={self.config.get('red_frame_height', 200)}\n")
            f.write("\n# Blue Frame Coordinates (Right Display Area)\n")
            f.write(f"blue_frame_x={self.config.get('blue_frame_x', 400)}\n")
            f.write(f"blue_frame_y={self.config.get('blue_frame_y', 350)}\n")
            f.write(f"blue_frame_width={self.config.get('blue_frame_width', 250)}\n")
            f.write(f"blue_frame_height={self.config.get('blue_frame_height', 150)}\n")
    
    def create_folders(self):
        """Create necessary folders"""
        folders = [
            self.config['pdf_input_folder'],
            self.config['pdf_output_folder'],
            self.config['log_output_folder'],
            self.config['ocr_image_folder']
        ]
        for folder in folders:
            os.makedirs(folder, exist_ok=True)
    
    def setup_tesseract(self):
        """Setup Tesseract path"""
        tesseract_paths = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            r"C:\Users\{}\AppData\Local\Programs\Tesseract-OCR\tesseract.exe".format(os.getenv('USERNAME'))
        ]
        
        for path in tesseract_paths:
            if os.path.exists(path):
                pytesseract.pytesseract.tesseract_cmd = path
                break
    
    def setup_ui(self):
        """Setup the user interface - 新仕様3分割レイアウト"""
        # Main frame
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Top frame for controls（3ボタンを同じ高さで横一列に配置）
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 0))
        top_frame.columnconfigure(0, weight=1)
        top_frame.columnconfigure(1, weight=1)
        top_frame.columnconfigure(2, weight=1)

        self.btn_input = ttk.Button(top_frame, text="入力フォルダ選択", command=self.select_input_folder, style='Large.TButton')
        self.btn_input.grid(row=0, column=0)

        self.btn_set_center = ttk.Button(top_frame, text="表示画像を設定（中）", command=self.set_center_area, style='Large.TButton')
        self.btn_set_center.grid(row=0, column=1)

        self.btn_set_right = ttk.Button(top_frame, text="表示画像を設定（右）", command=self.set_right_area, style='Large.TButton')
        self.btn_set_right.grid(row=0, column=2)
        
        # Middle frame for 3-column layout (pack-based)
        middle_frame = ttk.Frame(main_frame)
        middle_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # === 左側: PDFビューワ (1/3) ===
        left_frame = ttk.Frame(middle_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # PDF Canvas
        self.pdf_canvas = tk.Canvas(left_frame, bg='white', relief=tk.SUNKEN, bd=2)
        self.pdf_canvas.pack(fill=tk.BOTH, expand=True)
        self.pdf_canvas.bind("<Button-1>", self.on_canvas_click)
        self.pdf_canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.pdf_canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        # Re-render the page to keep full-screen fit on resize
        self.pdf_canvas.bind("<Configure>", self.on_pdf_canvas_configure)
        
        # === 中央: 表示画像（中） (1/3) ===
        center_frame = ttk.Frame(middle_frame)
        center_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 5))
        
        # 中央列のボタンは上段に移動済み（top_frame）。余白だけ確保（高さ0）
        ttk.Frame(center_frame).pack(fill=tk.X, pady=(0, 0))
        
        # 中央表示エリア（通常サイズに復元）
        center_display_frame = ttk.Frame(center_frame)
        center_display_frame.pack(fill=tk.BOTH, expand=True)
        
        self.center_canvas = tk.Canvas(center_display_frame, bg='lightgray', relief=tk.SUNKEN, bd=2)
        self.center_canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=0)
        # サイズ変化に応じて中央/右のプレビューを更新
        self.center_canvas.bind("<Configure>", self.on_side_canvas_configure)
        
        # 中央列ボトムコントロール（テキストボックス + 表示ボタン）- 固定高さで配置
        bottom_spacer_height = 80
        center_bottom_controls = ttk.Frame(center_frame, height=bottom_spacer_height)
        center_bottom_controls.pack(fill=tk.X)
        center_bottom_controls.pack_propagate(False)
        center_bottom_controls.columnconfigure(0, weight=1)

        vcmd = (self.root.register(self.validate_input), '%P')
        self.entry_var = getattr(self, 'entry_var', tk.StringVar())
        self.id_entry = ttk.Entry(center_bottom_controls, textvariable=self.entry_var,
                                  font=('Arial', 16), width=18, style='FocusEntry.TEntry',
                                  validate='key', validatecommand=vcmd)
        self.id_entry.grid(row=0, column=0, sticky='we', padx=(10, 10), pady=10)
        # 起動時に入力フォーカスをテキストボックスへ
        self.root.after(200, lambda: (self.id_entry.focus_set(), self.id_entry.icursor(tk.END)))
        # フォーカスが当たったら仮表示のキー値を初期化
        self.id_entry.bind('<FocusIn>', self.on_entry_focus_in)

        self.display_button = ttk.Button(center_bottom_controls, text="表示", command=self.on_display_click, style='Large.TButton', width=10)
        self.display_button.grid(row=0, column=1, padx=(0, 10), pady=10)
        # フォーカス移動（Enterキー）: Entry -> 表示(実行) -> 保存(可能なら) / 不可なら Entry
        self.id_entry.bind('<Return>', lambda e: (self.display_button.focus_set(), 'break'))
        self.display_button.bind('<Return>', self.on_display_enter)
        
        # === 右側: 表示画像（右） (1/3) ===
        right_frame = ttk.Frame(middle_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # 右列のボタンは上段に移動済み（top_frame）。余白だけ確保（高さ0）
        ttk.Frame(right_frame).pack(fill=tk.X, pady=(0, 0))
        
        # 右側表示エリア（通常サイズに復元）
        right_display_frame = ttk.Frame(right_frame)
        right_display_frame.pack(fill=tk.BOTH, expand=True)
        
        self.right_canvas = tk.Canvas(right_display_frame, bg='lightgray', relief=tk.SUNKEN, bd=2)
        self.right_canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=0)
        self.right_canvas.bind("<Configure>", self.on_side_canvas_configure)
        
        # 右列ボトム: 結果表示エリア + 「保存」ボタン
        right_bottom_area = ttk.Frame(right_frame, height=bottom_spacer_height)
        right_bottom_area.pack(fill=tk.X)
        right_bottom_area.pack_propagate(False)
        right_bottom_area.columnconfigure(0, weight=1)
        right_bottom_area.columnconfigure(1, weight=0)
        
        if not hasattr(self, 'result_var'):
            self.result_var = tk.StringVar(value="")
        self.result_label = ttk.Label(right_bottom_area, textvariable=self.result_var,
                                      anchor='w')
        self.result_label.grid(row=0, column=0, sticky='ew', padx=(10, 10), pady=10)
        
        self.save_button = ttk.Button(right_bottom_area, text="保存", command=self.on_save_click, style='Large.TButton', width=10)
        self.save_button.grid(row=0, column=1, sticky='e', padx=(0, 10), pady=10)
        # 保存(実行) -> 次のPDFへ -> Entryへフォーカス
        self.save_button.bind('<Return>', self.on_save_enter)
        # 保存ボタンの有効/無効はテキストボックスの8桁数字入力で制御
        # Entry の変更をフック
        try:
            self.entry_var.trace_add('write', self.update_save_button_state)
        except Exception:
            # 古いTkの場合の互換（trace）
            self.entry_var.trace('w', self.update_save_button_state)
        # 初期状態を反映（起動直後は無効が基本）
        self.update_save_button_state()
        
        # Navigation frame - 左右端寄せ + 中央にファイル情報
        nav_frame = ttk.Frame(main_frame)
        nav_frame.pack(fill=tk.X, pady=(0, 10))

        # 3カラムグリッド: [前へ] [ファイル情報] [次へ]
        nav_frame.columnconfigure(0, weight=1)
        nav_frame.columnconfigure(1, weight=0)
        nav_frame.columnconfigure(2, weight=1)

        self.prev_button = ttk.Button(nav_frame, text="<< 前へ", command=self.prev_pdf, style='Large.TButton', width=10)
        self.prev_button.grid(row=0, column=0, sticky='w', padx=(0, 10))

        self.file_info_label = ttk.Label(nav_frame, text="", font=('Arial', 16))
        self.file_info_label.grid(row=0, column=1)

        self.next_button = ttk.Button(nav_frame, text=">> 次へ", command=self.next_pdf, style='Large.TButton', width=10)
        # 2.5mm ≒ 10px（96DPI想定）だけ左に寄せる
        self.next_button.grid(row=0, column=2, sticky='e', padx=(10, 10))
        
        # Log frame
        log_frame = ttk.LabelFrame(main_frame, text="ログ")
        # ログ領域を約1/2の高さに抑える（固定行数・非expand）
        log_frame.pack(fill=tk.X, pady=(10, 0))
        self.log_text = tk.Text(log_frame, height=3)
        self.log_text.pack(fill=tk.X)

        # ウィンドウを閉じる際に設定を保存
        try:
            self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        except Exception:
            pass
        log_scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0), pady=5)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5, padx=(0, 5))
        
        # Configure button style
        style = ttk.Style()
        style.configure('Large.TButton', font=('Arial', 16))
        # Entry フォーカス時に常に #ffffe0（薄い黄色）にする
        style.configure('FocusEntry.TEntry', font=('Arial', 16), fieldbackground='white')
        style.map('FocusEntry.TEntry', fieldbackground=[('focus', '#ffffe0')])
        
        # 入力検証: 半角数字のみ、最大8文字
        # validate_input(self, value) で実装済み
    
    def validate_input(self, value):
        """Validate input to allow only digits and max 8 characters"""
        if len(value) <= 8 and (value.isdigit() or value == ""):
            return True
        return False
    
    def update_save_button_state(self, *_, **__):
        """Enable Save button only when Entry has exactly 8 digits."""
        try:
            value = self.entry_var.get().strip()
            if len(value) == 8 and value.isdigit():
                # enable
                try:
                    self.save_button.state(['!disabled'])
                except Exception:
                    self.save_button.configure(state='normal')
                # Tab移動対象に含める
                try:
                    self.save_button.configure(takefocus=1)
                except Exception:
                    pass
            else:
                # disable
                try:
                    self.save_button.state(['disabled'])
                except Exception:
                    self.save_button.configure(state='disabled')
                # Tab移動対象から除外
                try:
                    self.save_button.configure(takefocus=0)
                except Exception:
                    pass
                # もし現在フォーカスが保存ボタンならエントリへ戻す
                try:
                    if str(self.root.focus_get()) == str(self.save_button):
                        self.id_entry.focus_set()
                except Exception:
                    pass
        except Exception:
            # Fallback: keep disabled on any unexpected error
            try:
                self.save_button.state(['disabled'])
            except Exception:
                self.save_button.configure(state='disabled')
            try:
                self.save_button.configure(takefocus=0)
            except Exception:
                pass
    
    def on_display_click(self):
        """Handle display button click next to the Entry"""
        value = self.entry_var.get().strip()
        if len(value) == 8 and value.isdigit():
            self.log_message(f"表示ボタン: 入力={value}")
            # 将来: ODBCでMSSQL接続→valueをキーに検索→結果を右下のエリアへ表示
            # ここではプレースホルダとして仮表示
            if hasattr(self, 'result_var'):
                self.result_var.set(f"キー {value} の検索結果（仮表示）")
        else:
            self.log_message("8桁の半角数字を入力してください。")

    def on_entry_focus_in(self, event=None):
        """When Entry gets focus, clear the placeholder key display."""
        try:
            if hasattr(self, 'result_var') and self.result_var is not None:
                self.result_var.set("")
        except Exception:
            pass

    def on_display_enter(self, event=None):
        """Enter on Display: run display action, then move focus to Save if enabled, otherwise back to Entry."""
        try:
            self.on_display_click()
            # 直後の状態で保存可能かを判断
            self.update_save_button_state()
            try:
                state = self.save_button.state()
            except Exception:
                # Fallback for non-ttk state retrieval
                state = [] if str(self.save_button.cget('state')) == 'normal' else ['disabled']
            if 'disabled' not in state:
                self.save_button.focus_set()
            else:
                self.id_entry.focus_set()
        except Exception:
            # 何かあればエントリへ戻す
            try:
                self.id_entry.focus_set()
            except Exception:
                pass
        return 'break'
    
    def on_save_click(self):
        """Copy the currently viewed PDF to output folder with Entry value as filename."""
        # 1) Validate entry
        value = self.entry_var.get().strip() if hasattr(self, 'entry_var') else ""
        if not (len(value) == 8 and value.isdigit()):
            messagebox.showwarning("入力エラー", "8桁の半角数字を入力してください。")
            self.update_save_button_state()
            return

        # 2) Resolve current PDF path
        if not self.pdf_files:
            self.log_message("保存失敗: 表示中のPDFがありません。")
            messagebox.showerror("エラー", "表示中のPDFがありません。")
            return
        try:
            src_pdf = os.path.join(self.config.get('pdf_input_folder', ''), self.pdf_files[self.current_pdf_index])
        except Exception:
            src_pdf = None

        if not src_pdf or not os.path.isfile(src_pdf):
            self.log_message("保存失敗: 元PDFが見つかりません。")
            messagebox.showerror("エラー", "元PDFが見つかりません。")
            return

        # 3) Prepare destination path
        output_dir = self.config.get('pdf_output_folder') or 'pdf_output'
        try:
            os.makedirs(output_dir, exist_ok=True)
        except Exception as e:
            self.log_message(f"保存失敗: 出力フォルダ作成エラー: {e}")
            messagebox.showerror("エラー", f"出力フォルダの作成に失敗しました:\n{e}")
            return

        dest_pdf = os.path.join(output_dir, f"{value}.pdf")

        # 3.5) 現在ページに既存レコードがあるか確認（あれば更新モード）
        page_no = self.current_pdf_index + 1
        existing_row = self.get_csv_row_by_index(page_no)
        old_value_in_csv = None
        old_seq = None
        if existing_row:
            try:
                old_value_in_csv = existing_row[0] if len(existing_row) >= 1 else None
            except Exception:
                old_value_in_csv = None
            try:
                old_seq = int(existing_row[-1])
            except Exception:
                old_seq = None

        # 4) Overwrite confirmation if exists
        # 既存レコードがあり、旧ファイル名→新ファイル名の置換を行うモードでは確認ダイアログなしで実施
        if not existing_row:
            if os.path.exists(dest_pdf):
                if not messagebox.askyesno("上書き確認", f"既に存在します:\n{dest_pdf}\n上書きしますか？"):
                    self.log_message("保存をキャンセルしました（上書きしない）。")
                    return

        # 5) Copy / Replace
        try:
             # 旧PDFの削除（既存レコードがあり、旧値が存在し、新値と異なる場合）
            if existing_row and old_value_in_csv and old_value_in_csv != value:
                old_pdf_path = os.path.join(output_dir, f"{old_value_in_csv}.pdf")
                if os.path.exists(old_pdf_path):
                    try:
                        os.remove(old_pdf_path)
                        self.log_message(f"旧PDFを削除しました: {old_pdf_path}")
                    except Exception as de:
                        self.log_message(f"旧PDF削除エラー: {de}")
            shutil.copy2(src_pdf, dest_pdf)
            self.log_message(f"保存完了: {dest_pdf}")
            messagebox.showinfo("保存完了", f"保存しました:\n{dest_pdf}")
            # CSV: 既存レコードがあればその行を更新、なければ追記
            try:
                placeholder = ""
                try:
                    if hasattr(self, 'result_var') and self.result_var is not None:
                        placeholder = self.result_var.get()
                except Exception:
                    placeholder = ""
                if existing_row:
                    self.update_csv_row_by_index(page_no, value, placeholder, keep_seq=old_seq)
                else:
                    self.append_csv_log(value, placeholder)
            except Exception as e:
                self.log_message(f"CSVログ出力エラー: {e}")
            # 入力欄を初期化し、ボタン状態を更新
            try:
                self.entry_var.set("")
            except Exception:
                pass
            try:
                self.update_save_button_state()
            except Exception:
                pass
        except Exception as e:
            self.log_message(f"保存失敗: コピー中にエラー: {e}")
            messagebox.showerror("エラー", f"コピーに失敗しました:\n{e}")

    def append_csv_log(self, key_value: str, placeholder_text: str):
        """Append a CSV row 'key_value, placeholder_text, seq' into log_output/rename_log.csv with sequential numbering."""
        log_dir = self.config.get('log_output_folder') or 'log_output'
        os.makedirs(log_dir, exist_ok=True)
        csv_path = self.current_csv_path

        # Determine next sequence number by reading last valid row
        next_seq = 1
        if os.path.exists(csv_path):
            try:
                with open(csv_path, 'r', encoding='utf-8', newline='') as f:
                    reader = csv.reader(f)
                    last_seq = 0
                    for row in reader:
                        if not row:
                            continue
                        try:
                            # Use the last column as seq to support both 2-col and 3-col historical formats
                            last_seq = int(row[-1])
                        except Exception:
                            continue
                    next_seq = last_seq + 1 if last_seq >= 0 else 1
            except Exception as e:
                # 読み取り失敗時は1から（新規扱い）
                self.log_message(f"CSV読み込みエラー: {e} -> 連番を1から開始します")
                next_seq = 1

        # Append new row
        with open(csv_path, 'a', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            writer.writerow([key_value, placeholder_text, next_seq])
        self.log_message(f"CSV出力: {csv_path} に {key_value},{placeholder_text},{next_seq} を追記")

    def get_csv_row_by_index(self, index_1based: int):
        """Return the row (list[str]) at 1-based index from the current CSV file if exists, else None."""
        try:
            log_dir = self.config.get('log_output_folder') or 'log_output'
            csv_path = self.current_csv_path
            if not os.path.exists(csv_path):
                return None
            with open(csv_path, 'r', encoding='utf-8', newline='') as f:
                reader = csv.reader(f)
                for i, row in enumerate(reader, start=1):
                    if i == index_1based:
                        return row
            return None
        except Exception as e:
            self.log_message(f"CSV読み込みエラー(get_csv_row_by_index): {e}")
            return None

    def update_csv_row_by_index(self, index_1based: int, key_value: str, placeholder_text: str, keep_seq: int | None = None):
        """Update a specific 1-based row in the current CSV with new values while preserving sequence if provided.
        If the index is beyond current rows, does nothing and returns False.
        Returns True if updated, else False.
        """
        try:
            csv_path = self.current_csv_path
            if not os.path.exists(csv_path):
                return False
            rows: list[list[str]] = []
            with open(csv_path, 'r', encoding='utf-8', newline='') as f:
                rows = list(csv.reader(f))
            if index_1based < 1 or index_1based > len(rows):
                return False
            old_row = rows[index_1based - 1]
            seq_val = None
            if keep_seq is not None:
                seq_val = keep_seq
            else:
                try:
                    seq_val = int(old_row[-1])
                except Exception:
                    seq_val = index_1based  # フォールバック
            rows[index_1based - 1] = [key_value, placeholder_text, seq_val]
            with open(csv_path, 'w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f)
                writer.writerows(rows)
            self.log_message(f"CSV更新: {csv_path} の {index_1based} 行目を書き換え")
            return True
        except Exception as e:
            self.log_message(f"CSV更新エラー(update_csv_row_by_index): {e}")
            return False

    def on_save_enter(self, event=None):
        """Enter on Save: perform save, go to next PDF, and focus Entry."""
        try:
            self.on_save_click()
            # 次のPDFへ
            try:
                self.next_pdf()
            except Exception:
                pass
            # Entryへフォーカス
            try:
                self.id_entry.focus_set()
                self.id_entry.icursor(tk.END)
            except Exception:
                pass
        except Exception:
            # 例外時もEntryへ
            try:
                self.id_entry.focus_set()
            except Exception:
                pass
        return 'break'
    
    def on_close(self):
        """Save config on exit and close the app"""
        try:
            self.save_config()
            self.log_message("設定を保存して終了します。")
        except Exception as e:
            self.log_message(f"設定保存中にエラー: {e}")
        finally:
            try:
                self.root.destroy()
            except Exception:
                pass
    
    def log_message(self, message):
        """Add message to log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def select_input_folder(self):
        """Select input folder containing PDF files"""
        folder = filedialog.askdirectory(title="入力フォルダを選択")
        if folder:
            self.config['pdf_input_folder'] = folder
            self.save_config()
            self.load_pdf_files()
    
    def load_pdf_files(self):
        """Load PDF files from input folder"""
        input_folder = self.config['pdf_input_folder']
        if not os.path.exists(input_folder):
            self.log_message(f"入力フォルダが見つかりません: {input_folder}")
            return
        
        # Get all PDF files and sort them
        self.pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.pdf')]
        self.pdf_files.sort()
        
        if self.pdf_files:
            self.current_pdf_index = 0
            self.log_message(f"{len(self.pdf_files)}個のPDFファイルを読み込みました")
            self.load_current_pdf()
        else:
            self.log_message("PDFファイルが見つかりません")
    
    def load_current_pdf(self):
        """Load and display current PDF"""
        # ocr_get_imageフォルダ内のpngファイルを削除
        ocr_folder = self.config.get('ocr_image_folder')
        if ocr_folder and os.path.isdir(ocr_folder):
            for filename in os.listdir(ocr_folder):
                if filename.lower().endswith('.png'):
                    try:
                        os.remove(os.path.join(ocr_folder, filename))
                    except Exception as e:
                        self.log_message(f"PNGファイル削除エラー: {e}")

        if not self.pdf_files:
            return
        
        pdf_path = os.path.join(self.config['pdf_input_folder'], self.pdf_files[self.current_pdf_index])
        
        try:
            # Close previous document
            if self.current_pdf_doc:
                self.current_pdf_doc.close()

            # Open new document
            self.current_pdf_doc = fitz.open(pdf_path)

            # Render into viewer
            self.render_current_page()

            # Update file info and side images
            self.update_file_info()
            self.update_display_images()

            self.log_message(f"PDFを読み込みました: {self.pdf_files[self.current_pdf_index]}")

        except Exception as e:
            self.log_message(f"PDFの読み込みエラー: {str(e)}")

    def render_current_page(self):
        """Render the first page left-half and display filling the PDF canvas."""
        if not self.current_pdf_doc:
            return
        try:
            page = self.current_pdf_doc[0]
            # Render page to image
            mat = fitz.Matrix(2.0, 2.0)
            pix = page.get_pixmap(matrix=mat)
            img_data = pix.tobytes("ppm")

            pil_image = Image.open(io.BytesIO(img_data))

            # Crop to left half (app仕様に合わせて維持)
            width, height = pil_image.size
            left_half = pil_image.crop((0, 0, width // 2, height))

            # Canvas size
            canvas_width = max(1, self.pdf_canvas.winfo_width())
            canvas_height = max(1, self.pdf_canvas.winfo_height())

            # Scale to cover (fill) the canvas while preserving aspect ratio
            img_w, img_h = left_half.size
            scale = max(canvas_width / img_w, canvas_height / img_h)
            new_w = int(img_w * scale)
            new_h = int(img_h * scale)
            resized = left_half.resize((new_w, new_h), Image.Resampling.LANCZOS)

            # Center-crop to canvas size for true full-screen fill
            left = max(0, (new_w - canvas_width) // 2)
            top = max(0, (new_h - canvas_height) // 2)
            right = left + canvas_width
            bottom = top + canvas_height
            filled = resized.crop((left, top, right, bottom))

            # Save transform state
            self._render_scale = scale
            self._crop_left = left
            self._crop_top = top

            # Display
            self.pdf_image = ImageTk.PhotoImage(filled)
            self.pdf_canvas.delete("all")
            self.pdf_canvas.create_image(canvas_width//2, canvas_height//2, image=self.pdf_image)

            # Draw red and blue frames
            self.draw_frames()

        except Exception as e:
            self.log_message(f"PDF描画エラー: {str(e)}")

    def on_pdf_canvas_configure(self, event):
        """Re-render current page when the PDF canvas size changes."""
        # Avoid excessive redraws by using after_idle
        if hasattr(self, "_resize_after_id") and self._resize_after_id:
            try:
                self.root.after_cancel(self._resize_after_id)
            except Exception:
                pass
        self._resize_after_id = self.root.after(50, self.render_current_page)
    
    def draw_ocr_rectangle(self):
        """Draw red rectangle for OCR area"""
        if hasattr(self, 'pdf_image'):
            # Calculate rectangle position based on image scale
            canvas_width = self.pdf_canvas.winfo_width()
            canvas_height = self.pdf_canvas.winfo_height()
            
            # Scale OCR coordinates to canvas coordinates
            scale_x = canvas_width / (self.pdf_image.width() * 2)  # *2 because we show left half
            scale_y = canvas_height / self.pdf_image.height()
            
            x1 = self.config['ocr_x'] * scale_x
            y1 = self.config['ocr_y'] * scale_y
            x2 = (self.config['ocr_x'] + self.config['ocr_width']) * scale_x
            y2 = (self.config['ocr_y'] + self.config['ocr_height']) * scale_y
            
            self.pdf_canvas.create_rectangle(x1, y1, x2, y2, outline='red', width=3, tags='ocr_rect')
    
    def update_file_info(self):
        """Update file information display"""
        if self.pdf_files:
            current_file = self.pdf_files[self.current_pdf_index]
            info_text = f"{current_file} ({self.current_pdf_index + 1}/{len(self.pdf_files)})"
            self.file_info_label.config(text=info_text)
            # Prevボタンは1ページ目では無効化
            try:
                if self.current_pdf_index == 0:
                    self.prev_button.state(['disabled'])
                else:
                    self.prev_button.state(['!disabled'])
            except Exception:
                try:
                    self.prev_button.configure(state='disabled' if self.current_pdf_index == 0 else 'normal')
                except Exception:
                    pass
            # Nextボタンは「次ページが未保存(=CSVに行が無い)」または「最終ページ」で無効化
            try:
                if self.current_pdf_index >= len(self.pdf_files) - 1:
                    # 最終ページ
                    self.next_button.state(['disabled'])
                else:
                    next_page_no = self.current_pdf_index + 2  # 1始まり
                    has_row = False
                    try:
                        row = self.get_csv_row_by_index(next_page_no)
                        has_row = bool(row)
                    except Exception:
                        has_row = False
                    if has_row:
                        self.next_button.state(['!disabled'])
                    else:
                        self.next_button.state(['disabled'])
            except Exception:
                try:
                    if self.current_pdf_index >= len(self.pdf_files) - 1:
                        self.next_button.configure(state='disabled')
                    else:
                        next_page_no = self.current_pdf_index + 2
                        row = self.get_csv_row_by_index(next_page_no)
                        self.next_button.configure(state='normal' if row else 'disabled')
                except Exception:
                    pass
    
    def extract_ocr_text(self):
        """Extract text from OCR area"""
        if not self.current_pdf_doc:
            return
        try:
            page = self.current_pdf_doc[0]
            
            # Define OCR area rectangle
            rect = fitz.Rect(
                self.config['ocr_x'],
                self.config['ocr_y'],
                self.config['ocr_x'] + self.config['ocr_width'],
                self.config['ocr_y'] + self.config['ocr_height']
            )
            
            # Extract image from OCR area
            mat = fitz.Matrix(4.0, 4.0)  # High resolution for OCR
            pix = page.get_pixmap(matrix=mat, clip=rect)
            img_data = pix.tobytes("ppm")
            
            # Convert to OpenCV format for processing
            pil_image = Image.open(io.BytesIO(img_data))
            cv_image = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)
            
            # Image preprocessing for better OCR
            processed_image = self.preprocess_image_for_ocr(cv_image)
            
            # Save OCR image
            # Windowsのファイルシステムエンコーディング問題を回避するための処理
            base_name = os.path.splitext(self.pdf_files[self.current_pdf_index])[0]
            try:
                # エンコーディングの不整合を修正する
                corrected_name = base_name.encode('utf-8').decode('cp932')
            except (UnicodeEncodeError, UnicodeDecodeError):
                corrected_name = base_name # 変換に失敗した場合は元の名前を使用

            ocr_image_filename = f"{corrected_name}_ocr.png"
            ocr_image_path = os.path.join(self.config['ocr_image_folder'], ocr_image_filename)
            # cv2.imwriteは日本語ファイルパスをサポートしないため、エンコードしてから書き込む
            result, encoded_image = cv2.imencode('.png', processed_image)
            if result:
                with open(ocr_image_path, 'wb') as f:
                    f.write(encoded_image)
            
            # Display OCR image
            self.display_ocr_image(processed_image)
            
            # Perform OCR
            text = self.perform_ocr(processed_image)
            
            # Extract digits
            extracted_digits = self.extract_digits(text)
            
            # Update filename entry
            self.filename_entry.delete(0, tk.END)
            self.filename_entry.insert(0, extracted_digits)
            
            self.log_message(f"OCR結果: {text} -> 抽出: {extracted_digits}")
            
        except Exception as e:
            self.log_message(f"OCRエラー: {str(e)}")
            self.filename_entry.delete(0, tk.END)
    
    def preprocess_image_for_ocr(self, image):
        """Preprocess image for better OCR accuracy"""
        # Convert to grayscale
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        
        # Apply CLAHE (Contrast Limited Adaptive Histogram Equalization)
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        enhanced = clahe.apply(gray)
        
        # Median blur to reduce noise
        denoised = cv2.medianBlur(enhanced, 3)
        
        # Adaptive thresholding
        binary = cv2.adaptiveThreshold(
            denoised, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 10
        )
        
        # Morphological operations
        kernel = np.ones((2, 2), np.uint8)
        closed = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, kernel, iterations=1)
        
        # Dilate to thicken text
        dilated = cv2.dilate(closed, np.ones((1, 1), np.uint8), iterations=1)
        
        # Resize for better OCR
        height, width = dilated.shape
        resized = cv2.resize(dilated, (int(width * 1.8), int(height * 1.8)), 
                           interpolation=cv2.INTER_LANCZOS4)
        
        return resized
    
    def display_ocr_image(self, cv_image):
        """Display OCR image in the OCR canvas"""
        try:
            # Convert OpenCV image to PIL
            if len(cv_image.shape) == 3:
                pil_image = Image.fromarray(cv2.cvtColor(cv_image, cv2.COLOR_BGR2RGB))
            else:
                pil_image = Image.fromarray(cv_image)
            
            # Resize to fit OCR canvas
            canvas_width = self.ocr_canvas.winfo_width()
            canvas_height = self.ocr_canvas.winfo_height()
            
            if canvas_width > 1 and canvas_height > 1:
                pil_image.thumbnail((canvas_width - 10, canvas_height - 10), Image.Resampling.LANCZOS)
            
            # Convert to PhotoImage and display
            self.ocr_image = ImageTk.PhotoImage(pil_image)
            self.ocr_canvas.delete("all")
            self.ocr_canvas.create_image(canvas_width//2, canvas_height//2, image=self.ocr_image)
            
        except Exception as e:
            self.log_message(f"OCR画像表示エラー: {str(e)}")
    
    def perform_ocr(self, image):
        """Perform OCR on preprocessed image"""
        try:
            # OCR configuration for digits and hyphens only
            config = "--oem 1 --psm 7 -c tessedit_char_whitelist=0123456789- " \
                    "-c load_system_dawg=0 -c load_freq_dawg=0"
            
            text = pytesseract.image_to_string(image, config=config, lang='eng')
            return text.strip()
            
        except Exception as e:
            self.log_message(f"OCR実行エラー: {str(e)}")
            return ""
    
    def extract_digits(self, text):
        """Extract 8-digit + 999 pattern from OCR text"""
        # Remove all non-digit characters except hyphens
        cleaned = re.sub(r'[^0-9-]', '', text)
        
        # Look for 8-digit + 999 pattern
        pattern = r'(\d{8}).*?(\d{3})'
        match = re.search(pattern, cleaned)
        
        if match:
            return match.group(1)  # Return 8-digit part
        
        # Fallback: extract first 8 digits
        digits_only = re.sub(r'[^0-9]', '', cleaned)
        if len(digits_only) >= 8:
            return digits_only[:8]
        
        return digits_only
    
    def set_ocr_area(self):
        """Enable OCR area selection mode"""
        self.selecting_area = True
        self.log_message("OCR範囲を選択してください（ドラッグで矩形を描画）")
        self.pdf_canvas.config(cursor="crosshair")
    
    def on_canvas_click(self, event):
        """Handle canvas click for area selection"""
        if self.selecting_area:
            self.start_x = event.x
            self.start_y = event.y
    
    def on_canvas_drag(self, event):
        """Handle canvas drag for area selection"""
        if self.selecting_area:
            if hasattr(self, 'rect_id') and self.rect_id:
                self.pdf_canvas.delete(self.rect_id)

            self.rect_id = self.pdf_canvas.create_rectangle(
                self.start_x, self.start_y, event.x, event.y,
                outline=self.current_selection_color, width=2, tags="selection_rect"
            )
    
    def on_canvas_release(self, event):
        """Handle canvas release for area selection - 新仕様対応"""
        if hasattr(self, 'selecting_area') and self.selecting_area:
            area_type = self.selecting_area
            self.selecting_area = False
            self.pdf_canvas.config(cursor="")
            
            # Calculate area coordinates
            canvas_width = self.pdf_canvas.winfo_width()
            canvas_height = self.pdf_canvas.winfo_height()
            
            if hasattr(self, 'pdf_image'):
                # Convert canvas coordinates to PDF-left-half coordinates
                scale = self._render_scale if hasattr(self, '_render_scale') else 1.0
                off_x = self._crop_left if hasattr(self, '_crop_left') else 0
                off_y = self._crop_top if hasattr(self, '_crop_top') else 0

                c_x1 = min(self.start_x, event.x)
                c_y1 = min(self.start_y, event.y)
                c_x2 = max(self.start_x, event.x)
                c_y2 = max(self.start_y, event.y)

                # Inverse transform including render matrix scale (2.0):
                # pdf = (canvas + offset) / (matrix_scale * scale)
                matrix_scale = 2.0
                denom = (matrix_scale * scale) if (matrix_scale * scale) != 0 else 1.0
                x1 = (c_x1 + off_x) / denom
                y1 = (c_y1 + off_y) / denom
                x2 = (c_x2 + off_x) / denom
                y2 = (c_y2 + off_y) / denom
                
                # Update configuration based on area type
                if area_type == 'center':
                    self.config['red_frame_x'] = int(x1)
                    self.config['red_frame_y'] = int(y1)
                    self.config['red_frame_width'] = int(x2 - x1)
                    self.config['red_frame_height'] = int(y2 - y1)
                    self.log_message(f"中央表示エリア（赤枠）を更新: ({int(x1)}, {int(y1)}, {int(x2-x1)}, {int(y2-y1)})")
                elif area_type == 'right':
                    self.config['blue_frame_x'] = int(x1)
                    self.config['blue_frame_y'] = int(y1)
                    self.config['blue_frame_width'] = int(x2 - x1)
                    self.config['blue_frame_height'] = int(y2 - y1)
                    self.log_message(f"右側表示エリア（青枠）を更新: ({int(x1)}, {int(y1)}, {int(x2-x1)}, {int(y2-y1)})")
                
                # Save configuration
                self.save_config()
                
                # Redraw frames
                self.pdf_canvas.delete("selection_rect")
                self.draw_frames()
                
                # Update display images
                self.update_display_images()
    
    def save_file(self):
        """Save current PDF with new filename"""
        if not self.pdf_files or not self.current_pdf_doc:
            messagebox.showwarning("警告", "保存するPDFファイルがありません")
            return
        
        new_filename = self.filename_entry.get().strip()
        if not new_filename:
            messagebox.showwarning("警告", "ファイル名を入力してください")
            return
        
        if not new_filename.isdigit() or len(new_filename) != 8:
            messagebox.showwarning("警告", "ファイル名は8桁の数字である必要があります")
            return
        
        try:
            # Source and destination paths
            source_path = os.path.join(self.config['pdf_input_folder'], self.pdf_files[self.current_pdf_index])
            dest_path = os.path.join(self.config['pdf_output_folder'], f"{new_filename}.pdf")
            
            # Copy file
            shutil.copy2(source_path, dest_path)
            
            # Log to file
            self.log_to_file(self.pdf_files[self.current_pdf_index], new_filename)
            
            self.log_message(f"ファイルを保存しました: {new_filename}.pdf")
            
            # Move to next PDF
            self.next_pdf()
            
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの保存に失敗しました: {str(e)}")
    
    def log_to_file(self, original_filename, new_filename):
        """Log rename operation to file"""
        today = datetime.now().strftime("%Y%m%d")
        log_file = os.path.join(self.config['log_output_folder'], f"{today}.txt")
        
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(f"{new_filename}\n")
    
    def prev_pdf(self):
        """Go to previous PDF"""
        if self.pdf_files and self.current_pdf_index > 0:
            self.current_pdf_index -= 1
            self.load_current_pdf()
            # 前ページのページ番号（1始まり）に対応するCSV行をフォームへ反映
            try:
                page_no = self.current_pdf_index + 1
                row = self.get_csv_row_by_index(page_no)
                if row:
                    # row: [value] or [value, placeholder] or [value, placeholder, seq]
                    value = row[0] if len(row) >= 1 else ""
                    placeholder = row[1] if len(row) >= 2 else ""
                    # 反映
                    if hasattr(self, 'entry_var'):
                        self.entry_var.set(value)
                    if hasattr(self, 'result_var') and self.result_var is not None:
                        self.result_var.set(placeholder)
                    # 保存ボタンの状態を更新
                    try:
                        self.update_save_button_state()
                    except Exception:
                        pass
                else:
                    self.log_message(f"CSVにページ{page_no}の行が見つかりません")
            except Exception as e:
                self.log_message(f"CSV参照エラー: {e}")
    
    def next_pdf(self):
        """Go to next PDF"""
        if self.pdf_files and self.current_pdf_index < len(self.pdf_files) - 1:
            self.current_pdf_index += 1
            self.load_current_pdf()
            # 次ページのページ番号（1始まり）に対応するCSV行をフォームへ反映（prevと同様）
            try:
                page_no = self.current_pdf_index + 1
                row = self.get_csv_row_by_index(page_no)
                if row:
                    value = row[0] if len(row) >= 1 else ""
                    placeholder = row[1] if len(row) >= 2 else ""
                    if hasattr(self, 'entry_var'):
                        self.entry_var.set(value)
                    if hasattr(self, 'result_var') and self.result_var is not None:
                        self.result_var.set(placeholder)
                    try:
                        self.update_save_button_state()
                    except Exception:
                        pass
                else:
                    self.log_message(f"CSVにページ{page_no}の行が見つかりません")
            except Exception as e:
                self.log_message(f"CSV参照エラー: {e}")
        elif self.pdf_files and self.current_pdf_index == len(self.pdf_files) - 1:
            messagebox.showinfo("完了", "全てのPDFファイルの処理が完了しました")
    
    # === 新仕様対応メソッド ===
    def set_center_area(self):
        """Set center display area (red frame)"""
        self.selecting_area = 'center'
        self.current_selection_color = 'red'
        self.pdf_canvas.config(cursor="crosshair")
        self.log_message("中央表示エリアを設定してください（赤枠をドラッグ）")
    
    def set_right_area(self):
        """Set right display area (blue frame)"""
        self.selecting_area = 'right'
        self.current_selection_color = 'blue'
        self.pdf_canvas.config(cursor="crosshair")
        self.log_message("右側表示エリアを設定してください（青枠をドラッグ）")
    
    def draw_frames(self):
        """Draw red and blue frames on PDF canvas"""
        if hasattr(self, 'pdf_image'):
            # 古い枠を全て削除
            self.pdf_canvas.delete('red_frame')
            self.pdf_canvas.delete('blue_frame')
            
            canvas_width = self.pdf_canvas.winfo_width()
            canvas_height = self.pdf_canvas.winfo_height()
            
            # Use stored transform (cover mode on left-half)
            scale = self._render_scale if hasattr(self, '_render_scale') else 1.0
            off_x = self._crop_left if hasattr(self, '_crop_left') else 0
            off_y = self._crop_top if hasattr(self, '_crop_top') else 0
            
            # Draw red frame (center area)
            if all(key in self.config for key in ['red_frame_x', 'red_frame_y', 'red_frame_width', 'red_frame_height']):
                matrix_scale = 2.0
                x1 = self.config['red_frame_x'] * (matrix_scale * scale) - off_x
                y1 = self.config['red_frame_y'] * (matrix_scale * scale) - off_y
                x2 = (self.config['red_frame_x'] + self.config['red_frame_width']) * (matrix_scale * scale) - off_x
                y2 = (self.config['red_frame_y'] + self.config['red_frame_height']) * (matrix_scale * scale) - off_y
                self.pdf_canvas.create_rectangle(x1, y1, x2, y2, outline='red', width=3, tags='red_frame')
            
            # Draw blue frame (right area)
            if all(key in self.config for key in ['blue_frame_x', 'blue_frame_y', 'blue_frame_width', 'blue_frame_height']):
                matrix_scale = 2.0
                x1 = self.config['blue_frame_x'] * (matrix_scale * scale) - off_x
                y1 = self.config['blue_frame_y'] * (matrix_scale * scale) - off_y
                x2 = (self.config['blue_frame_x'] + self.config['blue_frame_width']) * (matrix_scale * scale) - off_x
                y2 = (self.config['blue_frame_y'] + self.config['blue_frame_height']) * (matrix_scale * scale) - off_y
                self.pdf_canvas.create_rectangle(x1, y1, x2, y2, outline='blue', width=3, tags='blue_frame')
    
    def update_display_images(self):
        """Update center and right display images based on frame areas"""
        if not self.current_pdf_doc:
            return
        # キャンバスの実サイズが未確定（初回起動直後など）の場合は少し待って再実行
        try:
            cw1 = self.center_canvas.winfo_width()
            ch1 = self.center_canvas.winfo_height()
            cw2 = self.right_canvas.winfo_width()
            ch2 = self.right_canvas.winfo_height()
            if min(cw1, ch1, cw2, ch2) <= 1:
                self.root.after(60, self.update_display_images)
                return
        except Exception:
            pass
        
        try:
            page = self.current_pdf_doc[0]
            
            # Update center display (red frame area)
            if all(key in self.config for key in ['red_frame_x', 'red_frame_y', 'red_frame_width', 'red_frame_height']):
                self.extract_and_display_area('center', 
                    self.config['red_frame_x'], self.config['red_frame_y'],
                    self.config['red_frame_width'], self.config['red_frame_height'])
            
            # Update right display (blue frame area)
            if all(key in self.config for key in ['blue_frame_x', 'blue_frame_y', 'blue_frame_width', 'blue_frame_height']):
                self.extract_and_display_area('right',
                    self.config['blue_frame_x'], self.config['blue_frame_y'],
                    self.config['blue_frame_width'], self.config['blue_frame_height'])
                    
        except Exception as e:
            self.log_message(f"画像表示エラー: {str(e)}")
    
    def on_side_canvas_configure(self, event):
        """Debounced update of side preview images on canvas resize."""
        if hasattr(self, "_side_resize_after_id") and self._side_resize_after_id:
            try:
                self.root.after_cancel(self._side_resize_after_id)
            except Exception:
                pass
        self._side_resize_after_id = self.root.after(50, self.update_display_images)
    
    def extract_and_display_area(self, area_type, x, y, width, height):
        """Extract and display image from specified area"""
        try:
            page = self.current_pdf_doc[0]
            
            # Define area rectangle
            rect = fitz.Rect(x, y, x + width, y + height)
            
            # Extract image from area
            mat = fitz.Matrix(2.0, 2.0)  # Scale factor
            pix = page.get_pixmap(matrix=mat, clip=rect)
            img_data = pix.tobytes("ppm")
            
            # Convert to PIL Image
            pil_image = Image.open(io.BytesIO(img_data))
            
            # Get target canvas
            target_canvas = self.center_canvas if area_type == 'center' else self.right_canvas
            
            # Resize to fit canvas
            canvas_width = target_canvas.winfo_width()
            canvas_height = target_canvas.winfo_height()
            
            if canvas_width > 1 and canvas_height > 1:
                pil_image.thumbnail((canvas_width, canvas_height), Image.Resampling.LANCZOS)
            
            # Convert to PhotoImage and display
            if area_type == 'center':
                self.center_image = ImageTk.PhotoImage(pil_image)
                self.center_canvas.delete("all")
                self.center_canvas.create_image(canvas_width//2, canvas_height//2, image=self.center_image)
            else:
                self.right_image = ImageTk.PhotoImage(pil_image)
                self.right_canvas.delete("all")
                self.right_canvas.create_image(canvas_width//2, canvas_height//2, image=self.right_image)
                
        except Exception as e:
            self.log_message(f"{area_type}エリア画像抽出エラー: {str(e)}")

if __name__ == "__main__":
    import io
    
    root = tk.Tk()
    app = PDFRenamerApp(root)
    root.mainloop()
