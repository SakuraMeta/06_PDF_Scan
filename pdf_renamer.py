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
        
        # Create folders
        self.create_folders()
        
        # Setup UI
        self.setup_ui()
        
        # Setup Tesseract path
        self.setup_tesseract()
    
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
        
        # Top frame for controls
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 入力フォルダ選択ボタン（左端）
        ttk.Button(top_frame, text="入力フォルダ選択", command=self.select_input_folder, 
                  style='Large.TButton').pack(side=tk.LEFT, padx=(0, 10))
        
        # Middle frame for 3-column layout (1:1:1)
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
        
        # === 中央: 表示画像（中） (1/3) ===
        center_frame = ttk.Frame(middle_frame)
        center_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 5))
        
        # 表示画像を設定（中）ボタン
        center_btn_frame = ttk.Frame(center_frame)
        center_btn_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Button(center_btn_frame, text="表示画像を設定（中）", command=self.set_center_area, 
                  style='Large.TButton').pack(side=tk.LEFT)
        
        # 中央表示エリア
        center_display_frame = ttk.LabelFrame(center_frame, text="表示画像（中）")
        center_display_frame.pack(fill=tk.BOTH, expand=True)
        
        self.center_canvas = tk.Canvas(center_display_frame, bg='lightgray', relief=tk.SUNKEN, bd=2)
        self.center_canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # === 右側: 表示画像（右） (1/3) ===
        right_frame = ttk.Frame(middle_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # 表示画像を設定（右）ボタン
        right_btn_frame = ttk.Frame(right_frame)
        right_btn_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Button(right_btn_frame, text="表示画像を設定（右）", command=self.set_right_area, 
                  style='Large.TButton').pack(side=tk.LEFT)
        
        # 右側表示エリア
        right_display_frame = ttk.LabelFrame(right_frame, text="表示画像（右）")
        right_display_frame.pack(fill=tk.BOTH, expand=True)
        
        self.right_canvas = tk.Canvas(right_display_frame, bg='lightgray', relief=tk.SUNKEN, bd=2)
        self.right_canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Navigation frame - 簡素化
        nav_frame = ttk.Frame(main_frame)
        nav_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Navigation buttons and file info (centered)
        nav_controls = ttk.Frame(nav_frame)
        nav_controls.pack(anchor=tk.CENTER)
        
        ttk.Button(nav_controls, text="<< 前へ", command=self.prev_pdf, 
                  style='Large.TButton').pack(side=tk.LEFT, padx=(0, 10))
        
        self.file_info_label = ttk.Label(nav_controls, text="", font=('Arial', 16))
        self.file_info_label.pack(side=tk.LEFT, padx=(10, 10))
        
        ttk.Button(nav_controls, text=">> 次へ", command=self.next_pdf, 
                  style='Large.TButton').pack(side=tk.LEFT, padx=(10, 0))
        
        # Log frame
        log_frame = ttk.LabelFrame(main_frame, text="ログ")
        log_frame.pack(fill=tk.X)
        
        self.log_text = tk.Text(log_frame, height=4, font=('Arial', 10))
        log_scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0), pady=5)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5, padx=(0, 5))
        
        # Configure button style
        style = ttk.Style()
        style.configure('Large.TButton', font=('Arial', 16))
    
    def validate_input(self, value):
        """Validate input to allow only digits and max 8 characters"""
        if len(value) <= 8 and (value.isdigit() or value == ""):
            return True
        return False
    
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
            page = self.current_pdf_doc[0]  # First page
            
            # Render page as image (left half only)
            mat = fitz.Matrix(2.0, 2.0)  # Scale factor
            pix = page.get_pixmap(matrix=mat)
            img_data = pix.tobytes("ppm")
            
            # Convert to PIL Image
            pil_image = Image.open(io.BytesIO(img_data))
            
            # Crop to left half
            width, height = pil_image.size
            left_half = pil_image.crop((0, 0, width // 2, height))
            
            # Resize to fit canvas
            canvas_width = self.pdf_canvas.winfo_width()
            canvas_height = self.pdf_canvas.winfo_height()
            
            if canvas_width > 1 and canvas_height > 1:
                left_half.thumbnail((canvas_width, canvas_height), Image.Resampling.LANCZOS)
            
            # Convert to PhotoImage and display
            self.pdf_image = ImageTk.PhotoImage(left_half)
            self.pdf_canvas.delete("all")
            self.pdf_canvas.create_image(canvas_width//2, canvas_height//2, image=self.pdf_image)
            
            # Draw red and blue frames
            self.draw_frames()
            
            # Update file info
            self.update_file_info()
            
            # Update display images
            self.update_display_images()
            
            self.log_message(f"PDFを読み込みました: {self.pdf_files[self.current_pdf_index]}")
            
        except Exception as e:
            self.log_message(f"PDFの読み込みエラー: {str(e)}")
    
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
            ocr_image_path = os.path.join(
                self.config['ocr_image_folder'],
                f"{os.path.splitext(self.pdf_files[self.current_pdf_index])[0]}_ocr.png"
            )
            cv2.imwrite(ocr_image_path, processed_image)
            
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
                # Convert canvas coordinates to PDF coordinates
                scale_x = (self.pdf_image.width() * 2) / canvas_width  # *2 because we show left half
                scale_y = self.pdf_image.height() / canvas_height
                
                x1 = min(self.start_x, event.x) * scale_x
                y1 = min(self.start_y, event.y) * scale_y
                x2 = max(self.start_x, event.x) * scale_x
                y2 = max(self.start_y, event.y) * scale_y
                
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
    
    def next_pdf(self):
        """Go to next PDF"""
        if self.pdf_files and self.current_pdf_index < len(self.pdf_files) - 1:
            self.current_pdf_index += 1
            self.load_current_pdf()
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
            
            # Scale coordinates to canvas coordinates
            scale_x = canvas_width / (self.pdf_image.width() * 2)  # *2 because we show left half
            scale_y = canvas_height / self.pdf_image.height()
            
            # Draw red frame (center area)
            if all(key in self.config for key in ['red_frame_x', 'red_frame_y', 'red_frame_width', 'red_frame_height']):
                x1 = self.config['red_frame_x'] * scale_x
                y1 = self.config['red_frame_y'] * scale_y
                x2 = (self.config['red_frame_x'] + self.config['red_frame_width']) * scale_x
                y2 = (self.config['red_frame_y'] + self.config['red_frame_height']) * scale_y
                self.pdf_canvas.create_rectangle(x1, y1, x2, y2, outline='red', width=3, tags='red_frame')
            
            # Draw blue frame (right area)
            if all(key in self.config for key in ['blue_frame_x', 'blue_frame_y', 'blue_frame_width', 'blue_frame_height']):
                x1 = self.config['blue_frame_x'] * scale_x
                y1 = self.config['blue_frame_y'] * scale_y
                x2 = (self.config['blue_frame_x'] + self.config['blue_frame_width']) * scale_x
                y2 = (self.config['blue_frame_y'] + self.config['blue_frame_height']) * scale_y
                self.pdf_canvas.create_rectangle(x1, y1, x2, y2, outline='blue', width=3, tags='blue_frame')
    
    def update_display_images(self):
        """Update center and right display images based on frame areas"""
        if not self.current_pdf_doc:
            return
        
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
