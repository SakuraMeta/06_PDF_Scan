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
        
        # Variables
        self.pdf_files = []
        self.current_pdf_index = 0
        self.current_pdf_doc = None
        self.ocr_area = None
        self.selecting_area = False
        self.start_x = 0
        self.start_y = 0
        self.rect_id = None
        
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
                        if key in ['ocr_x', 'ocr_y', 'ocr_width', 'ocr_height', 'digit_length']:
                            config[key] = int(value)
                        else:
                            config[key] = value
        except FileNotFoundError:
            # Default configuration
            config = {
                'pdf_input_folder': 'pdf_input',
                'pdf_output_folder': 'pdf_output',
                'log_output_folder': 'log_output',
                'ocr_image_folder': 'ocr_get_image',
                'ocr_x': 600,
                'ocr_y': 250,
                'ocr_width': 300,
                'ocr_height': 50,
                'digit_length': 8
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
            f.write("\n# OCR Area Coordinates (x, y, width, height)\n")
            f.write(f"ocr_x={self.config['ocr_x']}\n")
            f.write(f"ocr_y={self.config['ocr_y']}\n")
            f.write(f"ocr_width={self.config['ocr_width']}\n")
            f.write(f"ocr_height={self.config['ocr_height']}\n")
            f.write("\n# Target digit length\n")
            f.write(f"digit_length={self.config['digit_length']}\n")
    
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
        """Setup the user interface"""
        # Main frame
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Top frame for controls
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Control buttons
        btn_font = ('Arial', 16)
        
        ttk.Button(top_frame, text="入力フォルダ選択", command=self.select_input_folder, 
                  style='Large.TButton').pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(top_frame, text="OCR範囲を設定", command=self.set_ocr_area, 
                  style='Large.TButton').pack(side=tk.LEFT, padx=(0, 10))
        
        # Middle frame for PDF and OCR display
        middle_frame = ttk.Frame(main_frame)
        middle_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Left frame for PDF viewer (50% width)
        left_frame = ttk.Frame(middle_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # PDF Canvas
        self.pdf_canvas = tk.Canvas(left_frame, bg='white', relief=tk.SUNKEN, bd=2)
        self.pdf_canvas.pack(fill=tk.BOTH, expand=True)
        self.pdf_canvas.bind("<Button-1>", self.on_canvas_click)
        self.pdf_canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.pdf_canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        
        # Right frame (50% width)
        right_frame = ttk.Frame(middle_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # OCR Image frame (6/7 of right frame height)
        ocr_frame = ttk.LabelFrame(right_frame, text="OCRイメージ")
        ocr_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        self.ocr_canvas = tk.Canvas(ocr_frame, bg='white', relief=tk.SUNKEN, bd=2)
        self.ocr_canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Future use frame (1/7 of right frame height)
        future_frame = ttk.LabelFrame(right_frame, text="後ほど使用するエリア")
        future_frame.pack(fill=tk.X, pady=(5, 0))
        future_frame.configure(height=80)
        future_frame.pack_propagate(False)
        
        # Center frame for filename input
        center_frame = ttk.Frame(main_frame)
        center_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Filename input section (centered)
        filename_frame = ttk.Frame(center_frame)
        filename_frame.pack(anchor=tk.CENTER)
        
        ttk.Label(filename_frame, text="ファイル名", font=('Arial', 16)).pack(side=tk.LEFT, padx=(0, 10))
        
        # Text entry with validation for 8 digits, but wider display
        vcmd = (self.root.register(self.validate_input), '%P')
        self.filename_entry = ttk.Entry(filename_frame, font=('Arial', 16), width=16, validate='key', validatecommand=vcmd)
        self.filename_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(filename_frame, text="名前を付けて保存", command=self.save_file, 
                  style='Large.TButton').pack(side=tk.LEFT)
        
        # Navigation frame
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
            
            # Draw OCR area rectangle
            self.draw_ocr_rectangle()
            
            # Update file info
            self.update_file_info()
            
            # Extract OCR text
            self.extract_ocr_text()
            
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
            # Remove previous rectangle
            self.pdf_canvas.delete("selection_rect")
            
            # Draw new rectangle
            self.rect_id = self.pdf_canvas.create_rectangle(
                self.start_x, self.start_y, event.x, event.y,
                outline='red', width=2, tags="selection_rect"
            )
    
    def on_canvas_release(self, event):
        """Handle canvas release for area selection"""
        if self.selecting_area:
            self.selecting_area = False
            self.pdf_canvas.config(cursor="")
            
            # Calculate OCR area coordinates
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
                
                # Update configuration
                self.config['ocr_x'] = int(x1)
                self.config['ocr_y'] = int(y1)
                self.config['ocr_width'] = int(x2 - x1)
                self.config['ocr_height'] = int(y2 - y1)
                
                # Save configuration
                self.save_config()
                
                # Redraw OCR rectangle
                self.pdf_canvas.delete("selection_rect")
                self.draw_ocr_rectangle()
                
                # Re-extract OCR text
                self.extract_ocr_text()
                
                self.log_message(f"OCR範囲を更新しました: ({self.config['ocr_x']}, {self.config['ocr_y']}, {self.config['ocr_width']}, {self.config['ocr_height']})")
    
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
            timestamp = datetime.now().strftime("%H:%M:%S")
            f.write(f"[{timestamp}] {original_filename} -> {new_filename}.pdf\n")
    
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

if __name__ == "__main__":
    import io
    
    root = tk.Tk()
    app = PDFRenamerApp(root)
    root.mainloop()
