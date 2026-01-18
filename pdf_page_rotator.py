import PyPDF2
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import io
import fitz  # pymupdf

class PDFRotatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF页面旋转工具")
        self.root.geometry("900x700")
        
        # 设置图标（如果有的话）
        try:
            self.root.iconbitmap("pdf_icon.ico")
        except:
            pass
        
        # 变量初始化
        self.input_path = ""
        self.pdf_reader = None
        self.total_pages = 0
        self.rotations = {}  # 存储页码和旋转角度 {页码: 角度}
        self.page_previews = []  # 存储页面预览
        
        # 设置样式
        self.setup_styles()
        
        # 创建UI
        self.create_widgets()
        
    def setup_styles(self):
        """设置样式"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # 自定义颜色
        self.bg_color = "#f0f0f0"
        self.button_color = "#4a6fa5"
        self.highlight_color = "#6b9bd2"
        
        self.root.configure(bg=self.bg_color)
    
    def create_widgets(self):
        """创建界面组件"""
        # 标题
        title_frame = tk.Frame(self.root, bg=self.bg_color)
        title_frame.pack(pady=10)
        
        title_label = tk.Label(
            title_frame, 
            text="📄 PDF页面旋转工具", 
            font=("微软雅黑", 20, "bold"),
            bg=self.bg_color,
            fg="#2c3e50"
        )
        title_label.pack()
        
        subtitle_label = tk.Label(
            title_frame,
            text="旋转PDF页面方向，支持90°、180°、270°旋转",
            font=("微软雅黑", 10),
            bg=self.bg_color,
            fg="#7f8c8d"
        )
        subtitle_label.pack()
        
        # 分隔线
        ttk.Separator(self.root, orient='horizontal').pack(fill='x', padx=20, pady=10)
        
        # 主内容区域
        main_frame = tk.Frame(self.root, bg=self.bg_color)
        main_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # 左侧控制面板
        left_frame = tk.Frame(main_frame, bg=self.bg_color)
        left_frame.pack(side='left', fill='y', padx=(0, 10))
        
        # 文件选择区域
        file_frame = tk.LabelFrame(left_frame, text="文件操作", font=("微软雅黑", 11), 
                                   bg=self.bg_color, padx=10, pady=10)
        file_frame.pack(fill='x', pady=(0, 15))
        
        tk.Label(file_frame, text="选择PDF文件:", bg=self.bg_color, 
                font=("微软雅黑", 9)).pack(anchor='w', pady=(0, 5))
        
        # 文件路径显示
        self.file_path_var = tk.StringVar()
        self.file_path_entry = tk.Entry(
            file_frame, 
            textvariable=self.file_path_var,
            font=("微软雅黑", 9),
            state='readonly',
            width=30
        )
        self.file_path_entry.pack(side='left', fill='x', expand=True, padx=(0, 10))
        
        # 选择文件按钮
        self.select_btn = tk.Button(
            file_frame,
            text="浏览...",
            command=self.select_file,
            bg=self.button_color,
            fg="white",
            font=("微软雅黑", 9),
            relief="flat",
            padx=15,
            cursor="hand2"
        )
        self.select_btn.pack(side='right')
        
        # 文件信息显示
        self.info_frame = tk.Frame(left_frame, bg=self.bg_color)
        self.info_frame.pack(fill='x', pady=(0, 15))
        
        self.file_info_label = tk.Label(
            self.info_frame,
            text="未选择文件",
            bg=self.bg_color,
            font=("微软雅黑", 9),
            fg="#7f8c8d"
        )
        self.file_info_label.pack(anchor='w')
        
        # 批量操作区域
        batch_frame = tk.LabelFrame(left_frame, text="批量操作", font=("微软雅黑", 11), 
                                   bg=self.bg_color, padx=10, pady=10)
        batch_frame.pack(fill='x', pady=(0, 15))
        
        tk.Label(batch_frame, text="旋转所有页面至:", bg=self.bg_color, 
                font=("微软雅黑", 9)).pack(anchor='w', pady=(0, 5))
        
        # 批量旋转按钮
        batch_btn_frame = tk.Frame(batch_frame, bg=self.bg_color)
        batch_btn_frame.pack(fill='x')
        
        batch_buttons = [
            ("顺时针90°", 90),
            ("逆时针90°", 270),
            ("旋转180°", 180)
        ]
        
        for text, angle in batch_buttons:
            btn = tk.Button(
                batch_btn_frame,
                text=text,
                command=lambda a=angle: self.rotate_all_pages(a),
                bg="#e9ecef",
                fg="#495057",
                font=("微软雅黑", 9),
                relief="flat",
                padx=10,
                cursor="hand2"
            )
            btn.pack(side='left', padx=2, pady=5)
        
        # 保存按钮
        save_frame = tk.Frame(left_frame, bg=self.bg_color)
        save_frame.pack(fill='x', pady=(20, 0))
        
        self.save_btn = tk.Button(
            save_frame,
            text="💾 保存旋转后的PDF",
            command=self.save_pdf,
            bg="#27ae60",
            fg="white",
            font=("微软雅黑", 11, "bold"),
            relief="flat",
            padx=30,
            pady=10,
            cursor="hand2",
            state='disabled'
        )
        self.save_btn.pack(fill='x')
        
        # 右侧页面预览区域
        right_frame = tk.Frame(main_frame, bg=self.bg_color)
        right_frame.pack(side='right', fill='both', expand=True)
        
        # 页面预览标题
        preview_header = tk.Frame(right_frame, bg=self.bg_color)
        preview_header.pack(fill='x', pady=(0, 10))
        
        tk.Label(preview_header, text="页面预览与设置", font=("微软雅黑", 12, "bold"),
                bg=self.bg_color, fg="#2c3e50").pack(side='left')
        
        self.page_count_label = tk.Label(
            preview_header,
            text="共 0 页",
            font=("微软雅黑", 10),
            bg=self.bg_color,
            fg="#7f8c8d"
        )
        self.page_count_label.pack(side='right')
        
        # 创建滚动区域
        canvas_frame = tk.Frame(right_frame, bg=self.bg_color)
        canvas_frame.pack(fill='both', expand=True)
        
        # 创建Canvas和Scrollbar
        self.canvas = tk.Canvas(canvas_frame, bg=self.bg_color, highlightthickness=0)
        scrollbar = tk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        
        # 滚动区域内部框架
        self.scrollable_frame = tk.Frame(self.canvas, bg=self.bg_color)
        
        # 配置Canvas
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        # 鼠标滚轮绑定
        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        self.canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # 布局
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 状态栏
        self.status_bar = tk.Label(
            self.root,
            text="就绪",
            bd=1,
            relief=tk.SUNKEN,
            anchor=tk.W,
            font=("微软雅黑", 9),
            bg="#e9ecef",
            fg="#495057"
        )
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def select_file(self):
        """选择PDF文件"""
        file_path = filedialog.askopenfilename(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        
        if file_path:
            self.input_path = file_path
            self.file_path_var.set(os.path.basename(file_path))
            self.load_pdf()
    
    def load_pdf(self):
        """加载PDF文件"""
        try:
            with open(self.input_path, 'rb') as file:
                self.pdf_reader = PyPDF2.PdfReader(file)
                self.total_pages = len(self.pdf_reader.pages)
                
                # 重置旋转设置
                self.rotations = {}
                self.page_previews = []
                
                # 更新UI
                self.file_info_label.config(
                    text=f"文件: {os.path.basename(self.input_path)}\n"
                         f"大小: {os.path.getsize(self.input_path) // 1024} KB\n"
                         f"页数: {self.total_pages} 页"
                )
                
                self.page_count_label.config(text=f"共 {self.total_pages} 页")
                self.save_btn.config(state='normal')
                self.update_status(f"已加载PDF文件: {os.path.basename(self.input_path)}")
                
                # 创建页面预览
                self.create_page_previews()
                
        except Exception as e:
            messagebox.showerror("错误", f"无法加载PDF文件:\n{str(e)}")
            self.update_status("加载PDF文件失败")
    
    def create_page_previews(self):
        """创建页面预览"""
        # 清除旧的预览
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        self.page_previews = []
        
        # 打开PDF文件
        pdf_document = fitz.open(self.input_path)
        
        # 创建每个页面的预览项
        for page_num in range(min(self.total_pages, 50)):  # 限制预览页数
            # 获取页面缩略图
            page = pdf_document[page_num]
            pix = page.get_pixmap(matrix=fitz.Matrix(0.2, 0.2))  # 缩放因子0.2
            
            # 转换为PIL图像
            img_data = pix.tobytes("ppm")
            img = Image.open(io.BytesIO(img_data))
            img = img.resize((150, 200))  # 调整大小
            photo = ImageTk.PhotoImage(img)
            
            # 创建页面框架
            page_frame = tk.Frame(self.scrollable_frame, bg="white", relief="solid", bd=1)
            page_frame.pack(fill='x', pady=5, padx=5)
            
            # 页面标题
            page_header = tk.Frame(page_frame, bg="#f8f9fa")
            page_header.pack(fill='x', pady=(5, 0))
            
            tk.Label(
                page_header,
                text=f"第 {page_num + 1} 页",
                font=("微软雅黑", 10, "bold"),
                bg="#f8f9fa"
            ).pack(side='left', padx=10, pady=5)
            
            # 旋转控制
            control_frame = tk.Frame(page_header, bg="#f8f9fa")
            control_frame.pack(side='right', padx=10)
            
            # 旋转角度标签
            angle_label = tk.Label(
                control_frame,
                text="旋转: 0°",
                font=("微软雅黑", 9),
                bg="#f8f9fa",
                width=10
            )
            angle_label.pack(side='left', padx=5)
            
            # 旋转按钮
            btn_frame = tk.Frame(control_frame, bg="#f8f9fa")
            btn_frame.pack(side='left')
            
            # 存储页面数据（包括图像引用）
            page_data = {
                'frame': page_frame,
                'angle_label': angle_label,
                'page_num': page_num,
                'current_angle': 0,
                'image': photo  # 保存图像引用防止被垃圾回收
            }
            self.page_previews.append(page_data)
            
            # 创建旋转按钮
            buttons = [
                ("↶ 逆90°", -90),
                ("↷ 顺90°", 90),
                ("↻ 180°", 180),
                ("↺ 重置", 0)
            ]
            
            for text, angle_change in buttons:
                btn = tk.Button(
                    btn_frame,
                    text=text,
                    command=lambda pn=page_num, ac=angle_change: self.rotate_single_page(pn, ac),
                    bg="#e9ecef",
                    fg="#495057",
                    font=("微软雅黑", 8),
                    relief="flat",
                    padx=5,
                    cursor="hand2"
                )
                btn.pack(side='left', padx=2)
            
            # 显示缩略图
            img_label = tk.Label(
                page_frame,
                image=photo,
                bg="white",
                relief="groove",
                bd=1
            )
            img_label.pack(padx=10, pady=10)
        
        pdf_document.close()
    
    def rotate_single_page(self, page_num, angle_change):
        """旋转单个页面"""
        if page_num not in self.rotations:
            self.rotations[page_num] = 0
        
        # 计算新的角度
        new_angle = (self.rotations[page_num] + angle_change) % 360
        self.rotations[page_num] = new_angle
        
        # 更新UI
        page_data = self.page_previews[page_num]
        page_data['current_angle'] = new_angle
        page_data['angle_label'].config(text=f"旋转: {new_angle}°")
        
        # 高亮显示
        page_data['frame'].configure(bg="#e3f2fd")
        self.root.after(300, lambda: page_data['frame'].configure(bg="white"))
        
        self.update_status(f"第 {page_num + 1} 页设置为 {new_angle}° 旋转")
    
    def rotate_all_pages(self, angle):
        """旋转所有页面到指定角度"""
        if not self.pdf_reader:
            messagebox.showwarning("警告", "请先选择PDF文件")
            return
        
        # 确认操作
        if not messagebox.askyesno("确认", f"确定要将所有 {self.total_pages} 页旋转 {angle} 度吗？"):
            return
        
        # 设置所有页面的旋转角度
        for page_num in range(self.total_pages):
            self.rotations[page_num] = angle
            
            # 更新UI
            if page_num < len(self.page_previews):
                page_data = self.page_previews[page_num]
                page_data['current_angle'] = angle
                page_data['angle_label'].config(text=f"旋转: {angle}°")
        
        self.update_status(f"所有页面已设置为 {angle}° 旋转")
        messagebox.showinfo("完成", f"已设置所有页面旋转 {angle}°")
    
    def save_pdf(self):
        """保存旋转后的PDF"""
        if not self.input_path or not self.pdf_reader:
            messagebox.showwarning("警告", "请先选择PDF文件")
            return
        
        # 检查是否有旋转设置
        if not self.rotations:
            if not messagebox.askyesno("确认", "没有设置任何旋转，确定要继续吗？"):
                return
        
        # 选择保存位置
        default_name = os.path.splitext(os.path.basename(self.input_path))[0] + "_rotated.pdf"
        output_path = filedialog.asksaveasfilename(
            title="保存旋转后的PDF",
            initialfile=default_name,
            defaultextension=".pdf",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        
        if not output_path:
            return
        
        try:
            # 执行旋转
            with open(self.input_path, 'rb') as input_file:
                reader = PyPDF2.PdfReader(input_file)
                writer = PyPDF2.PdfWriter()
                
                for page_num in range(len(reader.pages)):
                    page = reader.pages[page_num]
                    
                    # 应用旋转（如果有的话）
                    if page_num in self.rotations and self.rotations[page_num] != 0:
                        page.rotate(self.rotations[page_num])
                    
                    writer.add_page(page)
                
                # 保存文件
                with open(output_path, 'wb') as output_file:
                    writer.write(output_file)
            
            # 成功消息
            rotation_count = sum(1 for angle in self.rotations.values() if angle != 0)
            messagebox.showinfo(
                "完成",
                f"PDF已成功保存！\n"
                f"文件: {os.path.basename(output_path)}\n"
                f"已旋转页面: {rotation_count} 页\n"
                f"保存位置: {output_path}"
            )
            
            self.update_status(f"PDF已保存: {os.path.basename(output_path)}")
            
            # 询问是否打开文件
            if messagebox.askyesno("打开文件", "是否打开保存的PDF文件？"):
                os.startfile(output_path)
            
        except Exception as e:
            messagebox.showerror("错误", f"保存PDF时出错:\n{str(e)}")
            self.update_status("保存失败")
    
    def update_status(self, message):
        """更新状态栏"""
        self.status_bar.config(text=f"状态: {message}")
        self.root.update()

def main():
    root = tk.Tk()
    app = PDFRotatorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()