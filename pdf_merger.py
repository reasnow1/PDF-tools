import tkinter as tk
from tkinter import filedialog, messagebox, Listbox, MULTIPLE
import PyPDF2
import os

class PDFMergerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF合并工具")
        self.root.geometry("600x500")
        
        self.files = []
        
        # 创建界面元素
        self.create_widgets()
    
    def create_widgets(self):
        # 标题
        title_label = tk.Label(self.root, text="PDF文件合并工具", font=("Arial", 16))
        title_label.pack(pady=10)
        
        # 按钮框架
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)
        
        # 添加文件按钮
        add_btn = tk.Button(btn_frame, text="添加PDF文件", command=self.add_files)
        add_btn.grid(row=0, column=0, padx=5)
        
        # 移除文件按钮
        remove_btn = tk.Button(btn_frame, text="移除选中", command=self.remove_file)
        remove_btn.grid(row=0, column=1, padx=5)
        
        # 清空列表按钮
        clear_btn = tk.Button(btn_frame, text="清空列表", command=self.clear_list)
        clear_btn.grid(row=0, column=2, padx=5)
        
        # 上移/下移按钮
        up_btn = tk.Button(btn_frame, text="上移", command=self.move_up)
        up_btn.grid(row=0, column=3, padx=5)
        
        down_btn = tk.Button(btn_frame, text="下移", command=self.move_down)
        down_btn.grid(row=0, column=4, padx=5)
        
        # 文件列表
        list_frame = tk.Frame(self.root)
        list_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.listbox = Listbox(list_frame, selectmode=MULTIPLE, 
                              yscrollcommand=scrollbar.set, height=15)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.listbox.yview)
        
        # 合并按钮
        merge_btn = tk.Button(self.root, text="合并PDF", command=self.merge_pdfs,
                             bg="green", fg="white", font=("Arial", 12))
        merge_btn.pack(pady=20)
        
        # 状态标签
        self.status_label = tk.Label(self.root, text="等待操作...", fg="blue")
        self.status_label.pack(pady=5)
    
    def add_files(self):
        files = filedialog.askopenfilenames(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        
        for file in files:
            if file not in self.files:
                self.files.append(file)
                self.listbox.insert(tk.END, os.path.basename(file))
        
        self.update_status()
    
    def remove_file(self):
        selected = self.listbox.curselection()
        for index in reversed(selected):
            self.files.pop(index)
            self.listbox.delete(index)
        self.update_status()
    
    def clear_list(self):
        self.files.clear()
        self.listbox.delete(0, tk.END)
        self.update_status()
    
    def move_up(self):
        selected = self.listbox.curselection()
        if not selected or selected[0] == 0:
            return
        
        for pos in selected:
            if pos > 0:
                self.files[pos], self.files[pos-1] = self.files[pos-1], self.files[pos]
        
        self.refresh_listbox()
        self.listbox.selection_set(selected[0]-1)
    
    def move_down(self):
        selected = self.listbox.curselection()
        if not selected or selected[-1] == len(self.files)-1:
            return
        
        for pos in reversed(selected):
            if pos < len(self.files)-1:
                self.files[pos], self.files[pos+1] = self.files[pos+1], self.files[pos]
        
        self.refresh_listbox()
        self.listbox.selection_set(selected[0]+1)
    
    def refresh_listbox(self):
        self.listbox.delete(0, tk.END)
        for file in self.files:
            self.listbox.insert(tk.END, os.path.basename(file))
    
    def update_status(self):
        count = len(self.files)
        self.status_label.config(text=f"已选择 {count} 个PDF文件")
    
    def merge_pdfs(self):
        if len(self.files) < 2:
            messagebox.showwarning("警告", "请至少选择2个PDF文件进行合并！")
            return
        
        # 获取第一个文件的名称（不含扩展名）作为默认名称
        if self.files:
            first_file_name = os.path.basename(self.files[0])
            base_name = os.path.splitext(first_file_name)[0]  # 移除扩展名
            default_name = base_name + "_merged.pdf"
        else:
            default_name = "merged.pdf"
        
        # 选择保存位置，并设置默认文件名
        output_file = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF文件", "*.pdf")],
            initialfile=default_name  # 设置默认文件名
        )
        
        if not output_file:
            return
        
        try:
            pdf_writer = PyPDF2.PdfWriter()
            
            for file in self.files:
                pdf_reader = PyPDF2.PdfReader(file)
                for page in range(len(pdf_reader.pages)):
                    pdf_writer.add_page(pdf_reader.pages[page])
            
            with open(output_file, 'wb') as out:
                pdf_writer.write(out)
            
            messagebox.showinfo("成功", f"PDF合并完成！\n保存至: {output_file}")
            self.status_label.config(text="合并完成！", fg="green")

            # 询问是否打开文件
            if messagebox.askyesno("打开文件", "是否打开保存的PDF文件？"):
                os.startfile(output_file)
            
        except Exception as e:
            messagebox.showerror("错误", f"合并失败：{str(e)}")
            self.status_label.config(text="合并失败！", fg="red")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFMergerGUI(root)
    root.mainloop()