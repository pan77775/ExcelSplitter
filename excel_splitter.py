import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os

class ExcelSplitter:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 分頁小幫手")
        self.root.geometry("600x600")
        self.root.resizable(False, False)  # 禁止調整視窗大小
        
        self.file_path = None
        self.df = None
        self.columns = []
        self.selected_columns = []
        self.checkboxes = {}  # 儲存所有勾選框的變數
        self.split_column = None  # 儲存分頁欄位
        
        # 建立主框架
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 設定主框架的網格權重
        root.grid_columnconfigure(0, weight=1)
        root.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)
        
        # 選擇檔案按鈕
        self.select_button = ttk.Button(self.main_frame, text="選擇 Excel 檔案", command=self.select_file)
        self.select_button.grid(row=0, column=0, pady=10)
        
        # 檔案路徑標籤
        self.file_label = ttk.Label(self.main_frame, text="尚未選擇檔案")
        self.file_label.grid(row=1, column=0, pady=5)
        
        # 狀態標籤（移到檔案路徑標籤下方）
        self.status_label = ttk.Label(self.main_frame, text="")
        self.status_label.grid(row=2, column=0, pady=5)
        
        # 分頁欄位選擇框架
        self.split_frame = ttk.LabelFrame(self.main_frame, text="分頁設定", padding="5")
        self.split_frame.grid(row=3, column=0, pady=10, sticky="ew")
        self.split_frame.grid_columnconfigure(0, weight=1)
        
        # 欄位選擇標籤
        self.column_label = ttk.Label(self.split_frame, text="請選擇要依據分頁的欄位：")
        self.column_label.grid(row=0, column=0, pady=5)
        
        # 欄位選擇下拉選單
        self.column_var = tk.StringVar()
        self.column_combo = ttk.Combobox(self.split_frame, textvariable=self.column_var, state="readonly", width=40)
        self.column_combo.grid(row=1, column=0, pady=5)
        self.column_combo.bind('<<ComboboxSelected>>', self.on_column_selected)
        
        # 輸出欄位選擇框架
        self.output_frame = ttk.LabelFrame(self.main_frame, text="輸出欄位選擇", padding="5")
        self.output_frame.grid(row=4, column=0, pady=10, sticky="nsew")
        self.output_frame.grid_columnconfigure(0, weight=1)
        self.output_frame.grid_rowconfigure(0, weight=1)
        
        # 建立滾動條
        self.canvas = tk.Canvas(self.output_frame, height=229)  # 設定固定高度
        self.scrollbar = ttk.Scrollbar(self.output_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        
        # 全選/取消全選按鈕
        self.select_all_var = tk.BooleanVar()
        self.select_all_check = ttk.Checkbutton(self.output_frame, text="全選", 
                                               variable=self.select_all_var,
                                               command=self.toggle_select_all)
        self.select_all_check.grid(row=1, column=0, columnspan=2, pady=5)
        
        # 執行分頁按鈕
        self.split_button = ttk.Button(self.main_frame, text="執行分頁", command=self.split_excel, state="disabled")
        self.split_button.grid(row=5, column=0, pady=10)
        
        # 分頁完成狀態標籤
        self.result_label = ttk.Label(self.main_frame, text="")
        self.result_label.grid(row=6, column=0, pady=5)
        
        # 綁定視窗大小改變事件
        self.root.bind('<Configure>', self.on_window_resize)
    
    def calculate_columns(self):
        if not self.columns:
            return 1
        
        # 獲取視窗寬度
        window_width = self.root.winfo_width()
        
        # 計算每個欄位需要的寬度（考慮padding和margin）
        column_width = 200  # 假設每個欄位需要200像素寬
        max_columns = max(1, window_width // column_width)
        
        # 根據欄位數量調整
        total_columns = len(self.columns)
        return min(max_columns, total_columns)
    
    def create_column_frames(self, num_columns):
        # 清除現有的框架
        for frame in self.column_frames:
            frame.destroy()
        self.column_frames.clear()
        
        # 建立新的框架
        for i in range(num_columns):
            frame = ttk.Frame(self.scrollable_frame)
            frame.grid(row=0, column=i, padx=5, sticky="nsew")
            self.column_frames.append(frame)
            self.scrollable_frame.grid_columnconfigure(i, weight=1)
    
    def on_window_resize(self, event):
        if hasattr(self, 'columns') and self.columns:
            num_columns = self.calculate_columns()
            if num_columns != len(self.column_frames):
                self.create_column_frames(num_columns)
                self.update_checkboxes()
    
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="選擇 Excel 檔案",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if file_path:
            self.file_path = file_path
            self.file_label.config(text=f"已選擇檔案：{os.path.basename(file_path)}")
            try:
                self.df = pd.read_excel(file_path)
                self.columns = self.df.columns.tolist()
                self.column_combo['values'] = self.columns
                self.column_combo.set('')
                self.split_column = None
                
                # 更新勾選框
                self.update_checkboxes()
                self.status_label.config(text="檔案讀取成功！請選擇分頁依據的欄位")
            except Exception as e:
                messagebox.showerror("錯誤", f"讀取檔案時發生錯誤：{str(e)}")
    
    def update_checkboxes(self):
        # 清除現有的勾選框
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.checkboxes.clear()
        
        if not self.columns:
            return
        
        # 固定每行顯示5個勾選框
        columns_per_row = 5
        
        # 建立新的勾選框
        for i, col in enumerate(self.columns):
            var = tk.BooleanVar()
            self.checkboxes[col] = var
            
            # 如果是分頁欄位，自動選中並禁用
            if col == self.split_column:
                var.set(True)
                cb = ttk.Checkbutton(self.scrollable_frame, text=col, variable=var, state='disabled')
            else:
                cb = ttk.Checkbutton(self.scrollable_frame, text=col, variable=var)
            
            # 計算行和列的位置
            row = i // columns_per_row
            col_index = i % columns_per_row
            cb.grid(row=row, column=col_index, padx=5, pady=2, sticky="w")
            
            # 設定列的權重
            self.scrollable_frame.grid_columnconfigure(col_index, weight=1)
    
    def on_column_selected(self, event):
        selected_column = self.column_var.get()
        if selected_column:
            self.split_column = selected_column
            # 更新勾選框狀態
            self.update_checkboxes()
            self.split_button.config(state="normal")
    
    def toggle_select_all(self):
        if not self.split_column:
            return
            
        for col, var in self.checkboxes.items():
            if col != self.split_column:  # 不影響分頁欄位
                var.set(self.select_all_var.get())
    
    def split_excel(self):
        if not self.file_path or self.df is None:
            messagebox.showerror("錯誤", "請先選擇有效的 Excel 檔案")
            return
        
        if not self.split_column:
            messagebox.showerror("錯誤", "請選擇要依據分頁的欄位")
            return
        
        # 獲取選擇的輸出欄位
        self.selected_columns = [col for col, var in self.checkboxes.items() if var.get()]
        if not self.selected_columns:
            messagebox.showerror("錯誤", "請至少選擇一個要輸出的欄位")
            return
        
        try:
            # 自動生成輸出檔案路徑
            file_dir = os.path.dirname(self.file_path)
            file_name = os.path.basename(self.file_path)
            name, ext = os.path.splitext(file_name)
            output_path = os.path.join(file_dir, f"{name}_分頁{ext}")
            
            # 檢查檔案是否已存在
            counter = 1
            while os.path.exists(output_path):
                output_path = os.path.join(file_dir, f"{name}_分頁_{counter}{ext}")
                counter += 1
            
            # 依據選擇的欄位進行分組
            grouped = self.df.groupby(self.split_column)
            
            # 建立 Excel writer
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # 將每個分組寫入不同的工作表
                for name, group in grouped:
                    # 移除工作表名稱中的非法字符
                    sheet_name = str(name)[:31]  # Excel 工作表名稱最長 31 個字符
                    # 只輸出選擇的欄位
                    group[self.selected_columns].to_excel(writer, sheet_name=sheet_name, index=False)
            
            self.status_label.config(text=f"分頁完成！檔案已儲存至：{output_path}")
            messagebox.showinfo("成功", "Excel 檔案分頁完成！")
            
        except Exception as e:
            messagebox.showerror("錯誤", f"分頁過程中發生錯誤：{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelSplitter(root)
    root.mainloop() 