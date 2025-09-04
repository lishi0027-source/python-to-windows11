import pandas as pd
import re
import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk
import threading

class ExcelProcessorApp:
    def __init__(self, root):
        """初始化应用程序"""
        self.root = root
        self.setup_ui_basics()  # 设置UI基础属性
        self.init_thread_vars()  # 初始化线程相关变量
        self.current_dir = self.get_current_dir()  # 获取当前目录
        self.create_widgets()    # 创建界面组件

    def setup_ui_basics(self):
        """设置UI基础属性：标题、大小、字体和关闭协议"""
        self.root.title("Excel数据处理工具")
        self.root.geometry("700x400")
        self.root.resizable(False, False)
        
        # 设置中文字体
        default_font = ('SimHei', 10)
        self.root.option_add("*Font", default_font)
        
        # 设置窗口关闭事件处理
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def init_thread_vars(self):
        """初始化线程管理变量"""
        self.processing_thread = None  # 处理线程
        self.running = False           # 处理状态标记

    def get_current_dir(self):
        """获取当前目录，兼容打包和开发环境"""
        try:
            if getattr(sys, 'frozen', False):
                return os.path.dirname(sys.executable)  # 打包后环境
            else:
                return os.path.dirname(os.path.abspath(__file__))  # 开发环境
        except:
            return os.getcwd()  # 兼容交互式环境

    def create_widgets(self):
        """创建所有界面组件"""
        # 程序信息区域
        self.create_info_labels()
        
        # 输入区域 - 三个表格的输入框
        self.create_input_fields()
        
        # 按钮区域
        self.create_buttons()
        
        # 状态区域
        self.create_status_label()

    def create_info_labels(self):
        """创建程序信息标签"""
        # 当前目录显示
        dir_label = ttk.Label(
            self.root, 
            text=f"程序当前目录：{self.current_dir}", 
            wraplength=650
        )
        dir_label.place(x=30, y=10)
        
        # 输入提示
        note_label = ttk.Label(
            self.root, 
            text="请输入文件名（无需后缀）和对应的工作表名称"
        )
        note_label.place(x=30, y=40)

    def create_input_fields(self):
        """创建所有输入字段"""
        # 焊接产量表输入
        self.original_name = self.create_table_input(
            label_text="焊接产量表：",
            sheet_label="工作表名（默认07）：",
            default_sheet="07",
            y_pos=80
        )
        
        # 料号索引表输入
        self.index_name = self.create_table_input(
            label_text="料号索引表：",
            sheet_label="工作表名（默认Sheet1）：",
            default_sheet="Sheet1",
            y_pos=140
        )
        
        # 成品编码表输入
        self.product_name = self.create_table_input(
            label_text="成品编码表：",
            sheet_label="工作表名（默认Sheet）：",
            default_sheet="Sheet",
            y_pos=200
        )

    def create_table_input(self, label_text, sheet_label, default_sheet, y_pos):
        """
        创建单个表格的输入区域
        参数:
            label_text: 主标签文本
            sheet_label: 工作表标签文本
            default_sheet: 工作表默认值
            y_pos: Y轴位置
        返回:
            文件名输入框的StringVar对象
        """
        # 主标签
        ttk.Label(self.root, text=label_text).place(x=30, y=y_pos)
        
        # 文件名输入框
        name_var = tk.StringVar()
        ttk.Entry(self.root, textvariable=name_var, width=30).place(x=120, y=y_pos)
        
        # 工作表标签
        ttk.Label(self.root, text=sheet_label).place(x=350, y=y_pos)
        
        # 工作表输入框
        sheet_var = tk.StringVar(value=default_sheet)
        setattr(self, f"{label_text.split('：')[0]}_sheet", sheet_var)  # 动态存储工作表变量
        ttk.Entry(self.root, textvariable=sheet_var, width=15).place(x=480, y=y_pos)
        
        return name_var

    def create_buttons(self):
        """创建功能按钮"""
        # 开始处理按钮
        self.start_btn = ttk.Button(
            self.root, 
            text="开始处理", 
            command=self.start_processing
        )
        self.start_btn.place(x=200, y=280)
        
        # 退出按钮
        self.exit_btn = ttk.Button(
            self.root, 
            text="退出", 
            command=self.on_close
        )
        self.exit_btn.place(x=400, y=280)

    def create_status_label(self):
        """创建状态显示标签"""
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(
            self.root, 
            textvariable=self.status_var
        ).place(x=30, y=340)

    # ------------------------------
    # 文件处理相关方法
    # ------------------------------
    def find_excel_file(self, base_name):
        """查找Excel文件，支持.xls和.xlsx格式"""
        if not base_name:
            return None
            
        extensions = ['.xlsx', '.xls']
        # 检查是否已包含后缀
        if any(base_name.endswith(ext) for ext in extensions):
            full_path = os.path.join(self.current_dir, base_name)
            return full_path if os.path.isfile(full_path) else None
        
        # 尝试添加后缀查找
        for ext in extensions:
            full_path = os.path.join(self.current_dir, f"{base_name}{ext}")
            if os.path.isfile(full_path):
                return full_path
                
        return None

    def get_engine(self, file_path):
        """根据文件后缀获取合适的Excel引擎"""
        ext = os.path.splitext(file_path)[1].lower()
        return 'xlrd' if ext == '.xls' else 'openpyxl'

    def extract_dimensions(self, specification):
        """从字符串中提取三维尺寸信息(支持x、X、*等分隔符)"""
        if pd.isna(specification):
            return None

        spec_str = str(specification).strip()
        pattern = re.compile(r"(\d+)[x×X*](\d+)[x×X*](\d+)")
        match = pattern.search(spec_str)
        return tuple(int(dim) for dim in match.groups()) if match else None

    def normalize_mark(self, mark):
        """标准化标记：统一"无标记"和"标记"为"无标"""
        if pd.isna(mark):
            return ""
        mark_str = str(mark).lower().strip()
        return "无标" if mark_str in ["无标记", "标记"] else mark_str

    # ------------------------------
    # 处理流程控制
    # ------------------------------
    def start_processing(self):
        """启动处理线程"""
        if self.running:
            return
            
        self.running = True
        self.start_btn.config(state="disabled")
        self.update_status("开始处理...")
        
        # 在新线程中执行处理，避免界面冻结
        self.processing_thread = threading.Thread(target=self.process_files)
        self.processing_thread.daemon = True  # 设为守护线程
        self.processing_thread.start()

    def process_files(self):
        """处理文件的实际逻辑"""
        try:
            # 获取输入值
            original_name = self.original_name.get().strip()
            index_name = self.index_name.get().strip()
            product_name = self.product_name.get().strip()
            
            # 获取工作表名称（使用动态存储的变量）
            target_sheet = self.焊接产量表_sheet.get().strip() or "07"
            index_sheet = self.料号索引表_sheet.get().strip() or "Sheet1"
            product_sheet = self.成品编码表_sheet.get().strip() or "Sheet"
            
            # 验证输入
            empty_fields = []
            if not original_name: empty_fields.append("焊接产量表文件名")
            if not index_name: empty_fields.append("料号索引表文件名")
            if not product_name: empty_fields.append("成品编码表文件名")
            
            if empty_fields:
                self.show_error("输入错误", f"以下字段不能为空：\n{', '.join(empty_fields)}")
                return
            
            # 查找文件
            original_file = self.find_excel_file(original_name)
            index_file = self.find_excel_file(index_name)
            product_file = self.find_excel_file(product_name)
            
            # 验证文件存在性
            if not original_file:
                self.show_error("文件未找到", f"未找到焊接产量表文件：{original_name}")
                return
            if not index_file:
                self.show_error("文件未找到", f"未找到料号索引表文件：{index_name}")
                return
            if not product_file:
                self.show_error("文件未找到", f"未找到成品编码表文件：{product_name}")
                return
            
            # 显示配置信息
            self.show_info("配置确认", (
                f"使用的配置信息：\n"
                f"焊接产量表：{os.path.basename(original_file)}，工作表：{target_sheet}\n"
                f"料号索引表：{os.path.basename(index_file)}，工作表：{index_sheet}\n"
                f"成品编码表：{os.path.basename(product_file)}，工作表：{product_sheet}"
            ))
            
            # 开始处理数据
            self.update_status("正在读取文件...")
            target_df = pd.read_excel(
                original_file, sheet_name=target_sheet,
                engine=self.get_engine(original_file), keep_default_na=False
            )
            
            index_df = pd.read_excel(
                index_file, sheet_name=index_sheet,
                engine=self.get_engine(index_file)
            )
            
            product_df = pd.read_excel(
                product_file, sheet_name=product_sheet,
                engine=self.get_engine(product_file)
            )
            
            # 清理列名
            self.update_status("正在处理数据...")
            
            target_df.columns = target_df.columns.str.replace(r'\s', '', regex=True)
            index_df.columns = index_df.columns.str.replace(r'\s', '',regex=True)
            product_df.columns = product_df.columns.str.replace(r'\s', '', regex=True)
            
            # 验证必要列
            required_cols = ['货号', '标记']
            if not all(col in target_df.columns for col in required_cols):
                missing = [col for col in required_cols if col not in target_df.columns]
                raise ValueError(f"目标工作表缺少必要列: {missing}")
            
            # 处理空值并创建索引字段
            target_df['货号'] = target_df['货号'].fillna('').astype(str)
            target_df['标记'] = target_df['标记'].fillna('').astype(str)
            target_df['索引字段'] = target_df['货号'] + target_df['标记']
            
            # 匹配料号索引表
            if '索引字段' not in index_df.columns or '料号' not in index_df.columns:
                raise ValueError("料号索引表缺少'索引字段'或'料号'列")
            
            index_map = pd.Series(index_df['料号'].values, index=index_df['索引字段']).to_dict()
        
            # 新增或更新料号列
            if '料号' in target_df.columns:
                mask = target_df['料号'].isna() | (target_df['料号'] == '')
                target_df.loc[mask, '料号'] = target_df.loc[mask, '索引字段'].map(index_map)
            else:
                index_pos = target_df.columns.get_loc('索引字段')
                target_df.insert(index_pos + 1, '料号', target_df['索引字段'].map(index_map))
         
            # 匹配成品编码表
            product_required = ['规格型号', '标记', '产品编号']
            if not all(col in product_df.columns for col in product_required):
                missing = [col for col in product_required if col not in product_df.columns]
                raise ValueError(f"成品编码表缺少必要列: {missing}")
            
            product_df['提取的尺寸'] = product_df['规格型号'].apply(self.extract_dimensions)
            product_df['标准化标记'] = product_df['标记'].apply(self.normalize_mark)
            target_df['标准化标记'] = target_df['标记'].apply(self.normalize_mark)
            
            # 统计匹配结果
            mask = (target_df['料号'].isna()) | (target_df['料号'] == '')
            count, count1 = 0, 0
            
            for idx, row in target_df[mask].iterrows():
                count1 += 1
                row_dim = self.extract_dimensions(row['货号'])
                if row_dim:
                    matches = product_df[
                        (product_df['提取的尺寸'] == row_dim) & 
                        (product_df['标准化标记'] == row['标准化标记'])
                    ]
                    if not matches.empty:
                        target_df.at[idx, '料号'] = matches.iloc[0]['产品编号']
                        count += 1
            
            # 清理临时列
           
            temp_cols = ['标准化标记', '提取的尺寸']
            target_df = target_df.drop(columns=[col for col in temp_cols if col in target_df.columns])
            product_df = product_df.drop(columns=[col for col in temp_cols if col in product_df.columns])
            
            # 处理料号格式
            target_df['料号'] = target_df['料号'].fillna('').astype(str).str.strip()
            
            # 写回文件
            self.update_status("正在写回文件...")
            ext = os.path.splitext(original_file)[1].lower()
            
            if ext == '.xlsx':
                with pd.ExcelWriter(
                    original_file, engine='openpyxl',
                    mode='a', if_sheet_exists='replace'
                ) as writer:
                    target_df.to_excel(writer, sheet_name=target_sheet, index=False)
            else:
                all_sheets = pd.read_excel(original_file, sheet_name=None, engine='xlrd')
                all_sheets[target_sheet] = target_df
                
                with pd.ExcelWriter(original_file, engine='xlwt') as writer:
                    for sheet_name, df in all_sheets.items():
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # 处理完成
            self.show_info("成功", (
                f"处理完成！\n已更新 {os.path.basename(original_file)} 的 '{target_sheet}' 工作表\n"
                f"精确匹配后未匹配: {count1}\n"
                f"尺寸和标记匹配: {count}\n"
                f"剩余未匹配: {count1-count}"
            ))
            self.update_status("处理完成")
            
        except Exception as e:
            if not self.running:    #如果是因为退出而中断
                return
            error_msg = f"处理错误: {str(e)}"
            print(error_msg)
            self.show_error("错误", error_msg)
            self.update_status("处理出错")
        finally:
            self.running = False
            self.root.after(0, lambda: self.start_btn.config(state="normal"))

    # ------------------------------
    # 界面交互辅助方法
    # ------------------------------
    def update_status(self, text):
        """更新状态标签（线程安全）"""
        self.status_var.set(text)
        self.root.update_idletasks()

    def show_info(self, title, message):
        """显示信息对话框（线程安全）"""
        self.root.after(0, lambda: messagebox.showinfo(title, message))
        self.root.after(0, lambda: self.update_status("就绪"))

    def show_error(self, title, message):
        """显示错误对话框（线程安全）"""
        self.root.after(0, lambda: messagebox.showerror(title, message))
        self.root.after(0, lambda: self.update_status("就绪"))

    def cleanup(self):
        """执行清理操作"""

        # 关闭所有可能的文件句柄
        import gc
        gc.collect()    # 强制垃圾回收

    def on_close(self):
        """处理窗口关闭事件"""
        if self.running:
            if not messagebox.askyesno("确认", "正在处理数据，确定要退出吗？"):
                return
        
        # 确保线程终止
        self.running = False
        if self.processing_thread and self.processing_thread.is_alive():
            self.processing_thread.join(timeout=1)
            if self.processing_thread.is_alive():
                self.root.destroy()
                sys.exit(0)
        
        self.cleanup()
        self.root.destroy()

if __name__ == "__main__":
    # 跨平台编码设置
    if os.name == 'nt':
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    
    # 启动应用
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()
    
