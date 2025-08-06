import pandas as pd
import re
import os
import sys
from openpyxl import load_workbook

def extract_dimensions(specification):
    """从字符串中提取三维尺寸信息(支持x、X、*等分隔符）"""
    if pd.isna(specification):
        return None

    spec_str = str(specification).strip()
    pattern = re.compile(r"(\d+)[x×X*](\d+)[x×X*](\d+)")
    match = pattern.search(spec_str)
    if match:
        return tuple(int(dim) for dim in match.groups())
    return None

def normalize_mark(mark):
    """标准化标记：将“无标记”和“无标”统一为“无标”，其他标记去空格"""
    mark_str = str(mark).lower().strip()
    if mark_str in ["无标记","标记"]:
        return "无标"
    else:
        return mark_str

def process_and_write_back(
    original_file, index_file, product_file,
    target_sheet="07",
    index_sheet="Sheet1",
    product_sheet="Sheet"
):
    """处理焊接产量表数据，匹配料号信息并写回原始文件"""
    try:
        # 跨平台兼容的引擎选择
        def get_engine(file_path):
            ext = os.path.splitext(file_path)[1].lower()
            if ext == '.xls':
                return 'xlrd'
            elif ext == '.xlsx':
                return 'openpyxl'
            return None

        # 读取原始文件中的目标工作表
        print(f"正在读取原始文件中的 '{target_sheet}' 工作表...")
        target_df = pd.read_excel(
            original_file,
            sheet_name=target_sheet,
            engine=get_engine(original_file),
            keep_default_na=False
        )

        # 读取辅助表格数据
        index_df = pd.read_excel(
            index_file,
            sheet_name=index_sheet,
            engine=get_engine(index_file)
        )

        product_df = pd.read_excel(
            product_file,
            sheet_name=product_sheet,
            engine=get_engine(product_file)
        )

        # 清理列名中的空格
        target_df.columns = target_df.columns.str.strip()
        index_df.columns = index_df.columns.str.strip()
        product_df.columns = product_df.columns.str.strip()

        # 创建索引字段（货号 + 标记）
        print("创建索引字段...")
        required_cols = ['货号', '标记']
        if not all(col in target_df.columns for col in required_cols):
            missing = [col for col in required_cols if col not in target_df.columns]
            raise ValueError(f"目标工作表缺少必要列: {missing}")

        # 处理空值并转换为字符串
        target_df['货号'] = target_df['货号'].fillna('').astype(str)
        target_df['标记'] = target_df['标记'].fillna('').astype(str)
        target_df['索引字段'] = target_df['货号'] + target_df['标记']

        # 匹配料号索引表获取料号
        print("匹配料号索引表...")
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

        # 匹配成品编码表补充未匹配的料号
        print("匹配成品编码表...")
        product_required = ['规格型号','标记', '产品编号']
        if not all(col in product_df.columns for col in product_required):
            missing = [col for col in product_required if col not in product_df.columns]
            raise ValueError(f"成品编码表缺少必要列: {missing}")

        product_df['提取的尺寸'] = product_df['规格型号'].apply(extract_dimensions)
        product_df['标准化标记'] = product_df['标记'].apply(normalize_mark)
        target_df['标准化标记'] = target_df['标记'].apply(normalize_mark)

        mask = (target_df['料号'].isna()) | (target_df['料号'] == '')

        count = 0
        count1 = 0
        for idx, row in target_df[mask].iterrows():
            count1 += 1
            row_dim = extract_dimensions(row['货号'])
            if row_dim:
                matches = product_df[(product_df['提取的尺寸'] == row_dim) & 
                            (product_df['标准化标记'] == row['标准化标记'])]
                if not matches.empty:
                    target_df.at[idx, '料号'] = matches.iloc[0]['产品编号']
                    count += 1

        print(f"精确匹配后，还有{count1}没有匹配！")
        print(f"尺寸和标记匹配了{count}条数据")
        print(f"还剩{count1-count}个索引字段找不到料号，为空值")

        # 清理临时列
        if '提取的尺寸' in target_df.columns:
            target_df = target_df.drop(columns=['提取的尺寸'])
        
        temp_cols = ['标准化标记', '提取的尺寸']
        target_df = target_df.drop(columns=[col for col in temp_cols if col in target_df.columns])
        product_df = product_df.drop(columns=[col for col in temp_cols if col in product_df.columns])
        # -------------------- 关键修改：处理料号格式，避免科学计数法 --------------------
        # 1. 将料号列强制转换为字符串（确保空值显示为空白，而非NaN）
        target_df['料号'] = target_df['料号'].fillna('').astype(str).str.strip()
        
        # 写回原始文件（跨平台兼容）
        print(f"正在写回原始文件的 '{target_sheet}' 工作表...")
        ext = os.path.splitext(original_file)[1].lower()

        if ext == '.xlsx':
            with pd.ExcelWriter(
                original_file,
                engine='openpyxl',
                mode='a',
                if_sheet_exists='replace'
            ) as writer:
                target_df.to_excel(writer, sheet_name=target_sheet, index=False)
        else:
            all_sheets = pd.read_excel(original_file, sheet_name=None, engine='xlrd')
            all_sheets[target_sheet] = target_df

            with pd.ExcelWriter(original_file, engine='xlwt') as writer:
                for sheet_name, df in all_sheets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"操作完成！已更新: {original_file} 的 '{target_sheet}' 工作表")
        return target_df

    except Exception as e:
        print(f"处理错误: {str(e)}")
        # 出错时暂停，避免闪退
        input("\n按任意键继续...")
        return None


def find_excel_file(base_name, current_dir):
    """
    查找当前目录下是否存在指定基础名称的Excel文件（.xls或.xlsx）
    返回找到的完整路径，未找到则返回None
    """
    # 打印调试信息
    print(f"查找文件: {base_name} 在目录: {current_dir}")
    
    # 可能的文件后缀
    extensions = ['.xlsx', '.xls']
    
    # 先检查用户是否已输入后缀（兼容旧方式）
    if any(base_name.endswith(ext) for ext in extensions):
        full_path = os.path.join(current_dir, base_name)
        if os.path.exists(full_path) and os.path.isfile(full_path):
            print(f"找到文件: {full_path}")
            return full_path
        print(f"未找到文件: {full_path}")
        return None
    
    # 自动尝试添加后缀查找
    for ext in extensions:
        full_path = os.path.join(current_dir, f"{base_name}{ext}")
        if os.path.exists(full_path) and os.path.isfile(full_path):
            print(f"找到文件: {full_path}")
            return full_path
        print(f"未找到文件: {full_path}")
    
    return None


def get_sheet_name(file_path, prompt, default_sheet, engine):
    """
    获取并验证用户输入的工作表名称
    支持输入quit退出程序，检测工作表是否存在
    """
    while True:
        sheet_name = input(prompt).strip() or default_sheet
        
        # 检查是否退出程序
        if sheet_name.lower() == 'quit':
            print("用户选择退出程序")
            sys.exit(0)
            
        # 检查工作表是否存在
        try:
            # 只获取工作表名称列表，不读取全部数据
            xl = pd.ExcelFile(file_path, engine=engine)
            if sheet_name in xl.sheet_names:
                return sheet_name
            else:
                print(f"错误：文件中不存在 '{sheet_name}' 工作表")
                print(f"该文件包含的工作表有：{', '.join(xl.sheet_names)}")
                print("请重新输入，或输入quit退出程序\n")
        except Exception as e:
            print(f"检查工作表时出错: {str(e)}")
            print("请重新输入，或输入quit退出程序\n")
            # 出错时暂停，避免闪退
            input("按任意键继续...")


def clear_screen():
    """跨平台清屏函数"""
    try:
        os.system('cls' if os.name == 'nt' else 'clear')
    except:
        pass  # 清屏失败不影响主程序


if __name__ == "__main__":
    # 全局异常捕获，防止闪退
    try:
        # 跨平台编码设置（解决中文显示问题）
        if os.name == 'nt':
            import io
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

        # 更可靠的路径获取方式
        if getattr(sys, 'frozen', False):
            # 打包后的环境
            current_dir = os.path.dirname(sys.executable)
        else:
            # 开发环境
            current_dir = os.path.dirname(os.path.abspath(__file__))

        clear_screen()
        print(f"程序当前目录：{current_dir}")
        print("请确保所有Excel文件与本程序放在同一文件夹中")
        print("输入过程中随时可以输入'quit'退出程序")
        print("输入文件名时无需添加后缀（程序会自动识别.xls和.xlsx格式）\n")
        
        # 获取焊接产量表（自动检测后缀）
        while True:
            original_name = input("请输入焊接产量表的文件名（无需后缀）：").strip()
            if original_name.lower() == 'quit':
                print("用户选择退出程序")
                sys.exit(0)
                
            original_file = find_excel_file(original_name, current_dir)
            if original_file:
                break
            print(f"错误：未找到文件 '{original_name}.xls' 或 '{original_name}.xlsx'")
            print("请重新输入，或输入quit退出程序\n")
            input("按任意键继续...")

        # 获取焊接产量表的工作表名称
        ext = os.path.splitext(original_file)[1].lower()
        engine = 'openpyxl' if ext == '.xlsx' else 'xlrd'
        target_sheet = get_sheet_name(
            original_file,
            f"请输入焊接产量表的工作表名称（默认07）：",
            "07",
            engine
        )
        
        # 获取料号索引表（自动检测后缀）
        while True:
            index_name = input("请输入料号索引表的文件名（无需后缀）：").strip()
            if index_name.lower() == 'quit':
                print("用户选择退出程序")
                sys.exit(0)
                
            index_file = find_excel_file(index_name, current_dir)
            if index_file:
                break
            print(f"错误：未找到文件 '{index_name}.xls' 或 '{index_name}.xlsx'")
            print("请重新输入，或输入quit退出程序\n")
            input("按任意键继续...")

        # 获取料号索引表的工作表名称
        ext = os.path.splitext(index_file)[1].lower()
        engine = 'openpyxl' if ext == '.xlsx' else 'xlrd'
        index_sheet = get_sheet_name(
            index_file,
            f"请输入料号索引表的工作表名称（默认Sheet1）：",
            "Sheet1",
            engine
        )

        # 获取成品编码表（自动检测后缀）
        while True:
            product_name = input("请输入成品编码表的文件名（无需后缀）：").strip()
            if product_name.lower() == 'quit':
                print("用户选择退出程序")
                sys.exit(0)
                
            product_file = find_excel_file(product_name, current_dir)
            if product_file:
                break
            print(f"错误：未找到文件 '{product_name}.xls' 或 '{product_name}.xlsx'")
            print("请重新输入，或输入quit退出程序\n")
            input("按任意键继续...")

        # 获取成品编码表的工作表名称
        ext = os.path.splitext(product_file)[1].lower()
        engine = 'openpyxl' if ext == '.xlsx' else 'xlrd'
        product_sheet = get_sheet_name(
            product_file,
            f"请输入成品编码表的工作表名称（默认Sheet）：",
            "Sheet",
            engine
        )
        
        # 显示最终使用的文件和工作表信息
        print("\n使用的配置信息：")
        print(f"焊接产量表：{original_file}，工作表：{target_sheet}")
        print(f"料号索引表：{index_file}，工作表：{index_sheet}")
        print(f"成品编码表：{product_file}，工作表：{product_sheet}\n")
        
        # 执行处理
        process_and_write_back(
            original_file=original_file,
            index_file=index_file,
            product_file=product_file,
            target_sheet=target_sheet,
            index_sheet=index_sheet,
            product_sheet=product_sheet
        )
        
    except Exception as e:
        # 捕获所有未处理的异常，防止闪退
        print(f"\n程序发生错误：{str(e)}")
    finally:
        # 无论是否出错，都暂停等待用户确认
        input("\n处理完成，按任意键退出...")
