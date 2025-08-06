import pandas as pd
import re
import os
from openpyxl import load_workbook

def extract_dimensions(specification):
    """
    从字符串中提取三维尺寸信息（支持x、×、*等分隔符）
    
    参数:
        specification: 可能包含尺寸信息的字符串
        
    返回:
        成功提取时返回尺寸元组(长, 宽, 高)，否则返回None
    """
    # 处理空值情况
    if pd.isna(specification):
        return None
    
    # 转换为字符串并清理空格
    spec_str = str(specification).strip()
    
    # 使用正则表达式提取尺寸信息
    pattern = re.compile(r"(\d+)[x×*](\d+)[x×*](\d+)")
    match = pattern.search(spec_str)
    
    if match:
        return tuple(int(dim) for dim in match.groups())
    return None

def process_and_write_back(
    original_file, index_file, product_file,
    target_sheet="07",
    index_sheet="Sheet1",
    product_sheet="Sheet"
):
    """
    处理焊接产量表数据，匹配料号信息并写回原始文件
    
    参数:
        original_file: 原始焊接产量表文件路径
        index_file: 料号索引表文件路径
        product_file: 成品编码表文件路径
        target_sheet: 要更新的原始工作表名称，默认"07"
        index_sheet: 料号索引表的工作表名称，默认"Sheet1"
        product_sheet: 成品编码表的工作表名称，默认"Sheet"
        
    返回:
        处理后的DataFrame，发生错误时返回None
    """
    try:
        # 根据文件后缀选择合适的引擎
        def get_engine(file_path):
            if file_path.endswith('.xls'):
                return 'xlrd'
            elif file_path.endswith('.xlsx'):
                return 'openpyxl'
            return None

        # 读取原始文件中的目标工作表
        print(f"正在读取原始文件中的 '{target_sheet}' 工作表...")
        target_df = pd.read_excel(
            original_file,
            sheet_name=target_sheet,
            engine=get_engine(original_file),
            keep_default_na=False  # 保留原始空值状态，不自动转换
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

        # 清理列名中的空格（不修改原始数据）
        target_df.columns = target_df.columns.str.strip()
        index_df.columns = index_df.columns.str.strip()
        product_df.columns = product_df.columns.str.strip()

        # 1. 创建索引字段（货号 + 标记）
        print("创建索引字段...")
        required_cols = ['货号', '标记']
        if not all(col in target_df.columns for col in required_cols):
            missing = [col for col in required_cols if col not in target_df.columns]
            raise ValueError(f"目标工作表缺少必要列: {missing}")
            
        # 处理空值并转换为字符串
        target_df['货号'] = target_df['货号'].fillna('').astype(str).str.strip()
        target_df['标记'] = target_df['标记'].fillna('').astype(str).str.strip()
        
        # 新增或更新索引字段列
        if '索引字段' in target_df.columns:
            target_df['索引字段'] = target_df['货号'] + target_df['标记']
        else:
            target_df['索引字段'] = target_df['货号'] + target_df['标记']

        # 2. 匹配料号索引表获取料号
        print("匹配料号索引表...")
        if '索引字段' not in index_df.columns or '料号' not in index_df.columns:
            raise ValueError("料号索引表缺少'索引字段'或'料号'列")
            
        # 创建索引字段到料号的映射
        index_map = pd.Series(index_df['料号'].values, index=index_df['索引字段']).to_dict()
        
        # 新增或更新料号列（确保在索引字段右侧）
        if '料号' in target_df.columns:
            # 只更新空值，保留已有数据
            mask = target_df['料号'].isna() | (target_df['料号'] == '')
            target_df.loc[mask, '料号'] = target_df.loc[mask, '索引字段'].map(index_map)
        else:
            # 插入到索引字段右侧
            index_pos = target_df.columns.get_loc('索引字段')
            target_df.insert(index_pos + 1, '料号', target_df['索引字段'].map(index_map))

        # 3. 匹配成品编码表补充未匹配的料号
        print("匹配成品编码表...")
        product_required = ['规格型号', '净重', '产品编号']
        if not all(col in product_df.columns for col in product_required):
            missing = [col for col in product_required if col not in product_df.columns]
            raise ValueError(f"成品编码表缺少必要列: {missing}")
            
        # 预处理成品编码表的尺寸信息
        product_df['提取的尺寸'] = product_df['规格型号'].apply(extract_dimensions)
        
        # 处理仍为空的料号
        mask = (target_df['料号'].isna()) | (target_df['料号'] == '')
        for idx, row in target_df[mask].iterrows():
            # 从货号中提取尺寸
            row_dim = extract_dimensions(row['货号'])
            if row_dim:
                # 匹配相同尺寸的记录
                dim_matches = product_df[product_df['提取的尺寸'] == row_dim]
                if not dim_matches.empty:
                    # 进一步匹配单重
                    weight_matches = dim_matches[dim_matches['净重'] == row['单重']]
                    if not weight_matches.empty:
                        target_df.at[idx, '料号'] = weight_matches.iloc[0]['产品编号']

        # 清理可能的临时列
        if '提取的尺寸' in target_df.columns:
            target_df = target_df.drop(columns=['提取的尺寸'])

        # 将结果写回原始文件
        print(f"正在写回原始文件的 '{target_sheet}' 工作表...")
        file_ext = original_file.split('.')[-1].lower()
        
        if file_ext == 'xlsx':
            # 处理xlsx格式：精确替换目标工作表
            with pd.ExcelWriter(
                original_file,
                engine='openpyxl',
                mode='a',
                if_sheet_exists='replace'
            ) as writer:
                target_df.to_excel(writer, sheet_name=target_sheet, index=False)
        else:
            # 处理xls格式：读取所有工作表后整体写回
            all_sheets = pd.read_excel(original_file, sheet_name=None, engine='xlrd')
            all_sheets[target_sheet] = target_df  # 替换目标工作表
            
            with pd.ExcelWriter(original_file, engine='xlwt') as writer:
                for sheet_name, df in all_sheets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"操作完成！已更新: {original_file} 的 '{target_sheet}' 工作表")
        return target_df

    except Exception as e:
        print(f"处理错误: {str(e)}")
        return None


def get_full_path(filename):
    """获取文件的完整路径（程序所在目录+文件名）"""
    # 获取当前脚本所在目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    # 拼接完整路径
    full_path = os.path.join(current_dir, filename)
    # 处理路径中的特殊字符
    return full_path.replace("\\", "/")  # 统一使用斜杠，避免Windows路径问题


def validate_file_exists(file_path, file_desc):
    """验证文件是否存在"""
    if not os.path.exists(file_path):
        print(f"错误：{file_desc} '{file_path}' 不存在！")
        return False
    if not os.path.isfile(file_path):
        print(f"错误：{file_desc} '{file_path}' 不是有效的文件！")
        return False
    return True


if __name__ == "__main__":
    # 获取程序所在目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"程序当前目录：{current_dir}")
    print("请确保所有Excel文件与本程序放在同一文件夹中")
    
    # 获取目标工作表名称
    target_sheet = input("请输入目标工作表名称（默认07）：") or "07"
    
    # 获取文件名并生成完整路径
    while True:
        original_name = input("请输入焊接产量表的文件名（例如：焊接产量表.xlsx）：")
        original_file = get_full_path(original_name)
        if validate_file_exists(original_file, "焊接产量表"):
            break
    
     
    # 获取目标工作表名称
    target_sheet = input("请输入目标工作表名称（默认07）：") or "07"

    while True:
        index_name = input("请输入料号索引表的文件名（例如：料号索引.xlsx）：")
        index_file = get_full_path(index_name)
        if validate_file_exists(index_file, "料号索引表"):
            break

    while True:
        product_name = input("请输入成品编码表的文件名（例如：成品编码.xls）：")
        product_file = get_full_path(product_name)
        if validate_file_exists(product_file, "成品编码表"):
            break
    
    # 显示最终使用的文件路径
    print("\n使用的文件路径：")
    print(f"焊接产量表：{original_file}")
    print(f"料号索引表：{index_file}")
    print(f"成品编码表：{product_file}")
    
    # 执行处理
    process_and_write_back(
        original_file=original_file,
        index_file=index_file,
        product_file=product_file,
        target_sheet=target_sheet
    )
    
    # 防止程序运行完毕后立即关闭（在部分环境中有用）
    input("\n处理完成，按任意键退出...")