import pandas as pd
from pathlib import Path
import os

def format_stock_code(code):
    """将股票代码格式化为6位字符串"""
    if pd.isna(code):
        return ''
    code_str = str(int(code)) if isinstance(code, float) else str(code)
    return code_str.zfill(6)

def merge_manual_input():
    """将手动输入的Excel文件合并到原始表格中"""
    # 基础目录
    base_dir = Path("..")
    failed_dir = base_dir / "failed_pdfs"
    
    # 遍历每个年份目录
    for year_dir in failed_dir.glob("*"):
        if not year_dir.is_dir():
            continue
            
        year = year_dir.name
        print(f"\nProcessing year: {year}")
        
        # 原始表格路径
        original_file = base_dir / f"{year}.xlsx"
        if not original_file.exists():
            print(f"Warning: Original file not found: {original_file}")
            continue
            
        # 手动输入文件路径
        manual_file = year_dir / f"{year}_manual_input.xlsx"
        if not manual_file.exists():
            print(f"Warning: Manual input file not found: {manual_file}")
            continue
            
        try:
            # 读取原始表格
            print(f"Reading original file: {original_file}")
            original_df = pd.read_excel(original_file)
            
            # 格式化原始表格中的股票代码
            original_df['stock_code'] = original_df['stock_code'].apply(format_stock_code)
            
            # 读取手动输入文件
            print(f"Reading manual input file: {manual_file}")
            manual_df = pd.read_excel(manual_file)
            
            # 格式化手动输入文件中的股票代码
            manual_df['stock_code'] = manual_df['stock_code'].apply(format_stock_code)
            
            # 重命名列名，确保一致
            if 'page_number' in manual_df.columns:
                manual_df = manual_df.rename(columns={'page_number': 'page_num'})
            
            # 确保列名一致
            if 'page_num' not in original_df.columns:
                original_df['page_num'] = ''
            
            # 更新原始表格中的页码
            for _, row in manual_df.iterrows():
                stock_code = row['stock_code']
                page_num = row['page_num']
                
                # 在原始表格中查找匹配的行
                mask = original_df['stock_code'] == stock_code
                if mask.any():
                    original_df.loc[mask, 'page_num'] = page_num
                    print(f"Updated page number for {stock_code}: {page_num}")
                else:
                    print(f"Warning: Stock code {stock_code} not found in original file")
            
            # 删除page_number列（如果存在）
            if 'page_number' in original_df.columns:
                original_df = original_df.drop(columns=['page_number'])
                print("Deleted page_number column")
            
            # 保存更新后的表格
            print(f"Saving updated file: {original_file}")
            original_df.to_excel(original_file, index=False)
            print(f"Successfully merged manual input for year {year}")
            
        except Exception as e:
            print(f"Error processing year {year}: {str(e)}")

if __name__ == "__main__":
    merge_manual_input() 