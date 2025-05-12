import pandas as pd
from pathlib import Path

def format_stock_code(code):
    """将股票代码格式化为6位字符串"""
    if pd.isna(code):
        return ''
    code_str = str(int(code)) if isinstance(code, float) else str(code)
    return code_str.zfill(6)

def merge_all_years():
    """合并2015-2023年的数据"""
    # 基础目录
    base_dir = Path("..")
    
    # 创建空的DataFrame来存储所有数据
    all_data = pd.DataFrame()
    
    # 遍历2015-2023年
    for year in range(2015, 2024):
        file_path = base_dir / f"{year}.xlsx"
        if not file_path.exists():
            print(f"Warning: File not found: {file_path}")
            continue
            
        try:
            print(f"\nProcessing year: {year}")
            # 读取Excel文件
            df = pd.read_excel(file_path)
            
            # 打印列名以便调试
            print(f"Columns in {year}.xlsx: {df.columns.tolist()}")
            
            # 格式化股票代码
            df['stock_code'] = df['stock_code'].apply(format_stock_code)
            
            # 确保列名一致
            if 'page_number' in df.columns:
                df = df.rename(columns={'page_number': 'page_num'})
            
            # 选择需要的列
            required_columns = ['stock_code', '证券简称', 'page_num']
            df = df[required_columns]
            
            # 添加年份列
            df['year'] = year
            
            # 重新排列列顺序
            df = df[['year', 'stock_code', '证券简称', 'page_num']]
            
            # 添加到总数据中
            all_data = pd.concat([all_data, df], ignore_index=True)
            print(f"Successfully processed year {year}, added {len(df)} rows")
            
        except Exception as e:
            print(f"Error processing year {year}: {str(e)}")
    
    if len(all_data) > 0:
        # 按年份和股票代码排序
        all_data = all_data.sort_values(['year', 'stock_code'])
        
        # 保存合并后的数据
        output_file = base_dir / "all_years.xlsx"
        all_data.to_excel(output_file, index=False)
        print(f"\nSuccessfully saved merged data to: {output_file}")
        
        # 打印数据概览
        print("\nData Overview:")
        print(f"Total rows: {len(all_data)}")
        print(f"Years covered: {sorted(all_data['year'].unique())}")
        print(f"Unique stock codes: {len(all_data['stock_code'].unique())}")
    else:
        print("\nNo data was processed successfully")

if __name__ == "__main__":
    merge_all_years() 