import os
import shutil
import pandas as pd
from pathlib import Path
import glob

def format_stock_code(code):
    """将股票代码格式化为6位字符串"""
    code_str = str(code)
    if len(code_str) < 6:
        code_str = code_str.zfill(6)
    return code_str

def process_failed_files():
    """处理失败的PDF文件"""
    # 创建failed_pdfs目录
    failed_dir = Path("../failed_pdfs")
    failed_dir.mkdir(exist_ok=True)
    print(f"Created directory: {failed_dir}")
    
    # 读取所有错误日志文件
    error_files = list(Path('..').glob('*_errors.xlsx'))
    print(f"Found {len(error_files)} error log files")
    
    for error_file in error_files:
        try:
            print(f"\nProcessing error file: {error_file}")
            # 读取错误日志
            df = pd.read_excel(error_file)
            print(f"Found {len(df)} failed files in {error_file}")
            print(f"Columns in {error_file.name}: {list(df.columns)}")
            
            # 获取年份
            year = error_file.stem.split('_')[0]
            year_dir = failed_dir / year
            year_dir.mkdir(exist_ok=True)
            
            # 复制文件
            success_count = 0
            for _, row in df.iterrows():
                try:
                    # 格式化股票代码
                    stock_code = format_stock_code(row['stock_code'])
                    
                    # 查找匹配的目录
                    year_pattern = f"../{year}_*份"
                    matching_dirs = glob.glob(year_pattern)
                    if not matching_dirs:
                        print(f"Warning: No matching directory found for pattern: {year_pattern}")
                        continue
                    source_dir = matching_dirs[0]

                    # 构建源文件路径
                    source_file = os.path.join(source_dir, row['file_name'])
                    if not os.path.exists(source_file):
                        print(f"Warning: File not found: {source_file}")
                        continue

                    # 构建目标文件路径
                    target_file = year_dir / f"{stock_code}_{row['company_name']}.pdf"
                    shutil.copy2(source_file, target_file)
                    print(f"Copied {row['file_name']} to {target_file}")
                    success_count += 1

                except Exception as e:
                    print(f"Error processing file for {row['stock_code']}: {str(e)}")
            
            print(f"Successfully copied {success_count} out of {len(df)} files for year {year}")
            
            # 创建手动输入文件
            manual_input = pd.DataFrame({
                'year': [year] * len(df),
                'stock_code': [format_stock_code(code) for code in df['stock_code']],
                'company_name': df['company_name'],
                'page_number': [''] * len(df)  # 空列用于手动输入
            })
            manual_input_file = year_dir / f"{year}_manual_input.xlsx"
            manual_input.to_excel(manual_input_file, index=False)
            print(f"Created manual input file: {manual_input_file}")
        
        except Exception as e:
            print(f"Error processing error file {error_file}: {str(e)}")

if __name__ == "__main__":
    process_failed_files() 