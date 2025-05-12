import os
import PyPDF2
import pandas as pd
import logging
import warnings
from pathlib import Path
from tqdm import tqdm
from multiprocessing import Pool, cpu_count
from datetime import datetime

# 过滤掉PyPDF2的警告信息
warnings.filterwarnings('ignore', category=UserWarning, module='PyPDF2')

# 设置日志
def setup_logging():
    """设置日志配置"""
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(log_dir, f"pdf_processing_{timestamp}.log")
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)

def count_pdf_pages(pdf_path):
    """统计PDF文件的页数"""
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            # 检查PDF是否加密
            if reader.is_encrypted:
                return -1, "PDF文件已加密"
            # 检查PDF是否损坏
            if len(reader.pages) == 0:
                return -1, "PDF文件为空或损坏"
            return len(reader.pages)
    except Exception as e:
        return -1, str(e)  # 返回错误状态和错误信息

def process_year_folder(year_folder):
    """处理单个年份文件夹"""
    # 从文件夹名称中提取年份和总文件数
    year, total_files = year_folder.split('_')[0], int(year_folder.split('_')[1].rstrip('份'))
    print(f"Started processing {year}, total files: {total_files}")
    
    # 准备数据列表和错误列表
    data = []
    errors = []
    
    # 获取文件夹中的所有PDF文件
    pdf_files = list(Path(year_folder).glob('*.pdf'))
    actual_files = len(pdf_files)
    
    if actual_files != total_files:
        print(f"Warning: Expected {total_files} files in {year_folder}, but found {actual_files}")
    
    # 使用tqdm创建进度条
    progress_bar = tqdm(pdf_files, desc=f"Processing {year}", unit="file", 
                       position=int(year) - 2015, leave=True)
    
    success_count = 0
    error_count = 0
    
    for pdf_path in progress_bar:
        # 从文件名中提取信息
        filename = pdf_path.name
        parts = filename.split('_')
        
        if len(parts) >= 3:
            stock_code = parts[0]
            company_name = parts[2]
            
            # 统计页数
            result = count_pdf_pages(pdf_path)
            if isinstance(result, tuple):  # 处理失败
                page_num = 0
                error_msg = result[1]
                errors.append({
                    'year': year,
                    'stock_code': stock_code,
                    'company_name': company_name,
                    'file_name': filename,
                    'error': error_msg
                })
                print(f"Failed to process {filename}: {error_msg}")
                error_count += 1
            else:  # 处理成功
                page_num = result
                success_count += 1
            
            # 添加到数据列表
            data.append({
                'year': year,
                'stock_code': stock_code,
                '证券简称': company_name,
                'page_num': page_num
            })
    
    # 创建DataFrame并保存为Excel
    if data:
        df = pd.DataFrame(data)
        excel_path = f"{year}.xlsx"
        df.to_excel(excel_path, index=False)
        print(f"Saved {excel_path} with {len(data)} records")
    
    # 如果有错误，保存错误日志到Excel
    if errors:
        error_df = pd.DataFrame(errors)
        error_excel_path = f"{year}_errors.xlsx"
        error_df.to_excel(error_excel_path, index=False)
        print(f"Saved error log to {error_excel_path} with {len(errors)} records")
    
    # 记录处理统计信息
    print(f"Year {year} processing completed:")
    print(f"- Total files: {actual_files}")
    print(f"- Successfully processed: {success_count}")
    print(f"- Failed to process: {error_count}")
    
    return year, len(data), len(errors)

def main():
    # 获取当前目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    # 切换到父目录（pdf目录）
    os.chdir(os.path.dirname(current_dir))
    
    # 获取所有年份文件夹并按年份排序
    year_folders = sorted([f for f in os.listdir('.') if f.endswith('份')])
    total_years = len(year_folders)
    
    print(f"Found {total_years} years to process")
    
    # 使用进程池，设置进程数为4
    with Pool(processes=4) as pool:
        # 并行处理所有年份文件夹
        results = pool.map(process_year_folder, year_folders)
        
        # 打印处理结果
        for year, processed, errors in results:
            print(f"Completed {year}: processed {processed} files, encountered {errors} errors")

if __name__ == "__main__":
    main() 