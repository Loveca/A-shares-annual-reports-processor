# A股上市公司年报处理工具

这是一个用于处理A股上市公司年报数据的Python工具集。该工具可以自动处理PDF格式的年报文件，提取关键信息，并将数据整理成结构化的Excel格式。

## 功能特点

- 📊 自动处理2015-2023年A股上市公司年报PDF文件
- 📑 自动统计年报页数
- 📈 数据整合与合并
- ❌ 失败文件处理机制
- 📝 详细的错误日志记录
- 🔄 支持手动数据输入和合并

## 系统要求

- Python 3.8+
- Windows/Linux/MacOS

## 安装说明

1. 克隆仓库：
```bash
git clone [your-repository-url]
cd [repository-name]
```

2. 安装依赖：
```bash
pip install -r pdf/Code/requirements.txt
```

## 项目结构

```
.
├── pdf/
│   ├── Code/
│   │   ├── merge_all_years.py      # 合并所有年份数据
│   │   ├── merge_manual_input.py   # 处理手动输入数据
│   │   ├── process_failed.py       # 处理失败文件
│   │   ├── count_pages.py          # 统计PDF页数
│   │   ├── requirements.txt        # 项目依赖
│   │   ├── failed_pdfs/           # 处理失败的文件
│   │   └── logs/                  # 日志文件
│   ├── 年报页数2015-2023.xlsx     # 汇总数据
│   ├── 2015.xlsx ~ 2023.xlsx      # 各年份数据
│   └── failed_pdfs/               # 处理失败的文件
```

## 使用方法

### 1. 统计年报页数

```bash
python pdf/Code/count_pages.py
```

### 2. 合并所有年份数据

```bash
python pdf/Code/merge_all_years.py
```

### 3. 处理失败的文件

```bash
python pdf/Code/process_failed.py
```

### 4. 合并手动输入数据

```bash
python pdf/Code/merge_manual_input.py
```

## 数据格式

### 输入数据
- PDF格式的年报文件
- 手动输入的Excel数据

### 输出数据
- 按年份分类的Excel文件（2015-2023）
- 汇总数据文件（年报页数2015-2023.xlsx）
- 错误记录文件（*_errors.xlsx）

## 依赖库

- PyPDF2==3.0.1：PDF文件处理
- pandas==2.0.3：数据处理
- openpyxl==3.1.2：Excel文件操作
- tqdm==4.66.1：进度条显示

## 错误处理

- 所有处理失败的文件会被移动到 `failed_pdfs` 目录
- 错误信息会被记录在对应年份的 `*_errors.xlsx` 文件中
- 详细的错误日志保存在 `logs` 目录

## 注意事项

1. 确保有足够的磁盘空间存储PDF文件和Excel文件
2. 处理大量文件时可能需要较长时间
3. 建议定期备份重要数据
4. 确保PDF文件格式正确且可读

