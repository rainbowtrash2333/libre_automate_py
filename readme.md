# LibreOffice Python Automation Toolkit

[中文](./readme_zh.md)

**Background**  
As a statistics professional handling voluminous tabular data and reports, I developed this tool to automate LibreOffice operations when restricted to Linux environments without Office VBA access.

## Key Features
- 📊 **Excel Automation**: Data read/write, formatting, cell merging, formula calculation
- 📝 **Word Automation**: Document content replacement, template population
- 🔄 **Batch Processing**: Multi-source file batch operations
- 🎯 **Specialized Reporting**: Optimized for banking risk monitoring reports
- 🧵 **Thread-Safe**: Singleton-managed LibreOffice connections

## Core Modules
### Workbook (`workbook.py`)
Core Excel operations class providing:
- Workbook create/open/save
- Data range read/write
- Cell merging & formatting
- pandas DataFrame integration
- Auto-sum functionality

### Document (`word.py`)
Word document handler supporting:
- Document create/open/save
- Batch text replacement
- Content search & location

### OfficeLoader (`officeLoader.py`)
LibreOffice connection manager (Singleton):
- Thread-safe instance management
- Automatic connection & resource cleanup
- Context manager support

## Installation
Managed via [poetry](https://python-poetry.org/). New users see [documentation](https://python-poetry.org/docs/basic-usage/).
```sh
# Install poetry
pip install poetry

# Clone repository
git clone https://github.com/rainbowtrash2333/libreoffice_py.git
cd libreoffice_py
```

### Windows Setup
```sh
# Create virtual environment
py -3.9 -m venv .\.venv
.\.venv\Scripts\activate

# Install dependencies
poetry install

# Switch to UNO environment
oooenv env -t && oooenv env -u

# Re-switch after poetry commands
oooenv env -t
poetry <command>
oooenv env -t
```

### Linux Setup
```sh
python3 -m venv .venv
source .venv/bin/activate
poetry install
```

### Environment Verification
```sh
python -c "import uno; print('Environment validated')"
```

## Usage Examples
### Basic Excel Operations
```python
from workbook import Workbook

with Workbook(filepath="data.xlsx", visible=False) as wb:
    data = wb.get_used_value(0, range_name='A1:C10')  # Read data
```

### Word Template Processing
```python
from word import Word

with Word(filepath="template.doc") as doc:
    doc.replace_words(
        labels=["$(Company)", "$(Date)"], 
        values=["ABC Group", "2025-07-30"]
    )
    doc.save()  # Auto-closes document
```

### Batch Report Generation
```sh
python gen_xls.py  # Execute main report generator
```

## Utility Functions (`myutil.py`)
- **Data Conversion**: `array2df()` - Tuple-to-DataFrame
- **Value Processing**: `process_value_to_str()` - Smart value formatting
- **Coordinate Conversion**: `convert_cell_name_to_list()` - Cell address parsing
- **File Validation**: `check_files_exist()` - Batch file existence check

## Path Configuration
Modify these paths according to your environment:
- `data_path`: Source data directory
- `template_path`: Template directory  
- `result_path`: Output directory

## Important Notes
1. Requires LibreOffice 7.0+ pre-installed
2. Use absolute paths for reliable file access
3. Actively close documents to release resources
4. Utilize OfficeLoader singleton for multithreading

## License
Distributed under the MIT License.

## Contribution
Issues and Pull Requests are welcome.

