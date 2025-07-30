# LibreOffice Python 自动化工具

**背景需求**  
作为统计从业人员，需频繁处理大量表格和报告。受限于 Linux 工作环境无法使用 Office VBA，特开发此工具实现 LibreOffice 自动化操作。

## 功能特性
- 📊 **Excel 自动化**：数据读写、格式处理、单元格合并、公式计算
- 📝 **Word 自动化**：文档内容替换、模板化填充
- 🔄 **批量处理**：支持多数据源文件的批量化操作
- 🎯 **专业报告**：针对银行业风险监测报告定制优化
- 🧵 **线程安全**：单例模式管理 LibreOffice 连接实例

## 核心模块
### 工作簿处理 (`workbook.py`)
Excel 操作核心类，提供：
- 工作簿创建/打开/保存
- 数据范围读写
- 单元格合并与格式化
- pandas DataFrame 集成
- 自动求和功能

### 文档处理 (`word.py`)
Word 文档操作类，支持：
- 文档创建/打开/保存
- 批量文本替换
- 内容检索定位

### 连接管理 (`officeLoader.py`)
LibreOffice 连接管理器（单例模式）：
- 线程安全的实例管理
- 自动连接与资源回收
- 上下文管理器支持

## 安装部署
使用 [poetry](https://python-poetry.org/) 管理依赖，新用户请参考[官方指南](https://python-poetry.org/docs/basic-usage/)。
```sh
# 安装 poetry
pip install poetry

# 克隆项目
git clone https://github.com/rainbowtrash2333/libreoffice_py.git
cd libreoffice_py
```

### Windows 环境
```sh
# 创建虚拟环境
py -3.9 -m venv .\.venv
.\.venv\Scripts\activate

# 安装依赖
poetry install

# 切换至 UNO 环境
oooenv env -t && oooenv env -u

# 执行 poetry 命令后需重新切换环境
oooenv env -t
poetry <command>
oooenv env -t
```

### Linux 环境
```sh
python3 -m venv .venv
source .venv/bin/activate
poetry install
```

### 环境验证
```sh
python -c "import uno; print('环境校验通过')"
```

## 使用示例
### Excel 基础操作
```python
from workbook import Workbook

with Workbook(filepath="data.xlsx", visible=False) as wb:
    data = wb.get_used_value(0, range_name='A1:C10')  # 读取数据
```

### Word 模板处理
```python
from word import Word

with Word(filepath="template.doc") as doc:
    doc.replace_words(
        labels=["$(公司)", "$(日期)"], 
        values=["ABC集团", "2025-07-30"]
    )
    doc.save()  # 自动关闭文档
```

### 批量报告生成
```sh
python gen_xls.py  # 执行报表生成主程序
```

## 工具函数 (`myutil.py`)
- **数据转换**：`array2df()` - 元组数据转 DataFrame
- **数值处理**：`process_value_to_str()` - 智能数值格式化
- **坐标转换**：`convert_cell_name_to_list()` - 单元格地址解析
- **文件校验**：`check_files_exist()` - 批量文件存在性检查

## 路径配置
根据实际环境修改以下路径：
- `data_path`：数据源目录
- `template_path`：模板文件目录  
- `result_path`：输出结果目录

## 注意事项
1. 需预装 LibreOffice 7.0+
2. 推荐使用绝对路径确保文件访问
3. 操作完成后主动关闭文档释放资源
4. 多线程场景使用 OfficeLoader 单例

## 许可协议
本项目采用 MIT 开源许可证。

## 项目贡献
欢迎通过 Issue 反馈问题或提交 Pull Request 参与改进。

