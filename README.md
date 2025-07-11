# MCP-Sheet-Parser

一个专注于HTML转换的表格解析工具，支持Excel、CSV等格式转换为美观的HTML表格。

## ✨ 特性

- 🚀 **快速转换**: 高效解析Excel、CSV等表格文件
- 🎨 **多种主题**: 内置4种精美主题（默认、极简、暗色、打印）
- 📊 **完整功能**: 支持合并单元格、样式、注释、超链接
- 🔒 **安全可靠**: 内置安全检查，防止XSS攻击
- 📱 **响应式设计**: 自动适配移动设备显示
- 🛠️ **简单易用**: 命令行界面，支持批量处理

## 🎯 支持格式

### 输入格式
- **Excel**: `.xlsx`, `.xls`, `.xlsm`, `.xltm`
- **CSV**: `.csv`
- **WPS**: `.et`, `.ett`, `.ets`

### 输出格式
- **HTML**: 完整的HTML文档或纯表格代码

## 📦 安装

### 克隆仓库
```bash
git clone https://github.com/your-username/MCP-Sheet-Parser.git
cd MCP-Sheet-Parser
```

### 安装依赖
```bash
pip install -r requirements.txt
```

## 🚀 快速开始

### 基本使用
```bash
# 转换单个文件
python main.py data.xlsx --html output.html

# 使用不同主题
python main.py data.csv --html output.html --theme dark

# 只输出表格（不含HTML头部）
python main.py data.xlsx --html table.html --table-only
```

### 批量处理
```bash
# 批量转换多个文件
python main.py *.xlsx --batch --output-dir ./html_files/

# 批量转换并使用特定主题
python main.py data/*.csv --batch --output-dir ./output/ --theme minimal
```

### 主题预览
```bash
# 查看所有可用主题
python main.py --list-themes
```

### 文件信息
```bash
# 查看文件信息
python main.py data.xlsx --info
```

## 🎨 主题展示

| 主题名称 | 描述 | 适用场景 |
|---------|------|----------|
| `default` | 标准样式，平衡美观与实用 | 一般用途 |
| `minimal` | 极简设计，清爽干净 | 简洁展示 |
| `dark` | 暗色主题，护眼舒适 | 夜间使用 |
| `print` | 打印优化，黑白适配 | 打印文档 |

## 📋 功能特性

### 样式支持
- ✅ 字体样式（粗体、斜体、字号、颜色）
- ✅ 背景颜色
- ✅ 文本对齐（水平、垂直）
- ✅ 边框样式
- ✅ 合并单元格

### 交互功能
- ✅ 单元格注释（悬停显示）
- ✅ 超链接（新窗口打开）
- ✅ 响应式布局
- ✅ 移动设备适配

### 安全特性
- ✅ HTML内容转义
- ✅ 文件路径验证
- ✅ 大小限制检查
- ✅ 恶意内容过滤

## 📚 API使用

### 编程接口
```python
from mcp_sheet_parser.parser import SheetParser
from mcp_sheet_parser.html_converter import HTMLConverter

# 解析文件
parser = SheetParser('data.xlsx')
sheets = parser.parse()

# 转换为HTML
converter = HTMLConverter(sheets[0], theme='dark')
html = converter.to_html()

# 保存文件
converter.export_to_file('output.html')
```

### 配置选项
```python
from mcp_sheet_parser.config import Config

config = Config()
config.INCLUDE_COMMENTS = False  # 不包含注释
config.INCLUDE_HYPERLINKS = False  # 不包含超链接
config.MAX_FILE_SIZE_MB = 50  # 最大文件大小
```

## 🧪 运行测试

```bash
# 运行所有测试
python -m pytest tests/ -v

# 运行特定测试
python -m pytest tests/test_parser.py -v

# 生成覆盖率报告
python -m pytest tests/ --cov=mcp_sheet_parser --cov-report=html
```

## 📁 项目结构

```
MCP-Sheet-Parser/
├── mcp_sheet_parser/           # 核心模块
│   ├── __init__.py
│   ├── config.py              # 配置和常量
│   ├── parser.py              # 文件解析器
│   ├── html_converter.py      # HTML转换器
│   ├── data_validator.py      # 数据验证和清理
│   ├── security.py            # 安全检查
│   └── utils.py               # 工具函数
├── tests/                     # 测试文件
│   ├── test_parser.py
│   ├── test_html_converter.py
│   └── test_utils.py
├── examples/                  # 示例文件
├── main.py                    # 命令行入口
└── requirements.txt           # 依赖列表
```

## 📖 命令行参数

### 基本参数
- `input` - 输入文件路径（支持通配符）
- `--html, -o` - 输出HTML文件路径
- `--output-dir, -d` - 批量处理输出目录

### 样式选项
- `--theme, -t` - 选择主题（default/minimal/dark/print）
- `--table-only` - 只输出表格HTML
- `--no-comments` - 不包含单元格注释
- `--no-hyperlinks` - 不包含超链接

### 处理选项
- `--batch, -b` - 批量处理模式
- `--sheet N` - 只处理指定工作表（从0开始）

### 信息选项
- `--info, -i` - 显示文件信息
- `--list-themes` - 显示可用主题
- `--verbose, -v` - 详细输出
- `--version` - 显示版本信息

## 🔧 自定义配置

### 修改默认设置
```python
# 在config.py中修改
class Config:
    MAX_FILE_SIZE_MB = 200          # 最大文件大小
    MAX_ROWS = 2000000             # 最大行数
    MAX_COLS = 32768               # 最大列数
    HTML_DEFAULT_ENCODING = 'utf-8' # 默认编码
    INCLUDE_COMMENTS = True         # 包含注释
    INCLUDE_HYPERLINKS = True       # 包含超链接
```

### 添加自定义主题
```python
# 在config.py的THEMES字典中添加
THEMES['custom'] = {
    'name': '自定义主题',
    'description': '我的专属主题',
    'body_style': 'font-family: "Microsoft YaHei"; margin: 10px;',
    'table_style': 'border-collapse: collapse; width: 100%;',
    'cell_style': 'border: 1px solid #ccc; padding: 6px;',
    'header_style': 'background-color: #4CAF50; color: white;'
}
```

## 🐛 故障排除

### 常见问题

1. **编码错误**
   ```bash
   # 对于特殊编码的CSV文件，工具会自动检测
   # 支持: utf-8, gbk, gb2312, utf-8-sig
   ```

2. **文件过大**
   ```bash
   # 调整配置中的大小限制
   config.MAX_FILE_SIZE_MB = 500
   ```

3. **权限问题**
   ```bash
   # 确保输出目录有写入权限
   chmod 755 output_directory
   ```

## 📝 更新日志

### v1.0.0 (2024-12-XX)
- ✨ 专注HTML转换功能
- 🎨 4种内置主题
- 🔒 安全性增强
- 📱 响应式设计
- 🧪 完整测试覆盖

## 🤝 贡献

1. Fork本项目
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 创建Pull Request

## 📄 许可证

本项目采用MIT许可证 - 查看 [LICENSE](LICENSE) 文件了解详情

## 💬 支持

- 🐛 报告Bug: [Issues](https://github.com/your-username/MCP-Sheet-Parser/issues)
- 💡 功能建议: [Discussions](https://github.com/your-username/MCP-Sheet-Parser/discussions)
- 📧 联系我们: your-email@example.com

---

⭐ 如果这个项目对你有帮助，请给它一个Star！
