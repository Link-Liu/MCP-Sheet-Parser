# main.py
# MCP-Sheet-Parser 程序入口

def main():
    print("欢迎使用 MCP-Sheet-Parser！")
    import sys
    from mcp_sheet_parser.parser import SheetParser
    from mcp_sheet_parser.html_converter import HTMLConverter
    if len(sys.argv) < 2:
        print("用法: python main.py <表格文件路径> [--html <输出HTML路径>]")
        return
    file_path = sys.argv[1]
    parser = SheetParser(file_path)
    try:
        sheets = parser.parse()
        for idx, sheet in enumerate(sheets):
            print(f"\nSheet名称: {sheet['sheet_name']}")
            print(f"行数: {sheet['rows']}，列数: {sheet['cols']}")
            print("部分数据预览:")
            for row in sheet['data'][:5]:
                print(row)
            if sheet['merged_cells']:
                print(f"合并单元格信息: {sheet['merged_cells']}")
        # 如果有 --html 参数，输出HTML
        if '--html' in sys.argv:
            html_idx = sys.argv.index('--html')
            html_path = sys.argv[html_idx+1] if html_idx+1 < len(sys.argv) else 'output.html'
            # 只导出第一个sheet为HTML
            converter = HTMLConverter(sheets[0])
            html = converter.to_html()
            with open(html_path, 'w', encoding='utf-8') as f:
                f.write(html)
            print(f"已导出HTML到: {html_path}")
    except Exception as e:
        print(f"解析失败: {e}")

if __name__ == "__main__":
    main()
