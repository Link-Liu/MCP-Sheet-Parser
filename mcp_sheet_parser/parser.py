# parser.py
# 表格解析核心逻辑

import os
import pandas as pd
import openpyxl
import xlrd
from pandas.errors import EmptyDataError
from openpyxl.styles.colors import Color
from .config import Config, ErrorMessages
from .utils import setup_logger, validate_file_path, get_file_extension, is_supported_format, clean_cell_value
from .security import validate_file_path as security_validate_file_path
from .performance import (
    StreamingExcelParser, PerformanceOptimizer, MemoryMonitor,
    create_progress_callback, benchmark_performance
)
from .formula_processor import FormulaProcessor
from .chart_converter import ChartConverter
from .exceptions import (
    ErrorHandler, ErrorContext, ErrorSeverity,
    FileNotFoundError, UnsupportedFormatError, FileSizeExceededError,
    WorksheetParsingError, CellProcessingError, ParsingError,
    error_handler, safe_execute, create_error_context
)


def _get_cell_style(cell):
    style = {}
    # 字体
    if cell.font:
        style['bold'] = cell.font.bold
        style['italic'] = cell.font.italic
        style['font_size'] = cell.font.sz
        style['font_name'] = cell.font.name
        style['underline'] = cell.font.underline
        style['strike'] = cell.font.strike
        # 字体颜色
        if cell.font.color and cell.font.color.type == 'rgb' and cell.font.color.rgb:
            style['font_color'] = f"#{cell.font.color.rgb[-6:]}"
    # 背景色
    if cell.fill and hasattr(cell.fill, 'fgColor') and cell.fill.fgColor.type == 'rgb' and cell.fill.fgColor.rgb and cell.fill.fgColor.rgb != '00000000':
        style['bg_color'] = f"#{cell.fill.fgColor.rgb[-6:]}"
    # 对齐
    if cell.alignment:
        if cell.alignment.horizontal:
            style['align'] = cell.alignment.horizontal
        if cell.alignment.vertical:
            style['valign'] = cell.alignment.vertical
        style['wrap_text'] = cell.alignment.wrap_text
    # 边框
    if cell.border:
        border = {}
        for side in ['top', 'bottom', 'left', 'right']:
            border_side = getattr(cell.border, side)
            if border_side and border_side.style and border_side.style != 'none':
                # 安全处理边框颜色
                border_color = '#000000'  # 默认黑色
                if border_side.color and border_side.color.rgb:
                    try:
                        # 处理RGB对象
                        rgb_value = border_side.color.rgb
                        if hasattr(rgb_value, '__getitem__'):  # 可索引对象
                            border_color = f"#{rgb_value[-6:]}"
                        elif hasattr(rgb_value, 'value'):  # RGB对象可能有value属性
                            border_color = f"#{rgb_value.value[-6:]}"
                        else:
                            border_color = f"#{str(rgb_value)[-6:]}"
                    except (TypeError, AttributeError, IndexError):
                        border_color = '#000000'  # 发生错误时使用默认颜色
                
                border[side] = {
                    'style': border_side.style,
                    'color': border_color
                }
        if border:
            style['border'] = border
    # 超链接
    if cell.hyperlink:
        style['hyperlink'] = str(cell.hyperlink.target)
    # 批注/注释
    if cell.comment:
        style['comment'] = cell.comment.text
    return style


class SheetParser:
    def __init__(self, file_path, config=None, progress_callback=None):
        self.file_path = file_path
        self.config = config or Config()
        self.progress_callback = progress_callback
        self.logger = setup_logger(__name__)
        self.sheets = []  # 存储所有sheet的结构化数据
        self.error_handler = ErrorHandler(self.logger)
        
        # 初始化性能优化组件
        self.performance_optimizer = PerformanceOptimizer()
        self.memory_monitor = MemoryMonitor(
            max_memory_mb=getattr(self.config, 'MAX_MEMORY_MB', 2048)
        ) if getattr(self.config, 'ENABLE_PERFORMANCE_MODE', False) else None
        
        # 初始化公式处理器
        self.formula_processor = FormulaProcessor(config) if getattr(self.config, 'ENABLE_FORMULA_PROCESSING', True) else None
        
        # 初始化图表转换器
        self.chart_converter = ChartConverter(config) if getattr(self.config, 'ENABLE_CHART_CONVERSION', True) else None
        
        # 验证文件
        self._validate_file()
        
    def _validate_file(self):
        """验证文件的完整流程"""
        context = create_error_context("文件验证", file_path=self.file_path)
        
        # 检查文件是否存在
        if not validate_file_path(self.file_path):
            raise FileNotFoundError(self.file_path, context=context)
        
        # 安全检查
        is_safe, warnings = security_validate_file_path(self.file_path)
        if not is_safe:
            from .exceptions import SecurityError
            raise SecurityError(f"文件安全检查失败: {'; '.join(warnings)}", context=context)
        
        # 检查文件格式支持
        if not is_supported_format(self.file_path):
            ext = get_file_extension(self.file_path)
            raise UnsupportedFormatError(self.file_path, ext, context=context)
        
        # 检查文件大小
        file_size_mb = os.path.getsize(self.file_path) / (1024 * 1024)
        if file_size_mb > self.config.MAX_FILE_SIZE_MB:
            raise FileSizeExceededError(
                self.file_path, file_size_mb, self.config.MAX_FILE_SIZE_MB, context=context
            )
        
        # 为大文件优化性能配置
        if file_size_mb > 100 and getattr(self.config, 'ENABLE_PERFORMANCE_MODE', False):
            optimized_config = self.performance_optimizer.optimize_for_file(self.file_path)
            self.config.CHUNK_SIZE = getattr(optimized_config, 'CHUNK_SIZE', 1000)
            self.logger.info(f"启用性能优化模式，块大小: {self.config.CHUNK_SIZE}")
            
            if self.memory_monitor:
                self.logger.info(self.memory_monitor.get_memory_stats())

    def parse(self):
        """
        解析表格文件，返回结构化数据
        """
        context = create_error_context("文件解析", file_path=self.file_path)
        
        try:
            ext = get_file_extension(self.file_path)
            self.logger.info(f"开始解析文件: {self.file_path} (格式: {ext})")
            
            if ext in self.config.SUPPORTED_CSV_FORMATS:
                return [self._parse_csv()]
            elif ext in self.config.SUPPORTED_EXCEL_FORMATS:
                if ext == '.xls':
                    return self._parse_xls()
                else:
                    return self._parse_excel()
            elif ext in self.config.SUPPORTED_WPS_FORMATS:
                # WPS格式目前按Excel处理
                return self._parse_excel()
            else:
                raise UnsupportedFormatError(self.file_path, ext, context=context)
                
        except Exception as e:
            # 使用统一错误处理器
            custom_error = self.error_handler.handle_error(e, context)
            raise custom_error

    def _parse_csv(self):
        """解析CSV文件"""
        context = create_error_context("CSV解析", file_path=self.file_path)
        
        try:
            # 尝试自动检测编码
            encodings = ['utf-8', 'gbk', 'gb2312', 'utf-8-sig']
            df = None
            
            for encoding in encodings:
                try:
                    df = pd.read_csv(self.file_path, header=None, dtype=str, encoding=encoding)
                    self.logger.info(f"成功使用编码 {encoding} 读取CSV文件")
                    break
                except UnicodeDecodeError:
                    continue
            
            if df is None:
                raise ParsingError("无法检测CSV文件编码", context=context)
            
            # 清理数据
            data = df.fillna('').astype(str).values.tolist()
            data = [[clean_cell_value(cell) for cell in row] for row in data]
            
        except EmptyDataError:
            data = []
        except Exception as e:
            custom_error = self.error_handler.handle_error(e, context)
            raise custom_error
        
        # 检查数据大小限制
        self._validate_data_size(data, context)
        
        # CSV无样式
        styles = [[{} for _ in row] for row in data]
        return {
            'sheet_name': os.path.basename(self.file_path),
            'rows': len(data),
            'cols': len(data[0]) if data else 0,
            'data': data,
            'styles': styles,
            'merged_cells': [],  # CSV无合并单元格
            'comments': {},      # CSV无注释
            'hyperlinks': {}     # CSV无超链接
        }

    def _validate_data_size(self, data, context):
        """验证数据大小是否超限"""
        if data and len(data) > self.config.MAX_ROWS:
            context.additional_info = {'rows': len(data), 'max_rows': self.config.MAX_ROWS}
            raise ParsingError(
                f"行数超过限制 ({len(data)} > {self.config.MAX_ROWS})",
                severity=ErrorSeverity.HIGH,
                context=context
            )
            
        if data and len(data[0]) > self.config.MAX_COLS:
            context.additional_info = {'cols': len(data[0]), 'max_cols': self.config.MAX_COLS}
            raise ParsingError(
                f"列数超过限制 ({len(data[0])} > {self.config.MAX_COLS})",
                severity=ErrorSeverity.HIGH,
                context=context
            )

    @error_handler(operation="Excel文件解析")
    def _parse_excel(self):
        """解析现代Excel文件 (.xlsx, .xlsm等)"""
        # 获取文件大小，决定是否使用流式处理
        file_size_mb = os.path.getsize(self.file_path) / (1024 * 1024)
        use_streaming = (file_size_mb > 50 and 
                       getattr(self.config, 'ENABLE_PERFORMANCE_MODE', False))
        
        if use_streaming:
            return self._parse_excel_streaming()
        else:
            return self._parse_excel_standard()
    
    def _parse_excel_standard(self):
        """标准Excel解析方法"""
        context = create_error_context("Excel标准解析", file_path=self.file_path)
        
        try:
            wb = openpyxl.load_workbook(self.file_path, data_only=False)  
            sheets = []
            
            for ws in wb.worksheets:
                try:
                    sheet_context = create_error_context(
                        "工作表解析", 
                        file_path=self.file_path, 
                        sheet_name=ws.title
                    )
                    
                    self.logger.info(f"解析工作表: {ws.title}")
                    
                    # 检查工作表大小
                    self._validate_worksheet_size(ws, sheet_context)
                    
                    # 解析工作表数据
                    sheet_result = self._parse_worksheet_data(ws, sheet_context)
                    
                    # 处理公式（如果启用）
                    if self.formula_processor:
                        sheet_result = safe_execute(
                            self.formula_processor.process_sheet_formulas,
                            sheet_result,
                            operation=f"公式处理-{ws.title}",
                            default_value=sheet_result
                        )
                        self.logger.info(f"工作表 {ws.title} 公式处理完成")
                    
                    # 处理图表（如果启用）
                    if self.chart_converter:
                        sheet_result = safe_execute(
                            self.chart_converter.process_sheet_charts,
                            sheet_result,
                            operation=f"图表转换-{ws.title}",
                            default_value=sheet_result
                        )
                        self.logger.info(f"工作表 {ws.title} 图表处理完成")
                    
                    sheets.append(sheet_result)
                    
                except Exception as e:
                    # 工作表级别的错误处理
                    error = WorksheetParsingError(ws.title, str(e), context=sheet_context, original_error=e)
                    self.error_handler._log_error(error)
                    # 继续处理其他工作表，不中断整体流程
                    continue
            
            return sheets
            
        except Exception as e:
            custom_error = self.error_handler.handle_error(e, context)
            raise custom_error
    
    def _validate_worksheet_size(self, ws, context):
        """验证工作表大小"""
        if ws.max_row > self.config.MAX_ROWS:
            context.additional_info = {'rows': ws.max_row, 'max_rows': self.config.MAX_ROWS}
            raise ParsingError(
                f"工作表行数超过限制 ({ws.max_row} > {self.config.MAX_ROWS})",
                severity=ErrorSeverity.HIGH,
                context=context
            )
            
        if ws.max_column > self.config.MAX_COLS:
            context.additional_info = {'cols': ws.max_column, 'max_cols': self.config.MAX_COLS}
            raise ParsingError(
                f"工作表列数超过限制 ({ws.max_column} > {self.config.MAX_COLS})",
                severity=ErrorSeverity.HIGH,
                context=context
            )
    
    def _parse_worksheet_data(self, ws, context):
        """解析工作表数据"""
        data = []
        styles = []
        comments = {}  # 注释字典
        hyperlinks = {}  # 超链接字典
        
        # 使用max_row和max_column确保遍历所有行和列
        for r in range(1, ws.max_row + 1):
            data_row = []
            style_row = []
            for c in range(1, ws.max_column + 1):
                try:
                    cell = ws.cell(row=r, column=c)
                    
                    # 获取单元格值，优先获取公式
                    if cell.data_type == 'f':  # 公式单元格
                        # 检查公式是否已经包含等号
                        formula_value = str(cell.value) if cell.value else ''
                        if formula_value and not formula_value.startswith('='):
                            cell_value = f"={formula_value}"
                        else:
                            cell_value = formula_value
                    else:
                        # 对于非公式单元格，直接使用值
                        cell_value = cell.value if cell.value is not None else ''
                    
                    data_row.append(clean_cell_value(cell_value))
                    
                    # 获取样式信息 - 使用安全执行
                    style = safe_execute(
                        _get_cell_style, 
                        cell,
                        operation=f"样式提取-{r},{c}",
                        default_value={}
                    )
                    style_row.append(style)
                    
                    # 提取注释和超链接到独立字典
                    cell_key = f"{r-1}_{c-1}"  # 转换为0-based索引
                    
                    if 'comment' in style and style['comment']:
                        comments[cell_key] = style['comment']
                    
                    if 'hyperlink' in style and style['hyperlink']:
                        hyperlinks[cell_key] = style['hyperlink']
                        
                except Exception as e:
                    # 单元格级别的错误处理
                    cell_pos = f"{r},{c}"
                    error = CellProcessingError(cell_pos, str(e), context=context, original_error=e)
                    self.error_handler._log_error(error)
                    
                    # 使用默认值继续处理
                    data_row.append('')
                    style_row.append({})
                        
            data.append(data_row)
            styles.append(style_row)
        
        # 处理合并单元格
        merged_cells = []
        for merged in ws.merged_cells.ranges:
            merged_cells.append((
                merged.min_row-1, merged.min_col-1, 
                merged.max_row-1, merged.max_col-1
            ))
        
        return {
            'sheet_name': ws.title,
            'rows': len(data),
            'cols': len(data[0]) if data else 0,
            'data': data,
            'styles': styles,
            'merged_cells': merged_cells,
            'comments': comments,
            'hyperlinks': hyperlinks
        }
    
    def _parse_excel_streaming(self):
        """流式Excel解析方法"""
        context = create_error_context("Excel流式解析", file_path=self.file_path)
        
        try:
            self.logger.info("使用流式解析模式处理大文件")
            
            # 创建流式解析器
            chunk_size = getattr(self.config, 'CHUNK_SIZE', 1000)
            streaming_parser = StreamingExcelParser(
                chunk_size=chunk_size,
                enable_memory_monitor=self.memory_monitor is not None
            )
            
            # 创建进度回调
            progress_callback = None
            if (self.progress_callback and 
                getattr(self.config, 'ENABLE_PROGRESS_TRACKING', True)):
                progress_callback = self.progress_callback
            
            # 收集所有数据块
            worksheets = {}
            
            for chunk in streaming_parser.parse_in_chunks(self.file_path, progress_callback):
                ws_name = chunk['worksheet_name']
                
                if ws_name not in worksheets:
                    worksheets[ws_name] = {
                        'sheet_name': ws_name,
                        'data': [],
                        'styles': [],
                        'comments': {},
                        'hyperlinks': {},
                        'merged_cells': []
                    }
                
                # 合并数据块
                ws_data = worksheets[ws_name]
                ws_data['data'].extend(chunk['data'])
                ws_data['styles'].extend(chunk['styles'])
                ws_data['comments'].update(chunk['comments'])
                ws_data['hyperlinks'].update(chunk['hyperlinks'])
                ws_data['merged_cells'].extend(chunk['merged_cells'])
                
                # 内存检查
                if self.memory_monitor and self.memory_monitor.should_trigger_gc():
                    self.memory_monitor.trigger_gc()
            
            # 转换为最终格式
            sheets = []
            for ws_data in worksheets.values():
                ws_data['rows'] = len(ws_data['data'])
                ws_data['cols'] = len(ws_data['data'][0]) if ws_data['data'] else 0
                sheets.append(ws_data)
            
            return sheets
            
        except Exception as e:
            custom_error = self.error_handler.handle_error(e, context)
            raise custom_error

    @error_handler(operation="XLS文件解析")
    def _parse_xls(self):
        """解析旧版Excel文件 (.xls)"""
        context = create_error_context("XLS解析", file_path=self.file_path)
        
        workbook = xlrd.open_workbook(self.file_path)
        sheets = []
        
        for sheet_index in range(workbook.nsheets):
            worksheet = workbook.sheet_by_index(sheet_index)
            sheet_name = worksheet.name
            
            sheet_context = create_error_context(
                "XLS工作表解析", 
                file_path=self.file_path, 
                sheet_name=sheet_name
            )
            
            try:
                self.logger.info(f"解析工作表: {sheet_name}")
                
                # 检查工作表大小
                if worksheet.nrows > self.config.MAX_ROWS:
                    raise ParsingError(
                        f"XLS工作表行数超过限制 ({worksheet.nrows} > {self.config.MAX_ROWS})",
                        severity=ErrorSeverity.HIGH,
                        context=sheet_context
                    )
                    
                if worksheet.ncols > self.config.MAX_COLS:
                    raise ParsingError(
                        f"XLS工作表列数超过限制 ({worksheet.ncols} > {self.config.MAX_COLS})",
                        severity=ErrorSeverity.HIGH,
                        context=sheet_context
                    )
                
                data = []
                styles = []  # XLS文件样式提取较复杂，暂时用空样式
                
                for row_idx in range(worksheet.nrows):
                    data_row = []
                    style_row = []
                    for col_idx in range(worksheet.ncols):
                        try:
                            cell_value = worksheet.cell_value(row_idx, col_idx)
                            # 处理日期
                            if worksheet.cell_type(row_idx, col_idx) == xlrd.XL_CELL_DATE:
                                if isinstance(cell_value, (int, float)):
                                    date_tuple = xlrd.xldate_as_tuple(cell_value, workbook.datemode)
                                    cell_value = f"{date_tuple[0]}-{date_tuple[1]:02d}-{date_tuple[2]:02d}"
                            data_row.append(clean_cell_value(cell_value))
                            style_row.append({})  # 暂时没有样式信息
                        except IndexError:
                            data_row.append('')
                            style_row.append({})
                    data.append(data_row)
                    styles.append(style_row)
                
                # XLS文件的合并单元格提取较复杂，暂时为空
                merged_cells = []
                
                sheets.append({
                    'sheet_name': sheet_name,
                    'rows': len(data),
                    'cols': len(data[0]) if data else 0,
                    'data': data,
                    'styles': styles,
                    'merged_cells': merged_cells,
                    'comments': {},      # XLS注释提取复杂，暂时为空
                    'hyperlinks': {}     # XLS超链接提取复杂，暂时为空
                })
                
            except Exception as e:
                error = WorksheetParsingError(sheet_name, str(e), context=sheet_context, original_error=e)
                self.error_handler._log_error(error)
                # 继续处理其他工作表
                continue
                
        return sheets
    
    def get_error_summary(self):
        """获取解析过程中的错误统计"""
        return self.error_handler.get_error_summary() 