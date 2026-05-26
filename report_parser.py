import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from typing import Dict, List, Tuple, Set
import threading

# 尝试导入PDF处理库
try:
    import pdfplumber

    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    import PyPDF2

    HAS_PYPDF2 = True
except ImportError:
    HAS_PYPDF2 = False


class PDFTextExtractor:
    """PDF文本提取器"""

    def extract_text_from_pdf(self, file_path: str) -> str:
        """从PDF文件中提取文本"""
        text = ""

        # 优先使用pdfplumber，因为它通常能提取更多文本
        if HAS_PDFPLUMBER:
            try:
                with pdfplumber.open(file_path) as pdf:
                    for page in pdf.pages:
                        # 尝试使用不同的提取策略
                        page_text = page.extract_text()
                        if not page_text or len(page_text.strip()) < 10:
                            # 如果文本提取不理想，尝试使用布局模式
                            page_text = page.extract_text(layout=True)
                        if page_text:
                            text += page_text + "\n"
                if text.strip():
                    return text
            except Exception as e:
                print(f"使用pdfplumber提取文本失败: {e}")

        # 回退到PyPDF2
        if HAS_PYPDF2:
            try:
                with open(file_path, 'rb') as file:
                    pdf_reader = PyPDF2.PdfReader(file)
                    for page in pdf_reader.pages:
                        page_text = page.extract_text()
                        if page_text:
                            text += page_text + "\n"
                return text
            except Exception as e:
                print(f"使用PyPDF2提取文本失败: {e}")

        # 如果两种方法都失败
        raise Exception("无法从PDF中提取文本，请确保已安装PyPDF2或pdfplumber")

    def extract_tables_from_pdf(self, file_path: str) -> List:
        """从PDF中提取表格数据"""
        tables = []
        if HAS_PDFPLUMBER:
            try:
                with pdfplumber.open(file_path) as pdf:
                    for page in pdf.pages:
                        page_tables = page.extract_tables()
                        if page_tables:
                            tables.extend(page_tables)
            except Exception as e:
                print(f"提取PDF表格时出错: {e}")
        return tables


class ChemicalReportAnalyzer:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("报告解析器")
        # 进一步减小界面大小
        self.root.geometry("700x500")

        # 检查PDF支持状态
        self.has_pdf_support = HAS_PDFPLUMBER or HAS_PYPDF2

        # 标记物
        self.markers = set()

        self.file_contents = {}

        # 初始化PDF提取器
        if self.has_pdf_support:
            self.pdf_extractor = PDFTextExtractor()
        else:
            self.pdf_extractor = None

        # 改进浓度匹配模式
        self.concentration_pattern = re.compile(r'(\d+\.\d+)\s*(ppm|mg/L|μg/mL|ng/ml)', re.IGNORECASE)
        self.concentration_pattern_int = re.compile(r'(\d+)\s*(ppm|mg/L|μg/mL|ng/ml)', re.IGNORECASE)
        self.nd_pattern = re.compile(r'N\.?\s*D\.?', re.IGNORECASE)
        self.uncorrected_pattern = re.compile(r'未校正', re.IGNORECASE)

        # 有效的化合物名称模式 - 使用更宽松的正则表达式
        self.valid_compound_pattern = re.compile(r'^[A-Za-z0-9\-\u4e00-\u9fff,\.\'\(\)\）\（\[\]_/]+$')

        self.setup_ui()

    def setup_ui(self):
        """设置用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="4")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 文件选择区域 - 更紧凑的设计
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="4")
        file_frame.pack(fill=tk.X, pady=(0, 4))

        # 文件选择按钮和开始分析按钮在同一行
        file_top_frame = ttk.Frame(file_frame)
        file_top_frame.pack(fill=tk.X, pady=1)

        ttk.Button(file_top_frame, text="选择报告文件",
                   command=self.select_files).pack(side=tk.LEFT, padx=(0, 8))

        self.file_count_label = ttk.Label(file_top_frame, text="未选择文件")
        self.file_count_label.pack(side=tk.LEFT, padx=(0, 10))

        # 开始分析按钮放在同一行右对齐，使用相同格式
        ttk.Button(file_top_frame, text="开始分析",
                   command=self.start_analysis).pack(side=tk.RIGHT)

        # 分析设置区域 - 更紧凑的设计
        settings_frame = ttk.LabelFrame(main_frame, text="分析设置", padding="4")
        settings_frame.pack(fill=tk.X, pady=(0, 4))

        # 阈值设置和标记物设置在同一行
        settings_top_frame = ttk.Frame(settings_frame)
        settings_top_frame.pack(fill=tk.X, pady=1)

        # 阈值设置
        threshold_frame = ttk.Frame(settings_top_frame)
        threshold_frame.pack(side=tk.LEFT, padx=(0, 15))

        ttk.Label(threshold_frame, text="检测阈值:").pack(side=tk.LEFT)
        self.threshold_var = tk.StringVar(value="0.1")
        ttk.Entry(threshold_frame, textvariable=self.threshold_var, width=6).pack(side=tk.LEFT, padx=3)
        ttk.Label(threshold_frame, text="(≥此值为阳性)").pack(side=tk.LEFT)

        # 标记物设置
        marker_frame = ttk.Frame(settings_top_frame)
        marker_frame.pack(side=tk.LEFT)

        self.marker_var = tk.BooleanVar()
        marker_check = ttk.Checkbutton(marker_frame, text="排除标记物",
                                       variable=self.marker_var, command=self.toggle_marker_input)
        marker_check.pack(side=tk.LEFT)

        self.marker_entry = ttk.Entry(marker_frame, width=18, state="disabled")
        self.marker_entry.pack(side=tk.LEFT, padx=3)
        self.marker_entry.bind("<Return>", self.update_markers)

        # 结果显示区域
        result_frame = ttk.LabelFrame(main_frame, text="分析结果", padding="4")
        result_frame.pack(fill=tk.BOTH, expand=True)

        # 创建标签页
        notebook = ttk.Notebook(result_frame)
        notebook.pack(fill=tk.BOTH, expand=True)

        # 摘要标签页 - 减小高度
        summary_tab = ttk.Frame(notebook, padding="4")
        notebook.add(summary_tab, text="摘要")

        self.summary_text = scrolledtext.ScrolledText(summary_tab, height=8, wrap=tk.WORD)
        self.summary_text.pack(fill=tk.BOTH, expand=True)

        # 详细结果标签页 - 减小高度
        detail_tab = ttk.Frame(notebook, padding="4")
        notebook.add(detail_tab, text="详细结果")

        self.detail_text = scrolledtext.ScrolledText(detail_tab, height=8, wrap=tk.WORD)
        self.detail_text.pack(fill=tk.BOTH, expand=True)

        # 化合物列表标签页 - 减小高度
        compound_tab = ttk.Frame(notebook, padding="4")
        notebook.add(compound_tab, text="所有化合物")

        self.compound_text = scrolledtext.ScrolledText(compound_tab, height=8, wrap=tk.WORD)
        self.compound_text.pack(fill=tk.BOTH, expand=True)

        # 调试信息标签页 - 减小高度
        debug_tab = ttk.Frame(notebook, padding="4")
        notebook.add(debug_tab, text="调试信息")

        self.debug_text = scrolledtext.ScrolledText(debug_tab, height=8, wrap=tk.WORD)
        self.debug_text.pack(fill=tk.BOTH, expand=True)

        # 状态栏 - 移除边框阴影
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief="flat")
        status_bar.pack(fill=tk.X, side=tk.BOTTOM, pady=(2, 0))

    def toggle_marker_input(self):
        """切换标记物输入框状态"""
        if self.marker_var.get():
            self.marker_entry.config(state="normal")
            self.status_var.set("请输入标记物名称，用逗号分隔")
        else:
            self.marker_entry.config(state="disabled")
            self.markers.clear()
            self.status_var.set("标记物排除已禁用")

    def update_markers(self, event=None):
        """更新标记物列表"""
        marker_text = self.marker_entry.get().strip()
        if marker_text:
            self.markers = {marker.strip() for marker in marker_text.split(",")}
            self.status_var.set(f"已设置 {len(self.markers)} 个标记物")
        else:
            self.markers.clear()
            self.status_var.set("标记物列表已清空")

    def select_files(self):
        """选择报告文件"""
        filetypes = [("文本文件", "*.txt"), ("所有文件", "*.*")]
        if self.has_pdf_support:
            filetypes.insert(0, ("PDF文件", "*.pdf"))

        files = filedialog.askopenfilenames(
            title="选择报告文件",
            filetypes=filetypes
        )

        if files:
            self.file_contents.clear()
            self.status_var.set(f"正在读取 {len(files)} 个文件...")

            # 在新线程中读取文件
            threading.Thread(target=self.read_files, args=(files,), daemon=True).start()

    def read_files(self, files):
        """读取文件内容"""
        successful_files = 0

        for file_path in files:
            try:
                # 检查文件扩展名
                _, ext = os.path.splitext(file_path)

                if ext.lower() == '.pdf':
                    # 如果没有PDF支持，跳过PDF文件
                    if not self.has_pdf_support:
                        continue

                    # 使用PDF提取器处理PDF文件
                    content = self.pdf_extractor.extract_text_from_pdf(file_path)
                else:
                    # 处理文本文件
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                        content = f.read()

                if content:
                    filename = os.path.basename(file_path)
                    self.file_contents[filename] = {
                        'content': content,
                        'path': file_path
                    }
                    successful_files += 1

            except Exception as e:
                print(f"读取文件 {file_path} 时出错: {e}")

        # 更新UI
        self.root.after(0, self.update_file_count, successful_files, len(files))

    def update_file_count(self, successful, total):
        """更新文件计数显示"""
        self.file_count_label.config(text=f"已选择 {successful}/{total} 个文件")
        if successful < total and not self.has_pdf_support:
            self.status_var.set(f"成功读取 {successful} 个文件 (跳过PDF文件)")
        else:
            self.status_var.set(f"成功读取 {successful} 个文件")

    def start_analysis(self):
        """开始分析"""
        if not self.file_contents:
            messagebox.showwarning("警告", "请先选择要分析的文件")
            return

        try:
            threshold = float(self.threshold_var.get())
        except ValueError:
            messagebox.showerror("错误", "请输入有效的阈值数字")
            return

        # 更新标记物列表
        if self.marker_var.get():
            self.update_markers()

        self.status_var.set("正在分析报告...")

        # 在新线程中进行分析
        threading.Thread(target=self.analyze_reports_thread, args=(threshold,), daemon=True).start()

    def analyze_reports_thread(self, threshold):
        """在新线程中分析报告"""
        try:
            detected, not_detected, details, debug_info = self.analyze_reports(self.file_contents, threshold)
            self.root.after(0, self.display_results, detected, not_detected, details, debug_info, threshold)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("分析错误", str(e)))

    def analyze_reports(self, file_contents: Dict[str, dict], threshold: float) -> Tuple[
        List[str], List[str], Dict, str]:
        """分析所有报告"""
        detected_reports = []
        not_detected_reports = []
        detailed_results = {}
        debug_info = "=== 调试信息 ===\n\n"

        for filename, file_info in file_contents.items():
            content = file_info['content']
            file_path = file_info['path']
            compounds, file_debug_info = self.parse_report_content(content, filename, file_path)
            debug_info += file_debug_info

            # 过滤掉标记物
            target_compounds = {}
            for name, conc in compounds.items():
                if not self.is_marker(name):
                    target_compounds[name] = conc

            # 检查是否有超过等于阈值的化合物
            detected_compounds = {name: conc for name, conc in target_compounds.items()
                                  if conc >= threshold}  # 改为大于等于

            if detected_compounds:
                detected_reports.append(filename)
            else:
                not_detected_reports.append(filename)

            detailed_results[filename] = {
                'all_compounds': compounds,
                'target_compounds': target_compounds,
                'detected_compounds': detected_compounds,
                'has_detection': len(detected_compounds) > 0
            }

        return detected_reports, not_detected_reports, detailed_results, debug_info

    def parse_report_content(self, content: str, filename: str, file_path: str) -> Tuple[Dict[str, float], str]:
        """解析报告内容，提取化合物和浓度信息"""
        compounds = {}
        debug_info = f"文件: {filename}\n"

        # 首先尝试MH的TSCA报告解析方法
        mh_compounds, mh_debug = self.parse_tsca_format(content, filename, file_path)
        debug_info += mh_debug

        if mh_compounds:
            compounds.update(mh_compounds)
            debug_info += f"  MH TSCA格式解析成功: {len(mh_compounds)} 个化合物\n"

        # 如果MH方法没有找到化合物，尝试原有的多种解析策略
        if not compounds:
            lines = content.split('\n')
            strategies = [
                self.parse_compact_format,  # 先尝试紧凑格式（PAE报告）
                self.parse_standard_quant_report,  # 再尝试标准定量报告格式
                self.parse_table_format,  # 再尝试表格格式
                self.parse_standard_format,  # 最后尝试标准格式
            ]

            for strategy in strategies:
                result, strategy_debug = strategy(lines, filename, file_path)
                debug_info += strategy_debug
                if result:
                    # 只添加新解析的化合物，避免覆盖
                    for name, conc in result.items():
                        if name not in compounds and self.is_valid_compound(name):
                            compounds[name] = conc
                    debug_info += f"  策略 {strategy.__name__} 成功解析了 {len(result)} 个化合物\n"

        debug_info += f"  最终解析结果: {len(compounds)} 个化合物\n\n"
        return compounds, debug_info

    def parse_tsca_format(self, content: str, filename: str, file_path: str) -> Tuple[Dict[str, float], str]:
        """解析TSCA报告格式（MH的解析方法）"""
        compounds = {}
        debug_info = f"  尝试MH TSCA格式解析...\n"

        # 提取化合物和浓度信息
        tsca_compounds, istd_info = self.extract_compounds_and_concentrations(content, file_path)
        debug_info += f"    提取到 {len(tsca_compounds)} 个化合物\n"

        # 转换为标准格式
        for compound_name, concentration_str in tsca_compounds.items():
            if self.is_valid_compound(compound_name):
                # 解析浓度值
                if concentration_str == "N.D.":
                    compounds[compound_name] = 0.0
                else:
                    # 提取数值部分
                    conc_match = re.search(r'(\d+\.\d+)\s*(mg/L|μg/mL|ng/ml)', concentration_str, re.IGNORECASE)
                    if conc_match:
                        try:
                            concentration = float(conc_match.group(1))
                            compounds[compound_name] = concentration
                        except ValueError:
                            compounds[compound_name] = 0.0
                    else:
                        compounds[compound_name] = 0.0

        return compounds, debug_info

    def extract_compounds_and_concentrations(self, content: str, file_path: str) -> Tuple[Dict[str, str], Dict]:
        """从TSCA报告中提取化合物和浓度信息"""
        compounds = {}
        istd_info = {}

        # 首先尝试从表格中提取数据
        tables = self.pdf_extractor.extract_tables_from_pdf(file_path)
        table_compounds = self.extract_from_tables(tables, content)
        compounds.update(table_compounds)

        # 如果表格提取失败，尝试从文本中提取
        if not compounds:
            text_compounds, istd_info = self.extract_from_text(content)
            compounds.update(text_compounds)

        # 对于外标法报告，使用专门的提取方法
        if '外标法' in content:
            external_compounds = self.extract_from_external_standard(content)
            compounds.update(external_compounds)

        # 清理无效的化合物名称
        compounds = self.clean_compounds(compounds)

        return compounds, istd_info

    def extract_from_tables(self, tables: List, content: str) -> Dict[str, str]:
        """从PDF表格中提取化合物和浓度"""
        compounds = {}

        for table in tables:
            if not table:
                continue

            # 查找表头行
            header_row = -1
            for i, row in enumerate(table):
                if row and any(
                        cell and any(keyword in str(cell) for keyword in ['名称', '化合物', 'ISTD', '浓度']) for cell in
                        row):
                    header_row = i
                    break

            if header_row == -1:
                continue

            # 确定名称和浓度所在的列
            name_col = -1
            conc_col = -1

            for j, cell in enumerate(table[header_row]):
                if cell and any(keyword in str(cell) for keyword in ['名称', '化合物']):
                    name_col = j
                elif cell and any(keyword in str(cell) for keyword in ['浓度', '最终浓度']):
                    conc_col = j

            # 如果没找到浓度列，尝试最后一列
            if conc_col == -1 and len(table[header_row]) > 1:
                conc_col = len(table[header_row]) - 1

            # 提取数据行
            for i in range(header_row + 1, len(table)):
                row = table[i]
                if not row or len(row) <= max(name_col, conc_col):
                    continue

                name_cell = row[name_col] if name_col < len(row) else None
                conc_cell = row[conc_col] if conc_col < len(row) else None

                if name_cell and self.is_valid_compound(name_cell):
                    # 处理浓度值
                    concentration = "N.D."
                    if conc_cell:
                        if 'N.D.' in str(conc_cell):
                            concentration = "N.D."
                        else:
                            conc_match = re.search(r'(\d+\.\d+)\s*(mg/L|μg/mL|ng/ml)', str(conc_cell), re.IGNORECASE)
                            if conc_match:
                                concentration = f"{conc_match.group(1)} {conc_match.group(2)}"

                    compounds[name_cell] = concentration

        return compounds

    def extract_from_text(self, content: str) -> Tuple[Dict[str, str], Dict]:
        """从文本中提取化合物和浓度"""
        compounds = {}
        istd_info = {}
        lines = content.split('\n')

        # 查找表格区域
        table_start = -1
        for i, line in enumerate(lines):
            # 内标法表头
            if '名称' in line and 'ISTD' in line and '最终浓度单位' in line:
                table_start = i
                break
            # 外标法表头
            elif '名称' in line and '浓度' in line:
                table_start = i
                break

        if table_start == -1:
            return compounds, istd_info

        # 提取内标信息（仅对内标法）
        if 'ISTD' in content:
            for i in range(max(0, table_start - 5), min(table_start + 5, len(lines))):
                line = lines[i]
                if 'ISTD浓度单位' in line:
                    # 提取内标浓度
                    conc_match = re.search(r'(\d+\.\d+)\s*(mg/L|μg/mL|ng/ml)', line, re.IGNORECASE)
                    if conc_match:
                        istd_info['concentration'] = f"{conc_match.group(1)} {conc_match.group(2)}"

        # 提取化合物数据
        for i in range(table_start + 1, min(table_start + 50, len(lines))):
            line = lines[i].strip()
            if not line or '生成时间' in line or '样品色谱图' in line:
                break

            # 尝试解析化合物行
            # 内标法格式: 化合物名称 ISTD名称 ISTD保留时间 ISTD离子对 ISTD响应 ISTD浓度单位 保留时间 离子对 响应 最终浓度
            # 外标法格式: 化合物名称 保留时间 离子对 响应 浓度
            parts = re.split(r'\s+', line)
            if len(parts) >= 3:
                compound = parts[0]

                # 查找浓度值（通常在最后一列或倒数第二列）
                concentration = "N.D."
                for part in parts[-3:]:  # 检查最后三列
                    if 'N.D.' in part:
                        concentration = "N.D."
                        break
                    else:
                        conc_match = re.search(r'(\d+\.\d+)\s*(mg/L|μg/mL|ng/ml)', part, re.IGNORECASE)
                        if conc_match:
                            concentration = f"{conc_match.group(1)} {conc_match.group(2)}"
                            break

                if self.is_valid_compound(compound):
                    compounds[compound] = concentration

        return compounds, istd_info

    def extract_from_external_standard(self, content: str) -> Dict[str, str]:
        """专门处理外标法报告的化合物提取"""
        compounds = {}
        lines = content.split('\n')

        # 查找外标法表格区域
        table_start = -1
        for i, line in enumerate(lines):
            if '名称' in line and 'RT' in line and '响应' in line and '离子对' in line and '最终浓度单位' in line:
                table_start = i
                break

        if table_start == -1:
            # 如果没有找到完整表头，尝试简化表头
            for i, line in enumerate(lines):
                if '名称' in line and '最终浓度单位' in line:
                    table_start = i
                    break

        if table_start == -1:
            return compounds

        # 从表格区域提取化合物
        for i in range(table_start + 1, min(table_start + 30, len(lines))):
            line = lines[i].strip()
            if not line or any(end_marker in line for end_marker in ['生成时间', '样品色谱图', '=====']):
                break

            # 尝试分割行数据
            parts = re.split(r'\s{2,}', line)  # 按多个空格分割
            if len(parts) < 5:  # 如果分割后列数不够，尝试按单个空格分割
                parts = line.split()

            if len(parts) >= 5:
                # 第一列是化合物名称，最后一列是浓度
                compound = parts[0]
                concentration = parts[-1]

                # 处理特殊情况：如果化合物名称包含逗号，可能需要合并
                if ',' in compound and len(parts) > 5:
                    # 尝试合并前几列作为化合物名称
                    compound = ' '.join(parts[0:2])
                    concentration = parts[-1]

                # 处理中文化合物名称
                if self.is_valid_compound(compound):
                    # 标准化浓度格式
                    if 'N.D.' in concentration:
                        compounds[compound] = "N.D."
                    else:
                        conc_match = re.search(r'([\d.]+)\s*(mg/L|μg/mL|ng/ml)', concentration, re.IGNORECASE)
                        if conc_match:
                            compounds[compound] = f"{conc_match.group(1)} {conc_match.group(2)}"

        return compounds

    def clean_compounds(self, compounds: Dict[str, str]) -> Dict[str, str]:
        """清理化合物字典，移除无效的化合物名称"""
        cleaned = {}
        for compound, concentration in compounds.items():
            # 排除反向文本和其他无效名称
            if self.is_valid_compound(compound) and not self.is_excluded_compound(compound):
                cleaned[compound] = concentration
        return cleaned

    def is_excluded_compound(self, text: str) -> bool:
        """检查是否是应该排除的化合物名称"""
        # 排除列表 - 包含常见无效名称和反向文本
        exclude_list = [
            'ISTD', 'RT', '响应', '离子对', '面积', 'Counts', 'min',
            '样品', '分析', '报告', '名称', '浓度', '最终浓度', '数据', '文件',
            '采集方法', '样品瓶', '其他', '稀释', '体积', '类型', '批处理',
            'stnuoC', 'niM', 'x10', 'x10^', 'TIC', 'EIC', 'Scan', 'BB',
            '比值', '生成时间', '样品色谱图', 'Selected', 'Ion', 'SIM'
        ]

        # 检查是否是反向文本
        reversed_text = text[::-1]
        if any(exclude in reversed_text for exclude in ['Counts', 'Min', 'RT']):
            return True

        if any(exclude in text for exclude in exclude_list):
            return True

        return False

    # 原有的解析方法保持不变，只添加file_path参数
    def parse_compact_format(self, lines, filename: str, file_path: str = None):
        """解析紧凑格式的报告（如PAE报告）"""
        compounds = {}
        debug_info = f"  尝试紧凑格式解析...\n"

        # 查找目标化合物区域
        in_target_section = False
        compound_count = 0
        parsing_finished = False

        for i, line in enumerate(lines):
            if parsing_finished:
                break

            line = line.strip()

            # 检测目标化合物区域
            if '目标化合物' in line or 'Target Compounds' in line:
                in_target_section = True
                debug_info += f"    第 {i + 1} 行: 进入目标化合物区域\n"
                continue

            # 检查报告解析结束标志
            if '----------' in line or '==========' in line or '定性峰超出范围' in line:
                if in_target_section:
                    debug_info += f"    第 {i + 1} 行: 遇到结束标志，停止解析\n"
                    parsing_finished = True
                    break
                continue

            # 跳过空行
            if not line:
                continue

            # 跳过注释行和其他非化合物行
            if self.should_skip_line(line):
                debug_info += f"    第 {i + 1} 行: 跳过非化合物行\n"
                continue

            # 只在目标化合物区域解析化合物行
            if in_target_section:
                # 检查是否是化合物行（包含数字和可能的浓度值）
                if re.search(r'\d+\)\s+', line) and (
                        'N.D.' in line or 'mg/L' in line or '未校正' in line or re.search(r'\d+\.\d+\s+mg/L', line)):
                    compound_name, concentration, parse_debug = self.parse_compact_compound_line(line)
                    debug_info += f"    第 {i + 1} 行: {parse_debug}\n"
                    if compound_name and self.is_valid_compound(compound_name) and compound_name not in compounds:
                        compounds[compound_name] = concentration
                        compound_count += 1

        debug_info += f"    解析了 {compound_count} 个化合物\n"
        return compounds, debug_info

    def should_skip_line(self, line: str) -> bool:
        """判断是否应该跳过该行（非化合物行）"""
        skip_patterns = [
            r'\(#\)', r'定性峰', r'手动积分', r'已求和的信号',
            r'PAES-CAL', r'SVHC-CAL', r'SCCP-MCCP-CAL', r'PCN-CAL', r'Fri', r'Mon', r'Tue', r'Wed', r'Thu', r'Sat',
            r'Sun',
            r'^\d{4}-\d{2}-\d{2}', r'\d{2}:\d{2}:\d{2}', r'^\s*$', r'^---', r'^===',
            r'\.M\s+\d', r'周一', r'周二', r'周三', r'周四', r'周五', r'周六', r'周日',
            r'内标', r'ISTD', r'Internal Standard'  # 跳过内标相关行
        ]

        for pattern in skip_patterns:
            if re.search(pattern, line, re.IGNORECASE):
                return True

        # 检查是否包含日期时间格式
        if re.search(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}', line) or re.search(r'\d{1,2}:\d{2}:\d{2}', line):
            return True

        return False

    def parse_compact_compound_line(self, line: str) -> Tuple[str, float, str]:
        """解析紧凑格式报告中的化合物行（如PAE报告）"""
        debug_info = f"解析紧凑格式行: {line[:50]}..."

        # 提取化合物名称
        compound_name = self.extract_compound_name_from_compact_line(line)
        if not compound_name:
            debug_info += " -> 无法提取化合物名称"
            return "", 0.0, debug_info

        # 使用新的浓度解析方法
        concentration, _, status = self.parse_concentration(line)

        if status in ['未检出', '未校正']:
            debug_info += f" -> {status}: {compound_name}"
            return compound_name, 0.0, debug_info

        debug_info += f" -> {compound_name}: {concentration}"
        return compound_name, concentration, debug_info

    def extract_compound_name_from_compact_line(self, line: str) -> str:
        """从紧凑格式行中提取化合物名称"""
        # 移除行首的数字和括号
        cleaned = re.sub(r'^\d+\)\s*', '', line)  # 移除 "1) " 格式

        # 特别处理希腊字母α的情况
        # 匹配包含希腊字母α的化合物名称
        greek_pattern = r'^([αβγδεζηθικλμνξοπρστυφχψω\u4e00-\u9fffA-Za-z0-9\-\u4e00-\u9fff,\.\'\(\)\）\（\[\]_/]+)'
        match = re.match(greek_pattern, cleaned)
        if match:
            compound_name = match.group(1).strip()
            # 进一步清理：确保名称合理
            if len(compound_name) >= 2 and not compound_name.replace('-', '').replace(',', '').replace('.', '').replace(
                    '\'', '').replace('(', '').replace(')', '').replace('（', '').replace('）', '').replace('[',
                                                                                                          '').replace(
                ']', '').replace('α', '').replace('β', '').replace('γ', '').replace('δ', '').replace('ε', '').replace(
                'ζ', '').replace('η', '').replace('θ', '').replace('ι', '').replace('κ', '').replace('λ', '').replace(
                'μ', '').replace('ν', '').replace('ξ', '').replace('ο', '').replace('π', '').replace('ρ', '').replace(
                'σ', '').replace('τ', '').replace('υ', '').replace('φ', '').replace('χ', '').replace('ψ', '').replace(
                'ω', '').isdigit():
                return compound_name

        # 如果上述方法失败，使用通用清理方法
        return self.clean_compound_name(cleaned)

    def parse_standard_quant_report(self, lines, filename: str, file_path: str = None):
        """解析标准定量报告格式（如TS25101323001-氯苯-检出.pdf）"""
        compounds = {}
        debug_info = f"  尝试标准定量报告解析...\n"

        # 查找目标化合物区域
        in_target_section = False
        compound_count = 0
        parsing_finished = False

        for i, line in enumerate(lines):
            if parsing_finished:
                break

            line = line.strip()

            # 检测目标化合物区域
            if '目标化合物' in line or 'Target Compounds' in line:
                in_target_section = True
                debug_info += f"    第 {i + 1} 行: 进入目标化合物区域\n"
                continue

            # 检查报告解析结束标志
            if '----------' in line or '==========' in line or '定性峰超出范围' in line:
                if in_target_section:
                    debug_info += f"    第 {i + 1} 行: 遇到结束标志，停止解析\n"
                    parsing_finished = True
                    break
                continue

            # 跳过空行
            if not line:
                continue

            # 跳过注释行和其他非化合物行
            if self.should_skip_line(line):
                debug_info += f"    第 {i + 1} 行: 跳过非化合物行\n"
                continue

            # 只在目标化合物区域解析化合物行
            if in_target_section:
                # 检查是否是化合物行（包含数字和单位）
                if re.search(r'\d+\.\d+\s+mg/L', line) or 'N.D.' in line or '0.000 mg/L' in line or re.search(
                        r'\d+\)\s+[\wα\-\[\]]', line):
                    compound_name, concentration, parse_debug = self.parse_quant_compound_line(line)
                    debug_info += f"    第 {i + 1} 行: {parse_debug}\n"
                    if compound_name and self.is_valid_compound(compound_name) and compound_name not in compounds:
                        compounds[compound_name] = concentration
                        compound_count += 1

        debug_info += f"    解析了 {compound_count} 个化合物\n"
        return compounds, debug_info

    def parse_quant_compound_line(self, line: str) -> Tuple[str, float, str]:
        """解析定量报告中的化合物行"""
        debug_info = f"解析: {line[:50]}..."

        # 提取化合物名称
        compound_name = self.extract_compound_name_from_quant_line(line)
        if not compound_name:
            debug_info += " -> 无法提取化合物名称"
            return "", 0.0, debug_info

        # 使用新的浓度解析方法
        concentration, _, status = self.parse_concentration(line)

        if status in ['未检出', '未校正']:
            debug_info += f" -> {status}: {compound_name}"
            return compound_name, 0.0, debug_info

        debug_info += f" -> {compound_name}: {concentration}"
        return compound_name, concentration, debug_info

    def extract_compound_name_from_quant_line(self, line: str) -> str:
        """从定量报告行中提取化合物名称"""
        # 移除行首的数字和括号
        cleaned = re.sub(r'^\d+\)\s*', '', line)  # 移除 "1) " 格式

        # 特别处理希腊字母α的情况
        # 匹配包含希腊字母α的化合物名称
        greek_pattern = r'^([αβγδεζηθικλμνξοπρστυφχψω\u4e00-\u9fffA-Za-z0-9\-\u4e00-\u9fff,\.\'\(\)\）\（\[\]_/]+)'
        match = re.match(greek_pattern, cleaned)
        if match:
            compound_name = match.group(1).strip()
            # 进一步清理：确保名称合理
            if len(compound_name) >= 2 and not compound_name.replace('-', '').replace(',', '').replace('.', '').replace(
                    '\'', '').replace('(', '').replace(')', '').replace('（', '').replace('）', '').replace('[',
                                                                                                          '').replace(
                ']', '').replace('α', '').replace('β', '').replace('γ', '').replace('δ', '').replace('ε', '').replace(
                'ζ', '').replace('η', '').replace('θ', '').replace('ι', '').replace('κ', '').replace('λ', '').replace(
                'μ', '').replace('ν', '').replace('ξ', '').replace('ο', '').replace('π', '').replace('ρ', '').replace(
                'σ', '').replace('τ', '').replace('υ', '').replace('φ', '').replace('χ', '').replace('ψ', '').replace(
                'ω', '').isdigit():
                return compound_name

        # 如果上述方法失败，使用通用清理方法
        return self.clean_compound_name(cleaned)

    def parse_table_format(self, lines, filename: str, file_path: str = None):
        """解析表格格式的报告（如TS25110765001A-PAHS-检出.pdf）"""
        compounds = {}
        debug_info = f"  尝试表格格式解析...\n"

        # 查找表格开始
        in_table = False
        header_found = False
        compound_count = 0
        in_target_section = False
        parsing_finished = False

        for i, line in enumerate(lines):
            if parsing_finished:
                break

            line = line.strip()

            # 检测目标化合物区域
            if '目标化合物' in line or 'Target Compounds' in line:
                in_target_section = True
                debug_info += f"    第 {i + 1} 行: 进入目标化合物区域\n"
                continue

            # 检查报告解析结束标志
            if '----------' in line or '==========' in line or '定性峰超出范围' in line:
                if in_target_section:
                    debug_info += f"    第 {i + 1} 行: 遇到结束标志，停止解析\n"
                    parsing_finished = True
                    break
                continue

            # 跳过注释行和其他非化合物行
            if self.should_skip_line(line):
                debug_info += f"    第 {i + 1} 行: 跳过非化合物行\n"
                continue

            # 检测表格开始
            if in_target_section and '| 化合物' in line and '保留时间' in line:
                in_table = True
                header_found = True
                debug_info += f"    第 {i + 1} 行: 进入表格区域\n"
                continue

            # 解析表格行 - 只在目标化合物区域解析
            if in_target_section and in_table and line and '|' in line:
                # 跳过表头行
                if header_found:
                    header_found = False
                    continue

                # 解析表格行
                compound_name, concentration, parse_debug = self.parse_table_row(line)
                debug_info += f"    第 {i + 1} 行: {parse_debug}\n"
                if compound_name and self.is_valid_compound(compound_name) and compound_name not in compounds:
                    compounds[compound_name] = concentration
                    compound_count += 1

        debug_info += f"    解析了 {compound_count} 个化合物\n"
        return compounds, debug_info

    def parse_table_row(self, line: str) -> Tuple[str, float, str]:
        """解析表格行"""
        debug_info = f"解析表格行: {line[:50]}..."

        # 提取化合物名称
        compound_name = self.extract_compound_name_from_table(line)
        if not compound_name:
            debug_info += " -> 无法提取化合物名称"
            return "", 0.0, debug_info

        # 使用新的浓度解析方法
        concentration, _, status = self.parse_concentration(line)

        if status in ['未检出', '未校正']:
            debug_info += f" -> {status}: {compound_name}"
            return compound_name, 0.0, debug_info

        debug_info += f" -> {compound_name}: {concentration}"
        return compound_name, concentration, debug_info

    def extract_compound_name_from_table(self, line: str) -> str:
        """从表格行中提取化合物名称"""
        # 分割表格列
        parts = [part.strip() for part in line.split('|') if part.strip()]

        if len(parts) >= 1:
            # 第一列是化合物名称，可能包含数字和括号
            name_part = parts[0]
            # 使用新的化合物名称清理方法
            return self.clean_compound_name(name_part)

        return ""

    def parse_standard_format(self, lines, filename: str, file_path: str = None):
        """解析标准格式的报告"""
        compounds = {}
        debug_info = f"  尝试标准格式解析...\n"

        # 查找目标化合物区域
        in_target_section = False
        compound_count = 0
        parsing_finished = False

        for i, line in enumerate(lines):
            if parsing_finished:
                break

            line = line.strip()

            # 检测目标化合物区域
            if '目标化合物' in line or 'Target Compounds' in line:
                in_target_section = True
                debug_info += f"    第 {i + 1} 行: 进入目标化合物区域\n"
                continue

            # 检查报告解析结束标志
            if '----------' in line or '==========' in line or '定性峰超出范围' in line:
                if in_target_section:
                    debug_info += f"    第 {i + 1} 行: 遇到结束标志，停止解析\n"
                    parsing_finished = True
                    break
                continue

            # 跳过表格分隔线
            if '---' in line or '===' in line or not line:
                continue

            # 跳过注释行和其他非化合物行
            if self.should_skip_line(line):
                debug_info += f"    第 {i + 1} 行: 跳过非化合物行\n"
                continue

            # 只在目标化合物区域解析化合物行
            if in_target_section:
                compound_name, concentration, parse_debug = self.parse_standard_line(line)
                debug_info += f"    第 {i + 1} 行: {parse_debug}\n"
                if compound_name and self.is_valid_compound(compound_name) and compound_name not in compounds:
                    compounds[compound_name] = concentration
                    compound_count += 1

        debug_info += f"    解析了 {compound_count} 个化合物\n"
        return compounds, debug_info

    def parse_standard_line(self, line: str) -> Tuple[str, float, str]:
        """解析标准格式的化合物行"""
        debug_info = f"解析标准行: {line[:50]}..."

        # 提取化合物名称
        compound_name = self.extract_compound_name_from_standard_line(line)
        if not compound_name:
            debug_info += " -> 无法提取化合物名称"
            return "", 0.0, debug_info

        # 使用新的浓度解析方法
        concentration, _, status = self.parse_concentration(line)

        if status in ['未检出', '未校正']:
            debug_info += f" -> {status}: {compound_name}"
            return compound_name, 0.0, debug_info

        debug_info += f" -> {compound_name}: {concentration}"
        return compound_name, concentration, debug_info

    def extract_compound_name_from_standard_line(self, line: str) -> str:
        """从标准格式行中提取化合物名称"""
        # 移除行首的数字和括号
        cleaned = re.sub(r'^\d+\)\s*', '', line)  # 移除 "1) " 格式

        # 特别处理希腊字母α的情况
        # 匹配包含希腊字母α的化合物名称
        greek_pattern = r'^([αβγδεζηθικλμνξοπρστυφχψω\u4e00-\u9fffA-Za-z0-9\-\u4e00-\u9fff,\.\'\(\)\）\（\[\]_/]+)'
        match = re.match(greek_pattern, cleaned)
        if match:
            compound_name = match.group(1).strip()
            # 进一步清理：确保名称合理
            if len(compound_name) >= 2 and not compound_name.replace('-', '').replace(',', '').replace('.', '').replace(
                    '\'', '').replace('(', '').replace(')', '').replace('（', '').replace('）', '').replace('[',
                                                                                                          '').replace(
                ']', '').replace('α', '').replace('β', '').replace('γ', '').replace('δ', '').replace('ε', '').replace(
                'ζ', '').replace('η', '').replace('θ', '').replace('ι', '').replace('κ', '').replace('λ', '').replace(
                'μ', '').replace('ν', '').replace('ξ', '').replace('ο', '').replace('π', '').replace('ρ', '').replace(
                'σ', '').replace('τ', '').replace('υ', '').replace('φ', '').replace('χ', '').replace('ψ', '').replace(
                'ω', '').isdigit():
                return compound_name

        # 如果上述方法失败，使用通用清理方法
        return self.clean_compound_name(cleaned)

    def parse_concentration(self, concentration_str: str) -> Tuple[float, str, str]:
        """
        解析浓度字符串，返回浓度值、单位和状态
        """
        if not concentration_str or concentration_str.strip() == '':
            return 0.0, '', '未检测'

        concentration_str = str(concentration_str).strip()

        # 检查是否为未检出
        if self.nd_pattern.search(concentration_str):
            return 0.0, '', '未检出'

        # 检查是否为未校正
        if self.uncorrected_pattern.search(concentration_str):
            return 0.0, '', '未校正'

        # 提取浓度值和单位
        # 首先尝试匹配小数+单位
        match = self.concentration_pattern.search(concentration_str)
        if match:
            value = float(match.group(1))
            unit = match.group(2)
            # 放宽浓度值范围限制（从0.00001到10000）
            if 0.00001 <= value <= 10000:
                return value, unit, '检出'

        # 特殊处理：如果文本被错误分割（如"0.3 6 mg/L"）
        parts = concentration_str.split()
        for i in range(len(parts) - 2):
            # 检查是否可能是被分割的小数
            # 模式：数字+点+空格+数字+空格+单位
            if ('.' in parts[i] and
                    parts[i].replace('.', '').replace(' ', '').isdigit() and
                    parts[i + 1].isdigit() and
                    parts[i + 2].lower() in ['mg/l', 'ppm', 'μg/ml', 'ng/ml']):
                try:
                    # 合并数字部分，去掉中间的空格
                    value_str = parts[i] + parts[i + 1]  # "0.3" + "6" = "0.36"
                    value = float(value_str)
                    unit = parts[i + 2]
                    # 放宽浓度值范围限制
                    if 0.00001 <= value <= 10000:
                        return value, unit, '检出'
                except ValueError:
                    continue

        # 尝试匹配整数+单位（只有在没有小数的情况下）
        match = self.concentration_pattern_int.search(concentration_str)
        if match:
            value = float(match.group(1))
            unit = match.group(2)
            # 放宽浓度值范围限制
            if 0.00001 <= value <= 10000:
                return value, unit, '检出'

        return 0.0, '', '未检出'

    def clean_compound_name(self, name: str) -> str:
        """
        清理化合物名称，移除编号和多余空格
        """
        # 移除类似 "1) "、"17) " 这样的编号
        name = re.sub(r'^\d+\)\s*', '', name)
        # 移除类似 "107) " 这样的编号
        name = re.sub(r'^\d+\)', '', name)
        # 移除保留时间等数字前缀
        name = re.sub(r'^\d+\.\d+\s+', '', name)
        # 移除多余空格
        name = re.sub(r'\s+', ' ', name).strip()

        # 进一步清理：移除末尾的点和空格
        name = re.sub(r'[.\s]+$', '', name)

        # 提取有效的化合物名称部分（包括希腊字母）
        greek_pattern = r'^([αβγδεζηθικλμνξοπρστυφχψω\u4e00-\u9fffA-Za-z0-9\-\u4e00-\u9fff,\.\'\(\)\）\（\[\]_/]+)'
        match = re.match(greek_pattern, name)
        if match:
            return match.group(1).strip()

        return name

    def is_valid_compound(self, compound_name: str) -> bool:
        """检查是否为有效的化合物名称"""
        if not compound_name or len(compound_name) < 2:
            return False

        # 检查是否符合化合物名称模式（包括希腊字母）
        greek_pattern = r'^[αβγδεζηθικλμνξοπρστυφχψω\u4e00-\u9fffA-Za-z0-9\-\u4e00-\u9fff,\.\'\(\)\）\（\[\]_/]+$'
        if re.match(greek_pattern, compound_name):
            return True

        # 检查是否为中文化合物名称
        if re.match(r'^[\u4e00-\u9fff]+$', compound_name):
            return True

        # 允许包含数字和特殊字符的化合物名称（包括希腊字母）
        extended_pattern = r'^[αβγδεζηθικλμνξοπρστυφχψω\u4e00-\u9fffA-Za-z][αβγδεζηθικλμνξοπρστυφχψω\u4e00-\u9fffA-Za-z0-9\-\u4e00-\u9fff,\.\'\(\)\）\（\[\]_/]*$'
        if re.match(extended_pattern, compound_name):
            return True

        # MH的化合物名称模式
        compound_patterns = [
            r'^[A-Z]{2,}-\d+$',  # 如 DIBP-223, DEHP-279
            r'^[A-Z]{3,}$',  # 如 MOCA, TSCA
            r'^[A-Z]{2,}\d*$',  # 如 BB, MDCA
            r'^[A-Z]+\(\d+:\d+\)$',  # 如 PIP(3:1)
            r'^[A-Z]+-\d+[A-Z]*$',  # 如 BBP-206
            r'^[\u4e00-\u9fff]+$',  # 纯中文化合物名称
            r'^[\u4e00-\u9fff][\u4e00-\u9fff\w\-\(\):,]+$',  # 中文混合，包含逗号
            r'^[\d,]+-[\u4e00-\u9fff]+$',  # 如 2,4,6-三叔丁基苯酚
            r'^[αβγδεζηθικλμνξοπρστυφχψω]+.*$',  # 以希腊字母开头的化合物
        ]

        for pattern in compound_patterns:
            if re.match(pattern, compound_name):
                return True

        return False

    def is_marker(self, compound_name: str) -> bool:
        """判断是否为标记物"""
        if not self.marker_var.get():
            return False
        return any(marker.lower() in compound_name.lower() for marker in self.markers)

    def display_results(self, detected: List[str], not_detected: List[str], details: Dict, debug_info: str,
                        threshold: float):
        """显示分析结果"""
        # 更新摘要标签页
        self.summary_text.delete(1.0, tk.END)

        summary = f"=== 分析结果 (阈值: {threshold}) ===\n\n"

        # 显示标记物排除状态
        if self.marker_var.get() and self.markers:
            summary += f"排除的标记物: {', '.join(self.markers)}\n\n"

        summary += f"✅ 阳性报告 ({len(detected)}):\n"
        for report in detected:
            detected_count = len(details[report]['detected_compounds'])
            summary += f"  ✓ {report} (检测到 {detected_count} 个化合物)\n"

        summary += f"\n❌ 阴性报告 ({len(not_detected)}):\n"
        for report in not_detected:
            summary += f"  ✗ {report}\n"

        self.summary_text.insert(1.0, summary)

        # 更新详细结果标签页
        self.detail_text.delete(1.0, tk.END)

        detail_content = f"=== 详细分析结果 ===\n\n"

        for filename, result in details.items():
            status = "✅ 阳性" if result['has_detection'] else "❌ 阴性"
            detail_content += f"📄 文件: {filename} [{status}]\n"

            if result['detected_compounds']:
                detail_content += "   检测到的化合物:\n"
                for name, conc in result['detected_compounds'].items():
                    detail_content += f"     ✓ {name}: {conc}\n"
            else:
                detail_content += "   未检测到超过阈值的化合物\n"

            detail_content += "\n"

        self.detail_text.insert(1.0, detail_content)

        # 更新化合物列表标签页
        self.compound_text.delete(1.0, tk.END)

        compound_content = "=== 所有检测到的化合物 ===\n\n"

        all_compounds = {}
        for filename, result in details.items():
            for name, conc in result['all_compounds'].items():
                if name not in all_compounds:
                    all_compounds[name] = []
                all_compounds[name].append((filename, conc))

        # 按化合物名称排序
        for compound in sorted(all_compounds.keys()):
            compound_content += f"🔬 {compound}:\n"
            for filename, conc in all_compounds[compound]:
                status = "标记物" if self.is_marker(compound) else "目标物"
                compound_content += f"   {filename}: {conc} [{status}]\n"
            compound_content += "\n"

        self.compound_text.insert(1.0, compound_content)

        # 更新调试信息标签页
        self.debug_text.delete(1.0, tk.END)
        self.debug_text.insert(1.0, debug_info)

        self.status_var.set(f"分析完成 - 阳性: {len(detected)}, 阴性: {len(not_detected)}")

    def run(self):
        """运行应用程序"""
        self.root.mainloop()


def _is_valid_compound_standalone(compound_name: str) -> bool:
    if not compound_name or len(compound_name) < 2:
        return False
    greek_pattern = r'^[αβγδεζηθικλμνξοπρστυφχψω一-鿿A-Za-z0-9\-一-鿿,\.\'\(\)\）\（\[\]_/]+$'
    if re.match(greek_pattern, compound_name):
        return True
    if re.match(r'^[一-鿿]+$', compound_name):
        return True
    extended_pattern = r'^[αβγδεζηθικλμνξοπρστυφχψω一-鿿A-Za-z][αβγδεζηθικλμνξοπρστυφχψω一-鿿A-Za-z0-9\-一-鿿,\.\'\(\)\）\（\[\]_/]*$'
    if re.match(extended_pattern, compound_name):
        return True
    return False


def _extract_from_tables_standalone(tables, content):
    """从PDF表格中提取化合物和浓度（独立于GUI类）"""
    compounds = {}
    for table in tables:
        if not table:
            continue
        header_row = -1
        for i, row in enumerate(table):
            if row and any(cell and any(keyword in str(cell) for keyword in ['名称', '化合物', 'ISTD', '浓度']) for cell in row):
                header_row = i
                break
        if header_row == -1:
            continue
        name_col = -1
        conc_col = -1
        for j, cell in enumerate(table[header_row]):
            if cell and any(keyword in str(cell) for keyword in ['名称', '化合物']):
                name_col = j
            elif cell and any(keyword in str(cell) for keyword in ['浓度', '最终浓度']):
                conc_col = j
        if conc_col == -1 and len(table[header_row]) > 1:
            conc_col = len(table[header_row]) - 1
        for i in range(header_row + 1, len(table)):
            row = table[i]
            if not row or len(row) <= max(name_col, conc_col):
                continue
            name_cell = row[name_col] if name_col < len(row) else None
            conc_cell = row[conc_col] if conc_col < len(row) else None
            if name_cell and _is_valid_compound_standalone(name_cell):
                concentration = "N.D."
                if conc_cell:
                    if 'N.D.' in str(conc_cell):
                        concentration = "N.D."
                    else:
                        conc_match = re.search(r'(\d+\.\d+)\s*(mg/L|μg/mL|ng/ml)', str(conc_cell), re.IGNORECASE)
                        if conc_match:
                            concentration = f"{conc_match.group(1)} {conc_match.group(2)}"
                compounds[name_cell] = concentration
    return compounds


def _extract_from_text_standalone(content):
    """从文本中提取化合物和浓度（独立于GUI类）"""
    compounds = {}
    lines = content.split('\n')
    table_start = -1
    for i, line in enumerate(lines):
        if '名称' in line and 'ISTD' in line and '最终浓度单位' in line:
            table_start = i
            break
        elif '名称' in line and '浓度' in line:
            table_start = i
            break
    if table_start == -1:
        return compounds
    for i in range(table_start + 1, min(table_start + 50, len(lines))):
        line = lines[i].strip()
        if not line or '生成时间' in line or '样品色谱图' in line:
            break
        parts = re.split(r'\s+', line)
        if len(parts) >= 3:
            compound = parts[0]
            concentration = "N.D."
            for part in parts[-3:]:
                if 'N.D.' in part:
                    concentration = "N.D."
                    break
                else:
                    conc_match = re.search(r'(\d+\.\d+)\s*(mg/L|μg/mL|ng/ml)', part, re.IGNORECASE)
                    if conc_match:
                        concentration = f"{conc_match.group(1)} {conc_match.group(2)}"
                        break
            if _is_valid_compound_standalone(compound):
                compounds[compound] = concentration
    return compounds


def _extract_from_external_standard_standalone(content):
    """处理外标法报告（独立于GUI类）"""
    compounds = {}
    lines = content.split('\n')
    table_start = -1
    for i, line in enumerate(lines):
        if '名称' in line and 'RT' in line and '响应' in line and '离子对' in line and '最终浓度单位' in line:
            table_start = i
            break
    if table_start == -1:
        for i, line in enumerate(lines):
            if '名称' in line and '最终浓度单位' in line:
                table_start = i
                break
    if table_start == -1:
        return compounds
    for i in range(table_start + 1, min(table_start + 30, len(lines))):
        line = lines[i].strip()
        if not line or any(m in line for m in ['生成时间', '样品色谱图', '=====']):
            break
        parts = re.split(r'\s{2,}', line)
        if len(parts) < 5:
            parts = line.split()
        if len(parts) >= 5:
            compound = parts[0]
            concentration = parts[-1]
            if ',' in compound and len(parts) > 5:
                compound = ' '.join(parts[0:2])
                concentration = parts[-1]
            if _is_valid_compound_standalone(compound):
                if 'N.D.' in concentration:
                    compounds[compound] = "N.D."
                else:
                    conc_match = re.search(r'([\d.]+)\s*(mg/L|μg/mL|ng/ml)', concentration, re.IGNORECASE)
                    if conc_match:
                        compounds[compound] = f"{conc_match.group(1)} {conc_match.group(2)}"
    return compounds


def parse_pdf_report(file_path: str) -> dict:
    """解析PDF定量报告，返回 {compound_name: (measured_value, unit)}。

    独立于GUI，可在Flask端点中直接调用。
    依次尝试：表格提取 → 文本提取 → 外标法提取。
    """
    extractor = PDFTextExtractor()
    text = extractor.extract_text_from_pdf(file_path)
    tables = extractor.extract_tables_from_pdf(file_path)

    compounds = {}

    # 1. 表格提取
    if tables:
        compounds.update(_extract_from_tables_standalone(tables, text))

    # 2. 文本提取
    if not compounds:
        compounds.update(_extract_from_text_standalone(text))

    # 3. 外标法
    if not compounds and '外标法' in text:
        compounds.update(_extract_from_external_standard_standalone(text))

    # 转为 (value, unit) 字典
    result = {}
    for name, conc_str in compounds.items():
        if conc_str == "N.D.":
            continue
        conc_match = re.search(r'([\d.]+)\s*(mg/L|μg/mL|ng/ml|ppm)', conc_str, re.IGNORECASE)
        if conc_match:
            try:
                result[name] = (float(conc_match.group(1)), conc_match.group(2))
            except ValueError:
                pass

    return result


def parse_epatemp_txt(file_path: str) -> dict:
    """解析 epatemp.txt 定量报告，返回 {compound_name: (measured_value, unit)}。"""
    import os
    if not os.path.isfile(file_path):
        return {}
    # 匹配目标化合物行: 编号) 名称 保留时间 定量离子 响应值  浓度 单位 [偏差]
    # 响应值可能是纯数字或带m后缀(手动积分)，浓度单位可能是ng/mg/L等
    row_pattern = re.compile(
        r'^\s*\d+\)\s+(.+?)\s{2,}'
        r'[\d.]+\s+\d+\s+\d+[m]?\s+([\d.]+)\s+(ng|mg/L|μg/mL|ug/mL)'
    )
    for encoding in ('gbk', 'gb18030', 'gb2312', 'utf-8', 'latin-1'):
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                content = f.read()
            if '�' in content:
                continue
            result = {}
            extracting = False
            for line in content.splitlines():
                if '目标化合物' in line:
                    extracting = True
                    continue
                if not extracting:
                    continue
                m = row_pattern.match(line)
                if m:
                    name = m.group(1).strip()
                    if name:
                        result[name] = (float(m.group(2)), m.group(3))
            if result:
                if encoding != 'latin-1':
                    return result
                if encoding == 'latin-1':
                    return result
        except (UnicodeDecodeError, UnicodeError):
            continue
    return {}


# 运行应用程序
    try:
        from ctypes import windll

        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass

    app = ChemicalReportAnalyzer()
    app.run()