import os
import chardet
import logging
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import black, blue
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph
from PyPDF2 import PdfReader, PdfWriter
import tkinter as tk
from tkinter import messagebox
import re

# 设置日志
logger = logging.getLogger('EpatempToPdfConverter')


class EpatempToPdfConverter:
    def __init__(self):
        self.chinese_font = self.register_chinese_font_pdf()
        self.unreviewed_files = []  # 存储未复核的文件信息
        self.extracted_info = {}  # 存储提取的信息

    def show_error_message(self, title, message):
        """显示错误弹窗"""
        try:
            # 创建隐藏的根窗口
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(title, message)
            root.destroy()
        except Exception:
            pass

    def get_epatemp_modify_time(self, epatemp_path):
        """获取epatemp.txt文件的修改时间"""
        try:
            if os.path.exists(epatemp_path):
                return datetime.fromtimestamp(os.path.getmtime(epatemp_path))
            return None
        except Exception:
            return None

    def parse_audit_time(self, time_str):
        """解析audit.txt中的时间格式"""
        try:
            # 尝试解析格式: "Fri Sep 12 10:00:28 2025"
            return datetime.strptime(time_str, "%a %b %d %H:%M:%S %Y")
        except ValueError:
            try:
                # 尝试其他可能的时间格式
                return datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")
            except ValueError:
                return None

    def show_warning_message(self, title, message):
        """显示警告弹窗"""
        try:
            # 创建隐藏的根窗口
            root = tk.Tk()
            root.withdraw()
            messagebox.showwarning(title, message)
            root.destroy()
        except Exception:
            pass

    def show_info_message(self, title, message):
        """显示信息弹窗"""
        try:
            # 创建隐藏的根窗口
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo(title, message)
            root.destroy()
        except Exception:
            pass

    def show_unreviewed_summary(self):
        """显示未复核文件的汇总弹窗"""
        if not self.unreviewed_files:
            return

        # 创建汇总消息
        if len(self.unreviewed_files) == 1:
            message = f"以下文件未复核:\n{self.unreviewed_files[0]}"
        else:
            message = f"以下 {len(self.unreviewed_files)} 个文件未复核:\n" + "\n".join(self.unreviewed_files)

        self.show_error_message("未复核文件", message)

        # 清空未复核文件列表
        self.unreviewed_files = []

    def extract_file_info(self, content, file_path=None):
        """
        从文件内容中提取"数据文件"和"样品"信息

        Args:
            content (str): 文件内容
            file_path (str): 文件路径

        Returns:
            dict: 包含提取信息的字典
        """
        info = {
            'data_file': None,
            'sample': None,
            'file_path': file_path,
            'extraction_success': False
        }

        try:
            lines = content.split('\n')

            for line in lines:
                line = line.strip()

                # 提取数据文件信息
                if "数据文件" in line and info['data_file'] is None:
                    # 处理多种格式
                    patterns = [
                        r'数据文件\s*[：:]\s*(.+)',
                        r'数据文件\s*(.+)',
                    ]

                    for pattern in patterns:
                        match = re.search(pattern, line)
                        if match:
                            data_file = match.group(1).strip()
                            # 清理可能的多余空格
                            data_file = re.sub(r'\s+', ' ', data_file)
                            # 如果以.D结尾，去掉.D
                            if data_file.endswith('.D'):
                                data_file = data_file[:-2]
                            info['data_file'] = data_file
                            break

                    # 如果没有匹配到，尝试直接分割
                    if info['data_file'] is None and ':' in line:
                        parts = line.split(':', 1)
                        if len(parts) > 1:
                            data_file = parts[1].strip()
                            if data_file.endswith('.D'):
                                data_file = data_file[:-2]
                            info['data_file'] = data_file

                # 提取样品信息
                if "样品" in line and info['sample'] is None:
                    # 处理多种格式
                    patterns = [
                        r'样品\s*[：:]\s*(.+)',
                        r'样品\s*(.+)',
                    ]

                    for pattern in patterns:
                        match = re.search(pattern, line)
                        if match:
                            sample = match.group(1).strip()
                            sample = re.sub(r'\s+', ' ', sample)
                            info['sample'] = sample
                            break

                    # 如果没有匹配到，尝试直接分割
                    if info['sample'] is None and ':' in line:
                        parts = line.split(':', 1)
                        if len(parts) > 1:
                            info['sample'] = parts[1].strip()

                # 如果已经找到两个信息，提前退出循环
                if info['data_file'] is not None and info['sample'] is not None:
                    break

            # 检查是否成功提取到信息
            if info['data_file'] is not None and info['sample'] is not None:
                info['extraction_success'] = True
                logger.debug(f"成功提取信息: 数据文件={info['data_file']}, 样品={info['sample']}")
            else:
                logger.debug(f"信息提取不完整: 数据文件={info['data_file']}, 样品={info['sample']}")

        except Exception as e:
            logger.error(f"提取文件信息时出错: {e}")
            info['extraction_success'] = False

        return info

    def get_extracted_info(self, file_path=None):
        """
        获取提取的信息

        Args:
            file_path (str): 可选，指定要获取信息的文件路径

        Returns:
            dict: 提取的信息字典，如果指定file_path则返回该文件的信息，
                  否则返回最后处理的文件信息
        """
        if file_path:
            return self.extracted_info.get(file_path, {})
        else:
            # 返回最后处理的文件信息（如果存在）
            if self.extracted_info:
                return list(self.extracted_info.values())[-1]
            return {}

    def clear_extracted_info(self):
        """清空提取的信息"""
        self.extracted_info = {}

    def clean_folder_name(self, folder_name):
        """清理文件夹名称，移除扩展名"""
        extensions = ['.D', '.d', '.TXT', '.txt', '.DAT', '.dat']
        for ext in extensions:
            if folder_name.endswith(ext):
                return folder_name[:-len(ext)]
        return folder_name

    def detect_file_encoding(self, file_path):
        """文件编码检测 - 改进版本"""
        try:
            with open(file_path, 'rb') as f:
                raw_data = f.read(4096)  # 读取更多数据提高检测准确性

            # 使用chardet检测
            result = chardet.detect(raw_data)
            encoding = result['encoding']
            confidence = result['confidence']

            print(f"chardet检测结果: {encoding}, 置信度: {confidence}")

            if encoding is None or confidence < 0.6:
                # 低置信度时，尝试通过内容判断
                try:
                    # 检查是否包含常见中文字符
                    test_gbk = raw_data.decode('gbk', errors='ignore')
                    test_utf8 = raw_data.decode('utf-8', errors='ignore')

                    gbk_chinese_count = sum(1 for char in test_gbk if '\u4e00' <= char <= '\u9fff')
                    utf8_chinese_count = sum(1 for char in test_utf8 if '\u4e00' <= char <= '\u9fff')

                    if gbk_chinese_count > utf8_chinese_count:
                        return 'gbk'
                    elif utf8_chinese_count > 0:
                        return 'utf-8'
                except:
                    pass

                return 'gbk'  # 默认回退到GBK

            encoding = encoding.lower()

            # 处理常见的误检测情况
            if encoding in ['windows-1252', 'cp1252', 'iso-8859-1']:
                # 检查是否可能实际上是GBK编码
                try:
                    test_content = raw_data.decode('gbk', errors='strict')
                    chinese_chars = sum(1 for char in test_content if '\u4e00' <= char <= '\u9fff')
                    if chinese_chars > 0:  # 如果包含中文字符，很可能是GBK
                        return 'gbk'
                except:
                    pass
                return 'windows-1252'
            elif encoding in ['gb2312', 'gb18030']:
                return 'gbk'
            elif encoding in ['big5', 'big5-hkscs']:
                return 'big5'

            return encoding

        except Exception as e:
            print(f"编码检测异常: {e}")
            return 'gbk'

    def convert_encoding(self, content, from_encoding, to_encoding='gb2312'):
        """将内容从一种编码转换为另一种编码"""
        try:
            decoded_content = content.decode(from_encoding)
            return decoded_content.encode(to_encoding)
        except:
            return content

    def register_chinese_font_pdf(self):
        """注册中文字体，确保支持中文"""
        font_candidates = [
            ('SimSun', 'simsun.ttc'),
            ('NSimSun', 'simsun.ttc'),
            ('KaiTi', 'simkai.ttf'),
            ('Courier', 'cour.ttf'),
            ('Courier New', 'courbd.ttf'),
            ('STSong', 'STSONG.TTF'),
            ('STHeiti', 'STHeiti.ttf'),
            ('WenQuanYi Micro Hei', '/usr/share/fonts/wenquanyi/wqy-microhei/wqy-microhei.ttc')
        ]

        for font_name, font_path in font_candidates:
            try:
                win_font_path = os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'Fonts', font_path)
                if os.path.exists(win_font_path):
                    pdfmetrics.registerFont(TTFont(font_name, win_font_path))
                    return font_name
                else:
                    try:
                        pdfmetrics.registerFont(TTFont(font_name, font_path))
                        return font_name
                    except:
                        continue
            except:
                continue

        try:
            from reportlab.pdfbase.cidfonts import UnicodeCIDFont
            pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
            return 'STSong-Light'
        except:
            return 'Helvetica'

    def remove_blank_pages(self, pdf_path):
        """删除PDF中的空白页（仅检查最后一页）"""
        try:
            reader = PdfReader(pdf_path)

            if len(reader.pages) <= 1:
                return

            last_page = reader.pages[-1]
            text = last_page.extract_text()

            if not text or not text.strip():
                writer = PdfWriter()

                for i in range(len(reader.pages) - 1):
                    writer.add_page(reader.pages[i])

                with open(pdf_path, 'wb') as f:
                    writer.write(f)

        except Exception:
            pass

    def parse_time_from_line(self, line, patterns):
        """从行中解析时间，尝试多种模式"""
        for pattern, time_format in patterns:
            try:
                match = re.search(pattern, line)
                if match:
                    time_str = match.group(1)
                    # 如果是报告时间的第二种模式，需要特殊处理
                    if time_format == "REPORT_CN_WEEKDAY":
                        # 处理中文星期格式: "2025-09-12 周五 08:38:27"
                        date_part = match.group(1)  # 2025-09-12
                        time_part = match.group(2)  # 08:38:27
                        time_str = f"{date_part} {time_part}"
                        time_format = "%Y-%m-%d %H:%M:%S"

                    result = datetime.strptime(time_str, time_format)
                    return result
            except Exception:
                continue
        return None

    def validate_epatemp_content(self, content, file_path=None):
        """验证epatemp文件内容是否符合要求"""
        lines = content.split('\n')

        # 检查第一行是否包含"未检查"
        if len(lines) > 0:
            first_line = lines[0].strip()
            if "未检查" in first_line:
                # 记录未复核文件信息 - 只记录文件夹名称
                if file_path:
                    folder_name = os.path.basename(os.path.dirname(file_path))
                    self.unreviewed_files.append(folder_name)  # 只添加文件夹名称，不添加括号内容
                return False, "定量报告未复核"

        # 提取定量时间和最后一行报告生成时间
        quant_time = None
        report_time = None

        # 定量时间解析模式
        quant_patterns = [
            # 格式: "定量时间  ：Sep 12 08:35:12 2025"
            (r'定量时间\s*[：:]\s*([A-Za-z]{3}\s+\d{1,2}\s+\d{2}:\d{2}:\d{2}\s+\d{4})', "%b %d %H:%M:%S %Y"),
            # 格式: "定量时间  : Aug 26 08:56:17 2025"
            (r'定量时间\s*:\s*([A-Za-z]{3}\s+\d{1,2}\s+\d{2}:\d{2}:\d{2}\s+\d{4})', "%b %d %H:%M:%S %Y"),
        ]

        # 报告生成时间解析模式
        report_patterns = [
            # 格式: "SCCP混标-CA...0200816-CI.M 2025-09-12 周五 08:38:27"
            (r'(\d{4}-\d{1,2}-\d{1,2})\s+周[一二三四五六日]\s+(\d{2}:\d{2}:\d{2})', "REPORT_CN_WEEKDAY"),
            # 格式: "SVHC-CAL-248项-新.M Tue Aug 26 08:59:07 2025"
            (r'([A-Za-z]{3}\s+[A-Za-z]{3}\s+\d{1,2}\s+\d{2}:\d{2}:\d{2}\s+\d{4})', "%a %b %d %H:%M:%S %Y"),
            # 格式: "SVHC-CAL-248项-新.M Aug 26 08:59:07 2025" (没有星期)
            (r'([A-Za-z]{3}\s+\d{1,2}\s+\d{2}:\d{2}:\d{2}\s+\d{4})', "%b %d %H:%M:%S %Y"),
        ]

        for line in lines:
            # 解析定量时间
            if "定量时间" in line and quant_time is None:
                quant_time = self.parse_time_from_line(line, quant_patterns)

            # 解析报告生成时间
            if ("SCCP混标" in line or "SVHC-CAL" in line) and report_time is None:
                report_time = self.parse_time_from_line(line, report_patterns)

        # 检查时间是否一致
        if quant_time and report_time:
            if quant_time == report_time:
                error_msg = f"定量报告未复核"

                # 记录未复核文件信息 - 只记录文件夹名称
                if file_path:
                    folder_name = os.path.basename(os.path.dirname(file_path))
                    self.unreviewed_files.append(folder_name)  # 只添加文件夹名称，不添加括号内容

                return False, error_msg
            else:
                # 时间不一致，验证通过
                return True, "文件验证通过"
        else:
            # 如果无法解析时间，允许继续转换但给出警告
            return True, "文件验证通过(时间解析不完整)"

    def txt_to_pdf(self, input_txt, output_pdf, folder_name=None):
        """将TXT文件转换为PDF，保留原始格式，并移除末尾空白行"""
        try:
            output_dir = os.path.dirname(output_pdf)
            if not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)

            font_name = self.chinese_font
            encoding = self.detect_file_encoding(input_txt)

            # 改进的编码读取和转换逻辑
            content = None
            encoding_attempts = [
                encoding,  # 首先使用检测到的编码
                'gbk',
                'gb2312',
                'utf-8',
                'windows-1252',
                'cp1252',
                'latin-1'
            ]

            # 尝试多种编码读取文件
            for enc in encoding_attempts:
                try:
                    with open(input_txt, 'r', encoding=enc, errors='replace') as f:
                        content = f.read()
                    print(f"成功使用编码 {enc} 读取文件")
                    break
                except (UnicodeDecodeError, LookupError) as e:
                    print(f"编码 {enc} 失败: {e}")
                    continue
                except Exception as e:
                    print(f"使用编码 {enc} 时发生错误: {e}")
                    continue

            # 如果所有编码尝试都失败，使用二进制读取和chardet
            if content is None:
                try:
                    with open(input_txt, 'rb') as f:
                        raw_content = f.read()

                    # 使用chardet检测实际编码
                    detected = chardet.detect(raw_content)
                    actual_encoding = detected['encoding'] or 'gbk'
                    confidence = detected['confidence'] or 0

                    print(f"chardet检测到编码: {actual_encoding}, 置信度: {confidence}")

                    # 尝试使用检测到的编码
                    try:
                        content = raw_content.decode(actual_encoding, errors='replace')
                    except:
                        # 如果检测到的编码失败，尝试常见编码
                        for enc in ['gbk', 'gb2312', 'utf-8', 'latin-1']:
                            try:
                                content = raw_content.decode(enc, errors='replace')
                                break
                            except:
                                continue

                    # 如果仍然失败，使用replace错误处理
                    if content is None:
                        content = raw_content.decode('gbk', errors='replace')

                except Exception as e:
                    error_msg = f"无法读取文件内容: {str(e)}"
                    return False, error_msg, {}

            # 特殊处理Windows-1252编码的内容
            if encoding in ['windows-1252', 'cp1252']:
                try:
                    # 先尝试直接解码为GBK
                    with open(input_txt, 'rb') as f:
                        raw_bytes = f.read()

                    # 尝试多种编码转换
                    conversion_attempts = [
                        ('windows-1252', 'gbk'),
                        ('cp1252', 'gbk'),
                        ('latin-1', 'gbk'),
                        ('windows-1252', 'utf-8'),
                        ('cp1252', 'utf-8')
                    ]

                    for from_enc, to_enc in conversion_attempts:
                        try:
                            temp_content = raw_bytes.decode(from_enc).encode(to_enc)
                            test_content = temp_content.decode(to_enc, errors='strict')
                            # 检查是否包含中文字符
                            if any('\u4e00' <= char <= '\u9fff' for char in test_content):
                                content = test_content
                                print(f"成功转换编码: {from_enc} -> {to_enc}")
                                break
                        except:
                            continue

                except Exception as e:
                    print(f"编码转换失败: {e}")
                    # 保持原始内容

            # 提取文件信息
            extracted_info = self.extract_file_info(content, input_txt)
            # 存储提取的信息
            self.extracted_info[input_txt] = extracted_info

            # 验证文件内容
            is_valid, validation_msg = self.validate_epatemp_content(content, input_txt)
            if not is_valid:
                return False, validation_msg, extracted_info

            # 其余代码保持不变...
            styles = getSampleStyleSheet()
            font_size = 12
            line_height = font_size
            styles.add(ParagraphStyle(
                name='EmptyLine',
                fontName=font_name,
                fontSize=12,
                leading=14,
                spaceBefore=0,
                spaceAfter=0,
                textColor='white'
            ))

            styles.add(ParagraphStyle(
                name='ChineseFixed',
                fontName=font_name,
                fontSize=12,
                leading=14,
                spaceBefore=0,
                spaceAfter=0
            ))

            doc = SimpleDocTemplate(
                output_pdf,
                pagesize=A4,
                rightMargin=40,
                leftMargin=0,
                topMargin=0,
                bottomMargin=0
            )
            story = []

            lines = content.split('\n')
            for i in range(len(lines) - 1, -1, -1):
                if lines[i].strip() == '':
                    lines.pop()
                else:
                    break

            for line in lines:
                stripped_line = line.rstrip()

                if stripped_line == '':
                    p = Paragraph('&nbsp;', styles['EmptyLine'])
                    story.append(p)
                else:
                    formatted_line = stripped_line.replace(' ', '&nbsp;')
                    p = Paragraph(formatted_line, styles['ChineseFixed'])
                    story.append(p)

            try:
                doc.build(story)
                self.remove_blank_pages(output_pdf)
                return True, "转换成功", extracted_info
            except Exception as e:
                error_msg = f"PDF生成失败: {str(e)}"
                return False, error_msg, extracted_info

        except Exception as e:
            error_msg = f"转换过程中出错: {str(e)}"
            return False, error_msg, {}

    def convert_epatemp_to_pdf(self, epatemp_path, output_pdf_path=None, folder_name=None):
        """将epatemp.txt文件转换为PDF的主函数"""
        try:
            if not os.path.exists(epatemp_path):
                error_msg = f"epatemp文件不存在: {epatemp_path}"
                return False, error_msg, {}

            if folder_name is None:
                folder_name = os.path.basename(os.path.dirname(epatemp_path))

            if output_pdf_path is None:
                output_dir = os.path.dirname(epatemp_path)
                clean_name = self.clean_folder_name(folder_name)
                output_pdf_path = os.path.join(output_dir, f"{clean_name}_report.pdf")

            success, message, extracted_info = self.txt_to_pdf(epatemp_path, output_pdf_path, folder_name)

            if success:
                return output_pdf_path, message, extracted_info
            else:
                return False, message, extracted_info

        except Exception as e:
            error_msg = f"转换epatemp到PDF时出错: {e}"
            return False, error_msg, {}


# 创建全局实例
converter = EpatempToPdfConverter()


# 提供简单的函数接口
def convert_epatemp(epatemp_path, output_pdf_path=None, folder_name=None):
    """简单的函数接口，用于转换epatemp文件到PDF"""
    return converter.convert_epatemp_to_pdf(epatemp_path, output_pdf_path, folder_name)


def batch_convert_epatemp(folder_path, output_dir=None):
    """批量转换文件夹中所有子文件夹的epatemp.txt文件"""
    import glob

    if not os.path.isdir(folder_path):
        error_msg = f"文件夹不存在: {folder_path}"
        converter.show_error_message("文件夹不存在", error_msg)
        return False, 0, 0, []

    if output_dir is None:
        output_dir = os.path.join(folder_path, "谱图报告")

    if not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    success_count = 0
    fail_count = 0
    failed_files = []
    all_extracted_info = []  # 存储所有提取的信息

    # 查找所有包含epatemp.txt的子文件夹
    for root, dirs, files in os.walk(folder_path):
        if "epatemp.txt" in files:
            epatemp_path = os.path.join(root, "epatemp.txt")
            folder_name = os.path.basename(root)

            if output_dir == os.path.dirname(epatemp_path):
                # 如果输出目录与输入目录相同，添加后缀
                output_pdf = os.path.join(output_dir, f"{folder_name}_report.pdf")
            else:
                output_pdf = os.path.join(output_dir, f"{folder_name}.pdf")

            result, message, extracted_info = convert_epatemp(epatemp_path, output_pdf, folder_name)

            if result:
                success_count += 1
                all_extracted_info.append(extracted_info)
            else:
                fail_count += 1
                failed_files.append((epatemp_path, message))

    # 批量转换完成后显示总结弹窗和未复核汇总
    if converter.unreviewed_files:
        converter.show_unreviewed_summary()

    if success_count > 0 and fail_count == 0:
        converter.show_info_message("批量转换完成", f"所有文件转换成功!\n成功转换: {success_count} 个文件")
    elif success_count > 0 and fail_count > 0:
        converter.show_warning_message("批量转换完成",
                                       f"批量转换完成!\n成功: {success_count} 个文件\n失败: {fail_count} 个文件")
    elif success_count == 0 and fail_count > 0:
        converter.show_error_message("批量转换完成",
                                     f"所有文件转换失败!\n失败: {fail_count} 个文件")

    return success_count, fail_count, failed_files, all_extracted_info


def get_chinese_font():
    """获取注册的中文字体名称"""
    return converter.chinese_font


def show_unreviewed_summary():
    """显示未复核文件汇总的公共接口"""
    converter.show_unreviewed_summary()


def clear_unreviewed_files():
    """清空未复核文件列表的公共接口"""
    converter.unreviewed_files = []


def get_extracted_info(file_path=None):
    """
    获取提取的文件信息

    Args:
        file_path (str): 可选，指定要获取信息的文件路径

    Returns:
        dict: 提取的信息字典，包含'data_file'和'sample'字段
    """
    return converter.get_extracted_info(file_path)


def extract_file_info_from_content(content, file_path=None):
    """
    直接从内容中提取文件信息的公共接口

    Args:
        content (str): 文件内容
        file_path (str): 可选，文件路径

    Returns:
        dict: 提取的信息字典
    """
    return converter.extract_file_info(content, file_path)


def clear_extracted_info():
    """清空所有提取的信息"""
    converter.clear_extracted_info()


# 在文件末尾添加
def get_epatemp_modify_time(epatemp_path):
    """获取epatemp.txt文件修改时间的公共接口"""
    return converter.get_epatemp_modify_time(epatemp_path)


# 示例使用方式函数
def example_usage():
    """
    示例使用方式，展示如何调用新增的功能
    """
    print("=== 示例使用方式 ===")

    # 1. 转换单个文件并获取提取的信息
    epatemp_path = "path/to/your/epatemp.txt"
    result, message, extracted_info = convert_epatemp(epatemp_path)

    if result:
        print(f"转换成功: {result}")
        print(f"提取的信息: {extracted_info}")

        # 使用提取的信息
        if extracted_info['extraction_success']:
            data_file = extracted_info['data_file']
            sample_name = extracted_info['sample']
            print(f"数据文件: {data_file}")
            print(f"样品名称: {sample_name}")

            # 可以将这些信息传递给其他模块使用
            # other_module.process_data(data_file, sample_name)
    else:
        print(f"转换失败: {message}")

    # 2. 批量转换并获取所有提取的信息
    folder_path = "path/to/your/folder"
    success_count, fail_count, failed_files, all_extracted_info = batch_convert_epatemp(folder_path)

    print(f"\n批量转换结果:")
    print(f"成功: {success_count} 个文件")
    print(f"失败: {fail_count} 个文件")

    # 处理所有提取的信息
    for i, info in enumerate(all_extracted_info):
        if info['extraction_success']:
            print(f"文件 {i + 1}: 数据文件={info['data_file']}, 样品={info['sample']}")

    # 3. 直接从内容中提取信息
    content = """
    数据文件  : D:\\Data\\2025\\09\\15\\sample001.d
    样品      : 测试样品001
    """
    info = extract_file_info_from_content(content)
    print(f"\n直接提取信息: {info}")

    # 4. 获取特定文件的提取信息
    specific_info = get_extracted_info("specific/file/path.txt")
    print(f"特定文件信息: {specific_info}")


if __name__ == "__main__":
    # 运行示例
    example_usage()
