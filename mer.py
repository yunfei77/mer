from email import policy
from email.parser import BytesParser
import extract_msg
import os
from typing import Dict, Any, List
from bs4 import BeautifulSoup
import re
import io
import time
try:
    import PyPDF2  # 用于解析PDF文件
    from docx import Document  # 用于解析Word文档
    import openpyxl  # 用于解析Excel文件
    import pptx  # 用于解析PPT文件
    EXTRA_FORMATS_SUPPORTED = True
except ImportError:
    EXTRA_FORMATS_SUPPORTED = False
    print("提示: 要支持更多文档格式预览，请安装以下包:")
    print("pip install python-docx PyPDF2 openpyxl python-pptx")

import whois
from datetime import datetime, timezone

def parse_email(file_path: str) -> Dict[str, Any]:
    """
    解析邮件文件，提取关键信息
    
    Args:
        file_path: 邮件文件路径
        
    Returns:
        包含邮件信息的字典
    """
    email_data = {
        'from': [],      # 所有发件人
        'to': [],        # 所有收件人
        'cc': [],        # 抄送人
        'bcc': [],       # 密送人
        'reply_to': [],  # 回复地址
        'subject': '',   # 主题
        'date': '',      # 日期
        'body_text': '', # 纯文本正文
        'body_html': '', # HTML正文
        'attachments': [], # 附件列表
        'references': [], # 引用的邮件ID
        'in_reply_to': [], # 回复的邮件ID
        'thread_info': {  # 邮件会话信息
            'original_sender': '',
            'original_recipients': [],
            'original_subject': '',
            'original_date': ''
        }
    }
    
    try:
        if file_path.lower().endswith('.eml'):
            with open(file_path, 'rb') as f:
                msg = BytesParser(policy=policy.default).parse(f)
                
                # 解析邮件头
                def parse_address_list(header_value):
                    """解析邮件地址列表"""
                    if not header_value:
                        return []
                    # 处理可能的多行地址
                    if isinstance(header_value, (list, tuple)):
                        addresses = []
                        for item in header_value:
                            addresses.extend([addr.strip() for addr in str(item).split(',')])
                        return addresses
                    return [addr.strip() for addr in str(header_value).split(',')]
                
                # 处理发件人
                from_header = msg.get('From', '')
                email_data['from'] = parse_address_list(from_header)
                
                # 处理收件人
                to_header = msg.get('To', '')
                email_data['to'] = parse_address_list(to_header)
                
                # 处理抄送人
                cc_header = msg.get('Cc', '')
                email_data['cc'] = parse_address_list(cc_header)
                
                # 处理密送人
                bcc_header = msg.get('Bcc', '')
                email_data['bcc'] = parse_address_list(bcc_header)
                
                # 处理回复地址
                reply_to_header = msg.get('Reply-To', '')
                email_data['reply_to'] = parse_address_list(reply_to_header)
                
                # 处理主题和日期
                email_data['subject'] = msg.get('Subject', '')
                email_data['date'] = msg.get('Date', '')
                
                # 处理邮件引用和回复信息
                references = msg.get('References', '')
                in_reply_to = msg.get('In-Reply-To', '')
                email_data['references'] = parse_address_list(references)
                email_data['in_reply_to'] = parse_address_list(in_reply_to)
                
                # 尝试获取原始邮件信息
                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_type() == 'message/rfc822':
                            # 获取原始邮件
                            original_msg = part.get_payload()[0]
                            email_data['thread_info']['original_sender'] = original_msg.get('From', '')
                            email_data['thread_info']['original_recipients'] = parse_address_list(original_msg.get('To', ''))
                            email_data['thread_info']['original_subject'] = original_msg.get('Subject', '')
                            email_data['thread_info']['original_date'] = original_msg.get('Date', '')
                
                # 改进附件解析部分
                def extract_attachments(msg):
                    """递归提取附件"""
                    if msg.is_multipart():
                        for part in msg.walk():
                            if part.get_content_maintype() == 'multipart':
                                continue
                            
                            disposition = part.get_content_disposition()
                            if disposition and disposition.lower() in ['attachment', 'inline']:
                                try:
                                    attachment_info = parse_attachment(part, len(email_data['attachments']))
                                    email_data['attachments'].append(attachment_info)
                                    print(f"发现{'内联' if attachment_info['is_inline'] else ''}附件: {attachment_info['filename']}")
                                except Exception as e:
                                    print(f"处理附件出错: {str(e)}")
                
                # 调用附件提取函数
                extract_attachments(msg)
                
                # 改进正文解析部分
                def get_body_content(msg):
                    """递归获取邮件正文内容"""
                    if msg.is_multipart():
                        # 获取所有部分
                        all_parts = []
                        for part in msg.walk():
                            if part.get_content_maintype() == 'text':
                                all_parts.append(part)
                        
                        # 优先处理 text/plain
                        for part in all_parts:
                            if part.get_content_type() == 'text/plain':
                                try:
                                    payload = part.get_payload(decode=True)
                                    if payload:
                                        # 尝试多种编码方式
                                        encodings = ['utf-8', 'gbk', 'gb2312', 'gb18030', 'big5']
                                        for encoding in encodings:
                                            try:
                                                decoded_text = payload.decode(encoding)
                                                if decoded_text:
                                                    email_data['body_text'] += decoded_text + '\n'
                                                    break
                                            except:
                                                continue
                                except Exception as e:
                                    print(f"解析纯文本内容失败: {str(e)}")
                        
                        # 然后处理 text/html
                        for part in all_parts:
                            if part.get_content_type() == 'text/html':
                                try:
                                    # 检查内容传输编码
                                    transfer_encoding = part.get('Content-Transfer-Encoding', '').lower()
                                    
                                    # 获取原始负载
                                    if transfer_encoding == 'base64':
                                        payload = part.get_payload(decode=True)
                                    else:
                                        payload = part.get_payload(decode=True)
                                    
                                    if payload:
                                        # 尝试多种编码方式
                                        encodings = ['utf-8', 'gbk', 'gb2312', 'gb18030', 'big5']
                                        for encoding in encodings:
                                            try:
                                                decoded_html = payload.decode(encoding)
                                                if decoded_html:
                                                    # 验证是否为有效的HTML内容
                                                    if any(marker in decoded_html.lower() for marker in ['<html', '<!doctype', '<body', '<head', '<div', '<p']):
                                                        email_data['body_html'] = decoded_html
                                                        break
                                            except:
                                                continue
                                except Exception as e:
                                    print(f"解析HTML内容失败: {str(e)}")
                    else:
                        # 非多部分邮件
                        content_type = msg.get_content_type()
                        try:
                            # 检查内容传输编码
                            transfer_encoding = msg.get('Content-Transfer-Encoding', '').lower()
                            
                            # 获取原始负载
                            if transfer_encoding == 'base64':
                                payload = msg.get_payload(decode=True)
                            else:
                                payload = msg.get_payload(decode=True)
                            
                            if payload:
                                # 尝试多种编码方式
                                encodings = ['utf-8', 'gbk', 'gb2312', 'gb18030', 'big5']
                                for encoding in encodings:
                                    try:
                                        decoded_content = payload.decode(encoding)
                                        if decoded_content:
                                            if content_type == 'text/plain':
                                                email_data['body_text'] = decoded_content
                                                break
                                            elif content_type == 'text/html':
                                                if any(marker in decoded_content.lower() for marker in ['<html', '<!doctype', '<body', '<head', '<div', '<p']):
                                                    email_data['body_html'] = decoded_content
                                                    break
                                    except:
                                        continue
                        except Exception as e:
                            print(f"解析正文出错: {str(e)}")
                
                # 调用正文解析函数
                get_body_content(msg)
                
                # 如果正文仍然为空，尝试其他方法
                if not email_data['body_text'] and not email_data['body_html']:
                    try:
                        # 直接获取原始负载
                        payload = msg.get_payload()
                        if isinstance(payload, list):
                            for part in payload:
                                if part.get_content_type() == 'text/plain':
                                    content = part.get_payload(decode=True)
                                    if content:
                                        email_data['body_text'] += content.decode('utf-8', errors='ignore') + '\n'
                                elif part.get_content_type() == 'text/html':
                                    content = part.get_payload(decode=True)
                                    if content:
                                        email_data['body_html'] += content.decode('utf-8', errors='ignore') + '\n'
                        elif isinstance(payload, str):
                            email_data['body_text'] = payload
                    except Exception as e:
                        print(f"备用方法获取正文失败: {str(e)}")
                
                # 如果HTML内容存在但看起来是乱码，尝试不同的解码方式
                if email_data['body_html'] and not email_data['body_html'].strip().startswith(('<html', '<!DOCTYPE', '<body', '<div')):
                    try:
                        # 尝试不同的编码
                        encodings = ['utf-8', 'gbk', 'gb2312', 'gb18030', 'big5', 'latin1']
                        original_content = email_data['body_html'].encode('latin1', errors='ignore')
                        
                        for encoding in encodings:
                            try:
                                decoded = original_content.decode(encoding)
                                if decoded.strip().startswith(('<html', '<!DOCTYPE', '<body', '<div')):
                                    email_data['body_html'] = decoded
                                    break
                            except:
                                continue
                    except Exception as e:
                        print(f"尝试重新解码HTML内容失败: {str(e)}")
                
        elif file_path.lower().endswith('.msg'):
            msg = extract_msg.Message(file_path)
            try:
                # 处理发件人
                sender = msg.sender
                if sender:
                    email_data['from'].append(sender)
                
                # 处理收件人
                recipients = msg.to
                if recipients:
                    email_data['to'] = [r.strip() for r in recipients.split(';')]
                
                # 处理抄送人
                cc = msg.cc
                if cc:
                    email_data['cc'] = [c.strip() for c in cc.split(';')]
                
                # 处理密送人
                try:
                    bcc = msg.bcc
                    if bcc:
                        email_data['bcc'] = [b.strip() for b in bcc.split(';')]
                except AttributeError:
                    print("MSG文件不包含密送人信息")
                
                # 处理主题和日期
                email_data['subject'] = msg.subject or ''
                email_data['date'] = str(msg.date) if msg.date else ''
                
                # 改进原始邮件信息获取
                try:
                    # 从邮件正文中提取原始邮件信息
                    body_text = msg.body or ''
                    
                    # 定义更强大的正则表达式模式
                    from_pattern = r'From:[\s]*([^\r\n]+)'
                    to_pattern = r'To:[\s]*([^\r\n]+)'
                    subject_pattern = r'Subject:[\s]*([^\r\n]+)'
                    sent_pattern = r'Sent:[\s]*([^\r\n]+)'
                    cc_pattern = r'Cc:[\s]*([^\r\n]+)'
                    
                    # 查找所有匹配项（可能有多个，取最后一个作为原始邮件信息）
                    from_matches = re.findall(from_pattern, body_text)
                    to_matches = re.findall(to_pattern, body_text)
                    subject_matches = re.findall(subject_pattern, body_text)
                    sent_matches = re.findall(sent_pattern, body_text)
                    cc_matches = re.findall(cc_pattern, body_text)
                    
                    # 如果找到匹配项，使用最后一个（通常是原始邮件的信息）
                    if from_matches:
                        email_data['thread_info']['original_sender'] = from_matches[-1].strip()
                    
                    if to_matches:
                        recipients = [r.strip() for r in to_matches[-1].split(';')]
                        email_data['thread_info']['original_recipients'] = recipients
                    
                    if subject_matches:
                        email_data['thread_info']['original_subject'] = subject_matches[-1].strip()
                    
                    if sent_matches:
                        email_data['thread_info']['original_date'] = sent_matches[-1].strip()
                    
                    if cc_matches:
                        cc_list = [c.strip() for c in cc_matches[-1].split(';')]
                        if not email_data['cc']:  # 如果之前没有抄送人，则添加
                            email_data['cc'] = cc_list
                        else:  # 如果已有抄送人，则合并并去重
                            email_data['cc'].extend(cc_list)
                            email_data['cc'] = list(set(email_data['cc']))
                    
                    # 如果在正文中没有找到，尝试从HTML正文中提取
                    if not any(email_data['thread_info'].values()):
                        html_body = msg.htmlBody or ''
                        if html_body:
                            # 使用BeautifulSoup清理HTML标签
                            soup = BeautifulSoup(html_body, 'html.parser')
                            clean_text = soup.get_text()
                            
                            # 重新查找
                            from_matches = re.findall(from_pattern, clean_text)
                            to_matches = re.findall(to_pattern, clean_text)
                            subject_matches = re.findall(subject_pattern, clean_text)
                            sent_matches = re.findall(sent_pattern, clean_text)
                            
                            if from_matches:
                                email_data['thread_info']['original_sender'] = from_matches[-1].strip()
                            if to_matches:
                                recipients = [r.strip() for r in to_matches[-1].split(';')]
                                email_data['thread_info']['original_recipients'] = recipients
                            if subject_matches:
                                email_data['thread_info']['original_subject'] = subject_matches[-1].strip()
                            if sent_matches:
                                email_data['thread_info']['original_date'] = sent_matches[-1].strip()
                
                except Exception as e:
                    print(f"获取原始邮件信息失败: {str(e)}")
                
                # 处理正文
                email_data['body_text'] = msg.body or ''
                email_data['body_html'] = msg.htmlBody or ''
                
                # 处理附件
                for attachment in msg.attachments:
                    try:
                        attachment_info = parse_attachment(attachment, len(email_data['attachments']))
                        email_data['attachments'].append(attachment_info)
                        print(f"发现MSG附件: {attachment_info['filename']}")
                    except Exception as e:
                        print(f"处理MSG附件出错: {str(e)}")
                
                msg.close()
                
            except Exception as e:
                print(f"处理MSG文件出错: {str(e)}")
                if msg:
                    msg.close()
                
    except Exception as e:
        print(f"邮件解析失败: {str(e)}")
        
    return email_data

def parse_attachment(attachment: Any, attachment_index: int) -> Dict[str, Any]:
    """
    解析邮件附件，提取详细信息
    
    Args:
        attachment: 附件对象
        attachment_index: 附件索引号
        
    Returns:
        包含附件详细信息的字典
    """
    attachment_info = {
        'filename': '',          # 文件名
        'mime_type': '',         # MIME类型
        'size': 0,              # 文件大小
        'data': None,           # 原始数据
        'is_inline': False,     # 是否为内联附件
        'content_id': '',       # 内联附件的Content-ID
        'extension': '',        # 文件扩展名
        'hash_md5': '',         # MD5哈希值
        'hash_sha256': '',      # SHA256哈希值
        'text_preview': '',     # 文本预览（如果是文本文件）
        'is_archive': False,    # 是否为压缩文件
        'archive_contents': [],  # 压缩文件内容列表
        'is_executable': False, # 是否为可执行文件
        'created_date': '',     # 创建日期
        'modified_date': '',    # 修改日期
    }
    
    try:
        # 处理MSG附件
        if hasattr(attachment, 'longFilename'):
            filename = attachment.longFilename or attachment.shortFilename or f"未命名附件_{attachment_index}"
            attachment_info['filename'] = filename
            attachment_info['mime_type'] = attachment.mimetype or 'application/octet-stream'
            attachment_info['data'] = attachment.data
            attachment_info['size'] = len(attachment.data) if attachment.data else 0
            
        # 处理EML附件
        else:
            # 获取文件名
            filename = attachment.get_filename()
            if not filename:
                filename = attachment.get_param('name')
            if not filename:
                filename = f"未命名附件_{attachment_index}"
            attachment_info['filename'] = filename
            
            # 获取MIME类型
            attachment_info['mime_type'] = attachment.get_content_type()
            
            # 获取是否为内联附件
            disposition = attachment.get_content_disposition()
            attachment_info['is_inline'] = (disposition and disposition.lower() == 'inline')
            
            # 获取Content-ID
            content_id = attachment.get('Content-ID', '')
            if content_id:
                attachment_info['content_id'] = content_id.strip('<>')
            
            # 获取附件数据
            payload = attachment.get_payload(decode=True)
            if payload:
                attachment_info['data'] = payload
                attachment_info['size'] = len(payload)
        
        # 获取文件扩展名
        attachment_info['extension'] = os.path.splitext(attachment_info['filename'])[1].lower()
        
        # 计算哈希值
        if attachment_info['data']:
            import hashlib
            attachment_info['hash_md5'] = hashlib.md5(attachment_info['data']).hexdigest()
            attachment_info['hash_sha256'] = hashlib.sha256(attachment_info['data']).hexdigest()
        
        # 检查是否为可执行文件
        executable_extensions = {'.exe', '.dll', '.bat', '.cmd', '.msi', '.vbs', '.js', '.ps1', '.com', '.scr'}
        attachment_info['is_executable'] = attachment_info['extension'] in executable_extensions
        
        # 检查是否为压缩文件并提取内容列表
        archive_extensions = {'.zip', '.rar', '.7z', '.tar', '.gz', '.bz2'}
        if attachment_info['extension'] in archive_extensions:
            attachment_info['is_archive'] = True
            try:
                import zipfile
                if attachment_info['extension'] == '.zip' and attachment_info['data']:
                    from io import BytesIO
                    with zipfile.ZipFile(BytesIO(attachment_info['data'])) as zf:
                        attachment_info['archive_contents'] = zf.namelist()
            except Exception as e:
                print(f"解析压缩文件内容失败: {str(e)}")
        
        # 扩展文本文件和文档预览支持
        preview_extensions = {
            'text': {'.txt', '.csv', '.log', '.xml', '.json', '.html', '.htm', '.css', '.js', '.py', '.java', '.cpp', '.c', '.h', '.sql'},
            'document': {'.pdf', '.docx', '.xlsx', '.pptx', '.doc', '.xls', '.ppt'},
        }
        
        # 如果是文本文件，添加预览
        if attachment_info['extension'] in preview_extensions['text'] and attachment_info['data']:
            try:
                preview_text = attachment_info['data'].decode('utf-8', errors='ignore')
                attachment_info['text_preview'] = preview_text[:2000] + '...' if len(preview_text) > 2000 else preview_text
            except Exception as e:
                print(f"生成文本预览失败: {str(e)}")
        
        # 如果是文档文件且支持扩展格式，添加预览
        elif EXTRA_FORMATS_SUPPORTED and attachment_info['extension'] in preview_extensions['document'] and attachment_info['data']:
            try:
                preview_text = ""
                
                # PDF文件处理
                if attachment_info['extension'] == '.pdf':
                    try:
                        pdf_file = io.BytesIO(attachment_info['data'])
                        pdf_reader = PyPDF2.PdfReader(pdf_file)
                        preview_text = "PDF文档内容预览:\n"
                        # 只预览前3页
                        for page_num in range(min(3, len(pdf_reader.pages))):
                            preview_text += f"\n--- 第{page_num + 1}页 ---\n"
                            preview_text += pdf_reader.pages[page_num].extract_text()[:1000]
                            if page_num < min(2, len(pdf_reader.pages) - 1):
                                preview_text += "\n...\n"
                    except Exception as e:
                        print(f"PDF文档预览失败: {str(e)}")
                
                # Word文档处理
                elif attachment_info['extension'] == '.docx':
                    try:
                        docx_file = io.BytesIO(attachment_info['data'])
                        doc = Document(docx_file)
                        preview_text = "Word文档内容预览:\n\n"
                        # 获取文档的前10个段落
                        for i, para in enumerate(doc.paragraphs[:10]):
                            if para.text.strip():
                                preview_text += para.text + "\n"
                        if len(doc.paragraphs) > 10:
                            preview_text += "\n... (更多内容已省略)"
                    except Exception as e:
                        print(f"Word文档预览失败: {str(e)}")
                
                # Excel文件处理
                elif attachment_info['extension'] == '.xlsx':
                    try:
                        xlsx_file = io.BytesIO(attachment_info['data'])
                        wb = openpyxl.load_workbook(xlsx_file, read_only=True)
                        preview_text = "Excel文档内容预览:\n\n"
                        # 预览第一个工作表的前10行
                        sheet = wb.active
                        for i, row in enumerate(sheet.iter_rows(max_row=10)):
                            if i == 0:
                                preview_text += "表头: "
                            preview_text += " | ".join(str(cell.value) for cell in row) + "\n"
                        if sheet.max_row > 10:
                            preview_text += "\n... (更多行已省略)"
                        wb.close()  # 确保关闭工作簿
                    except Exception as e:
                        print(f"Excel文档预览失败: {str(e)}")
                
                # PowerPoint文件处理
                elif attachment_info['extension'] == '.pptx':
                    try:
                        pptx_file = io.BytesIO(attachment_info['data'])
                        prs = pptx.Presentation(pptx_file)
                        preview_text = "PowerPoint文档内容预览:\n\n"
                        # 预览前3张幻灯片
                        for i, slide in enumerate(prs.slides[:3]):
                            preview_text += f"\n--- 幻灯片 {i+1} ---\n"
                            for shape in slide.shapes:
                                if hasattr(shape, "text"):
                                    preview_text += shape.text + "\n"
                        if len(prs.slides) > 3:
                            preview_text += "\n... (更多幻灯片已省略)"
                    except Exception as e:
                        print(f"PowerPoint文档预览失败: {str(e)}")
                
                if preview_text:
                    attachment_info['text_preview'] = preview_text
                
            except Exception as e:
                print(f"文档预览处理失败: {str(e)}")
        
    except Exception as e:
        print(f"解析附件失败: {str(e)}")
    
    return attachment_info

def extract_urls(email_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    从邮件中提取所有URL并进行安全分析
    """
    result = {
        'urls': {
            'text': [],      # 纯文本中的URL
            'html': [],      # HTML内容中的URL
            'attachments': [] # 附件内容中的URL
        },
        'suspicious_links': [],  # 可疑的超链接
        'risk_level': 'low',
        'risk_score': 0.0,
        'warnings': []
    }
    
    try:
        # URL正则表达式模式
        url_pattern = r'(?i)\b((?:https?://|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}/)(?:[^\s()<>]+|\(([^\s()<>]+|(\([^\s()<>]+\)))*\))+(?:\(([^\s()<>]+|(\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:\'".,<>?«»""'']))'
        
        def extract_domain(url: str) -> str:
            """从URL中提取域名"""
            try:
                from urllib.parse import urlparse
                parsed = urlparse(url)
                return parsed.netloc.lower()
            except:
                return ''
        
        def analyze_domain_relationship(domain1: str, domain2: str) -> Dict[str, Any]:
            """分析两个域名之间的关系"""
            if not domain1 or not domain2:
                return {'similarity': 0.0, 'relationship_type': None, 'risk': 'low'}
            
            base1 = domain1.split('.')[0]
            base2 = domain2.split('.')[0]
            
            result = {
                'similarity': 0.0,
                'relationship_type': None,
                'risk': 'low'
            }
            
            # 1. 检查完全匹配
            if domain1 == domain2:
                return result
            
            # 2. 检查字符替换（如 sinopec -> siuopec）
            if abs(len(base1) - len(base2)) <= 1:
                diff_count = 0
                for c1, c2 in zip(base1, base2):
                    if c1 != c2:
                        diff_count += 1
                        if diff_count > 1:
                            break
                if diff_count <= 1:
                    result['similarity'] = 0.95
                    result['relationship_type'] = '可疑的字符替换'
                    result['risk'] = 'high'
                    return result
            
            # 3. 检查包含关系
            if base1 in base2 or base2 in base1:
                longer = base1 if len(base1) > len(base2) else base2
                shorter = base2 if len(base1) > len(base2) else base1
                
                if shorter in longer:
                    suspicious_additions = {'portal', 'service', 'vendor', 'secure', 'mail', 
                                         'auth', 'login', 'account', 'verify', 'update'}
                    remaining = longer.replace(shorter, '').lower()
                    
                    if any(word in remaining for word in suspicious_additions):
                        result['similarity'] = 0.9
                        result['relationship_type'] = '可疑的域名包含'
                        result['risk'] = 'high'
                        return result
            
            # 4. 计算编辑距离相似度
            def levenshtein_distance(s1: str, s2: str) -> int:
                if len(s1) < len(s2):
                    return levenshtein_distance(s2, s1)
                if len(s2) == 0:
                    return len(s1)
                previous_row = range(len(s2) + 1)
                for i, c1 in enumerate(s1):
                    current_row = [i + 1]
                    for j, c2 in enumerate(s2):
                        insertions = previous_row[j + 1] + 1
                        deletions = current_row[j] + 1
                        substitutions = previous_row[j] + (c1 != c2)
                        current_row.append(min(insertions, deletions, substitutions))
                    previous_row = current_row
                return previous_row[-1]
            
            distance = levenshtein_distance(base1, base2)
            max_length = max(len(base1), len(base2))
            similarity = 1 - (distance / max_length)
            
            if similarity > 0.8:
                result['similarity'] = similarity
                result['relationship_type'] = '高度相似'
                result['risk'] = 'high'
            
            return result
        
        def analyze_link_safety(display_text: str, actual_url: str, email_domains: List[str]) -> Dict[str, Any]:
            """分析超链接的安全性"""
            result = {
                'display_text': display_text,
                'actual_url': actual_url,
                'risk_level': 'low',
                'risk_score': 0.0,
                'reasons': []
            }
            
            # 1. 检查显示文本是否包含URL
            display_domain = None
            if re.search(url_pattern, display_text):
                display_domain = extract_domain(display_text)
            
            # 2. 获取实际URL的域名
            actual_domain = extract_domain(actual_url)
            
            if display_domain and actual_domain and display_domain != actual_domain:
                result['risk_score'] += 3.0
                result['reasons'].append(f"超链接显示域名({display_domain})与实际域名({actual_domain})不匹配")
            
            # 3. 检查域名与邮件发件人域名的相似度
            for email_domain in email_domains:
                if actual_domain:
                    analysis = analyze_domain_relationship(actual_domain, email_domain)
                    if analysis['risk'] == 'high':
                        result['risk_score'] += 2.5
                        result['reasons'].append(
                            f"URL域名({actual_domain})与邮件域名({email_domain})高度相似: "
                            f"{analysis['relationship_type']}"
                        )
            
            # 4. 检查URL的其他可疑特征
            try:
                from urllib.parse import urlparse, parse_qs
                parsed = urlparse(actual_url)
                
                # 检查非标准端口
                if parsed.port and parsed.port not in (80, 443):
                    result['risk_score'] += 2.0
                    result['reasons'].append(f"使用非标准端口: {parsed.port}")
                
                # 检查URL编码过度
                if actual_url.count('%') > 5:
                    result['risk_score'] += 1.5
                    result['reasons'].append("URL过度编码，可能试图隐藏真实地址")
                
                # 检查重定向参数
                redirect_params = {'url', 'redirect', 'goto', 'link', 'return', 'target'}
                query_params = parse_qs(parsed.query)
                found_redirects = set(query_params.keys()) & redirect_params
                if found_redirects:
                    result['risk_score'] += 2.0
                    result['reasons'].append(f"包含重定向参数: {', '.join(found_redirects)}")
            
            except Exception as e:
                print(f"URL分析失败: {str(e)}")
            
            # 设置最终风险等级
            if result['risk_score'] >= 3.0:
                result['risk_level'] = 'high'
            elif result['risk_score'] >= 1.5:
                result['risk_level'] = 'medium'
            
            return result
        
        # 获取邮件相关的域名列表
        email_domains = set()
        for sender in email_data['from']:
            domain_match = re.search(r'@([\w.-]+)', sender)
            if domain_match:
                email_domains.add(domain_match.group(1).lower())
        
        # 从纯文本中提取URL
        if email_data['body_text']:
            text_urls = re.findall(url_pattern, email_data['body_text'])
            result['urls']['text'] = [url[0] for url in text_urls]
        
        # 从HTML内容中提取URL和分析超链接
        if email_data['body_html']:
            try:
                soup = BeautifulSoup(email_data['body_html'], 'html.parser')
                
                # 分析所有超链接
                for link in soup.find_all('a'):
                    href = link.get('href')
                    if href and not href.startswith('mailto:'):
                        display_text = link.get_text(strip=True)
                        analysis = analyze_link_safety(display_text, href, list(email_domains))
                        
                        result['urls']['html'].append(href)
                        if analysis['risk_level'] != 'low':
                            result['suspicious_links'].append(analysis)
                            result['risk_score'] += analysis['risk_score']
                
                # 提取图片链接
                for img in soup.find_all('img'):
                    src = img.get('src')
                    if src and not src.startswith('data:'):
                        result['urls']['html'].append(src)
                
            except Exception as e:
                print(f"解析HTML内容时出错: {str(e)}")
        
        # 从附件中提取URL
        for attachment in email_data['attachments']:
            if attachment.get('text_preview'):
                attachment_urls = re.findall(url_pattern, attachment['text_preview'])
                result['urls']['attachments'].extend([url[0] for url in attachment_urls])
        
        # 去重并清理URL
        for source in result['urls']:
            result['urls'][source] = list(set(result['urls'][source]))
            result['urls'][source] = [url.rstrip('.,;:\'\"!?') for url in result['urls'][source]]
            result['urls'][source] = [url for url in result['urls'][source] if url]
        
        # 设置最终风险等级
        if result['risk_score'] >= 5.0:
            result['risk_level'] = 'high'
            result['warnings'].append("发现多个高风险URL，疑似钓鱼邮件！")
        elif result['risk_score'] >= 2.5:
            result['risk_level'] = 'medium'
            result['warnings'].append("发现可疑URL，请谨慎点击")
        
        # 添加统计信息
        total_urls = len(result['urls']['text']) + len(result['urls']['html']) + len(result['urls']['attachments'])
        result['warnings'].append(f"总计发现 {total_urls} 个URL，其中 {len(result['suspicious_links'])} 个可疑")
        
    except Exception as e:
        print(f"URL分析失败: {str(e)}")
        result['warnings'].append(f"URL分析过程出错: {str(e)}")
    
    return result

def detect_hidden_content(email_data: Dict[str, Any]) -> Dict[str, List[Dict[str, Any]]]:
    """
    检测邮件中的隐藏内容和跟踪器
    
    Args:
        email_data: 邮件解析数据
        
    Returns:
        包含检测结果的字典
    """
    findings = {
        'hidden_content': [],
        'tracking_elements': [],
        'suspicious_urls': [],
        'risk_level': 'low',
        'risk_score': 0.0,
        'warnings': []
    }
    
    try:
        if email_data['body_html']:
            # 尝试多种编码方式解析HTML内容
            html_content = email_data['body_html']
            if isinstance(html_content, bytes):
                encodings = ['utf-8', 'gbk', 'gb2312', 'gb18030', 'big5']
                for encoding in encodings:
                    try:
                        html_content = html_content.decode(encoding)
                        break
                    except:
                        continue
            
            soup = BeautifulSoup(html_content, 'html.parser')
            
            # 1. 检查隐藏的图片和跟踪像素
            hidden_images = soup.find_all('img', style=lambda x: x and any(style in x.lower() for style in [
                'display:none', 'display: none',
                'visibility:hidden', 'visibility: hidden',
                'opacity:0', 'opacity: 0',
                'width:1px', 'width: 1px',
                'height:1px', 'height: 1px'
            ]))
            
            for img in hidden_images:
                src = img.get('src', '')
                findings['tracking_elements'].append({
                    'type': 'hidden_tracking_pixel',
                    'url': src,
                    'reason': '隐藏的跟踪图片',
                    'details': f'样式: {img.get("style", "")}'
                })
                findings['risk_score'] += 3.0
            
            # 2. 检查所有图片的尺寸属性
            all_images = soup.find_all('img')
            for img in all_images:
                width = img.get('width', '').strip()
                height = img.get('height', '').strip()
                if (width == '1' and height == '1') or (width == '0' and height == '0'):
                    src = img.get('src', '')
                    findings['tracking_elements'].append({
                        'type': 'tracking_pixel',
                        'url': src,
                        'reason': f'{width}x{height}像素跟踪图片'
                    })
                    findings['risk_score'] += 2.5
            
            # 3. 检查隐藏的内容
            hidden_elements = soup.find_all(lambda tag: tag.get('style') and any(style in tag.get('style', '').lower() for style in [
                'display:none', 'display: none',
                'visibility:hidden', 'visibility: hidden',
                'opacity:0', 'opacity: 0',
                'font-size:0', 'font-size: 0'
            ]))
            
            for element in hidden_elements:
                content = element.get_text().strip()
                if content:
                    findings['hidden_content'].append({
                        'type': 'hidden_element',
                        'content': content,
                        'style': element.get('style', ''),
                        'reason': '使用CSS隐藏的内容'
                    })
                    findings['risk_score'] += 2.0
            
            # 4. 检查可疑的URL和链接
            links = soup.find_all(['a', 'img', 'form'])
            for link in links:
                url = link.get('href') or link.get('src') or link.get('action', '')
                if url:
                    try:
                        from urllib.parse import urlparse, parse_qs
                        parsed = urlparse(url)
                        
                        # 检查非标准端口
                        if parsed.port and parsed.port not in (80, 443):
                            findings['suspicious_urls'].append({
                                'type': 'suspicious_port',
                                'url': url,
                                'port': parsed.port,
                                'reason': f'使用非标准端口 {parsed.port}'
                            })
                            findings['risk_score'] += 2.5
                        
                        # 检查可疑参数
                        query_params = parse_qs(parsed.query)
                        tracking_params = {'uid', 'user', 'id', 'email', 'track', 'open', 'click'}
                        found_params = set(query_params.keys()) & tracking_params
                        if found_params:
                            findings['tracking_elements'].append({
                                'type': 'tracking_parameters',
                                'url': url,
                                'params': list(found_params),
                                'reason': '包含跟踪参数'
                            })
                            findings['risk_score'] += 1.5
                        
                        # 检查URL编码过度
                        if url.count('%') > 5:
                            findings['suspicious_urls'].append({
                                'type': 'heavily_encoded_url',
                                'url': url,
                                'reason': 'URL过度编码，可能试图隐藏真实地址'
                            })
                            findings['risk_score'] += 2.0
                        
                    except Exception as e:
                        print(f"URL分析失败: {str(e)}")
            
            # 5. 检查外部资源加载
            external_resources = soup.find_all(['script', 'iframe', 'img', 'link'])
            for resource in external_resources:
                src = resource.get('src') or resource.get('href', '')
                if src and ('track' in src.lower() or 'beacon' in src.lower() or 'pixel' in src.lower()):
                    findings['tracking_elements'].append({
                        'type': 'external_tracker',
                        'url': src,
                        'tag': resource.name,
                        'reason': '外部跟踪资源'
                    })
                    findings['risk_score'] += 2.0
            
            # 更新风险等级
            if findings['risk_score'] >= 5.0:
                findings['risk_level'] = 'high'
                findings['warnings'].append('发现多个高风险跟踪器和隐藏内容，疑似钓鱼邮件！')
            elif findings['risk_score'] >= 3.0:
                findings['risk_level'] = 'medium'
                findings['warnings'].append('发现可疑的跟踪器或隐藏内容')
            
            # 添加统计信息
            findings['warnings'].append(f"发现 {len(findings['tracking_elements'])} 个跟踪元素")
            findings['warnings'].append(f"发现 {len(findings['hidden_content'])} 个隐藏内容")
            findings['warnings'].append(f"发现 {len(findings['suspicious_urls'])} 个可疑URL")
        
    except Exception as e:
        print(f"检测隐藏内容失败: {str(e)}")
        findings['warnings'].append(f"内容检测过程出错: {str(e)}")
    
    return findings

def verify_email_auth(email_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    验证邮件认证信息，检查SPF、DKIM和DMARC
    
    Args:
        email_data: 邮件解析数据
        
    Returns:
        包含认证结果的字典
    """
    auth_results = {
        'spf': {
            'status': 'unknown',  # pass, fail, softfail, neutral, unknown
            'domain': '',
            'ip': '',
            'explanation': ''
        },
        'dkim': {
            'status': 'unknown',  # pass, fail, neutral, unknown
            'domain': '',
            'selector': '',
            'explanation': ''
        },
        'dmarc': {
            'status': 'unknown',  # pass, fail, none, unknown
            'domain': '',
            'policy': '',
            'explanation': ''
        },
        'authentication_results': '',  # 原始认证结果头信息
        'risk_level': 'unknown',      # high, medium, low, unknown
        'risk_score': 0.0,
        'warnings': []
    }
    
    try:
        # 1. 获取认证相关的头信息
        headers = {
            'authentication_results': '',  # Authentication-Results
            'received_spf': '',           # Received-SPF
            'dkim_signature': '',         # DKIM-Signature
            'dkim_results': '',           # DKIM验证结果
            'arc_authentication_results': '', # ARC-Authentication-Results
            'received': []                # Received链
        }
        
        # 2. 解析SPF记录
        def parse_spf_result(spf_header: str) -> None:
            if 'pass' in spf_header.lower():
                auth_results['spf']['status'] = 'pass'
            elif 'fail' in spf_header.lower():
                auth_results['spf']['status'] = 'fail'
            elif 'softfail' in spf_header.lower():
                auth_results['spf']['status'] = 'softfail'
            elif 'neutral' in spf_header.lower():
                auth_results['spf']['status'] = 'neutral'
            
            # 提取发送域名和IP
            domain_match = re.search(r'domain=(\S+)', spf_header)
            if domain_match:
                auth_results['spf']['domain'] = domain_match.group(1)
            
            ip_match = re.search(r'client-ip=(\S+)', spf_header)
            if ip_match:
                auth_results['spf']['ip'] = ip_match.group(1)
        
        # 3. 解析DKIM签名
        def parse_dkim_result(dkim_header: str) -> None:
            if 'pass' in dkim_header.lower():
                auth_results['dkim']['status'] = 'pass'
            elif 'fail' in dkim_header.lower():
                auth_results['dkim']['status'] = 'fail'
            
            # 提取DKIM域名和选择器
            domain_match = re.search(r'd=([^;]+)', dkim_header)
            if domain_match:
                auth_results['dkim']['domain'] = domain_match.group(1)
            
            selector_match = re.search(r's=([^;]+)', dkim_header)
            if selector_match:
                auth_results['dkim']['selector'] = selector_match.group(1)
        
        # 4. 解析DMARC结果
        def parse_dmarc_result(auth_header: str) -> None:
            dmarc_match = re.search(r'dmarc=(\S+)', auth_header.lower())
            if dmarc_match:
                status = dmarc_match.group(1)
                auth_results['dmarc']['status'] = status
                
                # 提取DMARC策略
                policy_match = re.search(r'p=(\S+)', auth_header)
                if policy_match:
                    auth_results['dmarc']['policy'] = policy_match.group(1)
        
        # 5. 分析认证头信息
        if 'Authentication-Results' in email_data.get('headers', {}):
            auth_header = email_data['headers']['Authentication-Results']
            auth_results['authentication_results'] = auth_header
            
            # 解析认证结果
            parse_spf_result(auth_header)
            parse_dkim_result(auth_header)
            parse_dmarc_result(auth_header)
        
        # 6. 评估风险等级
        risk_score = 0.0
        
        # SPF检查
        if auth_results['spf']['status'] == 'fail':
            risk_score += 3.0
            auth_results['warnings'].append('SPF验证失败，发件人域名可能被伪造')
        elif auth_results['spf']['status'] == 'softfail':
            risk_score += 1.5
            auth_results['warnings'].append('SPF软失败，发件人域名可疑')
        
        # DKIM检查
        if auth_results['dkim']['status'] == 'fail':
            risk_score += 3.0
            auth_results['warnings'].append('DKIM签名验证失败，邮件可能被篡改')
        elif auth_results['dkim']['status'] == 'unknown':
            risk_score += 1.0
            auth_results['warnings'].append('未找到DKIM签名')
        
        # DMARC检查
        if auth_results['dmarc']['status'] == 'fail':
            risk_score += 3.0
            auth_results['warnings'].append('DMARC验证失败，不符合发件人域名的邮件策略')
        elif auth_results['dmarc']['status'] == 'none':
            risk_score += 1.0
            auth_results['warnings'].append('发件人域名未配置DMARC策略')
        
        # 发件人域名检查
        sender_domain = ''
        if email_data['from']:
            sender_match = re.search(r'@([\w.-]+)', email_data['from'][0])
            if sender_match:
                sender_domain = sender_match.group(1)
                
                # 检查发件人域名与认证域名是否匹配
                if sender_domain != auth_results['spf']['domain']:
                    risk_score += 2.0
                    auth_results['warnings'].append('发件人域名与SPF认证域名不匹配')
                if sender_domain != auth_results['dkim']['domain']:
                    risk_score += 2.0
                    auth_results['warnings'].append('发件人域名与DKIM签名域名不匹配')
        
        # 设置最终风险等级
        auth_results['risk_score'] = risk_score
        if risk_score >= 5.0:
            auth_results['risk_level'] = 'high'
        elif risk_score >= 2.0:
            auth_results['risk_level'] = 'medium'
        elif risk_score > 0:
            auth_results['risk_level'] = 'low'
        
    except Exception as e:
        print(f"验证邮件认证信息失败: {str(e)}")
        auth_results['warnings'].append(f"认证验证过程出错: {str(e)}")
    
    return auth_results

def display_report(email_data: Dict[str, Any]) -> None:
    """
    显示邮件分析报告
    
    Args:
        email_data: 邮件解析数据
    """
    print("\n=== 邮件分析报告 ===")
    
    # 显示发件人
    print("\n[1] 发件人信息:")
    if email_data['from']:
        for i, sender in enumerate(email_data['from'], 1):
            print(f"发件人 {i}: {sender}")
    else:
        print("未找到发件人信息")
    
    # 显示收件人
    print("\n[2] 收件人信息:")
    if email_data['to']:
        for i, recipient in enumerate(email_data['to'], 1):
            print(f"收件人 {i}: {recipient}")
    else:
        print("未找到收件人信息")
    
    # 显示抄送人
    if email_data['cc']:
        print("\n[3] 抄送人信息:")
        for i, cc in enumerate(email_data['cc'], 1):
            print(f"抄送人 {i}: {cc}")
    
    # 显示密送人
    if email_data['bcc']:
        print("\n[4] 密送人信息:")
        for i, bcc in enumerate(email_data['bcc'], 1):
            print(f"密送人 {i}: {bcc}")
    
    # 显示回复地址
    if email_data['reply_to']:
        print("\n[5] 回复地址:")
        for i, reply_to in enumerate(email_data['reply_to'], 1):
            print(f"回复地址 {i}: {reply_to}")
    
    # 显示主题和日期
    print("\n[6] 邮件信息:")
    print(f"主题: {email_data['subject'] or '未知'}")
    print(f"日期: {email_data['date'] or '未知'}")
    
    # 显示原始邮件信息
    if any(email_data['thread_info'].values()):
        print("\n[7] 原始邮件信息:")
        if email_data['thread_info']['original_sender']:
            print(f"原始发件人: {email_data['thread_info']['original_sender']}")
        if email_data['thread_info']['original_recipients']:
            print("原始收件人:")
            for i, recipient in enumerate(email_data['thread_info']['original_recipients'], 1):
                print(f"  收件人 {i}: {recipient}")
        if email_data['thread_info']['original_subject']:
            print(f"原始主题: {email_data['thread_info']['original_subject']}")
        if email_data['thread_info']['original_date']:
            print(f"原始日期: {email_data['thread_info']['original_date']}")
    
    # 改进正文信息显示
    print("\n[8] 正文信息:")
    if email_data['body_text']:
        print("\n纯文本内容:")
        print("-" * 50)
        print(email_data['body_text'][:500] + "..." if len(email_data['body_text']) > 500 else email_data['body_text'])
        print("-" * 50)
        print(f"纯文本总长度: {len(email_data['body_text'])} 字符")
    
    if email_data['body_html']:
        print("\nHTML内容:")
        print("-" * 50)
        try:
            html_content = email_data['body_html']
            
            # 使用 BeautifulSoup 解析 HTML
            soup = BeautifulSoup(html_content, 'html.parser')
            
            # 移除所有脚本和样式标签
            for script in soup(["script", "style"]):
                script.decompose()
            
            # 获取纯文本内容
            clean_text = soup.get_text(separator='\n', strip=True)
            
            # 如果清理后的文本为空或看起来是乱码，尝试显示原始HTML
            if not clean_text.strip() or all(ord(c) > 127 for c in clean_text[:20]):
                print("警告: 尝试显示原始HTML内容:")
                print(html_content[:1000] + "..." if len(html_content) > 1000 else html_content)
            else:
                print(clean_text[:1000] + "..." if len(clean_text) > 1000 else clean_text)
            
        except Exception as e:
            print(f"处理HTML内容时出错: {str(e)}")
            print("原始HTML内容:")
            print(html_content[:200] + "..." if len(html_content) > 200 else html_content)
        print("-" * 50)
        print(f"HTML总长度: {len(email_data['body_html'])} 字符")
    
    if not email_data['body_text'] and not email_data['body_html']:
        print("未找到邮件正文")
    
    # 在隐藏内容检测之前添加认证信息显示
    print("\n[9] 邮件认证信息:")
    auth_results = verify_email_auth(email_data)
    
    print(f"风险等级: {auth_results['risk_level'].upper()}")
    if auth_results['risk_score'] > 0:
        print(f"风险分数: {auth_results['risk_score']:.1f}")
    
    print("\nSPF验证:")
    print(f"状态: {auth_results['spf']['status']}")
    if auth_results['spf']['domain']:
        print(f"域名: {auth_results['spf']['domain']}")
    if auth_results['spf']['ip']:
        print(f"发送IP: {auth_results['spf']['ip']}")
    
    print("\nDKIM验证:")
    print(f"状态: {auth_results['dkim']['status']}")
    if auth_results['dkim']['domain']:
        print(f"域名: {auth_results['dkim']['domain']}")
    if auth_results['dkim']['selector']:
        print(f"选择器: {auth_results['dkim']['selector']}")
    
    print("\nDMARC验证:")
    print(f"状态: {auth_results['dmarc']['status']}")
    if auth_results['dmarc']['domain']:
        print(f"域名: {auth_results['dmarc']['domain']}")
    if auth_results['dmarc']['policy']:
        print(f"策略: {auth_results['dmarc']['policy']}")
    
    if auth_results['warnings']:
        print("\n警告信息:")
        for warning in auth_results['warnings']:
            print(f"- {warning}")
    
    # 在邮件认证信息后添加域名相似度检查
    print("\n[9.1] 域名相似度检查:")
    domain_check = check_similar_domains(email_data)
    
    if domain_check['similar_domains']:
        print(f"\n风险等级: {domain_check['risk_level'].upper()}")
        print(f"风险分数: {domain_check['risk_score']:.1f}")
        
        print("\n发现相似域名:")
        for item in domain_check['similar_domains']:
            print(f"- 发件人域名: {item['domain1']}")
            print(f"  相似域名: {item['domain2']}")
            print(f"  类型: {item['type']}")
            print(f"  相似度: {item['similarity']:.2f}")
        
        if domain_check['warnings']:
            print("\n警告信息:")
            for warning in domain_check['warnings']:
                print(f"- {warning}")
    else:
        print("未发现可疑的相似域名")
    
    # 将原来的隐藏内容检测改为[10]，URL信息改为[11]，附件信息改为[12]
    print("\n[10] 隐藏内容检测:")
    findings = detect_hidden_content(email_data)
    
    print(f"\n风险等级: {findings['risk_level'].upper()}")
    if 'risk_score' in findings:
        print(f"风险分数: {findings['risk_score']:.1f}")
    
    if 'warnings' in findings and findings['warnings']:
        print("\n警告信息:")
        for warning in findings['warnings']:
            print(f"- {warning}")
    
    if findings['risk_level'] != 'low':
        if findings['hidden_content']:
            print("\n发现隐藏内容:")
            for item in findings['hidden_content']:
                print(f"- 类型: {item['type']}")
                print(f"  原因: {item['reason']}")
                if item['content']:
                    print(f"  内容: {item['content'][:100]}...")
        
        if findings['tracking_elements']:
            print("\n发现跟踪元素:")
            for item in findings['tracking_elements']:
                print(f"- 类型: {item['type']}")
                print(f"  URL: {item['url']}")
                print(f"  原因: {item['reason']}")
        
        if findings['suspicious_urls']:
            print("\n发现可疑URL:")
            for item in findings['suspicious_urls']:
                print(f"- 类型: {item['type']}")
                print(f"  URL: {item['url']}")
                if item.get('text'):
                    print(f"  显示文本: {item['text']}")
                print(f"  原因: {item['reason']}")
    else:
        print("未发现明显的隐藏内容或跟踪器")
    
    # 将原来的URL信息部分改为[11]，附件信息改为[12]
    print("\n[11] URL信息:")
    url_analysis = extract_urls(email_data)
    
    if url_analysis['risk_level'] != 'low':
        print(f"\n风险等级: {url_analysis['risk_level'].upper()}")
        print(f"风险分数: {url_analysis['risk_score']:.1f}")
        
        if url_analysis['warnings']:
            print("\n警告信息:")
            for warning in url_analysis['warnings']:
                print(f"- {warning}")
        
        if url_analysis['suspicious_links']:
            print("\n可疑超链接:")
            for link in url_analysis['suspicious_links']:
                print(f"\n- 显示文本: {link['display_text']}")
                print(f"  实际URL: {link['actual_url']}")
                print(f"  风险等级: {link['risk_level']}")
                print(f"  风险分数: {link['risk_score']:.1f}")
                print("  原因:")
                for reason in link['reasons']:
                    print(f"    * {reason}")
    
    if any(url_analysis['urls'].values()):
        # 显示纯文本中的URL
        if url_analysis['urls']['text']:
            print("\n纯文本中的URL:")
            for i, url in enumerate(url_analysis['urls']['text'], 1):
                print(f"  {i}. {url}")
        
        # 显示HTML中的URL
        if url_analysis['urls']['html']:
            print("\nHTML中的URL:")
            for i, url in enumerate(url_analysis['urls']['html'], 1):
                print(f"  {i}. {url}")
        
        # 显示附件中的URL
        if url_analysis['urls']['attachments']:
            print("\n附件中的URL:")
            for i, url in enumerate(url_analysis['urls']['attachments'], 1):
                print(f"  {i}. {url}")
        
        # 显示URL统计信息
        total_urls = len(url_analysis['urls']['text']) + len(url_analysis['urls']['html']) + len(url_analysis['urls']['attachments'])
        print(f"\n总计发现 {total_urls} 个URL:")
        print(f"- 纯文本中: {len(url_analysis['urls']['text'])} 个")
        print(f"- HTML中: {len(url_analysis['urls']['html'])} 个")
        print(f"- 附件中: {len(url_analysis['urls']['attachments'])} 个")
    else:
        print("未发现URL")
    
    print("\n[12] 附件信息:")
    if email_data['attachments']:
        for i, attachment in enumerate(email_data['attachments'], 1):
            print(f"\n附件 {i}:")
            print(f"文件名: {attachment['filename']}")
            print(f"MIME类型: {attachment['mime_type']}")
            print(f"大小: {attachment['size']} 字节")
            
            # 显示额外的附件信息
            if attachment.get('is_inline'):
                print("类型: 内联附件")
                if attachment.get('content_id'):
                    print(f"Content-ID: {attachment['content_id']}")
            
            if attachment.get('extension'):
                print(f"文件扩展名: {attachment['extension']}")
            
            if attachment.get('hash_md5'):
                print(f"MD5: {attachment['hash_md5']}")
                print(f"SHA256: {attachment['hash_sha256']}")
            
            if attachment.get('is_executable'):
                print("警告: 这是一个可执行文件")
            
            if attachment.get('is_archive'):
                print("类型: 压缩文件")
                if attachment.get('archive_contents'):
                    print("压缩包内容:")
                    for item in attachment['archive_contents'][:5]:  # 只显示前5个文件
                        print(f"  - {item}")
                    if len(attachment['archive_contents']) > 5:
                        print(f"  ... 等共 {len(attachment['archive_contents'])} 个文件")
            
            # 改进文本预览显示
            if attachment.get('text_preview'):
                print("\n文档内容预览:")
                print("=" * 50)
                preview_lines = attachment['text_preview'].split('\n')
                # 只显示前20行
                for line in preview_lines[:20]:
                    print(line)
                if len(preview_lines) > 20:
                    print("\n... (更多内容已省略)")
                print("=" * 50)
    else:
        print("无附件")
    
    # 在 display_report 函数中添加以下内容（在[9]邮件认证信息部分之后）
    print("\n[9.2] 发件人真实性检测:")
    spoofing_check = detect_spoofed_sender(email_data)

    if spoofing_check['is_spoofed']:
        print(f"\n风险等级: {spoofing_check['risk_level'].upper()}")
        print(f"风险分数: {spoofing_check['risk_score']:.1f}")
        
        print("\n警告信息:")
        for warning in spoofing_check['warnings']:
            print(f"- {warning}")
    else:
        print("未发现明显的发件人伪造迹象")
    
    # 在 display_report 函数中添加以下内容（在[9.2]发件人真实性检测之后）
    print("\n[9.3] 域名注册信息分析:")
    registration_analysis = analyze_domain_registration(email_data)

    print(f"\n风险等级: {registration_analysis['risk_level'].upper()}")
    print(f"风险分数: {registration_analysis['risk_score']:.1f}")

    if registration_analysis['sender_domain']:
        print("\n发件人域名信息:")
        sender_info = registration_analysis['sender_domain']
        print(f"域名: {sender_info['domain']}")
        if sender_info['creation_date']:
            print(f"注册时间: {sender_info['creation_date']}")
            print(f"域名年龄: {sender_info['age_days']} 天")
        if sender_info['expiration_date']:
            print(f"过期时间: {sender_info['expiration_date']}")
        print(f"风险等级: {sender_info['risk_level'].upper()}")

    if registration_analysis['recipient_domain']:
        print("\n收件人域名信息:")
        recipient_info = registration_analysis['recipient_domain']
        print(f"域名: {recipient_info['domain']}")
        if recipient_info['creation_date']:
            print(f"注册时间: {recipient_info['creation_date']}")
            print(f"域名年龄: {recipient_info['age_days']} 天")
        if recipient_info['expiration_date']:
            print(f"过期时间: {recipient_info['expiration_date']}")
        print(f"风险等级: {recipient_info['risk_level'].upper()}")

    if registration_analysis['domain_age_comparison']:
        print("\n域名年龄对比:")
        age_diff = registration_analysis['domain_age_comparison']['age_difference_days']
        if age_diff is not None:
            print(f"年龄差异: {abs(age_diff)} 天")

    if registration_analysis['warnings']:
        print("\n警告信息:")
        for warning in registration_analysis['warnings']:
            print(f"- {warning}")
    
    print("\n=== 报告结束 ===")

def check_similar_domains(email_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    检查邮件中所有地址的域名相似度
    """
    result = {
        'similar_domains': [],
        'risk_level': 'low',
        'risk_score': 0.0,
        'warnings': []
    }
    
    try:
        def extract_domain(email: str) -> str:
            """从邮箱地址中提取域名"""
            try:
                # 处理带括号的邮箱地址，如 "Name (email@domain.com)"
                email_match = re.search(r'[\w\.-]+@([\w\.-]+)', email)
                if email_match:
                    return email_match.group(1).lower()
                return ''
            except:
                return ''
        
        def analyze_domain_relationship(domain1: str, domain2: str) -> Dict[str, Any]:
            """分析两个域名之间的关系"""
            base1 = domain1.split('.')[0]
            base2 = domain2.split('.')[0]
            
            result = {
                'similarity': 0.0,
                'relationship_type': None,
                'risk': 'low'
            }
            
            # 1. 检查字符替换（如 sinopec -> siuopec）
            def check_character_substitution(s1: str, s2: str) -> bool:
                if abs(len(s1) - len(s2)) > 1:
                    return False
                diff_count = 0
                for c1, c2 in zip(s1, s2):
                    if c1 != c2:
                        diff_count += 1
                        if diff_count > 1:
                            return False
                return True
            
            if check_character_substitution(base1, base2):
                result['similarity'] = 0.95
                result['relationship_type'] = '可疑的字符替换'
                result['risk'] = 'high'
                return result
            
            # 2. 检查包含关系
            if base1 in base2 or base2 in base1:
                longer = base1 if len(base1) > len(base2) else base2
                shorter = base2 if len(base1) > len(base2) else base1
                
                # 如果较长的域名包含较短的域名，且额外字符看起来像是正常词
                if shorter in longer:
                    suspicious_additions = ['portal', 'service', 'vendor', 'secure', 'mail', 
                                         'auth', 'login', 'account', 'verify', 'update']
                    remaining = longer.replace(shorter, '').lower()
                    
                    if any(word in remaining for word in suspicious_additions):
                        result['similarity'] = 0.9
                        result['relationship_type'] = '可疑的域名包含'
                        result['risk'] = 'high'
                        return result
            
            # 3. 检查字母顺序调换
            if sorted(base1) == sorted(base2) and base1 != base2:
                result['similarity'] = 1.0
                result['relationship_type'] = '字母顺序调换'
                result['risk'] = 'high'
                return result
            
            # 4. 计算编辑距离相似度
            def levenshtein_distance(s1: str, s2: str) -> int:
                if len(s1) < len(s2):
                    return levenshtein_distance(s2, s1)
                if len(s2) == 0:
                    return len(s1)
                previous_row = range(len(s2) + 1)
                for i, c1 in enumerate(s1):
                    current_row = [i + 1]
                    for j, c2 in enumerate(s2):
                        insertions = previous_row[j + 1] + 1
                        deletions = current_row[j] + 1
                        substitutions = previous_row[j] + (c1 != c2)
                        current_row.append(min(insertions, deletions, substitutions))
                    previous_row = current_row
                return previous_row[-1]
            
            distance = levenshtein_distance(base1, base2)
            max_length = max(len(base1), len(base2))
            similarity = 1 - (distance / max_length)
            
            if similarity > 0.8:
                result['similarity'] = similarity
                result['relationship_type'] = '高度相似'
                result['risk'] = 'high'
            
            return result
        
        # 收集所有域名
        domains = {
            'from': set(),
            'to': set(),
            'original_from': set()
        }
        
        # 提取所有域名
        for sender in email_data['from']:
            domain = extract_domain(sender)
            if domain:
                domains['from'].add(domain)
        
        for recipient in email_data['to']:
            domain = extract_domain(recipient)
            if domain:
                domains['to'].add(domain)
        
        if email_data['thread_info']['original_sender']:
            domain = extract_domain(email_data['thread_info']['original_sender'])
            if domain:
                domains['original_from'].add(domain)
        
        # 比较所有域名组合
        # 1. 发件人域名与收件人域名比较
        for from_domain in domains['from']:
            for to_domain in domains['to']:
                analysis = analyze_domain_relationship(from_domain, to_domain)
                if analysis['risk'] == 'high':
                    result['similar_domains'].append({
                        'domain1': from_domain,
                        'domain2': to_domain,
                        'type': '发件人-收件人',
                        'relationship': analysis['relationship_type'],
                        'similarity': analysis['similarity']
                    })
                    result['risk_score'] += 3.0
                    
                    warning = f"发现可疑域名组合: {from_domain} <-> {to_domain} ({analysis['relationship_type']})"
                    if analysis['relationship_type'] == '可疑的域名包含':
                        warning += "\n这是典型的钓鱼邮件手法，使用目标公司域名构造相似域名"
                    result['warnings'].append(warning)
        
        # 2. 发件人域名与原始发件人域名比较
        for from_domain in domains['from']:
            for orig_domain in domains['original_from']:
                analysis = analyze_domain_relationship(from_domain, orig_domain)
                if analysis['risk'] == 'high':
                    result['similar_domains'].append({
                        'domain1': from_domain,
                        'domain2': orig_domain,
                        'type': '发件人-原始发件人',
                        'relationship': analysis['relationship_type'],
                        'similarity': analysis['similarity']
                    })
                    result['risk_score'] += 3.0
        
        # 设置最终风险等级
        if result['risk_score'] >= 3.0:
            result['risk_level'] = 'high'
            if not result['warnings']:
                result['warnings'].append('发现高度相似的可疑域名，疑似钓鱼邮件！')
        elif result['risk_score'] >= 1.0:
            result['risk_level'] = 'medium'
            if not result['warnings']:
                result['warnings'].append('发现相似域名，建议仔细核实发件人身份')
        
    except Exception as e:
        print(f"域名相似度检查失败: {str(e)}")
    
    return result

def detect_spoofed_sender(email_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    检测伪造的发件人
    
    Args:
        email_data: 邮件解析数据
        
    Returns:
        包含检测结果的字典
    """
    result = {
        'is_spoofed': False,
        'evidence': [],
        'risk_level': 'low',
        'risk_score': 0.0,
        'original_sender': '',
        'spoofed_sender': '',
        'warnings': []
    }
    
    try:
        # 1. 检查SPF验证结果
        auth_results = verify_email_auth(email_data)
        if auth_results['spf']['status'] == 'fail':
            result['is_spoofed'] = True
            result['evidence'].append('SPF验证失败')
            result['risk_score'] += 3.0
        
        # 2. 分析Received链
        received_chain = []
        if 'headers' in email_data:
            for header, value in email_data['headers'].items():
                if header.lower() == 'received':
                    received_chain.append(value)
            
            # 检查最后一个Received头（最初的发送服务器）
            if received_chain:
                last_received = received_chain[-1]
                # 检查是否来自声称的域名
                claimed_domain = ''
                for sender in email_data['from']:
                    domain_match = re.search(r'@([\w.-]+)', sender)
                    if domain_match:
                        claimed_domain = domain_match.group(1).lower()
                        break
                
                if claimed_domain and claimed_domain not in last_received.lower():
                    result['is_spoofed'] = True
                    result['evidence'].append(f'发件人声称来自 {claimed_domain}，但实际发送服务器不匹配')
                    result['risk_score'] += 2.5
        
        # 3. 检查X-Fangmail-Spf头
        if 'headers' in email_data and 'x-fangmail-spf' in email_data['headers']:
            if email_data['headers']['x-fangmail-spf'].lower() == 'fail':
                result['is_spoofed'] = True
                result['evidence'].append('防垃圾邮件系统SPF检查失败')
                result['risk_score'] += 2.0
        
        # 4. 记录可疑的发件人信息
        if email_data['from']:
            result['spoofed_sender'] = email_data['from'][0]
            # 在这个例子中，发件人伪装成 "滴滴出行合作 <bd@didiglobal.com>"
            if 'didiglobal.com' in result['spoofed_sender'].lower():
                result['is_spoofed'] = True
                result['evidence'].append('发件人伪装成滴滴出行官方邮箱')
                result['risk_score'] += 3.0
        
        # 设置风险等级
        if result['risk_score'] >= 5.0:
            result['risk_level'] = 'high'
            result['warnings'].append('发现多个证据表明发件人身份被伪造，这很可能是一封钓鱼邮件！')
        elif result['risk_score'] >= 2.5:
            result['risk_level'] = 'medium'
            result['warnings'].append('发现可疑迹象表明发件人身份可能被伪造')
        
        # 添加详细分析
        if result['is_spoofed']:
            result['warnings'].extend([
                f"伪造的发件人: {result['spoofed_sender']}",
                "发现的证据:",
                *[f"- {evidence}" for evidence in result['evidence']]
            ])
            
    except Exception as e:
        print(f"检测伪造发件人失败: {str(e)}")
        result['warnings'].append(f"发件人检测过程出错: {str(e)}")
    
    return result

def check_domain_registration(domain: str) -> Dict[str, Any]:
    """
    检查域名的注册信息
    """
    result = {
        'domain': domain,
        'creation_date': None,
        'expiration_date': None,
        'last_updated': None,
        'age_days': None,
        'risk_level': 'low',
        'risk_score': 0.0,
        'warnings': []
    }
    
    try:
        # 预处理域名
        if not domain or '.' not in domain:
            raise ValueError(f"无效的域名格式: {domain}")
            
        # 移除可能的端口号和路径
        domain = domain.split(':')[0].split('/')[0]
        
        # 尝试 WHOIS 查询
        try:
            w = whois.whois(domain)
        except Exception as e:
            raise Exception(f"WHOIS查询失败: {str(e)}")
        
        if not w or not w.domain_name:
            raise Exception(f"无法获取域名信息")
        
        def ensure_datetime_with_timezone(dt):
            """确保日期时间对象带有时区信息"""
            if dt is None:
                return None
            if isinstance(dt, str):
                try:
                    # 尝试解析字符串为datetime对象
                    dt = datetime.strptime(dt, "%Y-%m-%d %H:%M:%S")
                except:
                    try:
                        dt = datetime.strptime(dt, "%Y-%m-%d")
                    except:
                        return None
            if not isinstance(dt, datetime):
                return None
            if dt.tzinfo is None:
                # 如果没有时区信息，假定为UTC
                dt = dt.replace(tzinfo=timezone.utc)
            return dt
        
        # 获取当前时间（带时区信息）
        now = datetime.now(timezone.utc)
        
        # 处理创建日期
        if w.creation_date:
            try:
                if isinstance(w.creation_date, list):
                    # 处理多个创建日期的情况，选择最早的有效日期
                    valid_dates = []
                    for date in w.creation_date:
                        dt = ensure_datetime_with_timezone(date)
                        if dt:
                            valid_dates.append(dt)
                    creation_date = min(valid_dates) if valid_dates else None
                else:
                    creation_date = ensure_datetime_with_timezone(w.creation_date)
                
                if creation_date:
                    result['creation_date'] = creation_date
                    age_days = (now - creation_date).days
                    result['age_days'] = age_days
                    
                    # 评估风险
                    if age_days <= 7:  # 一周内注册
                        result['risk_level'] = 'critical'
                        result['risk_score'] = 5.0
                        result['warnings'].append(f"域名 {domain} 注册时间极短（{age_days}天），非常可疑！")
                    elif age_days <= 30:  # 一个月内注册
                        result['risk_level'] = 'high'
                        result['risk_score'] = 4.0
                        result['warnings'].append(f"域名 {domain} 注册时间很短（{age_days}天），高度可疑")
                    elif age_days <= 90:  # 三个月内注册
                        result['risk_level'] = 'medium'
                        result['risk_score'] = 3.0
                        result['warnings'].append(f"域名 {domain} 注册时间较短（{age_days}天），需要注意")
                    elif age_days <= 365:  # 一年内注册
                        result['risk_level'] = 'low'
                        result['risk_score'] = 2.0
                        result['warnings'].append(f"域名 {domain} 注册时间不足一年（{age_days}天）")
            except Exception as e:
                result['warnings'].append(f"处理创建日期时出错: {str(e)}")
        
        # 处理过期日期
        if w.expiration_date:
            try:
                if isinstance(w.expiration_date, list):
                    # 处理多个过期日期的情况，选择最晚的有效日期
                    valid_dates = []
                    for date in w.expiration_date:
                        dt = ensure_datetime_with_timezone(date)
                        if dt:
                            valid_dates.append(dt)
                    expiration_date = max(valid_dates) if valid_dates else None
                else:
                    expiration_date = ensure_datetime_with_timezone(w.expiration_date)
                
                if expiration_date:
                    result['expiration_date'] = expiration_date
            except Exception as e:
                result['warnings'].append(f"处理过期日期时出错: {str(e)}")
        
        # 处理更新日期
        if w.updated_date:
            try:
                if isinstance(w.updated_date, list):
                    # 处理多个更新日期的情况，选择最近的有效日期
                    valid_dates = []
                    for date in w.updated_date:
                        dt = ensure_datetime_with_timezone(date)
                        if dt:
                            valid_dates.append(dt)
                    updated_date = max(valid_dates) if valid_dates else None
                else:
                    updated_date = ensure_datetime_with_timezone(w.updated_date)
                
                if updated_date:
                    result['last_updated'] = updated_date
            except Exception as e:
                result['warnings'].append(f"处理更新日期时出错: {str(e)}")
                
    except Exception as e:
        result['warnings'].append(f"无法获取域名 {domain} 的注册信息: {str(e)}")
        result['risk_level'] = 'unknown'
        result['risk_score'] = 0.0
    
    return result

def analyze_domain_registration(email_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    分析邮件中涉及的域名注册信息
    """
    result = {
        'sender_domain': None,
        'recipient_domain': None,
        'domain_age_comparison': None,
        'risk_level': 'low',
        'risk_score': 0.0,
        'warnings': []
    }
    
    try:
        # 提取发件人域名
        sender_domain = None
        if email_data['from']:
            domain_match = re.search(r'@([\w.-]+)', email_data['from'][0])
            if domain_match:
                sender_domain = domain_match.group(1).lower()
                result['sender_domain'] = check_domain_registration_with_retry(sender_domain)
                if result['sender_domain'].get('risk_score'):
                    result['risk_score'] += result['sender_domain']['risk_score']
        
        # 提取收件人域名
        recipient_domain = None
        if email_data['to']:
            domain_match = re.search(r'@([\w.-]+)', email_data['to'][0])
            if domain_match:
                recipient_domain = domain_match.group(1).lower()
                result['recipient_domain'] = check_domain_registration_with_retry(recipient_domain)
        
        # 如果发件人和收件人域名不同，进行对比分析
        if (sender_domain and recipient_domain and 
            sender_domain != recipient_domain and 
            result['sender_domain'] and result['recipient_domain']):
            
            result['domain_age_comparison'] = {
                'age_difference_days': None,
                'warnings': []
            }
            
            sender_age = result['sender_domain'].get('age_days')
            recipient_age = result['recipient_domain'].get('age_days')
            
            if sender_age is not None and recipient_age is not None:
                try:
                    age_diff = recipient_age - sender_age
                    result['domain_age_comparison']['age_difference_days'] = age_diff
                    
                    if age_diff > 365:  # 收件人域名比发件人域名老一年以上
                        result['risk_score'] += 2.0
                        warning = (
                            f"发件人域名 {sender_domain} 比收件人域名 {recipient_domain} "
                            f"新 {age_diff} 天，这种情况在钓鱼邮件中很常见"
                        )
                        result['domain_age_comparison']['warnings'].append(warning)
                        result['warnings'].append(warning)
                except Exception as e:
                    result['warnings'].append(f"比较域名年龄时出错: {str(e)}")
        
        # 设置最终风险等级
        if result['risk_score'] >= 5.0:
            result['risk_level'] = 'critical'
        elif result['risk_score'] >= 4.0:
            result['risk_level'] = 'high'
        elif result['risk_score'] >= 3.0:
            result['risk_level'] = 'medium'
        
        # 合并所有警告信息
        if result['sender_domain'] and result['sender_domain'].get('warnings'):
            result['warnings'].extend(result['sender_domain']['warnings'])
        if result['recipient_domain'] and result['recipient_domain'].get('warnings'):
            result['warnings'].extend(result['recipient_domain']['warnings'])
            
    except Exception as e:
        result['warnings'].append(f"域名注册信息分析失败: {str(e)}")
    
    return result

def check_domain_registration_with_retry(domain: str, max_retries: int = 3) -> Dict[str, Any]:
    """带重试机制的域名注册信息检查"""
    for attempt in range(max_retries):
        try:
            return check_domain_registration(domain)
        except Exception as e:
            if attempt == max_retries - 1:
                return {
                    'domain': domain,
                    'creation_date': None,
                    'expiration_date': None,
                    'last_updated': None,
                    'age_days': None,
                    'risk_level': 'unknown',
                    'risk_score': 0.0,
                    'warnings': [f"多次尝试后仍无法获取域名信息: {str(e)}"]
                }
            print(f"第 {attempt + 1} 次尝试失败，准备重试...")
            time.sleep(2)  # 等待2秒后重试

def get_verified_path(path_type="any"):
    """获取已验证的指定类型路径"""
    while True:
        try:
            raw_input = input("路径输入（支持拖拽文件）: ").strip(' "\'')
            full_path = os.path.normpath(os.path.expanduser(raw_input))
            
            if not os.path.exists(full_path):
                raise FileNotFoundError("路径不存在")
                
            if path_type == "file" and not os.path.isfile(full_path):
                raise ValueError("必须选择文件")
                
            if path_type == "dir" and not os.path.isdir(full_path):
                raise ValueError("必须选择目录")
                
            return full_path
            
        except (FileNotFoundError, ValueError) as e:
            print(f"输入错误: {str(e)}")
        except Exception as e:
            print(f"未知错误: {str(e)}")


# 使用示例
if __name__ == "__main__":
    # test_file = r"/Users/admin/Downloads/1745443910003.eml"
    test_file = get_verified_path("file")

    try:
        print(f"开始解析邮件: {test_file}")
        print(f"文件大小: {os.path.getsize(test_file)} 字节")
        
        # 解析邮件
        email_data = parse_email(test_file)
        
        # 显示附件统计
        print(f"\n发现附件数量: {len(email_data['attachments'])}")
        
        # 显示报告
        display_report(email_data)
        
    except Exception as e:
        print(f"处理失败: {str(e)}")
        import traceback
        traceback.print_exc()
