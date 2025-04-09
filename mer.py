from email import policy
from email.parser import BytesParser
import extract_msg
import os
from typing import Dict, Any, List
from bs4 import BeautifulSoup
import re
import io
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
                        for part in msg.walk():
                            if part.is_multipart():
                                continue
                            if part.get_content_maintype() == 'text':
                                content_type = part.get_content_type()
                                try:
                                    payload = part.get_payload(decode=True)
                                    if payload:
                                        charset = part.get_content_charset() or 'utf-8'
                                        content = payload.decode(charset, errors='ignore')
                                        if content_type == 'text/plain':
                                            email_data['body_text'] += content + '\n'
                                        elif content_type == 'text/html':
                                            email_data['body_html'] += content + '\n'
                                except Exception as e:
                                    print(f"解析正文部分出错: {str(e)}")
                            elif part.get_content_disposition() == 'attachment':
                                try:
                                    attachment = {
                                        'filename': part.get_filename() or "未知文件名",
                                        'mime_type': part.get_content_type(),
                                        'size': len(part.get_payload(decode=True)),
                                        'data': part.get_payload(decode=True)
                                    }
                                    email_data['attachments'].append(attachment)
                                except Exception as e:
                                    print(f"处理附件出错: {str(e)}")
                    else:
                        content_type = msg.get_content_type()
                        try:
                            payload = msg.get_payload(decode=True)
                            if payload:
                                charset = msg.get_content_charset() or 'utf-8'
                                content = payload.decode(charset, errors='ignore')
                                if content_type == 'text/plain':
                                    email_data['body_text'] = content
                                elif content_type == 'text/html':
                                    email_data['body_html'] = content
                        except Exception as e:
                            print(f"解析正文出错: {str(e)}")
                
                # 调用正文解析函数
                get_body_content(msg)
                
                # 如果正文为空，尝试其他方法获取
                if not email_data['body_text'] and not email_data['body_html']:
                    try:
                        # 尝试直接获取payload
                        payload = msg.get_payload()
                        if isinstance(payload, str):
                            email_data['body_text'] = payload
                        elif isinstance(payload, list):
                            for part in payload:
                                if part.get_content_type() == 'text/plain':
                                    email_data['body_text'] += part.get_payload() + '\n'
                                elif part.get_content_type() == 'text/html':
                                    email_data['body_html'] += part.get_payload() + '\n'
                    except Exception as e:
                        print(f"尝试获取正文备用方法失败: {str(e)}")
                
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

def extract_urls(email_data: Dict[str, Any]) -> Dict[str, List[str]]:
    """
    从邮件中提取所有URL
    
    Args:
        email_data: 邮件解析数据
        
    Returns:
        包含不同来源URL的字典
    """
    urls = {
        'text': [],      # 纯文本中的URL
        'html': [],      # HTML内容中的URL
        'attachments': [] # 附件内容中的URL
    }
    
    try:
        # URL正则表达式模式
        url_pattern = r'(?i)\b((?:https?://|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}/)(?:[^\s()<>]+|\(([^\s()<>]+|(\([^\s()<>]+\)))*\))+(?:\(([^\s()<>]+|(\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:\'".,<>?«»""'']))'
        
        # 从纯文本中提取URL
        if email_data['body_text']:
            text_urls = re.findall(url_pattern, email_data['body_text'])
            urls['text'] = [url[0] for url in text_urls]  # 取第一个分组作为完整URL
        
        # 从HTML内容中提取URL
        if email_data['body_html']:
            # 使用BeautifulSoup解析HTML
            soup = BeautifulSoup(email_data['body_html'], 'html.parser')
            
            # 提取所有链接标签的href属性
            for link in soup.find_all('a'):
                href = link.get('href')
                if href and not href.startswith('mailto:'):  # 排除邮件链接
                    urls['html'].append(href)
            
            # 提取HTML内容中的其他URL
            html_text = soup.get_text()
            html_urls = re.findall(url_pattern, html_text)
            urls['html'].extend([url[0] for url in html_urls])
            
            # 提取图片链接
            for img in soup.find_all('img'):
                src = img.get('src')
                if src and not src.startswith('data:'):  # 排除base64编码的图片
                    urls['html'].append(src)
        
        # 从附件中提取URL（如果是文本文件）
        for attachment in email_data['attachments']:
            if attachment.get('text_preview'):
                attachment_urls = re.findall(url_pattern, attachment['text_preview'])
                urls['attachments'].extend([url[0] for url in attachment_urls])
        
        # 去重并清理URL
        for source in urls:
            # 去除重复URL
            urls[source] = list(set(urls[source]))
            # 清理URL（移除尾部标点等）
            urls[source] = [url.rstrip('.,;:\'\"!?') for url in urls[source]]
            # 移除空URL
            urls[source] = [url for url in urls[source] if url]
        
    except Exception as e:
        print(f"提取URL失败: {str(e)}")
    
    return urls

def detect_hidden_content(email_data: Dict[str, Any]) -> Dict[str, List[Dict[str, Any]]]:
    """
    检测邮件中的隐藏内容和跟踪器
    
    Args:
        email_data: 邮件解析数据
        
    Returns:
        包含检测结果的字典
    """
    findings = {
        'hidden_content': [],    # 隐藏内容
        'tracking_elements': [], # 跟踪元素
        'suspicious_urls': [],   # 可疑URL
        'risk_level': 'low'     # 风险等级: low, medium, high
    }
    
    try:
        if email_data['body_html']:
            soup = BeautifulSoup(email_data['body_html'], 'html.parser')
            
            # 检测隐藏内容
            # 1. 检查隐藏的元素
            hidden_elements = soup.find_all(style=lambda x: x and ('display:none' in x.lower() or 
                                                                 'visibility:hidden' in x.lower() or 
                                                                 'opacity:0' in x.lower()))
            for element in hidden_elements:
                findings['hidden_content'].append({
                    'type': 'hidden_element',
                    'content': element.get_text().strip(),
                    'style': element.get('style', ''),
                    'reason': '使用CSS隐藏的内容'
                })
            
            # 2. 检查微小元素（1px或更小）
            tiny_elements = soup.find_all(style=lambda x: x and ('width:1px' in x.lower() or 
                                                               'height:1px' in x.lower() or 
                                                               'font-size:1px' in x.lower()))
            for element in tiny_elements:
                findings['hidden_content'].append({
                    'type': 'tiny_element',
                    'content': element.get_text().strip(),
                    'style': element.get('style', ''),
                    'reason': '微小元素可能用于隐藏内容'
                })
            
            # 3. 检查跟踪图片
            images = soup.find_all('img')
            tracking_keywords = {'track', 'pixel', 'beacon', 'analytics', 'monitor', 'stat'}
            for img in images:
                src = img.get('src', '')
                if src:
                    # 检查1x1像素图片
                    if img.get('width') == '1' and img.get('height') == '1':
                        findings['tracking_elements'].append({
                            'type': 'tracking_pixel',
                            'url': src,
                            'reason': '1x1像素跟踪图片'
                        })
                    
                    # 检查包含跟踪关键词的URL
                    if any(keyword in src.lower() for keyword in tracking_keywords):
                        findings['tracking_elements'].append({
                            'type': 'tracking_image',
                            'url': src,
                            'reason': '包含跟踪关键词的图片URL'
                        })
            
            # 4. 检查可疑URL
            links = soup.find_all('a')
            suspicious_keywords = {'click', 'track', 'redirect', 'goto', 'forward'}
            for link in links:
                href = link.get('href', '')
                if href:
                    # 检查重定向URL
                    if any(keyword in href.lower() for keyword in suspicious_keywords):
                        findings['suspicious_urls'].append({
                            'type': 'suspicious_link',
                            'url': href,
                            'text': link.get_text().strip(),
                            'reason': '包含可疑关键词的链接'
                        })
                    
                    # 检查URL编码过多的链接
                    if href.count('%') > 3:
                        findings['suspicious_urls'].append({
                            'type': 'encoded_url',
                            'url': href,
                            'text': link.get_text().strip(),
                            'reason': 'URL编码过多，可能试图隐藏真实地址'
                        })
            
            # 增强跟踪器检测
            # 1. 邮件服务提供商域名列表
            email_service_domains = {
                'sendgrid.net',
                'mailchimp.com',
                'salesforce.com',
                'marketo.com',
                'hubspot.com',
                'mailtrack.io',
                'constantcontact.com',
                'amazonses.com',
                'exacttarget.com',
                'pardot.com',
                'mailgun.com',
                'postmarkapp.com'
            }
            
            # 2. 跟踪相关路径关键词
            tracking_paths = {
                'wf/open',
                'open.php',
                'track',
                'tracking',
                'click',
                'view',
                'open',
                'beacon',
                'pixel',
                'analytics'
            }
            
            # 3. 检查所有链接和图片
            all_elements = soup.find_all(['a', 'img'])
            for element in all_elements:
                url = element.get('href') or element.get('src', '')
                if url:
                    # 解析URL
                    try:
                        from urllib.parse import urlparse, parse_qs
                        parsed_url = urlparse(url)
                        domain = parsed_url.netloc.lower()
                        path = parsed_url.path.lower()
                        
                        # 检查是否是已知的邮件服务提供商域名
                        if any(service in domain for service in email_service_domains):
                            findings['tracking_elements'].append({
                                'type': 'email_service_tracker',
                                'url': url,
                                'service': domain,
                                'reason': f'发现邮件服务商({domain})的跟踪链接'
                            })
                        
                        # 检查URL路径是否包含跟踪关键词
                        if any(track_path in path for track_path in tracking_paths):
                            findings['tracking_elements'].append({
                                'type': 'tracking_path',
                                'url': url,
                                'path': path,
                                'reason': '链接路径包含跟踪相关关键词'
                            })
                        
                        # 检查是否有可疑的查询参数
                        query_params = parse_qs(parsed_url.query)
                        suspicious_params = {'uid', 'user', 'id', 'email', 'u', 'upn', 'tracking'}
                        if any(param in suspicious_params for param in query_params):
                            findings['tracking_elements'].append({
                                'type': 'tracking_parameters',
                                'url': url,
                                'params': list(query_params.keys()),
                                'reason': '链接包含跟踪相关参数'
                            })
                        
                        # 检查编码和加密特征
                        if url.count('-') > 5 or url.count('=') > 2:
                            findings['tracking_elements'].append({
                                'type': 'encoded_tracking',
                                'url': url,
                                'reason': '链接包含大量编码或加密字符'
                            })
                        
                    except Exception as e:
                        print(f"URL解析失败: {str(e)}")
        
        # 评估整体风险等级
        risk_score = 0
        risk_score += len(findings['hidden_content']) * 2     # 隐藏内容权重
        risk_score += len(findings['tracking_elements']) * 1.5 # 跟踪元素权重
        risk_score += len(findings['suspicious_urls']) * 1.5   # 可疑URL权重
        
        # 根据跟踪器类型调整风险分数
        for item in findings['tracking_elements']:
            if item['type'] == 'email_service_tracker':
                risk_score += 2  # 已知邮件服务商的跟踪器权重更高
            elif item['type'] == 'encoded_tracking':
                risk_score += 1.5  # 编码跟踪链接权重较高
        
        if risk_score > 10:
            findings['risk_level'] = 'high'
        elif risk_score > 5:
            findings['risk_level'] = 'medium'
        
        findings['risk_score'] = risk_score
        
    except Exception as e:
        print(f"检测隐藏内容失败: {str(e)}")
    
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
        # 使用BeautifulSoup清理HTML标签
        soup = BeautifulSoup(email_data['body_html'], 'html.parser')
        clean_text = soup.get_text()
        print(clean_text[:500] + "..." if len(clean_text) > 500 else clean_text)
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
    
    # 将原来的隐藏内容检测改为[10]，URL信息改为[11]，附件信息改为[12]
    print("\n[10] 隐藏内容检测:")
    findings = detect_hidden_content(email_data)
    
    if findings['risk_level'] != 'low':
        print(f"\n风险等级: {findings['risk_level'].upper()}")
        print(f"风险分数: {findings['risk_score']:.1f}")
        
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
    urls = extract_urls(email_data)
    
    if any(urls.values()):
        # 显示纯文本中的URL
        if urls['text']:
            print("\n纯文本中的URL:")
            for i, url in enumerate(urls['text'], 1):
                print(f"  {i}. {url}")
        
        # 显示HTML中的URL
        if urls['html']:
            print("\nHTML中的URL:")
            for i, url in enumerate(urls['html'], 1):
                print(f"  {i}. {url}")
        
        # 显示附件中的URL
        if urls['attachments']:
            print("\n附件中的URL:")
            for i, url in enumerate(urls['attachments'], 1):
                print(f"  {i}. {url}")
        
        # 显示URL统计信息
        total_urls = len(urls['text']) + len(urls['html']) + len(urls['attachments'])
        print(f"\n总计发现 {total_urls} 个URL:")
        print(f"- 纯文本中: {len(urls['text'])} 个")
        print(f"- HTML中: {len(urls['html'])} 个")
        print(f"- 附件中: {len(urls['attachments'])} 个")
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
    
    print("\n=== 报告结束 ===")

# 使用示例
if __name__ == "__main__":
    test_file = r"C:\Users\v_ayfzhang\Documents\email\1744119309995.eml"

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
