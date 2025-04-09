# mer
恶意邮件文件分析

# 邮件解析工具 (Mail Extraction & Report Tool)

这是一个功能强大的邮件解析工具，可以解析EML和MSG格式的邮件文件，提取关键信息并生成详细的分析报告。

## 功能特点

1. **基础邮件信息提取**
   - 发件人、收件人、抄送、密送
   - 邮件主题和日期
   - 正文内容（纯文本和HTML格式）
   - 邮件会话信息（原始邮件信息）

2. **附件处理**
   - 支持多种文件格式的预览
   - 提取附件元数据（文件名、大小、类型等）
   - 计算文件哈希值（MD5、SHA256）
   - 检测可执行文件和压缩文件
   - 支持以下文档格式的预览：
     - PDF文件（前3页）
     - Word文档（前10段）
     - Excel表格（前10行）
     - PowerPoint（前3张幻灯片）
     - 文本文件

3. **安全分析**
   - 检测隐藏内容
   - 识别跟踪器和跟踪像素
   - 分析可疑URL
   - 验证邮件认证（SPF、DKIM、DMARC）

4. **URL分析**
   - 从正文提取URL
   - 从HTML内容提取链接
   - 从附件中提取URL
   - URL分类统计

## 安装要求

```bash
# 基础依赖
pip install extract-msg beautifulsoup4

# 文档预览支持（可选）
pip install python-docx PyPDF2 openpyxl python-pptx
```

## 使用方法

```python
from MER import parse_email, display_report

# 解析邮件文件
email_file = "path/to/your/email.eml"  # 或 .msg 文件
email_data = parse_email(email_file)

# 显示分析报告
display_report(email_data)
```

## 输出报告内容

分析报告包含以下部分：
1. 发件人信息
2. 收件人信息
3. 抄送人信息
4. 密送人信息
5. 回复地址
6. 邮件基本信息
7. 原始邮件信息
8. 正文内容
9. 邮件认证信息
10. 隐藏内容检测
11. URL信息
12. 附件信息

## 注意事项

1. 处理大型附件时可能需要较多内存
2. 某些文档格式的预览需要安装额外的依赖包
3. 部分邮件可能因格式问题无法完全解析
4. 建议在处理未知来源的邮件时注意安全性

## 错误处理

工具包含完善的错误处理机制：
- 文件读取错误处理
- 编码解析错误处理
- 附件提取错误处理
- 文档预览错误处理

## 贡献

欢迎提交问题和改进建议！

## 许可证

[选择合适的许可证]
