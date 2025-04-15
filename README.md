# 邮件风险分析工具 (Mail Evaluation & Risk Analysis Tool)

这是一个功能强大的邮件安全分析工具，专门用于检测和分析可疑邮件，帮助识别潜在的钓鱼攻击和安全威胁。

## 主要功能

1. **基础邮件解析**
   - 解析 EML 和 MSG 格式的邮件
   - 提取发件人、收件人、主题等基本信息
   - 支持多种字符编码（UTF-8, GBK, GB2312等）

2. **安全特征分析**
   - 发件人身份伪造检测
   - 域名相似度分析
   - SPF/DKIM/DMARC 验证
   - 域名注册信息检查
   - 隐藏内容和跟踪器检测
   - URL 安全性分析

3. **附件分析**
   - 支持多种文档格式（PDF, Word, Excel, PowerPoint）
   - 检测可执行文件和压缩包
   - 提供文件预览功能
   - 计算文件哈希值

## 安装要求

### 基本依赖
```bash
pip install python-whois beautifulsoup4 extract-msg
```

### 扩展文档支持（可选）
```bash
pip install python-docx PyPDF2 openpyxl python-pptx
```

## 使用方法

```python
from MER import parse_email, display_report

# 分析邮件
email_file = "path/to/your/email.eml"
email_data = parse_email(email_file)

# 显示分析报告
display_report(email_data)
```

## 分析报告内容

1. **基本信息**
   - 发件人/收件人信息
   - 邮件主题和日期
   - 邮件正文内容

2. **安全分析**
   - 邮件认证信息
   - 域名相似度检查
   - 发件人真实性检测
   - 域名注册信息分析

3. **内容分析**
   - 隐藏内容检测
   - URL信息分析
   - 附件详细信息

## 风险评估

工具使用多维度的风险评分系统：
- 域名注册时间评估
- 发件人身份验证
- URL安全性分析
- 隐藏内容检测
- 附件安全性评估

风险等级分为：
- Critical（严重）
- High（高）
- Medium（中）
- Low（低）
- Unknown（未知）

## 注意事项

1. 确保安装了所有必要的依赖包
2. 对于大型附件，分析可能需要较长时间
3. 某些功能（如WHOIS查询）可能受到网络限制
4. 建议在分析可疑邮件时使用隔离环境

## 贡献

欢迎提交问题报告和改进建议。如果您想贡献代码，请确保：
1. 遵循现有的代码风格
2. 添加适当的注释和文档
3. 确保所有测试通过

## 许可证

[添加您选择的许可证信息]
