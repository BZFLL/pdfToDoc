# PDF翻译转化Word文档

## 项目背景
本项目为解决国外留学期间课程讲义翻译需求而开发。
我用常用的翻译工具有些问题：
1. 专业术语翻译不准确

可以：
- 保留PDF原始内容
- 自定义翻译API服务商（支持OpenAI、DeepSeek等）
- 调整OCR识别参数

## 功能特性
- PDF转Word文档
- OCR识别
- 可扩展的API翻译集成

## 环境配置

### 系统要求
- Python 3.8+
- Poppler utils（macOS推荐使用Homebrew安装）
- Tesseract OCR 5.0+

```bash
# 安装依赖
pip install -r requirements.txt

# macOS安装poppler
brew install poppler
```

### 快速开始
1. 复制配置文件模板：
```bash
cp config.example.json config.json
```
2. 编辑配置文件：
```json
{
  "api_config": {
    "endpoint": "你的API端点",
    "key": "你的API密钥",
    "provider": "服务商名称"
  },
  "ocr_settings": {
    "dpi": 300
  },
  "poppler_config": {
    "path": "/opt/homebrew/Cellar/poppler/25.01.0/bin"
  }
}
```
3. 运行转换程序：
```bash
python pdfToDoc.py 输入文件.pdf 输出文档.docx
```

## 最佳实践
1. 敏感配置管理：
- 将config.json添加到.gitignore
- 使用环境变量管理API密钥
- 定期轮换访问凭证

2. 性能优化：
- 调整dpi设置（200-300最佳）
- 根据文档语言设置OCR参数
- 批量处理时启用缓存机制

3. 扩展开发：
- 在translate_text函数中实现自定义翻译逻辑
- 通过ImageProcessorApp类扩展GUI功能
- 添加PDF/A格式输出支持
