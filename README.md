# 论文翻译工具

这个工具帮助你将PDF格式的外文文献转换为中英文对照的Word文档，方便论文翻译工作。

## 功能

- 从PDF文件中提取文本
- 自动按页码组织文本
- 识别并高亮显示图表引用文本（如"Figure 1"）
- 标记图片位置，方便手动插入
- 生成包含原文和翻译的对照Word文档
- 支持多种翻译方式：
  - 手动翻译
  - TODO: Google翻译API（无需密钥）
  - Ollama本地模型翻译 (Recommended)
- 智能分段处理长文本，确保翻译完整性
- 自动重试机制，提高翻译成功率

## 文档特性

- 标准学术论文格式：
  - 中文：宋体（小四）
  - 英文：Times New Roman（小四）
- 文本按页码和段落组织，便于对照
- 标记出图片位置，便于后续手动插入图片
- 保持学术翻译风格
- 长文本自动分段处理，确保翻译完整

## 安装

1. 确保已安装Python 3.6+
2. 安装依赖：
```
pip install -r requirements.txt
```
3. 如果使用Ollama翻译，请确保已安装并运行Ollama：
   - 官方安装指南：https://github.com/ollama/ollama
   - 下载需要的模型，例如：`ollama pull llama3`

## 使用方法

### 基本用法（不带自动翻译）

```bash
python paper_translator.py your_paper.pdf
```

这将创建一个包含原文的Word文档，你可以手动添加翻译。

### 使用Google翻译API（推荐）

Google翻译通常提供较好的翻译质量，特别是对学术文本：

```bash
python paper_translator.py your_paper.pdf -t google
```

注意事项：
- 无需API密钥
- 自动处理长文本，分段翻译后合并
- 内置重试机制，确保翻译完整

### 使用Ollama本地翻译

确保Ollama服务正在运行，然后运行：

```bash
python paper_translator.py your_paper.pdf -t ollama -m your_model_name
```

可以通过`-m`参数指定使用的模型，例如：
```bash
python paper_translator.py your_paper.pdf -t ollama -m llama3
```

### 指定输出文件名

```bash
python paper_translator.py your_paper.pdf -o translation_output.docx -t google
```

## 文档格式说明

生成的Word文档采用以下格式规范：

- **字体设置**：
  - 所有中文内容：宋体（小四，12磅）
  - 所有英文内容：Times New Roman（小四，12磅）

- **文档结构**：
  - 标题：居中对齐
  - 页码标记：使用二级标题
  - 段落标记：使用三级标题
  - 原文与翻译：清晰分隔
  - 图片引用：以红色高亮显示
  - 图片位置：使用【图片位置】标记

## 翻译质量优化

- **长文本处理**：自动分段处理长文本（>5000字符），确保翻译完整
- **重试机制**：翻译失败时自动重试（最多3次），使用指数退避策略
- **结果验证**：检测翻译结果是否完整，不完整则重试
- **连接检查**：自动检测Ollama服务连接状态，提供切换选项

## 建议工作流程

1. 使用Google翻译生成初始翻译文档：
   ```bash
   python paper_translator.py your_paper.pdf -t google
   ```

2. 打开生成的Word文档，根据【图片位置】标记手动插入图片

3. 审核并优化翻译内容

4. 导出或直接使用最终文档

## TODO
1. 之前插入图片（AI agent)

2. 提升翻译效果