# AI文档生成器

一个基于AI的智能文档生成系统，支持从`.doc`模板文件生成定制化文档。

## 🌟 特性

- **五阶段工作流程**: DOC转换 → 模板分析 → JSON输入 → AI字段映射 → 文档生成
- **智能字段映射**: 使用AI自动匹配输入数据与模板字段，无需手动配置
- **通用化填充**: 不依赖硬编码字段名，支持任意模板和数据格式
- **自动转换**: 支持从`.doc`文件自动转换为`.docx`
- **详细日志**: 完整的过程记录和错误处理
- **AI驱动**: 使用Gemini 2.5 Pro进行智能分析和字段映射

## 🔧 工作流程

### 阶段0：DOC转换
- 使用LibreOffice将`.doc`文件转换为`.docx`
- 自动检测LibreOffice安装路径
- 支持macOS、Linux和Windows

### 阶段1：模板分析
- AI分析Word模板结构
- 自动识别所有字段标签
- 生成标准化的模板字段结构JSON

### 阶段2：JSON输入
- 从JSON文件加载待填入的数据
- 支持任意字段名和数据格式

### 阶段2.5：AI字段映射 🆕
- **核心创新功能**：使用AI智能映射字段名称
- 分析模板结构和输入数据的语义关系
- 自动将输入数据字段映射到模板字段
- 解决字段名不匹配的通用化问题

### 阶段3：智能填充
- 使用映射后的数据填充模板
- 支持多种填充模式（下一单元格、同一单元格等）
- 智能识别复杂表格结构

## 🚀 使用方法

### 安装依赖

```bash
pip install -r requirements.txt
```

### 运行程序

```bash
python main.py
```

### 文件结构

```
ai_docGen/
├── main.py                    # 主程序
├── template_test.doc          # DOC模板文件
├── sample_input.json          # 示例输入数据
├── requirements.txt           # 依赖包
└── README.md                 # 说明文档
```

## 📝 配置

### API配置
在`main.py`中设置OpenRouter API密钥：
```python
API_KEY = "your-openrouter-api-key"
```

### 输入数据格式
JSON文件支持任意字段名，AI会自动进行智能映射：

```json
{
  "编号": "GZ-FH-2025-001",
  "项目名称": "历史文物建筑修缮示范项目",
  "复核日期": "2025年1月25日",
  "original_condition_review": "详细的原形制复核内容...",
  "damage_assessment_review": "详细的病害评估内容...",
  "repair_plan_review": "详细的修缮方案内容...",
  "project_lead": "李建筑师",
  "reviewer": "王专家、张工程师"
}
```

## 🔍 AI字段映射详解

这是系统的核心创新功能，解决了传统模板填充中字段名硬编码的问题：

### 问题场景
- 模板AI分析得到：`{"serial_number": "", "project_name": ""}`
- 输入数据包含：`{"编号": "GZ-001", "项目名称": "修缮工程"}`
- 传统方法需要手动配置映射关系

### AI解决方案
- **语义分析**：AI理解"编号"和"serial_number"的语义关系
- **智能映射**：自动生成映射：`{"serial_number": "GZ-001", "project_name": "修缮工程"}`
- **通用化**：支持任意字段名组合，无需预设配置

### 映射优势
1. **零配置**：无需手动设置字段映射关系
2. **多语言**：支持中英文混合字段名
3. **容错性**：智能处理字段名变化和缺失
4. **可扩展**：适用于各种模板和数据格式

## 📊 运行示例

```
🚀 AI文档生成器 - 主程序
==================================================
18:10:12 - INFO - 🤖 AI生成器初始化完成
18:10:12 - INFO - 🚀 开始完整的AI文档生成流程

阶段0：DOC转换
18:10:14 - INFO - ✅ 转换成功: template_test_converted.docx

阶段1：模板分析
18:10:48 - INFO - ✅ 成功提取 8 个字段

阶段2：JSON输入
18:10:48 - INFO - ✅ 成功加载 8 个数据字段

阶段2.5：AI字段映射 🆕
18:11:01 - INFO - ✅ 成功映射 8 个字段

阶段3：智能填充
18:11:01 - INFO - ✅ 文档已成功生成
18:11:02 - INFO - 🎉 AI文档生成流程完成!
18:11:02 - INFO - ⏱️ 总用时: 49.93 秒
```

## 🛠️ 系统要求

- Python 3.7+
- LibreOffice（用于DOC转换）
- OpenRouter API密钥
- 网络连接（AI API调用）

## ✨ 技术亮点

1. **AI驱动的字段映射**：核心创新，解决通用化难题
2. **多阶段流水线**：清晰的处理步骤，易于维护和扩展
3. **智能错误处理**：完善的异常处理和降级策略
4. **详细日志系统**：全程记录，便于调试和监控
5. **跨平台支持**：自动检测系统环境，适配不同操作系统

## 🤖 AI能力

- **Gemini 2.5 Pro**：高性能的多模态AI模型
- **语义理解**：深度理解字段含义和关系
- **智能推理**：自动推断最佳字段映射策略
- **结构化输出**：确保JSON格式的准确性和一致性

## 🚀 快速使用

```bash
python main.py
```

## 📋 工作流程

1. **阶段0**：将 `.doc` 文件使用LibreOffice转换为 `.docx` 格式
2. **阶段1**：AI分析转换后的Word模板，提取字段结构
3. **阶段2**：读取JSON输入文件 (`sample_input.json`)，获取填充数据  
4. **阶段3**：智能填充模板，生成最终Word文档

## 📁 文件说明

- `main.py` - 主程序
- `template_test.doc` - **DOC模板文件（主要使用）**
- `template_test.docx` - DOCX模板文件（备用）
- `sample_input.json` - JSON输入数据
- `requirements.txt` - 依赖包

## 🔧 系统要求

**必须安装LibreOffice**：
- macOS: 从 [LibreOffice官网](https://www.libreoffice.org/) 下载安装
- 程序会自动查找LibreOffice路径：`/Applications/LibreOffice.app/Contents/MacOS/soffice`

## 🔧 修改输入数据

编辑 `sample_input.json` 文件来修改要填充的内容：

```json
{
  "serial_number": "项目编号",
  "project_name": "项目名称", 
  "review_date": "复核日期",
  "original_condition_review": "原形制复核情况",
  "damage_assessment_review": "病害和残损复核情况",
  "repair_plan_review": "修缮做法复核情况",
  "project_lead": "项目负责人",
  "reviewer": "复核人员"
}
```

## 📊 输出

程序会：
1. 显示详细的转换和处理日志
2. 生成中间转换文件：`template_test_converted.docx`
3. 生成最终输出文件：`AI生成文档_20250617_174440.docx`

## 🔍 工作流程日志示例

```
🔄 开始DOC到DOCX转换...
🔍 检查LibreOffice可用性...
✅ 找到LibreOffice: /Applications/LibreOffice.app/Contents/MacOS/soffice
📄 正在转换: template_test.doc -> template_test_converted.docx
✅ 转换成功: template_test_converted.docx
```
