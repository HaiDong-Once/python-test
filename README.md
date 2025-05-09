# DOCX转Markdown转换工具

这个项目实现了一个高质量的DOCX文档转Markdown格式的转换工具，支持多种复杂文档格式。

## 项目结构

```
├── utils/                  # 工具函数目录
│   ├── docx_to_md.py       # DOCX转Markdown核心代码（默认方法）
│   ├── cyrus_docx_converter.py # CYRUS-STUDIO转换方法（替代方法）
│   ├── docx_converter_selector.py # 转换方法选择器
│   └── ...                 # 其他工具函数
├── tests/                  # 测试目录
│   ├── test_docs/          # 测试文档
│   ├── test_outputs/       # 测试输出
│   ├── test_docx/          # DOCX转换相关测试
│   │   ├── test_basic.py   # 基本转换测试
│   │   ├── test_complex.py # 复杂文档测试
│   │   └── test_lists.py   # 列表处理测试
│   └── test_converter_methods.py # 转换方法比较测试
├── app.py                  # 主应用
├── run_tests.py            # 测试运行脚本
└── requirements.txt        # 项目依赖
```

## 核心功能

- DOCX文档转换为Markdown格式
- 支持复杂格式：标题、列表、表格、图片等
- 智能标题层级处理
- 图片自动提取与引用
- 格式优化
- 多种转换方式选择

## 转换方法

本工具支持两种转换方法：
https://gitee.com/comma-dong/image-projects/raw/master/20250509171906_chatGPT_outputs/image_1.png

1. **默认方法** - 我们自己开发的转换逻辑，针对格式和图片位置做了大量优化
2. **CYRUS方法** - 使用[CYRUS-STUDIO/docx2markdown](https://github.com/CYRUS-STUDIO/docx2markdown)开源库的转换功能（首次使用会自动安装）

在web界面上，用户可以通过每个文件旁边的标签页选择所需的转换方法，操作更加直观便捷。系统设计为按需安装CYRUS方法，用户首次选择此方法时会自动进行安装。

### 方法比较

两种方法的主要区别：

特性 | 默认方法 | CYRUS方法
--- | --- | ---
图片处理 | 精准的上下文位置匹配 | 基本的图片提取和引用
代码块识别 | 增强的代码块识别（支持灰色背景） | 基本的代码块识别 
样式保留 | 更多样式保留（粗体、斜体等） | 基本样式保留
表格转换 | 优化的表格转换 | 基本表格转换
特殊字符处理 | 完善的特殊字符转义 | 基本字符处理

## 安装与依赖

```bash
pip install -r requirements.txt
```

如需使用CYRUS方法，还需要安装：

```

## 安全说明

本项目已清理所有敏感信息，确保代码安全。