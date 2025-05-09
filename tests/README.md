# 测试目录

本目录包含项目的所有测试文件和测试相关资源。

## 目录结构

- `test_docs/` - 测试用的DOCX文档
- `test_outputs/` - 测试输出目录
- `test_docx/` - DOCX转换相关测试
  - `test_basic.py` - 基本转换测试
  - `test_complex.py` - 复杂文档转换测试
  - `test_lists.py` - 列表处理测试

## 运行测试

从项目根目录运行所有测试:

```bash
python run_tests.py
```

运行特定测试:

```bash
python run_tests.py tests/test_docx/test_basic.py
```

## 添加新测试

1. 在适当的子目录中创建`test_*.py`文件
2. 使用pytest风格编写测试函数
3. 使用`conftest.py`中定义的夹具来获取测试资源

示例:

```python
def test_my_feature(test_document_path, test_output_path):
    # 测试代码
    assert True
``` 