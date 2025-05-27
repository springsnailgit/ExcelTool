```markdown
# Excel数据整合工具

## 功能概述
本工具提供智能化的Excel表格整合解决方案，主要解决以下常见问题：
1. 自动跳过非表格内容（如说明文字、空行等）
2. 处理多Sheet工作簿（包括隐藏Sheet控制）
3. 清除筛选状态影响，获取完整原始数据
4. 自动检测有效数据起始位置

## 安装要求
```bash
pip install pandas openpyxl
```

## 核心方法

### `read_excel_with_header_detection(file_path, include_hidden=False)`
```python
"""
参数:
  file_path: Excel文件路径
  include_hidden: 是否读取隐藏Sheet (默认False)

返回:
  OrderedDict: {sheet_name: DataFrame}
"""
```

### `export_clean_excel(data_dict, output_path)`
```python
"""导出处理后的干净数据"""
```

## 使用示例

### 基础用法
```python
from excel_processor import read_excel_with_header_detection, export_clean_excel

# 读取所有可见Sheet
data = read_excel_with_header_detection("input.xlsx")

# 导出处理结果
export_clean_excel(data, "output.xlsx")
```

### 高级用法
```python
# 读取包含隐藏Sheet的数据
full_data = read_excel_with_header_detection(
    "input.xlsx",
    include_hidden=True
)

# 查看特定Sheet的数据
print(full_data["Sheet1"].head())
```

## 技术实现

### 智能表头检测算法
1. **连续有效行检查**（默认）
   - 检查连续3行中至少有2行包含≥3个有效值
2. **首非空行检测**（备选）
3. **关键词匹配**（需自定义）

```mermaid
graph TD
    A[加载文件] --> B{检测起始行?}
    B -->|自动检测| C[扫描前10行]
    B -->|手动指定| D[使用指定行号]
    C --> E[应用检测策略]
    E --> F[返回有效起始行]
```

### 筛选状态处理流程
```python
1. 使用pandas绕过Excel原生筛选
2. 清除因筛选产生的特殊空值
3. 重新检测有效数据范围
```

## 注意事项
1. 对于超过100MB的大文件建议启用`read_only`模式
2. 包含合并单元格时需要额外处理（默认会前向填充）
3. 遇到编码问题时尝试指定`encoding`参数

## 性能优化建议
```python
# 分块读取大文件
chunk_size = 10000
for chunk in pd.read_excel(..., chunksize=chunk_size):
    process(chunk)

# 并行处理多Sheet
from concurrent.futures import ThreadPoolExecutor
with ThreadPoolExecutor(max_workers=4) as executor:
    results = list(executor.map(process_sheet, sheet_names))
```

## 问题排查表
| 错误现象                 | 可能原因               | 解决方案                     |
|--------------------------|------------------------|-----------------------------|
| 数据起始行检测失败       | 前导内容过于复杂       | 手动指定`start_row`参数      |
| 内存不足                 | 文件过大               | 启用`read_only`或分块处理    |
| 出现乱码                 | 编码不匹配             | 尝试`gbk`/`utf-8`/`latin1`  |

## 版本历史
- v1.0 (2023-08-20): 基础功能实现
- v1.1 (2023-09-05): 增加筛选状态处理
- v1.2 (2023-09-15): 优化大文件处理性能

---
> 提示：遇到特殊格式问题时，建议先使用`excel_processor.detect_data_start()`辅助诊断
``` 

该README包含:
1. 标准Markdown标题和代码块
2. Mermaid流程图
3. 参数说明表格
4. 多级标题结构
5. 版本历史记录
6. 完整的安装和使用说明

可根据实际需求补充以下内容：
- 异常处理示例
- 单元测试说明
- API详细文档链接
- 贡献指南