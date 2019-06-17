# Python解析Excel导出json格式

## 1.使用路径
没有配置python的执行路径，需要在Excel的同级目录执行

## 2.配置

```python
# 此处做一些配置
configDic = {
    "sample.xlsx": "output_4.json"
}
manualHeaderrow = 0   # 手动配置标题行数，如果设为0表示自动识别标题行数，自动识别时按照合并单元格的最大行标row+1来作为标题行数
worksheetNum = 4    # 工作表编号，从0开始
```

## 3.嵌套结构的Excel配置
项目里的sample.xlsx是测试文件，分别对应生成0-4的json文件。
程序可以自动识别复杂表头的合并单元格，但需要基于从上往下包含的原则，即下层的标题不能比上层的标题更宽