# Excel数据分析师工具

## 项目简介

本工具是基于PyQt5开发的Excel数据处理与可视化工具，支持多文件批量处理、数据筛选、统计排行及图表生成等功能，适用于电商订单数据分析、销售统计等场景。

## 核心功能

1. 导入EXCEL
   - 选择文件夹自动加载所有Excel文件至左侧列表（项目提供测试文件在XS1中）
   - 支持.xlsx格式文件遍历与展示
2. 提取列数据
   - 提取关键字段：买家会员名、收货人姓名、联系手机、宝贝标题
   - 自动保存处理结果至原文件夹或自定义路径
3. 定向筛选
   - 合并多表后筛选特定商品（如"零基础学Python"）
   - 输出筛选结果并保存Excel
4. 多表合并
   - 智能合并文件夹内所有Excel文件
   - 保留原始索引与数据完整性
5. 多表统计排行
   - 按商品维度统计销售总量
   - 生成降序排列的销量排行榜
6. 生成图表
   - 贡献度分析（帕累托图）
   - 商品销售收入柱状图+累计比例曲线
   - 支持中文显示配置（华文细黑字体）

## 技术架构

- 界面框架：PyQt5（含QListView/QTextEdit组件）
- 数据处理：pandas + openpyxl
- 可视化库：matplotlib
- 文件操作：os/glob遍历，QFileDialog路径选择

## 安装与运行

####  依赖库(推荐使用镜像源安装)

```bash
bash

pip install pandas -i https://pypi.tuna.tsinghua.edu.cn/simple
pip install matplotlib -i https://pypi.tuna.tsinghua.edu.cn/simple
pip install PyQt5 -i https://pypi.tuna.tsinghua.edu.cn/simple
python -m pip install openpyxl
```

####  启动方式

```python
python

python dataEXCEL.py  # 主程序入口为show_MainWindow()
```

## 操作指南

1. 路径配置
   - 默认保存至源文件夹（单选按钮控制）
   - 点击"浏览"按钮可自定义输出目录
2. 数据处理流程
   - 左侧列表选择文件 → 右侧实时预览
   - 工具栏按钮触发对应处理功能
   - 处理结果自动保存并显示在文本区域
3. 图表生成
   - 执行"生成图表"功能后自动展示帕累托图
   - 图表包含：
     - 蓝色柱状图：各商品销售收入
     - 红色折线：累计收入占比曲线
     - 标注点：TOP10商品贡献度标记

## 注意事项

1. 路径分隔符适配：Windows系统使用`\`，代码中已通过`glob.glob(root+"\\*.xlsx")`兼容处理
2. 文件编码：确保Excel文件为UTF-8编码
3. 大文件处理：建议分批处理超过10万行的数据
4. 图表保存：当前版本为屏幕显示，如需保存图片可添加`plt.savefig('output.png')`

## 版本信息

- 开发环境：Python 3.8
- 测试平台：Windows 11
- 更新日期：2025年11月

## 贡献者

项目开发者：刘玲

联系方式：405800554@qq.com
