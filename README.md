# 文件信息导出工具

一个基于Python开发的文件信息导出工具，提供美观的图形界面，支持将文件夹中的文件信息导出为Excel或CSV格式。

## 功能特点

- 🎯 支持导出文件夹及其子文件夹中的文件信息
- 📊 支持Excel和CSV两种导出格式
- 🎨 美观的图形用户界面
- 📝 丰富的导出选项：
  - 文件大小
  - 创建时间
  - 修改时间
  - 文件类型
  - 完整路径
- 📈 自动生成文件统计信息
  - 文件总数
  - 文件夹总数
  - 文件夹总大小
  - 各类型文件数量统计
- 🚀 高性能文件扫描和导出
- 💫 实时进度显示

## 使用说明

1. 选择要导出信息的文件夹
2. 勾选是否包含子文件夹
3. 点击"开始扫描"按钮扫描文件
4. 选择需要导出的信息项（大小、时间等）
5. 选择导出格式（Excel/CSV）
6. 点击导出按钮完成导出

## 技术实现

- 使用Python 3开发
- GUI框架：tkinter
- Excel处理：openpyxl
- 文件系统操作：os模块
- 多线程支持：threading

## 安装依赖

```bash
pip install openpyxl
```

## 运行方式

```bash
python file2excel_beautified.py
```

## 打包说明

项目使用PyInstaller打包，配置文件为`file2excel_beautified.spec`。

## 贡献

欢迎提交Issue和Pull Request来帮助改进项目。

## 许可证

MIT License