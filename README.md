# 码上放心追溯码批量处理工具

一个用于批量处理"码上放心"平台追溯码的桌面工具，支持将追溯码整理成 Excel 表格。

## 功能特性

- 输入企业/医院名称
- 批量粘贴追溯码（每行一个）
- 自动去除重复追溯码
- 导出 Excel 表格（自动调整列宽）
- 自定义存储路径（自动保存）
- 完成后一键打开输出文件夹

## 环境要求

- **macOS**: 10.14+
- **Windows**: 10+
- **Python**: 3.8+

## 安装依赖

```bash
pip install PyQt5 pandas openpyxl
```

## 运行方式

### macOS

#### 方式一：直接运行源码

```bash
python trace_code_processor.py
```

#### 方式二：使用打包的应用

双击 `dist/码上放心追溯码工具.app` 即可启动。

如需重新打包：

```bash
pip install pyinstaller
pyinstaller --windowed --name "码上放心追溯码工具" trace_code_processor.py
```

### Windows

直接双击 `dist/码上放心追溯码工具.exe` 即可运行。

如需重新打包（在 Windows 环境中）：

```bash
pip install PyQt5 pandas openpyxl pyinstaller
pyinstaller --onefile --name "码上放心追溯码工具" trace_code_processor.py
```

## 使用说明

1. 打开应用
2. 输入企业/医院名称
3. 在文本框中粘贴追溯码（每行一个）
4. 点击"生成表格"
5. 选择是否清空输入继续处理下一个企业

输出文件默认保存在 `~/Documents/追溯码输出/`，文件命名格式：`企业名称_时间戳.xlsx`

## 文件结构

```
Small Tool/
├── trace_code_processor.py      # 主程序源码
├── README.md                     # 项目文档
├── .gitignore                    # Git 忽略配置
├── dist/                         # 打包输出目录
│   ├── 码上放心追溯码工具.app/    # macOS 应用包
│   └── 码上放心追溯码工具.exe     # Windows 可执行文件
└── build/                        # 打包临时文件
```

## 注意事项

- 追溯码会自动去重，保留首次出现的数据
- 存储路径会保存到系统配置中，下次打开自动恢复
- 生成的 Excel 文件包含两列：企业名称、追溯码
- macOS 和 Windows 版本功能完全一致

## 版本历史

- v1.0.0: 初始版本，基础功能完成
- v1.0.1: 添加存储路径持久化功能
- v1.0.2: 修复跨平台兼容性问题
  - 修复 DPI 设置顺序问题
  - 修复 Windows 打开文件夹命令 (open -> os.startfile)
  - 添加 macOS/Windows 双平台支持
