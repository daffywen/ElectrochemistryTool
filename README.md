# 电化学数据处理工具 v1.0.0

## 更新日期: 2025 年 5 月 21 日

电化学数据处理工具是一个用于自动化处理多种电化学测试数据的 Python 脚本工具。它旨在简化从原始数据文件到结构化 Excel 报告的转换过程，并提取关键的电化学参数。

支持的数据类型包括：

- **循环伏安法 (CV)**: 自动识别 CV 曲线，计算双电层电容 (Cdl) 等参数。
- **线性扫描伏安法 (LSV)**: 处理 LSV 数据，提取相关参数。
- **电化学阻抗谱 (EIS)**: 解析 EIS 数据，计算溶液电阻 (Rs)，并可生成 ZView 兼容格式的纯数据文件。

## 主要功能

- **自动文件识别**: 根据文件内容特征（如特定关键词）自动识别 CV, LSV, EIS 数据文件。
- **数据提取与处理**: 从文本文件中精确提取相关数据列。
  - CV: 电流、电压数据，Cdl 计算中的电流密度差 (Δj)。
  - LSV: 电流、电压数据。
  - EIS: 频率 (Freq), 实部阻抗 (Z'), 虚部阻抗 (Z'')。计算并输出 -Z''。
- **参数计算**:
  - CV: 双电层电容 (Cdl)。
  - EIS: 溶液电阻 (Rs)，通过查找-Z''与 Z'轴的交点或最接近点计算。
- **Tafel 图数据生成**: 结合 LSV 数据和 EIS 推导的 Rs 值，计算 `log(j)` 和 `Overpotential`，并将结果输出到专门的 "Tafel Data" 工作表。
- **格式化 Excel 报告**:
  - 将所有处理后的数据和计算结果统一输出到单个 Excel 工作簿中，包含 `CV Data`, `LSV Data`, `EIS Data`, `Tafel Data` 等详细数据工作表，以及一个汇总的 `Analysis Report` 工作表。
  - `Analysis Report` 工作表汇总了各项分析的关键参数，并作为打开 Excel 文件时的默认显示工作表。
  - 标准化的表头格式：3 行表头信息 + 1 行与表头格式相同的空行，数据从第 5 行开始 (适用于数据表)。
  - EIS 数据表中，Rs 交点处的 Z'和-Z''数据会以黄色背景和加粗字体高亮显示。
  - 支持在同一工作表中并列展示多个同类型文件的处理结果 (适用于 `CV Data`, `LSV Data`, `EIS Data`, `Tafel Data`)。
- **ZView 兼容文件生成**: 对于 EIS 数据，可选择生成纯数据格式的 `.txt` 文件 (命名为 `原文件名-ZView用.txt`)，方便导入 ZView 等专业拟合软件。这些文件将保存在用户选择的原始数据文件夹中。
- **日志记录**: 详细记录程序运行过程、警告和错误信息到 `logs` 文件夹下的日志文件，便于追踪和调试。
- **可执行文件打包**: 提供将 Python 应用程序打包为单个可执行文件 (`ElectrochemistryTool.exe`) 的支持，方便在没有 Python 环境的机器上运行。
- **用户友好的交互**:
  - 通过图形界面选择数据文件夹。
  - 清晰的终端进度提示。

## 安装与运行

### 环境要求

- Python 3.x
- `pip` (Python 包管理工具)

### 安装依赖

确保您的 Python 环境中已安装所有必要的库。在项目根目录 (`cursor` 目录) 下打开终端，运行：

```bash
pip install -r requirements.txt
```

`requirements.txt` 文件应包含以下内容 (至少):

```text
openpyxl
numpy
tqdm # 可选，用于显示进度条
```

### 运行程序

在项目根目录 (`cursor` 目录) 下打开终端，运行：

```bash
python run_electrochemistry.py
```

或者，如果提供了可执行版本 (`ElectrochemistryTool.exe`)，可以直接运行该程序。

## 使用方法

1. **启动程序**: 执行 `python run_electrochemistry.py` 或双击 `ElectrochemistryTool.exe`。
2. **选择文件夹**: 程序会弹出文件夹选择对话框。请选择包含原始电化学数据文件 (`.txt` 格式) 的文件夹。
3. **自动处理**: 程序将自动扫描选定文件夹中的文件：
    - 识别 CV, LSV, EIS 文件。
    - 对识别出的文件进行数据提取和计算。
    - 生成 ZView 兼容文件 (对于 EIS 数据)，保存在您选择的文件夹内。
    - 生成一个包含所有处理结果的 Excel 文件，保存在所选文件夹下的 `processed_data` 子目录中，文件名格式为 `文件夹名_processed_data_时间戳.xlsx`。
4. **查看结果**:
    - 打开生成的 Excel 文件查看详细数据和计算参数。默认打开 "Analysis Report" 工作表。
    - 检查原始文件夹中生成的 `-ZView用.txt` 文件。
    - 如有任何问题，请查看终端输出或 `logs` 文件夹中的日志文件。

## 文件结构说明

```text
cursor/
├── electrochemistry/         # 核心处理模块包
│   ├── common/               # 通用工具模块 (Excel, 文件操作)
│   │   ├── __init__.py
│   │   ├── excel_utils.py
│   │   └── file_utils.py
│   ├── __init__.py
│   ├── cv.py                 # CV数据处理模块
│   ├── eis.py                # EIS数据处理模块
│   ├── lsv.py                # LSV数据处理模块
│   ├── tafel.py              # Tafel分析模块
│   └── main.py               # 主控制逻辑
├── logs/                     # 日志文件存放目录
├── README.md                 # 本说明文件
├── requirements.txt          # Python依赖包列表
├── run_electrochemistry.py   # 程序主入口脚本
└── ElectrochemistryTool.spec # PyInstaller 配置文件

# 用户选择的文件夹 (示例)
selected_data_folder/
├── cv_data_1.txt
├── lsv_data_1.txt
├── eis_data_1.txt
├── eis_data_1-ZView用.txt    # <--- ZView文件会生成在这里
└── processed_data/           # <--- Excel报告会生成在这里
    └── selected_data_folder_processed_data_YYYYMMDD_HHMMSS.xlsx
```

## 注意事项

- **输入文件格式**: 当前主要支持 `.txt` 格式的原始数据文件。请确保您的数据文件结构与各模块（CV, LSV, EIS）的解析逻辑兼容。
  - CV/LSV: 通常需要包含明确的电压和电流数据列。
  - EIS: 需要包含 "A.C. Impedance" 关键词以供识别，并且数据区应有类似 "Freq/Hz, Z'/ohm, Z"/ohm,..." 的表头。
- **依赖安装**: 务必在运行前通过 `pip install -r requirements.txt` 安装所有依赖 (如果不是运行 `.exe` 版本)。
- **Excel 文件写入权限**: 确保程序对目标输出目录（即所选文件夹下的 `processed_data` 子目录）有写入权限。如果 Excel 文件已打开，可能导致保存失败。
- **防病毒软件误报**: 如果使用 `.exe` 版本，某些防病毒软件可能会对其进行标记 (例如，如果使用了 UPX 打包)。如果遇到这种情况，尝试使用未压缩的 `.exe` 版本或从源代码运行。
- **错误排查**: 如遇问题，首先检查终端的错误提示。更详细的信息可以查看 `logs` 文件夹下对应日期的日志文件。

## 未来可能的改进

- 支持更多数据格式和仪器型号。
- 增加数据可视化功能（如图表直接生成在 Excel 或独立图片文件）。
- 提供更多可配置的分析参数。
- 构建更完善的图形用户界面 (GUI)。

## 贡献

欢迎提出改进建议或参与代码贡献。

---

_该工具旨在提高电化学数据处理的效率，如有特定需求或发现 bug，请及时反馈。_
