# Vic_888
# 基于 Python API 的曲柄滑块机构自动化设计与运动学仿真系统

### 🛠️ 系统架构 (System Architecture)

* **驱动层 (Driver Layer)：** 基于 **Python 3.11** 开发，通过 `win32com` 库实现对 SolidWorks API 的深度调用，支持 **Conda** 环境一键配置。
* **执行层 (Execution Layer)：** 联动 **SolidWorks 2025** 参数化建模引擎，实现曲柄滑块机构的零部件自动生成、装配约束及约束自适应。
* **验证层 (Verification Layer)：** 集成 **Motion Study** 运动学仿真分析，通过 Python 脚本自动化导出仿真数据，验证机构运行逻辑。

---

### 🚀 项目状态 (Project Status)
- [x] 核心建模逻辑已跑通
- [x] 参数化驱动接口已就绪
- [ ] 部分 COM 接口调用（如运动仿真自动启停）尚在调试中 (In Progress)



# 近红外光谱玉米成分预测分析 (Corn Composition Prediction via NIR Spectroscopy)

本项目基于近红外光谱（NIR）数据，利用深度学习模型对玉米的成分进行预测分析。项目包含了完整的数据预处理、模型训练及自动化报告生成流程。

## 🌟 项目亮点 (Project Highlights)
- **自动化工作流**：支持从原始数据到可视化报告的自动生成。
- **深度学习驱动**：采用 PyTorch 框架构建预测模型，能够精准捕捉光谱特征。
- **可视化分析**：包含损失曲线、预测值 vs 真值对比图等关键指标。

## 📂 目录结构 (Directory Structure)
```text
.
├── code/               # 源代码 (Python scripts)
│   ├── Assignment.py    # 模型核心代码
│   └── generate_report.py # 报告生成脚本
├── docs/               # 项目文档 (Documentation)
│   └── 近红外光谱玉米成分预测分析报告.pdf
├── Results/            # 实验结果可视化 (Visualization)
│   ├── Loss_Curve.png
│   └── Prediction_vs_True.png
├── data.xlsx           # 实验数据集 (Dataset)
└── README.md           # 项目说明文档
