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
