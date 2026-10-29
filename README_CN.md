<h1 align="center">RotorHazard 全成绩导出插件（All Results Exporter）</h1>

<p align="center">
  <b>一键把整场赛事所有已保存比赛、飞手与逐圈时间导出为一个排版好的 Excel。</b><br>
  <a href="./README.md">[English]</a>
</p>

<p align="center">
  <a href="https://github.com/LeoFengFPV/RH-All-Results-Exporter/actions/workflows/rhfest.yml">
    <img src="https://github.com/LeoFengFPV/RH-All-Results-Exporter/actions/workflows/rhfest.yml/badge.svg" alt="RHFest">
  </a>
  <a href="https://github.com/LeoFengFPV/RH-All-Results-Exporter/releases">
    <img src="https://img.shields.io/github/v/release/LeoFengFPV/RH-All-Results-Exporter" alt="Release">
  </a>
  <a href="./LICENSE">
    <img src="https://img.shields.io/badge/License-MIT-yellow.svg" alt="License: MIT">
  </a>
  <img src="https://img.shields.io/badge/RotorHazard-4.4.0-blue" alt="Tested on RotorHazard 4.4.0">
  <img src="https://img.shields.io/badge/RHAPI-1.4-blue" alt="Tested on RHAPI 1.4">
</p>

---

## 项目背景与痛点

RotorHazard 自带的导出是 JSON，数据完整但普通人不方便打开和阅读；一些社区导出插件只能导出**最后一场**，逼操作员飞一场导一场。本插件解决这个痛点：**一键导出整场赛事**，输出排版好的电子表格，包含逐圈时间和每场比赛时实际使用的频点，方便比赛结尾复核。

## 功能

- 一键把**所有已保存比赛**汇总到单个工作表。
- 每位实际参赛飞手一行。
- 逐圈时间按列展开（第0圈、第1圈、第2圈……）。
- 比赛**频点**读取比赛时保存的记录，每位飞手只显示一个。
- 全空列自动隐藏。
- 类别、分组、轮次只在同一场实际比赛边界内合并单元格。
- 未分配飞手的空赛道位不产生成绩行。
- 以 `=` 开头的飞手名按文本保存，不会变成 Excel 公式。
- 按需执行，没有后台轮询或定时任务。

## 导出字段说明

| 字段 | 含义 |
|---|---|
| 类别 | 比赛类别名称 |
| 分组 | 分组/Heat 名称 |
| 轮次 | 轮次编号 |
| 名次 | 完赛名次 |
| 飞手 | 飞手呼号 |
| 团队 | 团队名（如有） |
| 频点 (MHz) | 该场为该飞手保存的频点 |
| 圈数 | 计圈数量 |
| 平均圈速 | 平均每圈用时 |
| 最快圈速 | 最快单圈用时 |
| 第0圈 … 第N圈 | 每一圈的单独用时 |
| 总时间 | 总用时（固定为最后一列） |

圈号与总时间规则：

- **Hole Shot（发令枪）** 与 **Staggered Start（错时发车）** 从**第0圈**开始编号（第0圈为起跑圈）。
- **First Lap（首圈制）** 从**第1圈**开始编号（没有第0圈）。
- **Staggered Start** 的总时间取 `total_time_laps`，不含起跑圈。

某列如果整列全空（例如一直没用的类别），就不会导出；类别、分组、轮次只在单场实际比赛内合并，绝不跨场次。

## 安装流程

1. 打开 RotorHazard，进入 **插件 → 上传**。
2. 选择 Release 附件 `all_results_exporter.zip` 上传。
3. 按提示重启 RotorHazard。
4. 在 **数据管理 → 导出器** 中出现 **Export All Results (XLSX)**。

依赖 [openpyxl](https://pypi.org/project/openpyxl/)。也可手动安装：把 `all_results_exporter` 文件夹放到 `~/rh-data/plugins/` 后重启。

## 使用路径

1. 正常飞行并**保存**比赛。
2. 赛事（或某一阶段）结束后，打开 **数据管理 → 导出器**。
3. 选择 **Export All Results (XLSX)** 下载文件。
4. 用 Excel、WPS 等打开复核。

导出范围仅包含**已保存比赛**；未保存的比赛或被丢弃的现场圈数据不在范围内。

## 兼容性

- 已在 **RotorHazard v4.4.0 / RHAPI 1.4** 测试。
- 因未测试更低版本，manifest 最低 RHAPI 版本保守设为 **1.4**。
- 标准环境包含 Python 3 与 `openpyxl` 依赖。

## 已验证内容

树莓派真实环境：

- RotorHazard v4.4.0 / RHAPI 1.4
- Raspberry Pi 4 Model B Rev 1.5、Python 3.13.5、openpyxl 3.1.5
- 真实赛事：**10 场已保存比赛、30 行飞手数据**
- 导出约 **0.2 秒**，页面响应正常
- Hole Shot 成绩对照真实硬件比赛逐格核对通过

`tests/` 下的自动化回归测试覆盖重名飞手、空赛道位、占位行、删圈、三种起跑方式、合并边界、公式型飞手名与文件重开，全部通过。

## 已知限制

- **First Lap** 与 **Staggered Start** 使用 RotorHazard **真实计分代码**验证，但圈速来自临时注入数据，并非实际飞行；首次在正式赛事使用这两种赛制时，请再核对起始圈号与总时间列。
- 原测试环境中**没有直接点击网页"插件上传"按钮**，而是用等效命令行安装加真实导出验证加载与输出。
- 未分配飞手的空赛道位（有模块频点、无飞手、无圈速）会被省略；这是给人看的成绩表的有意设计，尽管原生 JSON 备份会保留这些条目。

## 性能说明

- 真实赛事规模（10 场 / 30 行）：约 **0.2 秒**，内存增量可忽略，页面响应正常。
- 约 **4,000 行**的模拟压力测试（mock 数据，**非真实赛事**）约 **19.6 秒**；大型赛事请在非计时阶段导出。
- 不宣称零资源占用，时间与内存随已保存比赛和圈数增长。

## 故障排查

- **没有出现下载**：可能没有已保存比赛，或出现内部错误，请查看 RotorHazard 日志。
- **少了某位飞手**：确认该场已保存且该赛道位分配了飞手；空赛道位会被有意跳过。
- **少了某一圈**：被删除的圈不会导出；若手动删圈或补圈，请重新核对已保存比赛。
- **文件打不开**：确认下载完整，且计时器已安装 openpyxl。

## 隐私说明

仓库与 Release 不含真实比赛数据库、客户资料、飞手真实姓名、局域网 IP、日志、账号、Token 或 Cookie。示例数据与截图使用匿名飞手（如 **Pilot A**、**Pilot B**）。发布包在上传前会清理缓存与元数据。

## 开源协议

基于 [MIT 协议](./LICENSE)开源。
