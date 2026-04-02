# Skills-repo

存储流行的开源 AI 技能与自主开发的定制技能，包括提示词技能、智能体工作流技能、多模态实战技能以及迭代优化的案例存档。

## 概述

本仓库包含一系列旨在提高生产力和自动化各种任务的 AI 驱动技能。技能分为两大类：

- **开放技能（Open Skills）**：可自由使用和修改的公开技能
- **封闭技能（Close Skills）**：具有特定使用要求的专有或受限技能

## 目录结构

```
Skills-repo/
├── open-skills/              # 公开可用的技能
│   ├── chart-visualization/  # 数据可视化（26种图表类型）
│   ├── consulting-analysis/  # 专业研究报告生成
│   ├── data-analysis/        # Excel/CSV数据分析
│   ├── deep-research/        # 系统性网络研究方法论
│   ├── frontend-design/      # 生产级前端开发
│   ├── image-generation/     # AI图像生成
│   ├── podcast-generation/   # 文本转播客音频
│   ├── ppt-generation/       # PowerPoint演示文稿生成
│   ├── video-generation/     # AI视频生成
│   └── web-design-guidelines/# UI/UX合规审查
├── close-skills/             # 受限访问技能
│   └── docx-operate-skills/  # Word文档操作
├── README.md
└── README-zh.md
```

## 开放技能

### 1. 图表可视化

**位置**：[open-skills/chart-visualization](./open-skills/chart-visualization/)

智能数据可视化技能，可从26种可用图表类型中自动选择最合适的图表。

**支持的图表类型**：
| 类别 | 图表类型 |
|------|----------|
| 时间序列 | 折线图、面积图、双轴图 |
| 比较 | 条形图、柱状图、直方图 |
| 占比 | 饼图、矩形树图 |
| 关系 | 散点图、桑基图、韦恩图、关系图 |
| 地图 | 区域地图、标注地图、路径地图 |
| 层级 | 组织架构图、思维导图 |
| 专用 | 雷达图、漏斗图、水波图、词云图、箱线图、小提琴图、鱼骨图、流程图、表格 |

**依赖**：Node.js >= 18.0.0

### 2. 咨询分析

**位置**：[open-skills/consulting-analysis](./open-skills/consulting-analysis/)

生成符合麦肯锡/波士顿咨询标准的专业研究报告。

**能力**：
- 市场分析和消费者洞察
- 品牌战略和竞争情报
- 财务分析和投资研究
- 行业研究和宏观经济分析

### 3. 数据分析

**位置**：[open-skills/data-analysis](./open-skills/data-analysis/)

使用 DuckDB 分析 Excel/CSV 文件，实现高效数据处理。

**能力**：
- 模式检查和数据探索
- 基于 SQL 的查询
- 统计摘要
- 多工作表 Excel 支持
- 导出为 CSV/JSON/Markdown

### 4. 深度研究

**位置**：[open-skills/deep-research](./open-skills/deep-research/)

进行多角度全面网络研究的系统方法论。

**研究阶段**：
1. 广泛探索 - 了解整体情况
2. 深入挖掘 - 针对关键维度进行定向研究
3. 多样性与验证 - 多种信息来源
4. 综合检查 - 验证覆盖完整性

### 5. 前端设计

**位置**：[open-skills/frontend-design](./open-skills/frontend-design/)

创建具有高设计质量的独特、生产级前端界面。

**设计原则**：
- 具有明确意图的大胆美学方向
- 独特的字体和配色方案
- 有意义的动画和微交互
- 避免通用的"AI生成"美学

### 6. 图像生成

**位置**：[open-skills/image-generation](./open-skills/image-generation/)

使用结构化提示词和可选参考图像生成高质量图像。

**能力**：
- 角色设计和场景创建
- 使用参考图像进行风格引导生成
- 支持多种宽高比
- 结构化 JSON 提示词格式

### 7. 播客生成

**位置**：[open-skills/podcast-generation](./open-skills/podcast-generation/)

将文本内容转换为自然流畅的双主持人对话式播客音频。

**能力**：
- 文本转语音合成
- 男性和女性主持人声音
- 支持中英文
- 同时生成音频和文字稿

### 8. PPT生成

**位置**：[open-skills/ppt-generation](./open-skills/ppt-generation/)

生成带有 AI 生成幻灯片图像的专业 PowerPoint 演示文稿。

**演示风格**：
| 风格 | 适用场景 |
|------|----------|
| glassmorphism | 科技产品、AI/SaaS演示 |
| dark-premium | 高端产品、高管演示 |
| gradient-modern | 初创公司、创意机构 |
| neo-brutalist | 前卫品牌、Z世代目标群体 |
| 3d-isometric | 技术讲解、产品功能展示 |
| editorial | 年度报告、奢侈品牌 |
| minimal-swiss | 建筑、设计公司 |
| keynote | 主题演讲、产品发布 |

### 9. 视频生成

**位置**：[open-skills/video-generation](./open-skills/video-generation/)

使用结构化提示词和可选参考图像生成高质量视频。

**能力**：
- 参考图像作为引导帧
- 支持多种宽高比
- 结构化 JSON 提示词格式

### 10. 网页设计指南

**位置**：[open-skills/web-design-guidelines](./open-skills/web-design-guidelines/)

审查 UI 代码是否符合网页界面指南。

**能力**：
- 无障碍审计
- UX 最佳实践检查
- 设计模式验证

## 封闭技能

### DOCX 操作技能

**位置**：[close-skills/docx-operate-skills](./close-skills/docx-operate-skills/)

用于程序化操作 Microsoft Word 文档的轻量级技能。

**核心功能**：
- `replace_section_by_title` - 替换标题下的内容
- `clear_section_content` - 删除内容但保留标题
- `add_hyperlink_to_heading` - 为标题添加超链接
- `get_subtitles_under_heading` - 提取子标题
- `insert_heading_after_subtitles` - 插入新标题

## 使用方法

每个技能都包含一个 `SKILL.md` 文件，其中包含详细文档：
- 技能描述和触发条件
- 依赖和要求
- 分步工作流程
- 代码示例和参数

请导航到具体的技能目录以了解更多使用详情。

## 贡献

欢迎贡献！请随时提交 Pull Request 或创建 Issue 来改进现有技能或添加新技能。

## 许可证

每个技能可能有各自的许可证。请查看各技能的 `SKILL.md` 文件和附带的 `LICENSE` 文件以获取具体的许可信息。
