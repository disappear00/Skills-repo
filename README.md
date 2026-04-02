# Skills-repo

[中文文档](./README-zh.md) | English

Store popular open-source AI Skills & my original self-developed custom Skills, including prompt skills, agent workflow skills, multi-modal practical skills and iterative optimized case archives.

## Overview

This repository contains a collection of AI-powered skills designed to enhance productivity and automate various tasks. The skills are organized into two main categories:

- **Open Skills**: Publicly available skills that can be freely used and modified
- **Close Skills**: Proprietary or restricted skills with specific usage requirements

## Directory Structure

```
Skills-repo/
├── open-skills/              # Publicly available skills
│   ├── chart-visualization/  # Data visualization with 26 chart types
│   ├── consulting-analysis/  # Professional research report generation
│   ├── data-analysis/        # Excel/CSV data analysis
│   ├── deep-research/        # Systematic web research methodology
│   ├── frontend-design/      # Production-grade UI development
│   ├── image-generation/     # AI image generation
│   ├── podcast-generation/   # Text-to-podcast audio conversion
│   ├── ppt-generation/       # PowerPoint presentation generation
│   ├── video-generation/     # AI video generation
│   └── web-design-guidelines/# UI/UX compliance review
├── close-skills/             # Restricted access skills
│   └── docx-operate-skills/  # Word document manipulation
├── README.md
└── README-zh.md
```

## Open Skills

### 1. Chart Visualization

**Location**: [open-skills/chart-visualization](./open-skills/chart-visualization/)

Intelligent data visualization skill that automatically selects the most appropriate chart type from 26 available options.

**Supported Chart Types**:
| Category | Chart Types |
|----------|-------------|
| Time Series | Line, Area, Dual Axes |
| Comparisons | Bar, Column, Histogram |
| Part-to-Whole | Pie, Treemap |
| Relationships | Scatter, Sankey, Venn, Network Graph |
| Maps | District Map, Pin Map, Path Map |
| Hierarchies | Organization Chart, Mind Map |
| Specialized | Radar, Funnel, Liquid, Word Cloud, Boxplot, Violin, Fishbone, Flow Diagram, Spreadsheet |

**Dependencies**: Node.js >= 18.0.0

### 2. Consulting Analysis

**Location**: [open-skills/consulting-analysis](./open-skills/consulting-analysis/)

Generates professional, consulting-grade research reports following McKinsey/BCG standards.

**Capabilities**:
- Market analysis and consumer insights
- Brand strategy and competitive intelligence
- Financial analysis and investment research
- Industry research and macroeconomic analysis

### 3. Data Analysis

**Location**: [open-skills/data-analysis](./open-skills/data-analysis/)

Analyzes Excel/CSV files using DuckDB for efficient data processing.

**Capabilities**:
- Schema inspection and data exploration
- SQL-based querying
- Statistical summaries
- Multi-sheet Excel support
- Export to CSV/JSON/Markdown

### 4. Deep Research

**Location**: [open-skills/deep-research](./open-skills/deep-research/)

Systematic methodology for conducting thorough web research with multi-angle analysis.

**Research Phases**:
1. Broad Exploration - Understand the landscape
2. Deep Dive - Targeted research on key dimensions
3. Diversity & Validation - Multiple information sources
4. Synthesis Check - Comprehensive coverage verification

### 5. Frontend Design

**Location**: [open-skills/frontend-design](./open-skills/frontend-design/)

Creates distinctive, production-grade frontend interfaces with high design quality.

**Design Principles**:
- Bold aesthetic direction with intentionality
- Distinctive typography and color schemes
- Meaningful animations and micro-interactions
- Avoids generic "AI slop" aesthetics

### 6. Image Generation

**Location**: [open-skills/image-generation](./open-skills/image-generation/)

Generates high-quality images using structured prompts and optional reference images.

**Capabilities**:
- Character design and scene creation
- Style-guided generation with reference images
- Multiple aspect ratios support
- Structured JSON prompt format

### 7. Podcast Generation

**Location**: [open-skills/podcast-generation](./open-skills/podcast-generation/)

Converts text content into natural two-host conversational podcast audio.

**Capabilities**:
- Text-to-speech synthesis
- Male and female host voices
- Supports English and Chinese
- Generates transcript alongside audio

### 8. PPT Generation

**Location**: [open-skills/ppt-generation](./open-skills/ppt-generation/)

Generates professional PowerPoint presentations with AI-generated slide images.

**Presentation Styles**:
| Style | Best For |
|-------|----------|
| glassmorphism | Tech products, AI/SaaS demos |
| dark-premium | Premium products, executive presentations |
| gradient-modern | Startups, creative agencies |
| neo-brutalist | Edgy brands, Gen-Z targeting |
| 3d-isometric | Tech explainers, product features |
| editorial | Annual reports, luxury brands |
| minimal-swiss | Architecture, design firms |
| keynote | Keynotes, product reveals |

### 9. Video Generation

**Location**: [open-skills/video-generation](./open-skills/video-generation/)

Generates high-quality videos using structured prompts and optional reference images.

**Capabilities**:
- Reference image as guided frame
- Multiple aspect ratios
- Structured JSON prompt format

### 10. Web Design Guidelines

**Location**: [open-skills/web-design-guidelines](./open-skills/web-design-guidelines/)

Reviews UI code for Web Interface Guidelines compliance.

**Capabilities**:
- Accessibility audit
- UX best practices check
- Design pattern validation

## Close Skills

### DOCX Operate Skills

**Location**: [close-skills/docx-operate-skills](./close-skills/docx-operate-skills/)

Lightweight skill for manipulating Microsoft Word documents programmatically.

**Core Functions**:
- `replace_section_by_title` - Replace content under a heading
- `clear_section_content` - Remove content while keeping heading
- `add_hyperlink_to_heading` - Add hyperlinks to headings
- `get_subtitles_under_heading` - Extract subheadings
- `insert_heading_after_subtitles` - Insert new headings

## Usage

Each skill contains a `SKILL.md` file with detailed documentation including:
- Skill description and trigger conditions
- Dependencies and requirements
- Step-by-step workflow
- Code examples and parameters

Navigate to the specific skill directory to learn more about its usage.

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues to improve existing skills or add new ones.

## License

Each skill may have its own license. Please check the individual `SKILL.md` files and any accompanying `LICENSE` files for specific licensing information.
