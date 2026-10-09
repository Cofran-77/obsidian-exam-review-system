# Obsidian AI 错题复盘工作流

将错题录入、知识点归档、复习诊断和 Word 讲义导出串成一套工作流。支持 AI Agent 按 `AGENTS.md` 协同维护，采用 MIT 许可证。

## 快速开始

```bash
git clone https://github.com/Cofran-77/obsidian-exam-review-system.git
cd obsidian-exam-review-system
pip install -r requirements.txt
python format_docx.py
python -m unittest tests/test_format_docx.py
```

用 Obsidian 打开项目目录；安装社区插件 Dataview 后，复习看板可动态查询卡片。项目附有原创演示卡片，不包含个人错题记录或第三方题库。

## 工作流程

1. 将有权使用的试题和答案 PDF 放入 `试题/` 与 `答案/`。文件命名见 `helper_workflow.py` 的 `SUBJECT_MAP`；该工具用于文字提取和答案页渲染，不包含 OCR。
2. 告诉 AI 学科、题号、作答与错因，例如“马原第 1 题选 B，请按 AGENTS.md 录入”。没有 AI 时也可手动填写模板。
3. 按 `00-系统/模板/错题卡片模板.md` 创建 `{学科目录}/02-错题/*.md`，填好元数据、原题、解析和错因。
4. 创建或复用知识点，并挂载到专题。图谱遵循总纲 → 大纲 → 板块 → 专题 → 知识点；专题可连接错题，错题只连接知识点，知识点允许横向关联。
5. 执行编译与测试，生成根目录 `错题本.md` 和 `错题本.docx`。这两个文件是输出物，不要手改。
6. 复习后手动维护 `review_result`、`review_count`、`last_review`、`next_review` 与 `status`。系统不会自动安排间隔日期。

## 目录与维护

- `format_docx.py`：扫描原子卡片、校验元数据、生成 Markdown 与 Word。
- `helper_workflow.py`：PDF 定位、文本提取与临时答案截图。
- `00-系统/`：数据字典、模板、历史迁移和图谱修复脚本；批量维护前先备份。
- `00-复习看板与诊断中心/`：Dataview 查询。
- `01-马原/` 等五个学科目录：大纲、板块、错题、知识点、专题。
- `tests/`：独立于个人题库数量的测试。

Python 建议 3.10+。临时截图使用后立即删除。模板中的 `{{...}}` 占位符必须替换。

云端同步默认关闭。如需将输出复制到已存在的同步目录，设置 `WRONG_ANSWER_GOOGLE_DRIVE_DIR`。这是单向复制，不是双向同步。

开源仓库的 `.gitignore` 排除了 PDF、个人错题卡片、生成讲义与本地 Obsidian 配置；演示卡片除外。公开自己的知识库前，请自行检查题目授权与个人信息。

## 许可证

工作流代码、模板与原创示例采用 [MIT](LICENSE)。自行添加的试题、答案及资料版权属于各自权利人。
