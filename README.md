# Office for Obsidian

[中文](#中文) | [English](#english)

---

## 中文

**Office for Obsidian** 是一个 Obsidian 插件：在 Obsidian 中点击 vault 里的 Word、PowerPoint、Excel 附件，即可直接预览，无需安装 Office 或任何外部程序。全部解析与渲染均在本地完成，不联网、不上传任何数据。

- 插件 ID：`maic-for-office`
- 仅支持查看（只读），不支持编辑；暂不支持旧版二进制格式（`.doc / .ppt / .xls`）

### 功能

| 格式 | 打开方式 | 预览能力 |
| --- | --- | --- |
| `.docx` | 点击附件 | 标题 / 段落 / 字符样式、真实编号的列表、表格（合并单元格、底纹）、行内与独立图片（EMF/WMF 自动转 PNG）、浮动图片环绕、超链接、页眉页脚、目录行、脚注标记 |
| `.pptx` | 点击附件 | 幻灯片画布 + 缩略图侧栏，支持形状、图片、表格、图表、连接线、阴影、渐变等 |
| `.xlsx` | 点击附件 | 多工作表切换、冻结表头的表格网格，显示数值与公式计算结果 |

命令面板提供「Reload active office view」；设置页可调整幻灯片适宽与表格行列上限。

### 安装

**手动安装**：从 [Releases](https://github.com/Aqu-Lab/maic-for-office/releases/latest) 下载 `main.js`、`manifest.json`、`styles.css`，放入 `<vault>/.obsidian/plugins/maic-for-office/`，然后在 设置 → 第三方插件 中启用 **Office for Obsidian**。

上架社区插件库后可直接在插件市场搜索安装。

### 已知限制

- **仅限桌面版**：渲染引擎依赖 Node.js，Obsidian 移动端不可用。
- **Word 预览**为结构化保真排版（非逐像素分页）；浮动图片用近似方式环绕排版；源文档中失效的图片引用（外部链接图）不渲染。
- **PowerPoint 预览**的字体度量使用启发式估算，个别字体下的换行位置可能与 PowerPoint 略有差异。
- **Excel 预览**显示值与公式结果，不含单元格样式、图表与透视表；超大表格按设置页的行列上限截断。

### 本地开发

```bash
npm install
npm run build -w @genoffice/obsidian-plugin   # 产物：apps/obsidian-plugin/dist/
npm run test  -w @genoffice/obsidian-plugin   # 冒烟测试
```

设置环境变量 `OBSIDIAN_TEST_VAULT` 后，构建会自动安装到该 vault 并支持 `--watch` 增量构建。

### 许可

Apache-2.0（见 [LICENSE](LICENSE)）。本项目基于 Apache-2.0 开源办公引擎构建，引擎源码均在本仓库内。

---

## English

**Office for Obsidian** is an Obsidian plugin that lets you preview Word, PowerPoint, and Excel attachments directly inside Obsidian — no Office installation, no external apps. All parsing and rendering happen locally; nothing ever leaves your machine.

- Plugin ID: `maic-for-office`
- Viewing only (read-only); legacy binary formats (`.doc / .ppt / .xls`) are not supported.

### Features

| Format | How to open | Preview |
| --- | --- | --- |
| `.docx` | Click the attachment | Headings / paragraph / character styles, lists with real numbering, tables (merged cells, fills), inline & floating images (EMF/WMF auto-converted to PNG), hyperlinks, headers & footers, TOC lines, footnote marks |
| `.pptx` | Click the attachment | Slide canvas + thumbnail sidebar: shapes, pictures, tables, charts, connectors, shadows, gradients |
| `.xlsx` | Click the attachment | Sheet tabs, sticky-header grid, values and formula results |

The command palette offers "Reload active office view"; settings expose slide fit-width and sheet row/column limits.

### Installing

**Manual**: download `main.js`, `manifest.json`, and `styles.css` from the [latest release](https://github.com/Aqu-Lab/maic-for-office/releases/latest), put them into `<vault>/.obsidian/plugins/maic-for-office/`, then enable **Office for Obsidian** under Settings → Community plugins.

Once listed in the community plugin directory, it will be installable directly from the in-app marketplace.

### Known limitations

- **Desktop only**: the rendering engine relies on Node.js, unavailable on Obsidian Mobile.
- **Word preview** is structured-faithful, not pixel-perfect pagination; floating images use approximate CSS wrapping; broken image references in the source file are skipped.
- **PowerPoint preview** uses heuristic font metrics; line breaks may differ slightly from PowerPoint for some fonts.
- **Excel preview** shows values and formula results only — no cell styling, charts, or pivot tables; oversized sheets are capped by the settings limits.

### Development

```bash
npm install
npm run build -w @genoffice/obsidian-plugin   # output: apps/obsidian-plugin/dist/
npm run test  -w @genoffice/obsidian-plugin   # smoke tests
```

Set the `OBSIDIAN_TEST_VAULT` environment variable and builds auto-install into that vault, with `--watch` incremental rebuilds supported.

### License

Apache-2.0 (see [LICENSE](LICENSE)). Built on open-source (Apache-2.0) office engine packages, all included in this repository.
