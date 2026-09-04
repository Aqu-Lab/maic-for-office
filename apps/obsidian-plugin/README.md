# Office for Obsidian

在 Obsidian 中直接查看 vault 里的 `.docx`、`.pptx`、`.xlsx` 附件——解析与渲染全部在本地完成，不联网、不上传。

- 插件 ID / 仓库名：**`maic-for-office`**（Obsidian 官方规则要求 ID 不含 "obsidian"，故项目名 office-for-obsidian 体现在显示名与仓库描述中）
- 只做查看，不做编辑。旧格式 `.doc / .ppt / .xls` 暂不支持。

## 功能

| 格式 | 打开方式 | 渲染 |
| --- | --- | --- |
| `.docx` | 点击附件 | 结构化保真预览：标题/段落/字符样式、列表（真实编号定义）、表格（合并/底纹）、行内与独立图片（EMF/WMF 自动转 PNG）、浮动图片环绕、超链接、页眉页脚、TOC 行、脚注标记 |
| `.pptx` | 点击附件 | 幻灯片画布（Konva）+ 缩略图侧栏，复用 Slides 应用的静态渲染层（形状/图片/表格/图表/连接线/阴影/渐变） |
| `.xlsx` | 点击附件 | 工作表切换 + 冻结表头的表格网格，显示数值与公式缓存结果 |

命令面板提供 `Office for Obsidian: Reload active office view`；设置页可调整幻灯片默认适宽与表格行列上限。

## 本地开发

```bash
npm install
npm run build -w @genoffice/obsidian-plugin    # 产出 dist/{main.js,manifest.json,styles.css}
npm run test   -w @genoffice/obsidian-plugin   # 管线冒烟测试（docx/pptx/xlsx）
```

开发时自动安装：设置环境变量 `OBSIDIAN_TEST_VAULT` 指向测试 vault（或在 `package.json` 里加 `"officeViewer": { "testVault": "/path/to/vault" }`），`npm run build` / `npm run dev`（watch）会自动把产物拷进 `<vault>/.obsidian/plugins/maic-for-office/`，重载插件即可看到效果。

## 发布流程（维护者）

版本发布走 GitHub Actions（`.github/workflows/release.yml`）：

1. 更新 `apps/obsidian-plugin/manifest.json` 的 `version` 与 `apps/obsidian-plugin/versions.json`（键为新版本，值为最低支持的 Obsidian 版本）；
2. 提交并打 tag：`git tag -a <version> -m "<version>" && git push origin <version>`（tag 必须与 manifest 的 version 一致，不带 `v` 前缀）；
3. workflow 自动构建并把 `main.js`、`manifest.json`、`styles.css` 上传到对应 GitHub Release。

## 提交到 Obsidian 社区插件库

1. Fork [obsidianmd/obsidian-releases](https://github.com/obsidianmd/obsidian-releases)；
2. 在 `community-plugins.json` 末尾追加：

   ```json
   {
     "id": "maic-for-office",
     "name": "Office for Obsidian",
     "author": "Aqu-Lab",
     "description": "Preview Word (.docx), PowerPoint (.pptx) and Excel (.xlsx) attachments right inside Obsidian. Fully local, read-only.",
     "repo": "Aqu-Lab/maic-for-office"
   }
   ```

3. 向官方库发起 PR，按模板勾选清单（ID 与仓库名一致、ID 不含 "obsidian"、release 资产齐全、无混淆、无遥测等）。

**与其他 Office 插件冲突**：Obsidian 只允许一个插件注册同一扩展名（按插件加载顺序，先到先得）。本插件启用后，同 vault 里也注册 `.docx/.pptx/.xlsx` 的插件（如 xlsx-viewer、local-office-preview）会加载失败或报警告——属于预期行为，请按需停用其中一个。

## 架构

```
src/
├── main.ts                 # Plugin 入口：registerExtensions + registerView + 设置
├── file-io.ts              # TFile → vault.adapter.readBinary → Uint8Array
├── settings.ts             # 设置页
└── views/
    ├── base-view.ts        # FileView 基类：加载/错误态、onLoadFile 渲染触发、清理
    ├── docx-view.ts        # parseDocx → DocxRenderer
    │   └── render/docx-renderer.ts   # 块树 → 标准只读 DOM（无 Obsidian/React 依赖）
    ├── pptx-view.ts        # openPptx → buildRenderSlide → React root
    │   ├── pptx-canvas.tsx # 只读 Stage（缩略图侧栏 + 主画布）
    │   └── render/pptx/    # 从 apps/slides 复制的静态渲染闭包：
    │                       # konva-adapter / NodeBody(StaticNode) / ChartBody /
    │                       # image-loader / media mime·JPEG·TIFF 辅助 / media-resolver
    └── xlsx-view.ts        # SheetJS 解析 → 表格视图
        └── render/xlsx-table.ts
```

引擎包（`@genoffice/docx-engine`、`@genoffice/pptx-engine`、`@genoffice/pptx-render`）以 workspace 源码直接进 esbuild。`node:*` 内置模块与 `obsidian` 保持 external——Obsidian 桌面版的 Electron renderer 在运行时提供。

### 引擎来源

本项目基于 Apache-2.0 开源办公引擎构建，三个格式共用在仓库内的引擎包：

- **docx**：解析层为通用 docx 引擎；渲染层是为只读场景新写的 DOM 输出。
- **pptx**：渲染树构建与静态 Konva 绘制层（`render/pptx/`）来自仓库内的幻灯片渲染引擎。
- **xlsx**：桌面版原实现的解析走 Rust sidecar（无法随插件分发），此处改用 SheetJS CE，样式保真度有所取舍。

## 已知限制

- **桌面版限定**：引擎有 `node:fs/crypto` 引用，Obsidian 移动端没有 Node，`manifest.json` 标了 `isDesktopOnly: true`。
- **docx 无分页测量**：内容连续排版，不是 Word 级逐像素分页；浮动图片用 CSS float 近似环绕；源文档中失效的图片引用（外部链接图）不渲染。
- **pptx 字体 metrics 为启发式**：当前用启发式度量估算行宽；接入系统字体 + opentype.js（参照 `apps/slides/src/main/fonts.ts`）是下一步保真度重点。
- **xlsx 只显示值**：样式、图表、透视表不渲染；超大表按设置页上限截断。

## 许可

Apache-2.0。构建产物 `main.js` 内嵌了仓库内引擎包的编译代码，对应源码均在本仓库内。
