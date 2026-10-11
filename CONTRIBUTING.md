# 参与贡献

感谢关注 OnyxPDF。反馈问题和提交代码前，请先阅读以下约定。

## 反馈问题（Issue）

提交 Issue 时请包含：

- OnyxPDF 版本号（关于对话框中可见）与 Windows 版本；
- 复现步骤：打开什么文件、点了什么、期望结果与实际结果；
- 如涉及渲染错误，请附最小复现 PDF（注意先去除文档中的隐私内容）。

## 提交代码（Pull Request）

1. Fork 仓库并从 `main` 创建功能分支；
2. 本地跑通再提交：

```bash
npm install
npm start
```

3. 保持改动聚焦，一个 PR 解决一个问题；
4. 提交信息使用简短祈使句（参考现有提交历史，如 `Add PDF decryption module`）；
5. 不提交 `dist_out/`、`node_modules/`、`HANDOFF.md` 等本地内容（已在 `.gitignore` 中）。

## 开发环境

- Node.js（建议 LTS 版本）与 npm；
- 依赖与项目结构见 [BUILD.md](BUILD.md) 与 README 的「开发」章节。
