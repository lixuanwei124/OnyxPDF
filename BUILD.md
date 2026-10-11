# OnyxPDF 构建说明

## 本地构建 Windows 安装包（.exe）
```bash
npm install
npm run dist        # 产出 dist_out/OnyxPDF Setup 1.0.2.exe
```

## 代码签名（可选）
未签名的安装包在 Windows 上会触发 SmartScreen 提示（点"仍要运行"即可），
macOS 需右键 → 打开 绕过 Gatekeeper。正式分发建议配置代码签名证书：
- Windows: 在 package.json 的 build.win 中增加 `certificateFile` / `certificatePassword`
- macOS: 配置 `CSC_LINK` / `CSC_KEY_PASSWORD` 环境变量

## 输出
打包完成后，`dist_out` 只保留最终 Windows 安装包。
