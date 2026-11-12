# 发布到 GitHub

本项目源码采用 MIT 许可；发布文件包含的第三方代码遵循各自许可，见 `THIRD_PARTY_NOTICES.md`。此文档不表示仓库或 Release 已经发布；发布需要使用仓库所有者的 GitHub 登录状态。

## 发布前核验与使用者试用

发布者应先运行自动测试、浏览器流程与发布包校验；使用者的试用不作为继续开发、核验或准备发布的前提。Windows 桌面 Office 已验证过的样本和版本见 [测试记录](TESTING.md)。发布说明应列明实际测过的项目。

使用者可以在 `release/OfficeBridge.html` 中用两个合成样例完成检查、选择字体建议和下载，再用自己的文件核对副本仍可编辑。目标 Windows Office 的显示、公式和播放效果应在实际交接前核对。具体步骤见 [试用指南](QUICKSTART.md)。

## 准备发布包

在源码目录运行：

```sh
npm ci
npm run release:local
npm run release:verify
```

命令先运行自动测试并重新构建，再生成以下内容：

- `release/OfficeBridge.html`：单文件离线应用，已内嵌第三方许可。
- `release/START-HERE.txt`：便携包使用说明，也单独列入校验清单。
- `release/OfficeBridge-v1.1.0-Windows.zip` 和 `release/OfficeBridge-v1.1.0-Mac.zip`：对应系统的应用与本机 Office 核对助手。
- `release/OfficeBridge-v1.1.0-portable.zip`：应用、合成样例、说明和可选 Windows 验证脚本。
- `release/OfficeBridge-v1.1.0-source.zip`：完整可构建源码、测试和许可，不含依赖安装目录。
- `release/SHA256SUMS.txt` 与 `release/release-manifest.json`：产物校验信息及源码文件清单。

打包程序只收集 Git 已跟踪、且位于明确许可范围内的源文件。新增源文件需要先暂存再打包；拒绝符号链接、缓存和工作目录。`work/`、`node_modules/` 和 `release/` 不进入源码包。不要把真实文稿、私人截图或密钥放进公开源码目录；发布前检查 `release-manifest.json` 中的文件清单及内容。

源码 ZIP 已包含锁文件、测试、合成样例及构建脚本。解压进入 `OfficeBridge-v1.1.0` 后，可按 README 执行 `npm ci`、`npm test`、`npm run build`、`npm run build:portable`。普通构建不需要 Python、Office 或 Playwright；Python 只用于重新生成已附带的合成样例，Playwright 只用于单独的浏览器回归。

`release:local` 会重新生成同版本的本地发布文件；需要保留旧发布包时，先保存到其他位置。文档或源码修改后应重新打包并验证，不沿用旧 ZIP 的校验值。

## 创建公开仓库

在安装了 GitHub CLI 的环境中登录自己的账号：

```sh
gh auth login
gh auth status
```

不要把令牌或验证码发进聊天、源码或 Issue。确认当前目录是本项目，查看即将公开的内容并提交：

```sh
git status --short
git add .github .gitattributes .gitignore AGENTS.md CHANGELOG.md CONTRIBUTING.md LICENSE README.md README.en.md OfficeBridge-Mac.command OfficeBridge-Windows.cmd SECURITY.md THIRD_PARTY_NOTICES.md docs index.html package.json package-lock.json public scripts src tests vite.config.js
git diff --cached --stat
git commit -m "Release OfficeBridge 1.1.0"
```

如果本地已有相同内容的提交，跳过提交步骤。若 Git 尚未设置作者，在这个仓库内设置自己的 Git 姓名和邮箱，不需要更改全局配置。

下面的命令创建当前登录用户下名为 `officebridge` 的**公开**仓库；如该名字已占用，先选一个不同名字，不要覆盖或替换现有仓库：

```sh
gh repo create officebridge --public --source=. --remote=origin --push
```

若项目已连接正确的仓库，使用 `git push`，无需再次创建。GitHub Actions 会在提交后测试三个操作系统的构建；它不启动 Windows Office，也不自动发布或上传测试文稿。

## 添加可下载的 Release

在源码目录创建版本 Release。下列命令只上传明确列出的公开发布产物：

```sh
gh release create v1.1.0 release/OfficeBridge.html release/START-HERE.txt release/OfficeBridge-v1.1.0-Windows.zip release/OfficeBridge-v1.1.0-Mac.zip release/OfficeBridge-v1.1.0-portable.zip release/OfficeBridge-v1.1.0-source.zip release/SHA256SUMS.txt release/release-manifest.json --title "OfficeBridge 1.1.0" --notes-file CHANGELOG.md
```

如果这个版本已经存在，先查看现有 Release，不要使用覆盖参数来替换别人已下载的版本。发布后从 GitHub 下载离线包再试一次，并检查 README 中的下载入口指向真实地址。

应用没有服务端，也不需要部署服务器。需要在线入口时可另行配置静态托管，但单 HTML 已能离线使用。
