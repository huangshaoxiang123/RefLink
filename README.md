<p align="center">
  <img src="./banner.svg" alt="RefLink — Zotero 安全交叉引用跳转插件" width="100%">
</p>

# RefLink — Zotero 安全交叉引用跳转插件（Word & WPS 全平台）

> 点击论文中的 Zotero 引文角标，一键直达文末对应参考文献；**角标不变蓝、无下划线、不破坏 Zotero Refresh，全程本地解析不上传文档**。

用 Zotero 写学术论文时，点击正文里的 `[1]`、`[2,5]` 或 `(Smith, 2024)` 引文角标默认无法跳到文末参考文献；而无论是网上常见的 **VBA 开源宏脚本**，还是 **WPS 端 Zotero 插件自带的 `Citation Hyperlinks` 功能**，都会强行往引文域里塞入超链接——不仅让正文引文角标**全部变成蓝色并带下划线（破坏学位论文和期刊排版规范）**，而且 `Citation Hyperlinks` **仅限 WPS 端（Microsoft Word 官方插件不支持）**，点击 **Zotero Refresh** 刷新文献后还容易失效或引发冲突。

**RefLink** 采用安全的**书签 + 选区监听跳转机制**：只在文末参考文献段落建立轻量书签（`RLX_`），**绝不修改 Zotero 引文域内容**，零蓝字、零下划线、完整保留 Zotero 原生刷新与排版能力。

---

## 一、核心优势对比

| 特性 | 传统 VBA 开源宏脚本 | Zotero（WPS 插件）自带 `Citation Hyperlinks` | **RefLink 插件** |
| :--- | :--- | :--- | :--- |
| **支持平台** | 仅部分桌面版 Word（Mac/WPS 常报错） | **仅限 WPS**（Word 官方插件不支持） | **Word & WPS（Mac / Windows 全平台覆盖）** |
| **角标外观与论文排版** | **运行后变蓝并带下划线**，破坏论文排版 | **运行后全部变蓝并带下划线**，不符合期刊/学位论文格式规范 | **完全保持原字体、小上标、黑色、无下划线（零样式污染）** |
| **Zotero Refresh 刷新兼容** | 容易失效或破坏引文域 | 刷新后易丢失或需反复处理 | **原生兼容，绝不向引文域写入 HYPERLINK** |
| **拖动框选文字编辑** | 经常误触发跳转 | 容易误触超链接 | **智能防误跳，框选整段文字修改不干扰编辑** |
| **合并/范围/作者年份引文** | 经常漏跳或跳错编号 | 合并引文如 `[25,26]` 常整段混成一条蓝链接 | **精准二次校验（编号 + 标题 / DOI / 作者年份），合并引文可逐项分别跳转** |
| **脚注 / 尾注引文** | 多不支持 | 支持有限 | **WPS 版完整支持正文、脚注与尾注** |
| **安装与卸载体验** | 手动进宏编辑器粘贴代码 | 依赖第三方 WPS 插件菜单 | **一键安装/卸载程序，自动备份，绝不动 Zotero** |
| **隐私安全** | — | — | **全程本地解析，绝不上传论文正文、文件名或 Zotero 数据库** |

---

## 二、官方下载（安装包 + 图文 PDF 教程）

👉 前往 **[Releases 最新版本页](https://github.com/huangshaoxiang123/RefLink/releases/latest)** 或直接点击下表对应平台下载：

| 对应平台 | 官方安装包下载 | 图文使用教程 | 适用环境与说明 |
| :--- | :--- | :--- | :--- |
| **Mac · WPS 文字** | [RefLink-0.5.0-macOS.zip](https://github.com/huangshaoxiang123/RefLink/releases/latest/download/RefLink-0.5.0-macOS.zip) (152 KB) | [图文 PDF 教程](https://github.com/huangshaoxiang123/RefLink/releases/latest/download/RefLink-Mac-WPS-.pdf) | macOS 12+，支持正文 / 脚注 / 尾注，纯离线运行 |
| **Windows · WPS 文字** | [RefLink-WPS-0.5.0-Windows.zip](https://github.com/huangshaoxiang123/RefLink/releases/latest/download/RefLink-WPS-0.5.0-Windows.zip) (129 KB) | [图文 PDF 教程](https://github.com/huangshaoxiang123/RefLink/releases/latest/download/RefLink-Windows-WPS-.pdf) | Windows 10/11 + WPS 2019+，支持正文 / 脚注 / 尾注，纯离线运行 |
| **Mac · Microsoft Word** | [RefLink-Word-0.5.0-macOS.zip](https://github.com/huangshaoxiang123/RefLink/releases/latest/download/RefLink-Word-0.5.0-macOS.zip) (31 KB) | [图文 PDF 教程](https://github.com/huangshaoxiang123/RefLink/releases/latest/download/RefLink-Mac-Word-.pdf) | macOS 12+ + Microsoft 365 / Word 16.x（支持 WordApi 1.5），正文引文 |
| **Windows · Microsoft Word** | [RefLink-Word-0.5.0-Windows.zip](https://github.com/huangshaoxiang123/RefLink/releases/latest/download/RefLink-Word-0.5.0-Windows.zip) (10 KB) | [图文 PDF 教程](https://github.com/huangshaoxiang123/RefLink/releases/latest/download/RefLink-Windows-Word-.pdf) | Windows 10/11 + Microsoft 365 / 现代 Office 2021（支持 WordApi 1.5），正文引文 |

> 💡 **国内网络备用下载（百度网盘）**：如果校园网直连 GitHub Releases 下载缓慢，可通过百度网盘获取同款安装包与教程：
> - 链接：[https://pan.baidu.com/s/1d8JRiYEYE_Tomy9YgLw8yA?pwd=7ndz](https://pan.baidu.com/s/1d8JRiYEYE_Tomy9YgLw8yA?pwd=7ndz) &nbsp;|&nbsp; 提取码：`7ndz`
> - 网盘文件与 GitHub Release 完全一致，可通过 `.sha256` 核验完整性。

---

## 三、先装先测，满意再激活（一机一码 · 离线永久版）

你可以**免费下载安装**并在自己的论文副本上试用 **“检查文档 / 检测”** 功能，确认插件能 100% 准确识别你论文中的 Zotero 引文域与参考文献，再获取本机永久激活码：

1. **下载安装**：根据电脑系统（Mac / Windows）与软件（WPS / Word）下载对应安装包与 PDF 教程，按教程 2 分钟完成安装。
   - **Mac 用户注意**：首次双击安装器如提示“Apple 无法验证”，请勿点“移到废纸篓”，按 PDF 教程第二步进入 **系统设置 → 隐私与安全性 → 仍要打开** 即可放行。
   - **Windows 用户注意**：务必先**完整解压压缩包**再双击 `install.cmd`；如弹出蓝色 SmartScreen 提示，点 **更多信息 → 仍要运行** 即可。
2. **免费检测文档**：打开一篇保留 Zotero 活动引文域的 `.docx` 论文，点击顶部 `RefLink 交叉引用` → **检查文档（WPS）/ 检测（Word）**，查看引文域数量与匹配报告。
3. **永久离线激活**：点击 **启用/更新跳转（WPS）/ 创建/更新（Word）**，复制弹窗中的**本机机器码**（格式如 `RLD-XXXX-XXXX-XXXX`），添加作者微信获取本机专属永久激活码 / 激活文件（`.txt`）。
   - 激活采用 **ECDSA 离线加密校验**，绑定单台设备**永久有效**，不按年收费，后续小版本更新自动保留授权。
   - **Mac WPS 特别体验**：直接点击激活窗口里的 **“导入激活文件 (.txt)”** 选中作者发来的文件即可一步自动完成永久激活。

---

## 四、支持范围与隐私声明

- **支持引文格式**：GB/T 7714、AMA、Nature、Vancouver、APA、Harvard 等各类数字制（含单个、合并 `[2,5]`、范围 `[1–3]`）与作者-年份制（含 `2020a/b` 后缀、中英文无日期文献）。
- **文档要求**：需要保留 Zotero 活动引文域的 `.docx` 文件（已执行 *Unlink Citations* 变为纯文本的文档无法解析）。
- **零上传隐私承诺**：
  - **WPS 版（Mac / Windows）**：100% 纯离线运行，不发起任何网络请求。
  - **Word 版（Mac / Windows）**：采用微软官方 Office 加载项架构，仅在首次加载时拉取静态前端页面；**你的论文内容、标题、文件名和 Zotero 数据库全程只留在你自己的电脑内存中，绝不上传任何服务器**。

---

## 五、常见问题速览

<details>
<summary><b>1. 已经用了 RefLink，以后还能继续用 Zotero 添加/删除文献并 Refresh 吗？</b></summary>
<br>
完全可以。RefLink 不会向 Zotero 引文中写入 <code>HYPERLINK</code> 域，因此不会破坏 Zotero 刷新。增删文献并点击 Zotero Refresh 后，如果文末编号发生变化，再点一次 <b>启用/更新跳转（WPS）</b> 或 <b>创建/更新（Word）</b> 同步书签即可。
</details>

<details>
<summary><b>2. 卸载 RefLink 会不会删掉我的论文引用或把 Zotero 插件弄坏？</b></summary>
<br>
绝对不会。安装器和卸载器仅操作 RefLink 自己的加载项目录与独立注册记录，安装前会自动备份原有配置，绝不修改 Zotero 插件，也不会删除论文中的 Zotero 引文和参考文献。
</details>

<details>
<summary><b>3. Word 版和 WPS 版有什么区别？我应该装哪个？</b></summary>
<br>
日常用哪个软件写论文就装哪个版本：
<ul>
  <li><b>WPS 版</b>：支持正文、脚注与尾注引文，纯离线运行、跳转零延迟，兼容 WPS 2019 及以上桌面版。</li>
  <li><b>Word 版</b>：支持正文引文（受微软 WordApi 1.5 接口限制，不支持旧版 Office 2016/2019 独立买断版），兼容 Microsoft 365 与现代 Office 2021 桌面版。</li>
</ul>
</details>

<details>
<summary><b>4. 遇到兼容问题如何反馈？</b></summary>
<br>
请在插件顶部点击 <b>诊断信息（WPS）</b> 或 <b>关于/诊断（Word）</b>，点击 <b>复制内容</b>（诊断报告已完全脱敏，不含论文正文与文件名），附带系统/软件版本号提交 <a href="https://github.com/huangshaoxiang123/RefLink/issues">GitHub Issue</a> 或直接联系作者微信。
</details>

---

## 六、获取永久激活码与技术支持

- **作者微信**：`hsx15299177861`（添加请备注 **RefLink**）
- **问题反馈**：欢迎在本仓库 [Issues](https://github.com/huangshaoxiang123/RefLink/issues) 提交使用建议或兼容性反馈。

---

## 版权与许可声明

本软件（RefLink）为**闭源商业授权软件（All Rights Reserved）**，详见 [LICENSE.txt](./LICENSE.txt)。本 GitHub 仓库仅用于官方安装包分发、使用文档展示与用户问题反馈，**未经作者书面许可，禁止对安装包进行逆向工程、反混淆、修改、二次打包或商业转售**。

> “Zotero” 是 Corporation for Digital Scholarship 的注册商标；“Microsoft Word” 和 “WPS” 分别是微软和金山办公的商标。RefLink 为独立第三方工具，与上述官方无隶属或背书关系。
