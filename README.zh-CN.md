<div align="center">

# B站视频PPT提取器

**把B站视频里的幻灯片直接转成可编辑打开的 `.pptx` 文件，全程在浏览器内完成。**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Chrome extension](https://img.shields.io/badge/platform-Chrome%20%2F%20Edge-4285F4.svg)](README.zh-CN.md#安装)
[![Manifest V3](https://img.shields.io/badge/manifest-v3-34A853.svg)](manifest.json)
[![Zero dependencies](https://img.shields.io/badge/dependencies-none-FB7299.svg)](README.zh-CN.md#工作原理)

[English](README.md) | **简体中文**（以 [English 版](README.md) 为准）

</div>

## 功能

在B站视频页面上，插件自动逐段跳转抽帧、识别每一次翻页，把每页幻灯片导出为 PPTX 或图片 ZIP。它直接复用页面正在播放的视频流：不下载视频、不经过服务器、不依赖任何第三方库，导出画质即你看到的画质。

## 特性

- **自动翻页识别**：每帧缩小为灰度签名对比，差值超过可调阈值即判定为新的一页。
- **重复页标记**：讲师回翻的旧页会标记为「疑似重复」并默认不勾选。
- **画面裁剪**：人像、水印、角标遮挡画面时，拖框只保留幻灯片区域，识别与导出均只处理该区域。
- **起止时间**：只扫描指定片段（比如跳过片头的正片部分）。
- **采样可调**：采样间隔 1 到 10 秒，翻页阈值按视频灵活调整。
- **双格式导出**：可直接打开的 PPTX，或带时间戳文件名的 JPEG 图片 ZIP。

## 安装

1. 下载或克隆本仓库。
2. 打开 `chrome://extensions`（Edge 为 `edge://extensions`）。
3. 开启右上角 **开发者模式**。
4. 点击 **加载已解压的扩展程序**，选择仓库目录。

在受支持的视频页，页面右下角会出现悬浮按钮；工具栏图标同样可以打开或收起面板。

## 使用

1. 打开B站视频，**先把清晰度切到最高**：抽帧分辨率等于当前播放分辨率（宽度上限 1920）。
2. 点击页面右下角粉色 **提取PPT** 悬浮按钮。
3. 按需调整设置：
   - **采样间隔**（默认 3 秒）：翻页快就调小，视频长就调大。
   - **翻页阈值**（默认 5）：画面小动静造成误判就调大，漏页就调小。
   - **画面裁剪**：播放到幻灯片可见处，点 **选取区域…** 拖框。
   - **起止时间**：支持 `90`、`1:30`、`1:00:30` 格式，终点留空表示扫到结尾。
4. 点 **开始扫描**。视频会暂停，插件开始逐段抽帧，识别结果实时出现在缩略图区。
5. 单击缩略图勾选或取消，双击放大预览，不想要的页直接取消勾选。
6. 点 **⬇ 导出 PPTX**，文件即刻下载。

扫描期间请停留在本页面，不要切换清晰度或分P；插件检测到这类变化会自动清空结果。

## 工作原理

插件完全运行在视频页面的内容脚本里：

1. 程序化驱动页面自身的 `<video>` 元素：先暂停播放器，再按设定范围逐步设置 `currentTime`。
2. 每个采样点抓取一帧：计算 48 x 27 灰度签名用于变化检测，同时保留全分辨率 JPEG（宽度上限 1920）作为候选页。
3. 相邻签名按平均差值比较，超过阈值即开启新片段；每个片段用**最后一帧**（下一次变化前刚稳定的画面）作为该页。
4. 导出时在浏览器内合成 PPTX：`pptx-builder.js` 内置 CRC32、store-only ZIP 写入器和最小但完整的 OOXML 包（presentation、slideMaster、slideLayout、theme、每页一图），零依赖。弹幕和播放器 UI 是 DOM 覆盖层，帧直接取自视频流，不会出现在导出结果里。

## 权限说明

| 权限 | 用途 |
| --- | --- |
| `activeTab` | 仅在你点击工具栏图标时临时访问当前标签页，用于兜底注入内容脚本，无常驻主机权限。 |
| `scripting` | 上述兜底注入的程序化注入能力。 |
| `storage` | 把设置（间隔、阈值、裁剪、时间范围）保存在 `chrome.storage.local`，不读取任何站点数据。 |

本扩展没有 `host_permissions`、无远程代码、无网络请求。内容脚本按 `www.bilibili.com/video/*`、`/bangumi/play/*`、`/festival/*` 静态声明。

## 目录结构

```
bili-ppt-extractor/
├── manifest.json           # MV3 清单
├── background.js           # service worker：工具栏点击与兜底注入
├── content.js              # 扫描、翻页检测、去重、裁剪、界面与导出
├── pptx-builder.js         # CRC32 + store-only ZIP + 最小 OOXML PPTX 生成器
├── icons/                  # 扩展图标
├── dev-assets/             # 图标生成脚本（Pillow）
└── dev-test/               # 测试套件，见「开发与测试」
```

## 开发与测试

依赖：Node.js 18+（仅测试需要，扩展本身零依赖）、Python 3（需 `python-pptx` 与 `Pillow`）。

### 结构测试

用生产的 `pptx-builder.js` 在 Node 里构建 PPTX，再用 python-pptx 打开并校验包结构：

```powershell
python dev-test/make_slides.py    # Pillow 生成 4 张测试幻灯片
node   dev-test/build_test_pptx.js # 用真实构建器构建
python dev-test/validate_pptx.py  # zip 完整性 + python-pptx 校验
```

### 真实视频端到端测试

`dev-test/e2e/` 提供零依赖的 CDP 驱动（`run_e2e.mjs`）：启动真实无头 Edge 并加载扩展，打开真实B站视频，通过 DevTools 协议驱动真实的面板界面，探测视频流找出静态幻灯片时间窗，执行真实扫描，并从浏览器下载目录捕获导出的 PPTX。

自动化提示：正式版 Google Chrome 137 及以上在稳定通道会静默忽略 `--load-extension`，因此测试框架改用 Edge。手动加载扩展在两个浏览器中效果完全相同。

当前状态：测试框架已能以无头 Edge 加载扩展并正常渲染B站页面；完整自动流程（扫描并捕获下载产物）尚未跑通，验收仍以「使用」一节的手动步骤为准。

## 常见问题

**导出的图模糊。**
扫描前把清晰度切到最高，抽帧分辨率等于播放分辨率。

**重复页或误判太多。**
回翻页默认标记且不勾选。摄像头、动态水印造成的误判，可调大翻页阈值或裁剪到幻灯片区域。

**扫描很慢。**
1 小时视频、3 秒间隔约需 1200 次跳转，通常几分钟。可调大间隔或限定时间范围。

**弹幕会进画面吗？**
不会。帧直接取自视频流，弹幕与播放器 UI 是 DOM 覆盖层。

**直播、互动视频支持吗？**
不支持，扫描需要有限时长的可跳转视频流。

## 参与贡献

欢迎 Issue 与 PR。提交信息使用 Conventional Commits（`feat:`、`fix:`、`docs:`、`chore:`）。

## 许可证

[MIT](LICENSE)，Copyright (c) 2026 chenxuhl
