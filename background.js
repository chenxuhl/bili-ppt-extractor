// B站视频PPT提取器 — MV3 service worker
// 点击扩展图标：切换面板；若内容脚本未注入（少数边缘页面），则程序化注入兜底。
chrome.action.onClicked.addListener(async (tab) => {
  if (!tab || !tab.id) return;
  try {
    await chrome.tabs.sendMessage(tab.id, { type: 'BILI_PPT_TOGGLE' });
  } catch (e) {
    try {
      await chrome.scripting.executeScript({
        target: { tabId: tab.id },
        files: ['pptx-builder.js', 'content.js'],
      });
    } catch (e2) {
      console.warn('[BiliPPT] 注入失败，请确认当前页面是B站视频页：', e2);
    }
  }
});
