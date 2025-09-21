// Enable side panel behavior and limit it to the BRACU Connect mark entry pages.
chrome.runtime.onInstalled.addListener(() => {
  if (chrome.sidePanel && chrome.sidePanel.setPanelBehavior) {
    chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true });
  }
});

function updateSidePanelForTab(tabId, url) {
  try {
    if (!chrome.sidePanel || !chrome.sidePanel.setOptions) return;
    const enable = /^https?:\/\/connect\.bracu\.ac\.bd\/app\/exam-controller\/mark-entry\//.test(url || "");
    chrome.sidePanel.setOptions({
      tabId,
      path: "sidepanel.html",
      enabled: enable
    });
  } catch (e) {
    
  }
}

chrome.tabs.onUpdated.addListener((tabId, info, tab) => {
  if (info.status === "complete" && tab?.url) {
    updateSidePanelForTab(tabId, tab.url);
  }
});

chrome.tabs.onActivated.addListener(async ({ tabId }) => {
  try {
    const tab = await chrome.tabs.get(tabId);
    if (tab?.url) updateSidePanelForTab(tabId, tab.url);
  } catch (e) {}
});
