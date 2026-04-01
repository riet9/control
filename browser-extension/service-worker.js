const BRIDGE_URL = "http://127.0.0.1:38945/activity";

function detectBrowser() {
  const ua = navigator.userAgent || "";
  if (ua.includes("Edg/")) {
    return "edge";
  }
  if (ua.includes("OPR/")) {
    return "opera";
  }
  if (ua.includes("Brave")) {
    return "brave";
  }
  return "chromium";
}

function getDomain(url) {
  try {
    const parsed = new URL(url);
    return parsed.hostname.replace(/^www\./, "");
  } catch (error) {
    return "";
  }
}

function shouldReport(url) {
  if (!url) {
    return false;
  }

  return !(
    url.startsWith("chrome://") ||
    url.startsWith("edge://") ||
    url.startsWith("about:") ||
    url.startsWith("chrome-extension://")
  );
}

async function reportActiveTab() {
  const tabs = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
  const tab = tabs && tabs.length > 0 ? tabs[0] : null;
  if (!tab || !shouldReport(tab.url)) {
    return;
  }

  const payload = {
    browser: detectBrowser(),
    title: tab.title || "",
    url: tab.url || "",
    domain: getDomain(tab.url || "")
  };

  try {
    await fetch(BRIDGE_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });
  } catch (error) {
  }
}

chrome.runtime.onInstalled.addListener(() => {
  reportActiveTab();
});

chrome.runtime.onStartup.addListener(() => {
  reportActiveTab();
});

chrome.tabs.onActivated.addListener(() => {
  reportActiveTab();
});

chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
  if (tab.active && (changeInfo.url || changeInfo.title || changeInfo.status === "complete")) {
    reportActiveTab();
  }
});

chrome.windows.onFocusChanged.addListener(() => {
  reportActiveTab();
});
