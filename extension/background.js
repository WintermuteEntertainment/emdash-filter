// Toolbar button opens the app in a full tab (room to review each dash).
const api = typeof browser !== "undefined" ? browser : chrome;
api.action.onClicked.addListener(() => {
  api.tabs.create({ url: api.runtime.getURL("app.html") });
});
