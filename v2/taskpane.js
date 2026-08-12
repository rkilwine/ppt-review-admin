/* global Office, document */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    showFileName();
  }
});

function showFileName() {
  const el = document.getElementById("file-name");
  try {
    Office.context.document.getFilePropertiesAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value.url) {
        const url = result.value.url;
        const name = url.substring(url.lastIndexOf("\\") + 1).substring(url.lastIndexOf("/") + 1);
        el.textContent = name || "Unsaved presentation";
      } else {
        el.textContent = "Unsaved presentation";
      }
    });
  } catch (e) {
    el.textContent = "";
  }
}
