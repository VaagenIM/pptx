const viewerParams = new URLSearchParams(window.location.search);
let localPresentationFile = null;

if (viewerParams.get('embed') === '1') {
  document.documentElement.classList.add('pptx-embedded');
}

const downloadButton = document.createElement('button');
downloadButton.className = 'text-button';
downloadButton.type = 'button';
downloadButton.textContent = 'Download';
downloadButton.disabled = true;
document.querySelector('.view-controls').insertBefore(downloadButton, elements.fullscreenButton);

elements.fileInput.addEventListener('change', (event) => {
  localPresentationFile = event.target.files?.[0] ?? null;
}, true);

const getRemotePresentationUrl = () => new URLSearchParams(window.location.search).get('url');
const getDownloadName = (url) => {
  try {
    return decodeURIComponent(new URL(url).pathname.split('/').pop() || 'presentation.pptx');
  } catch {
    return 'presentation.pptx';
  }
};
const syncDownloadButton = () => {
  downloadButton.disabled = elements.app.classList.contains('is-empty') || (!localPresentationFile && !getRemotePresentationUrl());
};

downloadButton.addEventListener('click', async () => {
  const remoteUrl = getRemotePresentationUrl();
  const href = localPresentationFile ? URL.createObjectURL(localPresentationFile) : remoteUrl;
  if (!href) return;
  const link = document.createElement('a');
  let objectUrl = localPresentationFile ? href : null;
  link.href = href;
  link.download = localPresentationFile?.name || getDownloadName(remoteUrl);
  link.rel = 'noopener';
  try {
    if (remoteUrl && !localPresentationFile) {
      const response = await fetch(remoteUrl, { credentials: 'omit' });
      if (!response.ok) throw new Error(`Download returned HTTP ${response.status}.`);
      objectUrl = URL.createObjectURL(await response.blob());
      link.href = objectUrl;
    }
    link.click();
  } catch (error) {
    link.href = remoteUrl;
    link.removeAttribute('download');
    elements.status.textContent = `Direct download unavailable; opening the source URL. ${error.message}`;
    link.click();
  } finally {
    if (objectUrl) window.setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
  }
});

const downloadObserver = new MutationObserver(syncDownloadButton);
downloadObserver.observe(elements.app, { attributes: true, attributeFilter: ['class'] });
syncDownloadButton();

if (['0', 'false', 'hidden'].includes(viewerParams.get('sidebar'))) {
  const collapseSidebar = () => {
    if (
      !elements.app.classList.contains('is-empty') &&
      !elements.app.classList.contains('sidebar-closed') &&
      !elements.sidebarToggle.disabled
    ) {
      elements.sidebarToggle.click();
    }
  };
  const sidebarObserver = new MutationObserver(collapseSidebar);
  sidebarObserver.observe(elements.app, { attributes: true, attributeFilter: ['class'] });
  sidebarObserver.observe(elements.sidebarToggle, { attributes: true, attributeFilter: ['disabled'] });
  collapseSidebar();
}

let fullscreenZoom;

document.addEventListener('fullscreenchange', () => {
  if (!viewer) return;
  if (document.fullscreenElement === elements.stage) {
    fullscreenZoom = viewer.zoomPercent;
    const fitScale = elements.viewerContainer.clientWidth / viewer.slideWidth;
    const heightScale = elements.stage.clientHeight / viewer.slideHeight;
    void viewer.setZoom((heightScale / fitScale) * 100);
  } else if (fullscreenZoom !== undefined) {
    void viewer.setZoom(fullscreenZoom);
    fullscreenZoom = undefined;
  }
});
