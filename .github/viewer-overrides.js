const viewerParams = new URLSearchParams(window.location.search);
let localPresentationFile = null;

if (viewerParams.get('embed') === '1') {
  document.documentElement.classList.add('pptx-embedded');
  document.querySelector('.brand')?.removeAttribute('href');
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
let fullscreenFitMode;
let boundViewer = null;

const syncFullscreenSlide = () => {
  if (!viewer || document.fullscreenElement !== elements.stage) return;
  const currentIndex = String(viewer.currentSlideIndex);
  document.querySelectorAll('.viewer-container > [data-slide-index]').forEach((item) => {
    item.classList.toggle('pptx-fullscreen-slide', item.dataset.slideIndex === currentIndex);
  });
};

const bindViewerEvents = () => {
  if (!viewer || boundViewer === viewer) return;
  boundViewer = viewer;
  viewer.addEventListener('slidechange', syncFullscreenSlide);
};

const fitFullscreenSlide = async () => {
  if (!viewer || document.fullscreenElement !== elements.stage) return;
  const frameWidth = Math.min(elements.stage.clientWidth, elements.stage.clientHeight * (16 / 9));
  const frameHeight = frameWidth * (9 / 16);
  const slideScale = Math.min(frameWidth / viewer.slideWidth, frameHeight / viewer.slideHeight);
  await viewer.setFitMode('none');
  await viewer.setZoom(slideScale * 100);
  syncFullscreenSlide();
  await new Promise((resolve) => requestAnimationFrame(resolve));
  const activeSlide = document.querySelector('.pptx-fullscreen-slide > div > *');
  if (!activeSlide) return;
  const rendered = activeSlide.getBoundingClientRect();
  const correction = Math.min(frameWidth / rendered.width, frameHeight / rendered.height);
  if (Number.isFinite(correction) && Math.abs(correction - 1) > 0.001) {
    await viewer.setZoom(viewer.zoomPercent * correction);
  }
};

document.addEventListener('fullscreenchange', () => {
  if (!viewer) return;
  bindViewerEvents();
  if (document.fullscreenElement === elements.stage) {
    fullscreenZoom = viewer.zoomPercent;
    fullscreenFitMode = viewer.fitMode;
    requestAnimationFrame(() => {
      syncFullscreenSlide();
      void fitFullscreenSlide();
    });
  } else if (fullscreenZoom !== undefined) {
    document.querySelectorAll('.pptx-fullscreen-slide').forEach((item) => item.classList.remove('pptx-fullscreen-slide'));
    void viewer.setFitMode(fullscreenFitMode ?? 'contain').then(() => viewer.setZoom(fullscreenZoom));
    fullscreenZoom = undefined;
    fullscreenFitMode = undefined;
  }
});

const viewerObserver = new MutationObserver(() => {
  bindViewerEvents();
  syncFullscreenSlide();
});
viewerObserver.observe(elements.viewerContainer, { childList: true, subtree: true, attributes: true, attributeFilter: ['aria-busy'] });
