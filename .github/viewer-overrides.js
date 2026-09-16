const viewerParams = new URLSearchParams(window.location.search);
let localPresentationFile = null;

for (const name of [
  'pptx-primary',
  'pptx-primary-strong',
  'pptx-primary-ink',
  'pptx-primary-wash',
  'pptx-text-color',
  'pptx-text-soft',
  'pptx-text-muted',
  'pptx-logo-color',
  'pptx-font-family',
  'pptx-mono-font-family',
  'pptx-font-size-body',
  'pptx-font-size-meta',
  'pptx-font-size-control',
  'pptx-font-size-heading',
  'pptx-font-size-brand',
  'pptx-surface-deep',
  'pptx-background',
  'pptx-surface',
  'pptx-surface-raised',
  'pptx-surface-hover',
  'pptx-border',
  'pptx-border-soft',
  'pptx-topbar-background',
  'pptx-stage-background',
  'pptx-sidebar-background',
  'pptx-dialog-background',
  'pptx-dialog-shadow',
  'pptx-shadow',
  'pptx-danger',
  'pptx-fullscreen-background',
  'pptx-slide-background',
]) {
  const value = viewerParams.get(name);
  if (value) document.documentElement.style.setProperty(`--${name}`, value);
}

window.addEventListener('message', (event) => {
  if (event.data?.type !== 'pptx-theme' || !event.data.variables) return;
  for (const [name, value] of Object.entries(event.data.variables)) {
    if (/^--pptx-[a-z-]+$/.test(name) && typeof value === 'string') {
      document.documentElement.style.setProperty(name, value);
    }
  }
});

if (viewerParams.get('embed') === '1') {
  document.documentElement.classList.add('pptx-embedded');
  const brand = document.querySelector('.brand');
  brand?.removeAttribute('href');
  const brandLabel = brand?.querySelector('span:last-child');
  if (brandLabel) brandLabel.textContent = 'pptx';
}

if (viewerParams.get('compact') === '1') {
  document.documentElement.classList.add('pptx-compact');
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
const youtubeHosts = new Set(['youtube.com', 'www.youtube.com', 'm.youtube.com', 'youtube-nocookie.com', 'www.youtube-nocookie.com']);
const getYoutubeEmbedUrl = (value) => {
  try {
    const url = new URL(value);
    return youtubeHosts.has(url.hostname.toLowerCase()) && url.pathname.startsWith('/embed/') ? url.href : null;
  } catch {
    return null;
  }
};
const upgradeYoutubeVideos = (root) => {
  root.querySelectorAll('video').forEach((video) => {
    const embedUrl = getYoutubeEmbedUrl(video.currentSrc || video.getAttribute('src') || '');
    if (!embedUrl) return;
    const iframe = document.createElement('iframe');
    iframe.className = video.className;
    iframe.style.cssText = video.style.cssText;
    iframe.src = embedUrl;
    iframe.title = 'YouTube video player';
    iframe.frameBorder = '0';
    iframe.allow = 'accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share';
    iframe.referrerPolicy = 'strict-origin-when-cross-origin';
    iframe.allowFullscreen = true;
    video.replaceWith(iframe);
  });
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
    if (elements.app.classList.contains('is-empty')) return;
    elements.app.classList.add('sidebar-closed');
    elements.sidebarToggle.setAttribute('aria-expanded', 'false');
    elements.sidebarToggle.setAttribute('aria-label', 'Show slide thumbnails');
    elements.sidebarToggleTooltip.textContent = 'Show slide thumbnails';
    sidebarObserver.disconnect();
  };
  const sidebarObserver = new MutationObserver(collapseSidebar);
  sidebarObserver.observe(elements.app, { attributes: true, attributeFilter: ['class'] });
  collapseSidebar();
}

let fullscreenZoom;
let fullscreenFitMode;
let fullscreenOperation = 0;
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

const enterFullscreenFit = async () => {
  if (!viewer || document.fullscreenElement !== elements.stage) return;
  const operation = ++fullscreenOperation;
  const frameWidth = Math.min(elements.stage.clientWidth, elements.stage.clientHeight * (16 / 9));
  const frameHeight = frameWidth * (9 / 16);
  const slideScale = Math.min(frameWidth / viewer.slideWidth, frameHeight / viewer.slideHeight);
  await viewer.setFitMode('none');
  if (operation !== fullscreenOperation || document.fullscreenElement !== elements.stage) return;
  await viewer.setZoom(slideScale * 100);
  if (operation !== fullscreenOperation || document.fullscreenElement !== elements.stage) return;
  syncFullscreenSlide();
};

document.addEventListener('fullscreenchange', () => {
  if (!viewer) return;
  fullscreenOperation += 1;
  bindViewerEvents();
  if (document.fullscreenElement === elements.stage) {
    fullscreenZoom = viewer.zoomPercent;
    fullscreenFitMode = viewer.fitMode;
    requestAnimationFrame(() => {
      syncFullscreenSlide();
      void enterFullscreenFit();
    });
  } else if (fullscreenZoom !== undefined) {
    document.querySelectorAll('.pptx-fullscreen-slide').forEach((item) => item.classList.remove('pptx-fullscreen-slide'));
    const operation = fullscreenOperation;
    void viewer.setFitMode(fullscreenFitMode ?? 'contain').then(() => {
      if (operation === fullscreenOperation) return viewer.setZoom(fullscreenZoom);
      return undefined;
    });
    fullscreenZoom = undefined;
    fullscreenFitMode = undefined;
  }
});

const viewerObserver = new MutationObserver(() => {
  upgradeYoutubeVideos(elements.viewerContainer);
  bindViewerEvents();
  syncFullscreenSlide();
});
viewerObserver.observe(elements.viewerContainer, { childList: true, subtree: true, attributes: true, attributeFilter: ['aria-busy'] });
upgradeYoutubeVideos(elements.viewerContainer);
