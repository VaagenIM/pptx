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

const presentationMapUrl = 'https://raw.githubusercontent.com/VaagenIM/pptx/public/map.json';
const publicPresentationsUrl = 'https://raw.githubusercontent.com/VaagenIM/pptx/public/';
let resolvedPresentationUrl = null;
let resolvedPresentationName = null;
let presentationResolution = null;

const getPresentationId = () => new URLSearchParams(window.location.search).get('id');
const getRemotePresentationUrl = () => new URLSearchParams(window.location.search).get('url') || resolvedPresentationUrl;
const originalReplaceState = history.replaceState.bind(history);
history.replaceState = (state, title, url) => {
  if (url != null && getPresentationId()) {
    const nextUrl = new URL(url, window.location.href);
    if (nextUrl.searchParams.has('url')) {
      nextUrl.searchParams.delete('url');
      nextUrl.searchParams.set('id', getPresentationId());
      url = nextUrl.href;
    }
  }
  return originalReplaceState(state, title, url);
};
const getDownloadName = (url) => {
  try {
    return decodeURIComponent(new URL(url).pathname.split('/').pop() || 'presentation.pptx');
  } catch {
    return 'presentation.pptx';
  }
};
const resolvePresentationForDownload = async () => {
  const directUrl = new URLSearchParams(window.location.search).get('url');
  if (directUrl) return { url: directUrl, name: getDownloadName(directUrl) };
  if (!getPresentationId()) return null;
  presentationResolution ??= fetch(presentationMapUrl, { credentials: 'omit' })
    .then(async (response) => {
      if (!response.ok) throw new Error(`Presentation map returned HTTP ${response.status}.`);
      const map = await response.json();
      const path = map[getPresentationId()];
      if (typeof path !== 'string' || !path.startsWith('powerpoints/') || !path.endsWith('.pptx')) {
        throw new Error('The presentation ID was not found.');
      }
      const url = new URL(`powerpoints/${getPresentationId()}.pptx`, publicPresentationsUrl).href;
      return { url, name: path.split('/').pop() || 'presentation.pptx' };
    });
  const result = await presentationResolution;
  resolvedPresentationUrl = result.url;
  resolvedPresentationName = result.name;
  syncPresentationDisplayName();
  syncDownloadButton();
  return result;
};
const openMappedPresentation = async () => {
  if (!getPresentationId() || new URLSearchParams(window.location.search).get('url')) return;
  const presentation = await resolvePresentationForDownload();
  elements.urlInput.value = presentation.url;
  elements.urlForm.requestSubmit();
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
  downloadButton.disabled = elements.app.classList.contains('is-empty')
    || (!localPresentationFile && (!getRemotePresentationUrl() && !getPresentationId()));
};
const syncPresentationDisplayName = () => {
  if (!resolvedPresentationName || !getPresentationId()) return;
  const publicName = `${getPresentationId()}.pptx`;
  if (elements.documentTitle.textContent === publicName) {
    elements.documentTitle.textContent = resolvedPresentationName;
    elements.documentTitle.title = resolvedPresentationName;
  }
  if (elements.status.textContent.includes(publicName)) {
    elements.status.textContent = elements.status.textContent.replace(publicName, resolvedPresentationName);
  }
};

downloadButton.addEventListener('click', async () => {
  let objectUrl = null;
  let remoteUrl = null;
  let link = null;
  try {
    const resolved = await resolvePresentationForDownload();
    remoteUrl = resolved?.url || getRemotePresentationUrl();
    const href = localPresentationFile ? URL.createObjectURL(localPresentationFile) : remoteUrl;
    if (!href) return;
    link = document.createElement('a');
    objectUrl = localPresentationFile ? href : null;
    link.href = href;
    link.download = localPresentationFile?.name || resolved?.name || resolvedPresentationName || getDownloadName(remoteUrl);
    link.rel = 'noopener';
    if (remoteUrl && !localPresentationFile) {
      const response = await fetch(remoteUrl, { credentials: 'omit' });
      if (!response.ok) throw new Error(`Download returned HTTP ${response.status}.`);
      objectUrl = URL.createObjectURL(await response.blob());
      link.href = objectUrl;
    }
    link.click();
  } catch (error) {
    if (link && remoteUrl) {
      link.href = remoteUrl;
      link.removeAttribute('download');
      link.click();
    }
    elements.status.textContent = `Direct download unavailable; opening the source URL. ${error.message}`;
  } finally {
    if (objectUrl) window.setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
  }
});

const downloadObserver = new MutationObserver(syncDownloadButton);
downloadObserver.observe(elements.app, { attributes: true, attributeFilter: ['class'] });
const presentationNameObserver = new MutationObserver(syncPresentationDisplayName);
presentationNameObserver.observe(elements.documentTitle, { childList: true, characterData: true, subtree: true });
presentationNameObserver.observe(elements.status, { childList: true, characterData: true, subtree: true });
syncDownloadButton();
openMappedPresentation().catch((error) => {
  elements.status.textContent = `Could not resolve presentation ID. ${error.message}`;
});

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
