"use strict";

const CACHE_NAME = "schedule-images-v1";
const IMAGE_DIRECTORY = new URL("images/", self.registration.scope).pathname;
const MAX_PARALLEL_DOWNLOADS = 3;

self.addEventListener("install", (event) => {
  event.waitUntil(self.skipWaiting());
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    (async () => {
      const cacheNames = await caches.keys();
      await Promise.all(
        cacheNames
          .filter(
            (name) =>
              name.startsWith("schedule-images-") && name !== CACHE_NAME
          )
          .map((name) => caches.delete(name))
      );
      await self.clients.claim();
    })()
  );
});

function isScheduleImage(requestUrl) {
  return (
    requestUrl.origin === self.location.origin &&
    requestUrl.pathname.startsWith(IMAGE_DIRECTORY) &&
    requestUrl.pathname.endsWith(".png")
  );
}

async function saveImage(cache, request, response) {
  if (!response.ok) return false;
  try {
    await cache.put(request, response.clone());
    return true;
  } catch (error) {
    console.warn("편성표 이미지를 브라우저 저장소에 보관하지 못했습니다.", error);
    return false;
  }
}

self.addEventListener("fetch", (event) => {
  const requestUrl = new URL(event.request.url);
  if (event.request.method !== "GET" || !isScheduleImage(requestUrl)) return;

  event.respondWith(
    (async () => {
      try {
        const cache = await caches.open(CACHE_NAME);
        const cached = await cache.match(event.request);
        if (cached) return cached;

        const response = await fetch(event.request);
        await saveImage(cache, event.request, response);
        return response;
      } catch (error) {
        console.warn("편성표 이미지 캐시를 사용할 수 없어 네트워크에서 불러옵니다.", error);
        return fetch(event.request);
      }
    })()
  );
});

async function cacheYearImages(urls, year) {
  const cache = await caches.open(CACHE_NAME);
  const desiredUrls = new Set(
    urls.map((url) => new URL(url, self.registration.scope).href)
  );
  const cachedRequests = await cache.keys();
  await Promise.all(
    cachedRequests.map((request) => {
      const requestUrl = new URL(request.url);
      const match = requestUrl.pathname.match(
        /(?:dept|team)-(\d{4})-\d{2}-\d+(?:--(?:v1|v2|confirmed|final))?\.png$/
      );
      if (match && Number(match[1]) === year && !desiredUrls.has(request.url)) {
        return cache.delete(request);
      }
      return Promise.resolve(false);
    })
  );

  let nextIndex = 0;
  let cachedCount = 0;
  let failedCount = 0;
  let storageFailed = false;
  let networkFailed = false;

  async function fillCache() {
    while (nextIndex < urls.length) {
      if (storageFailed || networkFailed) return;
      const url = urls[nextIndex];
      nextIndex += 1;

      try {
        const request = new Request(url, { credentials: "same-origin" });
        if (await cache.match(request)) {
          cachedCount += 1;
          continue;
        }

        const response = await fetch(request);
        if (response.ok) {
          if (await saveImage(cache, request, response)) {
            cachedCount += 1;
          } else {
            failedCount += 1;
            storageFailed = true;
          }
        } else {
          failedCount += 1;
        }
      } catch (error) {
        failedCount += 1;
        networkFailed = true;
        console.warn("편성표 이미지 사전 저장에 실패했습니다.", url, error);
      }
    }
  }

  await Promise.all(
    Array.from(
      { length: Math.min(MAX_PARALLEL_DOWNLOADS, urls.length) },
      fillCache
    )
  );

  const clients = await self.clients.matchAll({
    type: "window",
    includeUncontrolled: true,
  });
  clients.forEach((client) => {
    client.postMessage({
      type: "IMAGE_CACHE_COMPLETE",
      year,
      cachedCount,
      failedCount,
    });
  });
}

self.addEventListener("message", (event) => {
  if (event.data?.type !== "PRECACHE_YEAR_IMAGES") return;

  const urls = Array.isArray(event.data.urls)
    ? event.data.urls.filter((url) => {
        try {
          return isScheduleImage(new URL(url, self.registration.scope));
        } catch (_error) {
          return false;
        }
      })
    : [];
  const year = Number(event.data.year);
  if (!Number.isInteger(year)) return;

  event.waitUntil(
    cacheYearImages(urls, year).catch((error) => {
      console.warn("연도별 편성표 이미지 캐시를 완료하지 못했습니다.", error);
    })
  );
});
