"use strict";

const CACHE_NAME = "family-chores-shell-v3";
const APP_URL = new URL("./", self.location.href);
const INDEX_URL = new URL("./index.html", self.location.href);
const APP_SHELL = [
  APP_URL.href,
  INDEX_URL.href,
  new URL("./manifest.webmanifest", self.location.href).href,
  new URL("./icon.svg", self.location.href).href
];

self.addEventListener("install", event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(APP_SHELL))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys.filter(key => key.startsWith("family-chores-shell-") && key !== CACHE_NAME)
          .map(key => caches.delete(key))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", event => {
  if (event.request.method !== "GET") return;

  const requestUrl = new URL(event.request.url);
  if (requestUrl.origin !== self.location.origin) return;

  if (event.request.mode === "navigate") {
    event.respondWith(
      fetch(event.request)
        .then(response => {
          if (response.ok) {
            const copy = response.clone();
            event.waitUntil(
              caches.open(CACHE_NAME).then(cache => cache.put(INDEX_URL.href, copy))
            );
          }
          return response;
        })
        .catch(() => caches.match(INDEX_URL.href))
    );
    return;
  }

  event.respondWith(
    caches.match(event.request)
      .then(cached => cached || fetch(event.request))
  );
});
