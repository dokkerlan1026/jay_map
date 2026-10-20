import {
  buildGoogleMapsUrl,
  buildOsrmRouteUrl,
  buildOsrmTableUrl,
  createRoutePlan,
} from "./route-planner.js";
import { stores, TAIPEI_MAIN_STATION } from "./stores.js";

const listElement = document.querySelector("#store-list");
const statusElement = document.querySelector("#route-status");
const planButton = document.querySelector("#plan-route");
const routeSummaryElement = document.querySelector("#route-summary");
const markers = new Map();
const rows = new Map();
let currentOrigin = null;
let routeLayer = null;
let originMarker = null;

if (!window.L) {
  statusElement.textContent = "地圖暫時無法載入，店家地址仍可正常查看。";
  statusElement.dataset.tone = "warning";
}

const map = window.L
  ? window.L.map("map", {
      zoomControl: true,
      scrollWheelZoom: false,
    }).setView([25.0445, 121.535], 13)
  : null;

if (map) {
  window.L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 19,
    attribution:
      '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>',
  }).addTo(map);
}

function popupContent(store) {
  return `
    <strong>${store.name}</strong>
    <span class="popup-address">${store.address}</span>
  `;
}

function selectStore(storeId, { focusMap = true } = {}) {
  for (const [id, row] of rows) {
    const isSelected = id === storeId;
    row.classList.toggle("is-selected", isSelected);
    row.setAttribute("aria-current", isSelected ? "true" : "false");
    row.querySelector(".selection-copy").hidden = !isSelected;
  }

  const store = stores.find(({ id }) => id === storeId);
  const marker = markers.get(storeId);
  if (!store) return;

  statusElement.textContent = `已選擇 ${store.name}。`;
  statusElement.dataset.tone = "neutral";

  if (map && marker) {
    if (focusMap) {
      map.flyTo([store.lat, store.lng], 15, { duration: 0.7 });
    }
    marker.openPopup();
  }
}

function createStoreRow(store, index) {
  const row = document.createElement("tr");
  row.className = "store-row";
  row.tabIndex = 0;
  row.dataset.storeId = store.id;
  row.setAttribute("aria-current", "false");
  row.setAttribute("aria-label", `${store.name}，${store.address}`);
  row.innerHTML = `
    <td>
      <span class="store-name-wrap">
        <span class="store-index" aria-hidden="true">${String(index + 1).padStart(2, "0")}</span>
        <span>
          <strong class="store-name">${store.name}</strong>
          <span class="selection-copy" hidden>目前選取</span>
        </span>
      </span>
    </td>
    <td>
      <span class="store-address">${store.address}</span>
      <button class="navigate-button" type="button" aria-label="導航到 ${store.name}">
        導航
        <span aria-hidden="true">↗</span>
      </button>
    </td>
  `;

  row.addEventListener("click", () => selectStore(store.id));
  row.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      selectStore(store.id);
    }
  });
  row.querySelector(".navigate-button").addEventListener("click", (event) => {
    event.stopPropagation();
    const origin = currentOrigin?.isFallback ? null : currentOrigin;
    window.open(buildGoogleMapsUrl(store, origin), "_blank", "noopener,noreferrer");
  });

  rows.set(store.id, row);
  return row;
}

for (const [index, store] of stores.entries()) {
  listElement.append(createStoreRow(store, index));

  if (map) {
    const marker = window.L.marker([store.lat, store.lng])
      .addTo(map)
      .bindPopup(popupContent(store));
    marker.on("click", () => selectStore(store.id, { focusMap: false }));
    markers.set(store.id, marker);
  }
}

if (map && markers.size) {
  const bounds = window.L.featureGroup([...markers.values()]).getBounds();
  map.fitBounds(bounds.pad(0.16));
}

function setStatus(message, tone = "neutral") {
  statusElement.textContent = message;
  statusElement.dataset.tone = tone;
}

function getOrigin() {
  if (currentOrigin) {
    return Promise.resolve(currentOrigin);
  }

  if (!navigator.geolocation) {
    currentOrigin = { ...TAIPEI_MAIN_STATION, isFallback: true };
    return Promise.resolve(currentOrigin);
  }

  return new Promise((resolve) => {
    navigator.geolocation.getCurrentPosition(
      ({ coords }) => {
        currentOrigin = {
          name: "目前位置",
          lat: coords.latitude,
          lng: coords.longitude,
          isFallback: false,
        };
        resolve(currentOrigin);
      },
      () => {
        currentOrigin = { ...TAIPEI_MAIN_STATION, isFallback: true };
        resolve(currentOrigin);
      },
      { enableHighAccuracy: false, timeout: 8_000, maximumAge: 60_000 },
    );
  });
}

async function fetchJson(url) {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Request failed with ${response.status}`);
  }

  const payload = await response.json();
  if (payload.code !== "Ok") {
    throw new Error(payload.message || "Route service returned an error");
  }
  return payload;
}

function clearRoute() {
  if (!map) return;
  if (routeLayer) map.removeLayer(routeLayer);
  if (originMarker) map.removeLayer(originMarker);
  routeLayer = null;
  originMarker = null;
}

function drawOrigin(origin) {
  if (!map) return;
  originMarker = window.L.circleMarker([origin.lat, origin.lng], {
    radius: 9,
    weight: 4,
    color: "#ffffff",
    fillColor: "#17352d",
    fillOpacity: 1,
  })
    .addTo(map)
    .bindPopup(`<strong>${origin.name}</strong><span class="popup-address">路線起點</span>`);
}

function drawFallbackRoute(orderedPoints) {
  if (!map) return;
  routeLayer = window.L.polyline(
    orderedPoints.map(({ lat, lng }) => [lat, lng]),
    { color: "#e6613e", weight: 5, dashArray: "10 10", opacity: 0.86 },
  ).addTo(map);
}

function drawRoadRoute(geometry) {
  if (!map) return;
  routeLayer = window.L.geoJSON(geometry, {
    style: { color: "#e6613e", weight: 6, opacity: 0.9 },
  }).addTo(map);
}

function updateRouteOrder(origin, orderedStores) {
  routeSummaryElement.replaceChildren();

  const startItem = document.createElement("li");
  startItem.innerHTML = `<span class="route-stop">起點</span><strong>${origin.name}</strong>`;
  routeSummaryElement.append(startItem);

  for (const [index, store] of orderedStores.entries()) {
    const row = rows.get(store.id);
    row.querySelector(".store-index").textContent = String(index + 1).padStart(2, "0");
    listElement.append(row);

    const item = document.createElement("li");
    const badge = document.createElement("span");
    const name = document.createElement("strong");
    badge.className = "route-stop";
    badge.textContent = String(index + 1);
    name.textContent = store.name;
    item.append(badge, name);
    routeSummaryElement.append(item);
  }
}

function fitRoute(orderedPoints) {
  if (!map) return;
  const bounds = window.L.latLngBounds(
    orderedPoints.map(({ lat, lng }) => [lat, lng]),
  );
  map.fitBounds(bounds.pad(0.14));
}

async function planBestRoute() {
  planButton.disabled = true;
  planButton.classList.add("is-loading");
  setStatus("正在取得起點並計算最佳順序…");

  try {
    const origin = await getOrigin();
    const allPoints = [origin, ...stores];
    if (origin.isFallback) {
      setStatus("無法取得目前位置，改由台北車站出發規劃。", "warning");
    } else {
      setStatus("已取得目前位置，正在比較道路距離…");
    }

    const tablePayload = await fetchJson(buildOsrmTableUrl(allPoints));
    const plan = createRoutePlan(allPoints, tablePayload.distances);
    const orderedStores = plan.order.slice(1).map((index) => stores[index - 1]);

    clearRoute();
    drawOrigin(origin);

    let usedEstimatedGeometry = false;
    let roadDistance = null;
    try {
      const routePayload = await fetchJson(buildOsrmRouteUrl(plan.orderedPoints));
      const route = routePayload.routes?.[0];
      if (!route?.geometry) throw new Error("Missing route geometry");
      drawRoadRoute(route.geometry);
      roadDistance = route.distance;
    } catch {
      usedEstimatedGeometry = true;
      drawFallbackRoute(plan.orderedPoints);
    }

    updateRouteOrder(origin, orderedStores);
    fitRoute(plan.orderedPoints);

    const distanceText = Number.isFinite(roadDistance)
      ? `，全程約 ${(roadDistance / 1000).toFixed(1)} 公里`
      : "";
    if (usedEstimatedGeometry || plan.usedFallback) {
      setStatus(`已排出建議順序${distanceText}；部分路段以直線距離估算。`, "warning");
    } else {
      setStatus(`最佳路線已完成${distanceText}。`, "success");
    }
  } catch (error) {
    console.error(error);
    setStatus("目前無法取得道路路線，仍可使用各店導航。", "warning");
  } finally {
    planButton.disabled = false;
    planButton.classList.remove("is-loading");
  }
}

planButton.addEventListener("click", planBestRoute);
