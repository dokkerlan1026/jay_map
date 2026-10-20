import { stores } from "./stores.js";

const listElement = document.querySelector("#store-list");
const statusElement = document.querySelector("#route-status");
const markers = new Map();
const rows = new Map();

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
      <button class="navigate-button" type="button" disabled aria-label="導航到 ${store.name}">
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
