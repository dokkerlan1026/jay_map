import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

import { stores } from "../dist/stores.js";

test("page exposes the required interactive regions", async () => {
  const html = await readFile(
    new URL("../dist/index.html", import.meta.url),
    "utf8",
  );

  for (const id of [
    "store-list",
    "map",
    "plan-route",
    "route-status",
    "route-summary",
  ]) {
    assert.match(html, new RegExp(`id=["']${id}["']`));
  }
  assert.match(html, /aria-live=["']polite["']/);
  assert.match(html, /leaflet\.css/);
  assert.match(html, /leaflet\.js/);
  assert.match(html, /src=["']\.\/app\.js["']/);
});

test("five Taipei sample stores have coordinates and addresses", () => {
  assert.equal(stores.length, 5);
  for (const store of stores) {
    assert.ok(store.name && store.address);
    assert.ok(Number.isFinite(store.lat) && Number.isFinite(store.lng));
  }
});
