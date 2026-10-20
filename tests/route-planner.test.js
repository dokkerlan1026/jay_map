import test from "node:test";
import assert from "node:assert/strict";

import {
  buildGoogleMapsUrl,
  buildOsrmRouteUrl,
  buildOsrmTableUrl,
  createEffectiveMatrix,
  findOptimalOrder,
} from "../dist/route-planner.js";

test("findOptimalOrder chooses the shortest open path from index 0", () => {
  const matrix = [
    [0, 2, 9, 9],
    [2, 0, 2, 9],
    [9, 2, 0, 2],
    [9, 9, 2, 0],
  ];

  assert.deepEqual(findOptimalOrder(matrix), [0, 1, 2, 3]);
});

test("createEffectiveMatrix replaces missing road distances", () => {
  const points = [
    { lat: 25, lng: 121 },
    { lat: 25.01, lng: 121.01 },
  ];

  const result = createEffectiveMatrix([[0, null], [null, 0]], points);

  assert.equal(result.usedFallback, true);
  assert.ok(result.matrix[0][1] > 0);
});

test("URL builders preserve longitude-latitude order", () => {
  const points = [
    { lat: 25.1, lng: 121.5 },
    { lat: 25.2, lng: 121.6 },
  ];

  assert.match(buildOsrmTableUrl(points), /121\.5,25\.1;121\.6,25\.2/);
  assert.match(buildOsrmRouteUrl(points), /overview=full/);
  assert.match(
    buildGoogleMapsUrl(points[1], points[0]),
    /destination=25\.2%2C121\.6/,
  );
});
