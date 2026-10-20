import test from "node:test";
import assert from "node:assert/strict";

import {
  buildScalableRouteOrder,
  buildGoogleMapsUrl,
  buildOsrmRouteUrl,
  buildOsrmTableUrl,
  createEffectiveMatrix,
  createRoutePlan,
  findOptimalOrder,
} from "../dist/route-planner.js";

test("buildScalableRouteOrder visits 25 destinations once without returning", () => {
  const size = 26;
  const matrix = Array.from({ length: size }, (_, from) =>
    Array.from({ length: size }, (_, to) => Math.abs(from - to)),
  );

  const order = buildScalableRouteOrder(matrix);

  assert.equal(order.length, size);
  assert.equal(order[0], 0);
  assert.equal(new Set(order).size, size);
  assert.deepEqual(
    [...order].sort((a, b) => a - b),
    [...Array(size).keys()],
  );
  assert.notEqual(order.at(-1), 0);
});

test("findOptimalOrder chooses the shortest open path from index 0", () => {
  const matrix = [
    [0, 2, 9, 9],
    [2, 0, 2, 9],
    [9, 2, 0, 2],
    [9, 9, 2, 0],
  ];

  assert.deepEqual(findOptimalOrder(matrix), [0, 1, 2, 3]);
});

test("createRoutePlan returns an open ordered path and fallback state", () => {
  const points = [
    { id: "origin", lat: 25, lng: 121 },
    { id: "a", lat: 25.01, lng: 121.01 },
    { id: "b", lat: 25.02, lng: 121.02 },
  ];
  const matrix = [
    [0, 1, 2],
    [1, 0, 1],
    [2, 1, 0],
  ];

  const plan = createRoutePlan(points, matrix);

  assert.deepEqual(plan.order, [0, 1, 2]);
  assert.deepEqual(
    plan.orderedPoints.map(({ id }) => id),
    ["origin", "a", "b"],
  );
  assert.equal(plan.usedFallback, false);
  assert.notEqual(plan.order.at(-1), 0);
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
