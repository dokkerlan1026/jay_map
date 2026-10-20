const OSRM_BASE = "https://router.project-osrm.org";

export function distanceMeters(a, b) {
  const toRadians = (value) => (value * Math.PI) / 180;
  const deltaLat = toRadians(b.lat - a.lat);
  const deltaLng = toRadians(b.lng - a.lng);
  const lat1 = toRadians(a.lat);
  const lat2 = toRadians(b.lat);
  const haversine =
    Math.sin(deltaLat / 2) ** 2 +
    Math.cos(lat1) * Math.cos(lat2) * Math.sin(deltaLng / 2) ** 2;

  return (
    6_371_000 *
    2 *
    Math.atan2(Math.sqrt(haversine), Math.sqrt(1 - haversine))
  );
}

export function createEffectiveMatrix(matrix, points) {
  let usedFallback = false;
  const effectiveMatrix = points.map((from, fromIndex) =>
    points.map((to, toIndex) => {
      const roadDistance = matrix?.[fromIndex]?.[toIndex];
      if (Number.isFinite(roadDistance)) {
        return roadDistance;
      }

      usedFallback = true;
      return distanceMeters(from, to);
    }),
  );

  return { matrix: effectiveMatrix, usedFallback };
}

export function findOptimalOrder(matrix) {
  if (matrix.length <= 1) {
    return matrix.length ? [0] : [];
  }

  const destinations = Array.from(
    { length: matrix.length - 1 },
    (_, index) => index + 1,
  );
  let bestOrder = null;
  let bestDistance = Number.POSITIVE_INFINITY;

  function visit(prefix, remaining, totalDistance) {
    if (totalDistance >= bestDistance) {
      return;
    }

    if (remaining.length === 0) {
      bestOrder = [0, ...prefix];
      bestDistance = totalDistance;
      return;
    }

    const currentIndex = prefix.length ? prefix.at(-1) : 0;
    for (const nextIndex of remaining) {
      visit(
        [...prefix, nextIndex],
        remaining.filter((value) => value !== nextIndex),
        totalDistance + matrix[currentIndex][nextIndex],
      );
    }
  }

  visit([], destinations, 0);
  return bestOrder;
}

export function createRoutePlan(points, roadMatrix) {
  const { matrix, usedFallback } = createEffectiveMatrix(roadMatrix, points);
  const order = findOptimalOrder(matrix);

  return {
    order,
    orderedPoints: order.map((index) => points[index]),
    usedFallback,
  };
}

function coordinates(points) {
  return points.map(({ lng, lat }) => `${lng},${lat}`).join(";");
}

export function buildOsrmTableUrl(points) {
  return `${OSRM_BASE}/table/v1/driving/${coordinates(points)}?annotations=distance`;
}

export function buildOsrmRouteUrl(points) {
  return `${OSRM_BASE}/route/v1/driving/${coordinates(points)}?overview=full&geometries=geojson&steps=false`;
}

export function buildGoogleMapsUrl(destination, origin) {
  const params = new URLSearchParams({
    api: "1",
    destination: `${destination.lat},${destination.lng}`,
    travelmode: "driving",
  });

  if (origin) {
    params.set("origin", `${origin.lat},${origin.lng}`);
  }

  return `https://www.google.com/maps/dir/?${params}`;
}
