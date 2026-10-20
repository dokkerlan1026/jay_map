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

  if (matrix.length > 10) {
    return buildScalableRouteOrder(matrix);
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

function routeDistance(order, matrix) {
  let total = 0;
  for (let index = 1; index < order.length; index += 1) {
    total += matrix[order[index - 1]][order[index]];
  }
  return total;
}

export function buildScalableRouteOrder(matrix) {
  if (matrix.length <= 1) {
    return matrix.length ? [0] : [];
  }

  const unvisited = new Set(
    Array.from({ length: matrix.length - 1 }, (_, index) => index + 1),
  );
  let current = 0;
  let order = [0];

  while (unvisited.size) {
    let nearest = null;
    let nearestDistance = Number.POSITIVE_INFINITY;
    for (const candidate of unvisited) {
      if (matrix[current][candidate] < nearestDistance) {
        nearest = candidate;
        nearestDistance = matrix[current][candidate];
      }
    }
    order.push(nearest);
    unvisited.delete(nearest);
    current = nearest;
  }

  let bestDistance = routeDistance(order, matrix);
  for (let pass = 0; pass < 100; pass += 1) {
    let improvedOrder = null;
    let improvedDistance = bestDistance;

    for (let start = 1; start < order.length - 1; start += 1) {
      for (let end = start + 1; end < order.length; end += 1) {
        const candidate = [
          ...order.slice(0, start),
          ...order.slice(start, end + 1).reverse(),
          ...order.slice(end + 1),
        ];
        const candidateDistance = routeDistance(candidate, matrix);
        if (candidateDistance + 0.001 < improvedDistance) {
          improvedOrder = candidate;
          improvedDistance = candidateDistance;
        }
      }
    }

    if (!improvedOrder) break;
    order = improvedOrder;
    bestDistance = improvedDistance;
  }

  return order;
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
