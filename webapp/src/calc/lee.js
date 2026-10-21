// Spreading resistance of a square source on a square plate cooled on its far face.
// Lee, Song, Au, Moran, "Constriction/spreading resistance model for electronics packaging" (1995).
// Square areas are turned into circles of the same area.
//
//   sourceMm2  heat source footprint (mm2)
//   plateMm2   plate area the source owns (mm2)
//   tMm        plate thickness (mm)
//   k          plate conductivity in plane (W/(m K))
//   h          h on the far face (W/(m2 K)); for a layer stack, the next layer's k / t
//
// Returns the terms and R (C/W). All inputs must be > 0.

export function leeSpreading({ sourceMm2, plateMm2, tMm, k, h }) {
  const a = Math.sqrt(sourceMm2 / Math.PI);           // source radius (mm)
  const b = Math.sqrt(plateMm2 / Math.PI);            // plate radius (mm)
  const eps = Math.min(1, Math.sqrt(sourceMm2 / plateMm2));
  const tau = tMm / b;
  const bi = (h * b) / 1000 / k;
  const lambda = Math.PI + 1 / (eps * Math.sqrt(Math.PI));
  const th = Math.tanh(lambda * tau);
  const phi = (th + lambda / bi) / (1 + (lambda / bi) * th);
  const psi = (eps * tau) / Math.sqrt(Math.PI) + ((1 - eps) * phi) / Math.sqrt(Math.PI);
  const r = psi / ((Math.sqrt(Math.PI) * k * a) / 1000);
  return { eps, tau, bi, lambda, phi, r };
}
