export function slugifyCarrier(name) {
  return (name || "")
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

export function findCarrierBySlug(carriers, slug) {
  const target = (slug || "").toUpperCase();
  return (carriers || []).find((c) => slugifyCarrier(c) === target) || null;
}

export function buildPartnerPath(carrierName) {
  return `/parceiro/${slugifyCarrier(carrierName)}`;
}
