// Shared helpers used by page-specific scripts.

const SEASONS = { 1: "Spring", 2: "Summer", 3: "Fall", 4: "Winter" };
const WEATHER = { 1: "Clear", 2: "Mist / Cloudy", 3: "Light Rain / Snow", 4: "Severe Weather" };

async function apiGet(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error("Request failed: " + res.status);
  return res.json();
}

async function apiSend(url, method, body) {
  const res = await fetch(url, {
    method,
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const err = new Error(data.error || (data.errors || []).join(", ") || "Request failed");
    err.details = data;
    throw err;
  }
  return data;
}

function fmtNumber(n) {
  return Number(n).toLocaleString();
}
