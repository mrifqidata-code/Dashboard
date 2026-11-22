export const EXBOARD_SHEET_URL = "https://docs.google.com/spreadsheets/d/1CLm1NYROYjPEGisRGw-A2ZMEk7OP6d3pDJywxTBJDXA/export?format=csv&gid=70575035";
export async function fetchExboardData() { const res = await fetch(EXBOARD_SHEET_URL, { cache: "no-store" }); if (!res.ok) throw new Error("Failed to fetch Exboard data"); const csv = await res.text();

const lines = csv.trim().split(/\r?\n/); const headers = lines[0]?.split(",") ?? []; const data = lines.slice(1).map((line) => { const cols = line.split(","); const obj = {}; headers.forEach((h, i) => (obj[h.trim()] = cols[i]?.trim())); return obj; });

return data; }
