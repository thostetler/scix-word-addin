/* global fetch, URLSearchParams, window */

export interface SearchResult {
  bibcode: string;
  title: string[];
  author: string[];
  year: string;
  pub: string;
}

export interface PaperDetail extends SearchResult {
  abstract?: string;
  citation_count?: number;
  doi?: string[];
  aff?: string[];
}

interface SearchResponse {
  response: {
    docs: SearchResult[];
    numFound: number;
  };
  nextCursorMark?: string;
}

export interface SearchResultWithCursor {
  docs: SearchResult[];
  nextCursorMark: string | null;
  numFound: number;
}

interface ExportResponse {
  export: string;
}

export type ExportFormat = string;

// Use local proxy in dev, direct API in production (via scixplorer proxy)
function getApiBaseUrl(): string {
  const hostname = window.location.hostname;
  if (hostname === "localhost" || hostname.endsWith(".ts.net")) {
    return "http://localhost:3002/v1/search/query";
  }
  // Production: proxy through scixplorer.org
  return "https://api.adsabs.harvard.edu/v1/search/query";
}

function getExportApiBaseUrl(): string {
  const hostname = window.location.hostname;
  if (hostname === "localhost" || hostname.endsWith(".ts.net")) {
    return "http://localhost:3002/v1/export";
  }
  return "https://api.adsabs.harvard.edu/v1/export";
}

const DEFAULT_FIELDS = "bibcode,title,author,year,pub";
const DETAIL_FIELDS = "bibcode,title,author,year,pub,abstract,citation_count,doi,aff";

export async function searchADS(
  query: string,
  token: string,
  rows = 10,
  cursorMark?: string
): Promise<SearchResultWithCursor> {
  const params = new URLSearchParams({
    q: query,
    fl: DEFAULT_FIELDS,
    rows: String(rows),
    sort: "score desc, id desc",
    cursorMark: cursorMark || "*",
  });

  const apiUrl = getApiBaseUrl();
  const response = await fetch(`${apiUrl}?${params}`, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!response.ok) {
    if (response.status === 401) {
      throw new Error("Invalid API token");
    }
    throw new Error(`ADS API error: ${response.status}`);
  }

  const data: SearchResponse = await response.json();
  const nextCursor = data.nextCursorMark;
  // If nextCursorMark equals what we sent, we've reached the end
  const hasMore = nextCursor && nextCursor !== cursorMark;

  return {
    docs: data.response.docs,
    nextCursorMark: hasMore ? nextCursor : null,
    numFound: data.response.numFound,
  };
}

export async function validateToken(token: string): Promise<boolean> {
  try {
    await searchADS("bibcode:2024ApJ", token, 1);
    return true;
  } catch {
    return false;
  }
}

export async function exportCitation(
  bibcode: string,
  format: ExportFormat,
  token: string
): Promise<string> {
  return exportCitations([bibcode], format, token);
}

export async function exportCitations(
  bibcodes: string[],
  format: ExportFormat,
  token: string
): Promise<string> {
  const apiUrl = `${getExportApiBaseUrl()}/${format}`;

  const response = await fetch(apiUrl, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ bibcode: bibcodes }),
  });

  if (!response.ok) {
    if (response.status === 401) {
      throw new Error("Invalid API token");
    }
    throw new Error(`Export API error: ${response.status}`);
  }

  const data: ExportResponse = await response.json();
  return data.export.trim();
}

export async function fetchPaperDetail(
  bibcode: string,
  token: string
): Promise<PaperDetail | null> {
  const params = new URLSearchParams({
    q: `bibcode:${bibcode}`,
    fl: DETAIL_FIELDS,
    rows: "1",
  });

  const apiUrl = getApiBaseUrl();
  const response = await fetch(`${apiUrl}?${params}`, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!response.ok) {
    if (response.status === 401) {
      throw new Error("Invalid API token");
    }
    throw new Error(`ADS API error: ${response.status}`);
  }

  interface DetailResponse {
    response: {
      docs: PaperDetail[];
      numFound: number;
    };
  }

  const data: DetailResponse = await response.json();
  return data.response.docs[0] || null;
}

export async function fetchReferences(
  bibcode: string,
  token: string,
  rows = 25
): Promise<SearchResult[]> {
  const params = new URLSearchParams({
    q: `references(bibcode:${bibcode})`,
    fl: DEFAULT_FIELDS,
    rows: String(rows),
  });

  const apiUrl = getApiBaseUrl();
  const response = await fetch(`${apiUrl}?${params}`, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!response.ok) {
    if (response.status === 401) {
      throw new Error("Invalid API token");
    }
    throw new Error(`ADS API error: ${response.status}`);
  }

  const data: SearchResponse = await response.json();
  return data.response.docs;
}

export interface ExportFormatInfo {
  name: string;
  type: string;
  route: string;
  extension: string;
}

export async function fetchExportManifest(token: string): Promise<ExportFormatInfo[]> {
  const apiUrl = `${getExportApiBaseUrl()}/manifest`;

  const response = await fetch(apiUrl, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!response.ok) {
    throw new Error(`Manifest API error: ${response.status}`);
  }

  return response.json();
}
