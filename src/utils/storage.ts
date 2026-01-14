/* global localStorage */

const TOKEN_KEY = "scix_ads_api_token";
const BIBLIOGRAPHY_KEY = "scix_bibliography";

export function getToken(): string | null {
  return localStorage.getItem(TOKEN_KEY);
}

export function setToken(token: string): void {
  localStorage.setItem(TOKEN_KEY, token);
}

export function clearToken(): void {
  localStorage.removeItem(TOKEN_KEY);
}

export function hasToken(): boolean {
  const token = getToken();
  return token !== null && token.length > 0;
}

// Bibliography storage

export interface BibliographyEntry {
  bibcode: string;
  title: string;
  authors: string;
  year: string;
  addedAt: number;
}

export function getBibliography(): BibliographyEntry[] {
  const data = localStorage.getItem(BIBLIOGRAPHY_KEY);
  if (!data) return [];
  try {
    return JSON.parse(data);
  } catch {
    return [];
  }
}

export function addToBibliography(entry: BibliographyEntry): boolean {
  const bib = getBibliography();
  if (bib.some((e) => e.bibcode === entry.bibcode)) {
    return false; // Already exists
  }
  bib.push(entry);
  localStorage.setItem(BIBLIOGRAPHY_KEY, JSON.stringify(bib));
  return true;
}

export function removeFromBibliography(bibcode: string): void {
  const bib = getBibliography().filter((e) => e.bibcode !== bibcode);
  localStorage.setItem(BIBLIOGRAPHY_KEY, JSON.stringify(bib));
}

export function clearBibliography(): void {
  localStorage.removeItem(BIBLIOGRAPHY_KEY);
}

export function getBibliographyCount(): number {
  return getBibliography().length;
}

export function isInBibliography(bibcode: string): boolean {
  return getBibliography().some((e) => e.bibcode === bibcode);
}
