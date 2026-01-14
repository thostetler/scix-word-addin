/* global document, Office, Word, HTMLInputElement, HTMLButtonElement, HTMLElement, HTMLSelectElement, KeyboardEvent, setTimeout, NodeListOf, Element, confirm */

import {
  searchADS,
  validateToken,
  SearchResult,
  PaperDetail,
  exportCitation,
  exportCitations,
  fetchPaperDetail,
  fetchReferences,
  fetchExportManifest,
  ExportFormatInfo,
} from "../api/ads";
import { formatInlineCitation, formatResultDisplay } from "../utils/citation";
import {
  getToken,
  setToken,
  clearToken,
  hasToken,
  BibliographyEntry,
  getBibliography,
  addToBibliography,
  removeFromBibliography,
  clearBibliography,
  isInBibliography,
} from "../utils/storage";

type CitationFormat = "inline" | "full" | "bibtex";

let selectedFormat: CitationFormat = "inline";
let expandedBibcode: string | null = null;
const detailCache: Map<string, PaperDetail> = new Map();

// Pagination state
let currentQuery: string | null = null;
let nextCursorMark: string | null = null;

Office.onReady(() => {
  initUI();
  bindEventHandlers();
  updateBibliographyUI();
  loadExportFormats();
});

function initUI(): void {
  updateAppState();
}

function updateAppState(): void {
  const settingsPanel = document.getElementById("settings-panel");
  const tokenStatus = document.getElementById("token-status");
  const app = document.getElementById("app");

  if (hasToken()) {
    app?.classList.remove("no-token");
    if (settingsPanel) settingsPanel.classList.add("collapsed");
    if (tokenStatus) {
      tokenStatus.textContent = "Token configured";
      tokenStatus.className = "status-badge success";
    }
  } else {
    app?.classList.add("no-token");
    if (settingsPanel) settingsPanel.classList.remove("collapsed");
    if (tokenStatus) {
      tokenStatus.textContent = "No token configured";
      tokenStatus.className = "status-badge warning";
    }
  }
}

function bindEventHandlers(): void {
  // Settings panel toggle
  const settingsBtn = document.getElementById("settings-btn");
  if (settingsBtn) settingsBtn.onclick = toggleSettingsPanel;

  // Token handlers
  const saveTokenBtn = document.getElementById("save-token-btn");
  const clearTokenBtn = document.getElementById("clear-token-btn");
  const toggleVisibilityBtn = document.getElementById("toggle-token-visibility");

  if (saveTokenBtn) saveTokenBtn.onclick = handleSaveToken;
  if (clearTokenBtn) clearTokenBtn.onclick = handleClearToken;
  if (toggleVisibilityBtn) toggleVisibilityBtn.onclick = toggleTokenVisibility;

  // Search handlers
  const searchBtn = document.getElementById("search-btn");
  const useSelectionBtn = document.getElementById("use-selection-btn");
  const searchInput = document.getElementById("search-input") as HTMLInputElement;

  if (searchBtn) searchBtn.onclick = handleSearch;
  if (useSelectionBtn) useSelectionBtn.onclick = handleUseSelection;

  if (searchInput) {
    searchInput.onkeydown = (e: KeyboardEvent) => {
      if (e.key === "Enter") {
        handleSearch();
      }
    };
  }

  // Quick field chips
  const fieldChips = document.querySelectorAll(".chip-field");
  fieldChips.forEach((chip) => {
    chip.addEventListener("click", () => {
      const field = chip.getAttribute("data-field");
      if (field && searchInput) {
        searchInput.value = field + searchInput.value;
        searchInput.focus();
      }
    });
  });

  // Format picker handlers
  const formatBtns = document.querySelectorAll(".format-btn");
  formatBtns.forEach((btn) => {
    btn.addEventListener("click", () => {
      const format = btn.getAttribute("data-format") as CitationFormat;
      if (format && !btn.hasAttribute("disabled")) {
        handleFormatChange(format, formatBtns);
      }
    });
  });

  // Bibliography handlers
  const bibliographyToggle = document.getElementById("bibliography-toggle");
  if (bibliographyToggle) {
    bibliographyToggle.onclick = () => togglePanel(bibliographyToggle, "#bibliography-panel");
  }

  const exportAllBtn = document.getElementById("export-all-btn");
  if (exportAllBtn) exportAllBtn.onclick = handleExportAll;

  const clearBibBtn = document.getElementById("clear-bib-btn");
  if (clearBibBtn) clearBibBtn.onclick = handleClearBibliography;
}

function toggleSettingsPanel(): void {
  const settingsPanel = document.getElementById("settings-panel");
  if (settingsPanel) {
    settingsPanel.classList.toggle("collapsed");
  }
}

function togglePanel(toggle: HTMLElement, panelSelector: string): void {
  const panel = document.querySelector(panelSelector);
  if (panel) {
    const isExpanded = toggle.getAttribute("aria-expanded") === "true";
    toggle.setAttribute("aria-expanded", String(!isExpanded));
    panel.classList.toggle("collapsed");
  }
}

function handleFormatChange(format: CitationFormat, buttons: NodeListOf<Element>): void {
  selectedFormat = format;
  buttons.forEach((btn) => {
    btn.classList.toggle("active", btn.getAttribute("data-format") === format);
  });
  updateResultPreviews();
}

function updateResultPreviews(): void {
  const previews = document.querySelectorAll(".result-preview");
  previews.forEach((preview) => {
    const card = preview.closest(".result-card");
    const bibcode = card?.getAttribute("data-bibcode");
    if (bibcode) {
      (preview as HTMLElement).textContent = getPreviewText();
    }
  });
}

function getPreviewText(): string {
  switch (selectedFormat) {
    case "inline":
      return ""; // Will be filled with actual citation in renderResults
    case "full":
      return "Full citation";
    case "bibtex":
      return "BibTeX entry";
  }
}

function toggleTokenVisibility(): void {
  const tokenInput = document.getElementById("token-input") as HTMLInputElement;
  if (tokenInput) {
    tokenInput.type = tokenInput.type === "password" ? "text" : "password";
  }
}

async function handleSaveToken(): Promise<void> {
  const tokenInput = document.getElementById("token-input") as HTMLInputElement;
  const tokenStatus = document.getElementById("token-status");
  const saveBtn = document.getElementById("save-token-btn") as HTMLButtonElement;

  const token = tokenInput?.value?.trim();
  if (!token) {
    if (tokenStatus) {
      tokenStatus.textContent = "Please enter a token";
      tokenStatus.className = "status-badge error";
    }
    return;
  }

  if (saveBtn) {
    saveBtn.disabled = true;
    saveBtn.innerHTML = `
      <svg class="spinner" viewBox="0 0 24 24"></svg>
      Validating...
    `;
  }

  const isValid = await validateToken(token);

  if (saveBtn) {
    saveBtn.disabled = false;
    saveBtn.innerHTML = `
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M19 21H5a2 2 0 01-2-2V5a2 2 0 012-2h11l5 5v11a2 2 0 01-2 2z"/>
        <polyline points="17,21 17,13 7,13 7,21"/>
        <polyline points="7,3 7,8 15,8"/>
      </svg>
      Save Token
    `;
  }

  if (isValid) {
    setToken(token);
    if (tokenInput) tokenInput.value = "";
    if (tokenStatus) {
      tokenStatus.textContent = "Token saved successfully";
      tokenStatus.className = "status-badge success";
    }
    // Load export formats now that we have a valid token
    loadExportFormats();
    // Update app state and collapse settings panel after success
    setTimeout(() => {
      updateAppState();
    }, 1000);
  } else {
    if (tokenStatus) {
      tokenStatus.textContent = "Invalid token - please check and try again";
      tokenStatus.className = "status-badge error";
    }
  }
}

function handleClearToken(): void {
  clearToken();
  const tokenInput = document.getElementById("token-input") as HTMLInputElement;
  const tokenStatus = document.getElementById("token-status");

  if (tokenInput) tokenInput.value = "";
  if (tokenStatus) {
    tokenStatus.textContent = "Token cleared";
    tokenStatus.className = "status-badge warning";
  }
  updateAppState();
}

async function handleSearch(): Promise<void> {
  const searchInput = document.getElementById("search-input") as HTMLInputElement;
  const resultsStatus = document.getElementById("results-status");
  const resultsCount = document.getElementById("results-count");

  const query = searchInput?.value?.trim();
  if (!query) {
    if (resultsStatus) {
      resultsStatus.textContent = "Enter a search query";
      resultsStatus.className = "results-status warning";
    }
    return;
  }

  const token = getToken();
  if (!token) {
    if (resultsStatus) {
      resultsStatus.textContent = "Please configure your API token first";
      resultsStatus.className = "results-status error";
    }
    // Open settings panel
    const settingsPanel = document.getElementById("settings-panel");
    if (settingsPanel) settingsPanel.classList.remove("collapsed");
    return;
  }

  if (resultsStatus) {
    resultsStatus.textContent = "Searching...";
    resultsStatus.className = "results-status loading";
  }
  if (resultsCount) resultsCount.textContent = "";
  clearResults();

  // Reset pagination state
  currentQuery = query;
  nextCursorMark = null;

  try {
    const result = await searchADS(query, token);
    if (result.docs.length === 0) {
      if (resultsStatus) {
        resultsStatus.textContent = "No results found";
        resultsStatus.className = "results-status warning";
      }
    } else {
      if (resultsStatus) resultsStatus.textContent = "";
      if (resultsCount) resultsCount.textContent = `(${result.numFound.toLocaleString()})`;
      nextCursorMark = result.nextCursorMark;
      renderResults(result.docs);
      updateLoadMoreButton();
    }
  } catch (err) {
    if (resultsStatus) {
      resultsStatus.textContent = err instanceof Error ? err.message : "Search failed";
      resultsStatus.className = "results-status error";
    }
  }
}

async function handleLoadMore(): Promise<void> {
  if (!currentQuery || !nextCursorMark) return;

  const token = getToken();
  if (!token) return;

  const loadMoreBtn = document.getElementById("load-more-btn") as HTMLButtonElement;
  if (loadMoreBtn) {
    loadMoreBtn.disabled = true;
    loadMoreBtn.textContent = "Loading...";
  }

  try {
    const result = await searchADS(currentQuery, token, 10, nextCursorMark);
    nextCursorMark = result.nextCursorMark;
    appendResults(result.docs);
    updateLoadMoreButton();
  } catch {
    showToast("Failed to load more results");
  } finally {
    if (loadMoreBtn) {
      loadMoreBtn.disabled = false;
      loadMoreBtn.textContent = "Load more";
    }
  }
}

async function handleUseSelection(): Promise<void> {
  const searchInput = document.getElementById("search-input") as HTMLInputElement;
  const resultsStatus = document.getElementById("results-status");

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      const text = selection.text?.trim();
      if (text && searchInput) {
        searchInput.value = text;
        await handleSearch();
      } else {
        if (resultsStatus) {
          resultsStatus.textContent = "No text selected";
          resultsStatus.className = "results-status warning";
        }
      }
    });
  } catch {
    if (resultsStatus) {
      resultsStatus.textContent = "Could not read selection";
      resultsStatus.className = "results-status error";
    }
  }
}

function clearResults(): void {
  const resultsList = document.getElementById("results-list");
  if (resultsList) {
    // Keep the empty state, remove result cards and load more button
    const cards = resultsList.querySelectorAll(".result-card");
    cards.forEach((card) => card.remove());
    const loadMoreBtn = document.getElementById("load-more-btn");
    if (loadMoreBtn) loadMoreBtn.remove();
  }
  // Reset pagination
  currentQuery = null;
  nextCursorMark = null;
}

function renderResults(results: SearchResult[]): void {
  const resultsList = document.getElementById("results-list");
  if (!resultsList) return;

  // Hide empty state
  const emptyState = resultsList.querySelector(".empty-state");
  if (emptyState) (emptyState as HTMLElement).style.display = "none";

  // Clear expanded state
  expandedBibcode = null;

  results.forEach((doc) => {
    const card = createResultCard(doc);
    resultsList.appendChild(card);
  });
}

function appendResults(results: SearchResult[]): void {
  const resultsList = document.getElementById("results-list");
  if (!resultsList) return;

  // Insert before load more button if it exists
  const loadMoreBtn = document.getElementById("load-more-btn");

  results.forEach((doc) => {
    const card = createResultCard(doc);
    if (loadMoreBtn) {
      resultsList.insertBefore(card, loadMoreBtn);
    } else {
      resultsList.appendChild(card);
    }
  });
}

function updateLoadMoreButton(): void {
  const resultsList = document.getElementById("results-list");
  if (!resultsList) return;

  let loadMoreBtn = document.getElementById("load-more-btn") as HTMLButtonElement;

  if (nextCursorMark) {
    // Create button if it doesn't exist
    if (!loadMoreBtn) {
      loadMoreBtn = document.createElement("button");
      loadMoreBtn.id = "load-more-btn";
      loadMoreBtn.className = "btn btn-ghost load-more-btn";
      loadMoreBtn.textContent = "Load more";
      loadMoreBtn.onclick = handleLoadMore;
      resultsList.appendChild(loadMoreBtn);
    }
  } else {
    // Remove button if no more results
    if (loadMoreBtn) loadMoreBtn.remove();
  }
}

function createResultCard(doc: SearchResult): HTMLElement {
  const display = formatResultDisplay(doc);
  const citation = formatInlineCitation(doc);

  const card = document.createElement("div");
  card.className = "result-card";
  card.setAttribute("data-bibcode", doc.bibcode);

  // Header section (clickable for expansion)
  const header = document.createElement("div");
  header.className = "result-header";
  header.setAttribute("role", "button");
  header.setAttribute("aria-expanded", "false");

  const titleEl = document.createElement("div");
  titleEl.className = "result-title";
  titleEl.textContent = display.title;

  const authorsEl = document.createElement("div");
  authorsEl.className = "result-authors";
  authorsEl.textContent = display.authors;

  const metaEl = document.createElement("div");
  metaEl.className = "result-meta";
  metaEl.textContent = [display.year, display.publication].filter(Boolean).join(" · ");

  // Actions row
  const actionsEl = document.createElement("div");
  actionsEl.className = "result-actions";

  const previewEl = document.createElement("div");
  previewEl.className = "result-preview";
  if (selectedFormat === "inline") {
    previewEl.textContent = `-> ${citation}`;
  } else if (selectedFormat === "full") {
    previewEl.textContent = "-> Full citation";
  } else {
    previewEl.textContent = "-> BibTeX entry";
  }

  const addBibBtn = document.createElement("button");
  addBibBtn.className = "action-btn add-bib-btn";
  addBibBtn.title = "Add to Bibliography";
  addBibBtn.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
    <path d="M19 21l-7-5-7 5V5a2 2 0 012-2h10a2 2 0 012 2z"/>
  </svg>`;
  if (isInBibliography(doc.bibcode)) {
    addBibBtn.classList.add("added");
  }
  addBibBtn.onclick = (e) => {
    e.stopPropagation();
    handleAddToBibliography(doc);
    addBibBtn.classList.add("added");
  };

  const insertBtn = document.createElement("button");
  insertBtn.className = "insert-btn";
  insertBtn.textContent = "Insert";
  insertBtn.onclick = (e) => {
    e.stopPropagation();
    insertCitation(doc, insertBtn);
  };

  actionsEl.appendChild(previewEl);
  actionsEl.appendChild(addBibBtn);
  actionsEl.appendChild(insertBtn);

  header.appendChild(titleEl);
  header.appendChild(authorsEl);
  header.appendChild(metaEl);
  header.appendChild(actionsEl);

  // Detail section (hidden by default)
  const detail = document.createElement("div");
  detail.className = "result-detail collapsed";
  detail.innerHTML = '<div class="detail-loading">Loading details...</div>';

  // Click handler for expansion
  header.onclick = (e) => {
    if ((e.target as HTMLElement).closest("button")) return;
    toggleDetailView(doc.bibcode, card, detail, header);
  };

  card.appendChild(header);
  card.appendChild(detail);
  return card;
}

async function toggleDetailView(
  bibcode: string,
  card: HTMLElement,
  detailEl: HTMLElement,
  header: HTMLElement
): Promise<void> {
  const isExpanded = header.getAttribute("aria-expanded") === "true";

  // Collapse if already expanded
  if (isExpanded) {
    header.setAttribute("aria-expanded", "false");
    detailEl.classList.add("collapsed");
    expandedBibcode = null;
    return;
  }

  // Collapse any other expanded card first
  if (expandedBibcode && expandedBibcode !== bibcode) {
    const prevCard = document.querySelector(`[data-bibcode="${expandedBibcode}"]`);
    const prevHeader = prevCard?.querySelector(".result-header");
    const prevDetail = prevCard?.querySelector(".result-detail");
    prevHeader?.setAttribute("aria-expanded", "false");
    prevDetail?.classList.add("collapsed");
  }

  // Expand this card
  header.setAttribute("aria-expanded", "true");
  detailEl.classList.remove("collapsed");
  expandedBibcode = bibcode;

  // Fetch detail if not cached
  if (!detailCache.has(bibcode)) {
    await fetchAndRenderDetail(bibcode, detailEl);
  } else {
    const cached = detailCache.get(bibcode);
    if (cached) renderDetail(cached, detailEl);
  }
}

async function fetchAndRenderDetail(bibcode: string, detailEl: HTMLElement): Promise<void> {
  const token = getToken();
  if (!token) {
    detailEl.innerHTML = '<div class="detail-error">API token required</div>';
    return;
  }

  try {
    const detail = await fetchPaperDetail(bibcode, token);
    if (detail) {
      detailCache.set(bibcode, detail);
      renderDetail(detail, detailEl);
    } else {
      detailEl.innerHTML = '<div class="detail-error">Paper not found</div>';
    }
  } catch {
    detailEl.innerHTML = '<div class="detail-error">Failed to load details</div>';
  }
}

function renderDetail(detail: PaperDetail, detailEl: HTMLElement): void {
  const abstractText = detail.abstract || "No abstract available.";
  const citations = detail.citation_count ?? 0;
  const doi = detail.doi?.[0] || null;
  const affiliations = detail.aff || [];

  // Truncate abstract for display
  const maxAbstractLength = 500;
  const truncatedAbstract =
    abstractText.length > maxAbstractLength
      ? abstractText.substring(0, maxAbstractLength) + "..."
      : abstractText;

  let affiliationsHtml = "";
  if (affiliations.length > 0) {
    const uniqueAffs = [...new Set(affiliations.filter((a) => a && a !== "-"))];
    if (uniqueAffs.length > 0) {
      affiliationsHtml = `
        <details class="detail-affiliations">
          <summary>Affiliations (${uniqueAffs.length})</summary>
          <ul class="aff-list">
            ${uniqueAffs
              .slice(0, 5)
              .map((a) => `<li>${a}</li>`)
              .join("")}
            ${uniqueAffs.length > 5 ? `<li>+${uniqueAffs.length - 5} more</li>` : ""}
          </ul>
        </details>
      `;
    }
  }

  detailEl.innerHTML = `
    <div class="detail-content">
      <div class="detail-abstract">
        <h4>Abstract</h4>
        <p class="abstract-text">${truncatedAbstract}</p>
      </div>
      <div class="detail-meta-row">
        <span class="detail-citations">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="14" height="14">
            <path d="M3 21c3 0 7-1 7-8V5c0-1.25-.756-2.017-2-2H4c-1.25 0-2 .75-2 1.972V11c0 1.25.75 2 2 2 1 0 1 0 1 1v1c0 1-1 2-2 2s-1 .008-1 1.031V21z"/>
            <path d="M15 21c3 0 7-1 7-8V5c0-1.25-.757-2.017-2-2h-4c-1.25 0-2 .75-2 1.972V11c0 1.25.75 2 2 2h.75c0 2.25.25 4-2.75 4v3z"/>
          </svg>
          ${citations} citations
        </span>
        ${doi ? `<a href="https://doi.org/${doi}" target="_blank" rel="noopener" class="detail-doi">${doi}</a>` : ""}
      </div>
      ${affiliationsHtml}
      <div class="detail-actions">
        <button class="btn btn-ghost references-btn">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="14" height="14">
            <line x1="8" y1="6" x2="21" y2="6"/>
            <line x1="8" y1="12" x2="21" y2="12"/>
            <line x1="8" y1="18" x2="21" y2="18"/>
            <line x1="3" y1="6" x2="3.01" y2="6"/>
            <line x1="3" y1="12" x2="3.01" y2="12"/>
            <line x1="3" y1="18" x2="3.01" y2="18"/>
          </svg>
          References
        </button>
        <button class="btn btn-ghost detail-add-bib-btn">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="14" height="14">
            <path d="M19 21l-7-5-7 5V5a2 2 0 012-2h10a2 2 0 012 2z"/>
          </svg>
          Add to Bib
        </button>
      </div>
    </div>
  `;

  // Bind handlers
  const refBtn = detailEl.querySelector(".references-btn");
  const addBibBtn = detailEl.querySelector(".detail-add-bib-btn");

  if (refBtn) {
    refBtn.addEventListener("click", () => showReferences(detail.bibcode));
  }
  if (addBibBtn) {
    addBibBtn.addEventListener("click", () => {
      handleAddToBibliography(detail);
      showToast("Added to bibliography");
    });
  }
}

async function getCitationText(doc: SearchResult): Promise<string> {
  if (selectedFormat === "inline") {
    return formatInlineCitation(doc);
  }

  const token = getToken();
  if (!token) {
    throw new Error("API token required for this format");
  }

  const format = selectedFormat === "full" ? "apsj" : "bibtex";
  return exportCitation(doc.bibcode, format, token);
}

async function insertCitation(doc: SearchResult, insertBtn: HTMLButtonElement): Promise<void> {
  const resultsStatus = document.getElementById("results-status");
  const originalText = insertBtn.textContent;

  try {
    // Show loading state for async formats
    if (selectedFormat !== "inline") {
      insertBtn.disabled = true;
      insertBtn.textContent = "...";
    }

    const citationText = await getCitationText(doc);

    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(citationText, Word.InsertLocation.replace);
      await context.sync();
    });

    if (resultsStatus) {
      const preview =
        citationText.length > 50 ? citationText.substring(0, 50) + "..." : citationText;
      resultsStatus.textContent = `Inserted: ${preview}`;
      resultsStatus.className = "results-status success";
    }
  } catch (err) {
    if (resultsStatus) {
      resultsStatus.textContent = err instanceof Error ? err.message : "Failed to insert citation";
      resultsStatus.className = "results-status error";
    }
  } finally {
    insertBtn.disabled = false;
    insertBtn.textContent = originalText;
  }
}

// References panel

async function showReferences(bibcode: string): Promise<void> {
  const token = getToken();
  if (!token) {
    showToast("API token required");
    return;
  }

  // Create or show references overlay
  let refPanel = document.getElementById("references-panel");
  if (!refPanel) {
    refPanel = createReferencesPanel();
    document.getElementById("app")?.appendChild(refPanel);
  }

  refPanel.setAttribute("data-parent-bibcode", bibcode);
  refPanel.classList.remove("collapsed");

  const refList = refPanel.querySelector(".references-list");
  const refCount = refPanel.querySelector(".references-count");
  const refStatus = refPanel.querySelector(".references-status");

  if (refList) refList.innerHTML = "";
  if (refStatus) {
    refStatus.textContent = "Loading references...";
    (refStatus as HTMLElement).className = "references-status loading";
  }

  try {
    const refs = await fetchReferences(bibcode, token);
    if (refCount) refCount.textContent = `(${refs.length})`;
    if (refStatus) refStatus.textContent = "";

    if (refs.length === 0) {
      if (refList) {
        refList.innerHTML = '<div class="references-empty">No references found</div>';
      }
    } else {
      refs.forEach((ref) => {
        const refCard = createReferenceCard(ref);
        refList?.appendChild(refCard);
      });
    }
  } catch {
    if (refStatus) {
      refStatus.textContent = "Failed to load references";
      (refStatus as HTMLElement).className = "references-status error";
    }
  }
}

function createReferencesPanel(): HTMLElement {
  const panel = document.createElement("div");
  panel.id = "references-panel";
  panel.className = "references-panel collapsed";
  panel.innerHTML = `
    <div class="references-header">
      <button class="back-btn" title="Back">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <polyline points="15 18 9 12 15 6"/>
        </svg>
      </button>
      <span class="references-title">References</span>
      <span class="references-count">(0)</span>
    </div>
    <div class="references-list"></div>
    <div class="references-status"></div>
  `;

  const backBtn = panel.querySelector(".back-btn");
  backBtn?.addEventListener("click", () => {
    panel.classList.add("collapsed");
  });

  return panel;
}

function createReferenceCard(doc: SearchResult): HTMLElement {
  const display = formatResultDisplay(doc);
  const card = document.createElement("div");
  card.className = "reference-card";
  card.setAttribute("data-bibcode", doc.bibcode);

  card.innerHTML = `
    <div class="ref-title">${display.title}</div>
    <div class="ref-authors">${display.authors}</div>
    <div class="ref-meta">${display.year}${display.publication ? " · " + display.publication : ""}</div>
    <div class="ref-actions">
      <button class="btn-sm add-bib-btn" title="Add to Bibliography">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M19 21l-7-5-7 5V5a2 2 0 012-2h10a2 2 0 012 2z"/>
        </svg>
      </button>
      <button class="btn-sm insert-btn">Insert</button>
    </div>
  `;

  const addBtn = card.querySelector(".add-bib-btn") as HTMLButtonElement;
  const insertBtn = card.querySelector(".insert-btn") as HTMLButtonElement;

  if (isInBibliography(doc.bibcode)) {
    addBtn?.classList.add("added");
  }

  addBtn?.addEventListener("click", () => {
    handleAddToBibliography(doc);
    addBtn.classList.add("added");
    showToast("Added to bibliography");
  });

  insertBtn?.addEventListener("click", () => {
    insertCitation(doc, insertBtn);
  });

  return card;
}

// Bibliography management

function handleAddToBibliography(doc: SearchResult | PaperDetail): void {
  const display = formatResultDisplay(doc);
  const entry: BibliographyEntry = {
    bibcode: doc.bibcode,
    title: display.title,
    authors: display.authors,
    year: display.year,
    addedAt: Date.now(),
  };

  const added = addToBibliography(entry);
  if (added) {
    updateBibliographyUI();
  }
}

function updateBibliographyUI(): void {
  const bib = getBibliography();
  const count = bib.length;

  // Update count badge
  const countEl = document.getElementById("bibliography-count");
  if (countEl) countEl.textContent = String(count);

  // Toggle empty state vs list
  const emptyEl = document.getElementById("bibliography-empty");
  const listEl = document.getElementById("bibliography-list");
  const actionsEl = document.getElementById("bibliography-actions");

  if (count === 0) {
    if (emptyEl) emptyEl.style.display = "block";
    if (listEl) listEl.style.display = "none";
    if (actionsEl) actionsEl.style.display = "none";
  } else {
    if (emptyEl) emptyEl.style.display = "none";
    if (listEl) {
      listEl.style.display = "block";
      renderBibliographyList(bib, listEl);
    }
    if (actionsEl) actionsEl.style.display = "block";
  }
}

async function loadExportFormats(): Promise<void> {
  const token = getToken();
  if (!token) return;

  try {
    const formats = await fetchExportManifest(token);
    populateExportDropdown(formats);
  } catch {
    // Fall back to default options if manifest fails
  }
}

function populateExportDropdown(formats: ExportFormatInfo[]): void {
  const select = document.getElementById("bibliography-format") as HTMLSelectElement;
  if (!select) return;

  // Filter to useful formats for Word users (exclude XML, CSL, and custom)
  const excludedTypes = ["XML", "CSL", "custom"];
  const usefulFormats = formats.filter((f) => !excludedTypes.includes(f.type));

  // Group by type for better organization
  const grouped = usefulFormats.reduce(
    (acc, f) => {
      const type = f.type;
      if (!acc[type]) acc[type] = [];
      acc[type].push(f);
      return acc;
    },
    {} as Record<string, ExportFormatInfo[]>
  );

  // Clear existing options
  select.innerHTML = "";

  // Define preferred order for groups
  const typeOrder = ["HTML", "tagged", "LaTeX", "other"];

  // Add options grouped by type
  for (const type of typeOrder) {
    const typeFormats = grouped[type];
    if (!typeFormats || typeFormats.length === 0) continue;

    const optgroup = document.createElement("optgroup");
    optgroup.label = type === "HTML" ? "Citations" : type === "tagged" ? "Tagged" : type;

    for (const format of typeFormats) {
      const option = document.createElement("option");
      // Extract route name (remove leading slash)
      option.value = format.route.replace("/", "");
      option.textContent = format.name;
      optgroup.appendChild(option);
    }

    select.appendChild(optgroup);
  }

  // Handle any remaining types not in the preferred order
  for (const type of Object.keys(grouped)) {
    if (typeOrder.includes(type)) continue;
    const typeFormats = grouped[type];
    if (!typeFormats || typeFormats.length === 0) continue;

    const optgroup = document.createElement("optgroup");
    optgroup.label = type;

    for (const format of typeFormats) {
      const option = document.createElement("option");
      option.value = format.route.replace("/", "");
      option.textContent = format.name;
      optgroup.appendChild(option);
    }

    select.appendChild(optgroup);
  }

  // Set default selection to APS Journals (apsj) if available
  const defaultFormat = "apsj";
  if (select.querySelector(`option[value="${defaultFormat}"]`)) {
    select.value = defaultFormat;
  }
}

function renderBibliographyList(entries: BibliographyEntry[], container: HTMLElement): void {
  container.innerHTML = "";

  // Sort by addedAt descending (most recent first)
  const sorted = [...entries].sort((a, b) => b.addedAt - a.addedAt);

  sorted.forEach((entry) => {
    const item = document.createElement("div");
    item.className = "bibliography-item";
    item.setAttribute("data-bibcode", entry.bibcode);

    item.innerHTML = `
      <div class="bib-info">
        <div class="bib-title">${entry.title}</div>
        <div class="bib-meta">${entry.authors} (${entry.year})</div>
      </div>
      <button class="bib-remove-btn" title="Remove">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <line x1="18" y1="6" x2="6" y2="18"/>
          <line x1="6" y1="6" x2="18" y2="18"/>
        </svg>
      </button>
    `;

    const removeBtn = item.querySelector(".bib-remove-btn");
    removeBtn?.addEventListener("click", () => {
      removeFromBibliography(entry.bibcode);
      updateBibliographyUI();
    });

    container.appendChild(item);
  });
}

async function handleExportAll(): Promise<void> {
  const bib = getBibliography();
  if (bib.length === 0) {
    showToast("Bibliography is empty");
    return;
  }

  const token = getToken();
  if (!token) {
    showToast("API token required");
    return;
  }

  const formatSelect = document.getElementById("bibliography-format") as HTMLSelectElement;
  const format = formatSelect?.value || "apsj";

  const exportBtn = document.getElementById("export-all-btn") as HTMLButtonElement;
  const originalHtml = exportBtn?.innerHTML;

  try {
    if (exportBtn) {
      exportBtn.disabled = true;
      exportBtn.textContent = "Exporting...";
    }

    const bibcodes = bib.map((e) => e.bibcode);
    const citationText = await exportCitations(bibcodes, format, token);

    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(citationText, Word.InsertLocation.replace);
      await context.sync();
    });

    showToast(`Inserted ${bib.length} citations`);
  } catch {
    showToast("Export failed");
  } finally {
    if (exportBtn && originalHtml) {
      exportBtn.disabled = false;
      exportBtn.innerHTML = originalHtml;
    }
  }
}

function handleClearBibliography(): void {
  if (confirm("Remove all papers from bibliography?")) {
    clearBibliography();
    updateBibliographyUI();
    showToast("Bibliography cleared");
  }
}

// Toast notification

function showToast(message: string): void {
  // Remove existing toast
  const existing = document.querySelector(".toast");
  if (existing) existing.remove();

  const toast = document.createElement("div");
  toast.className = "toast";
  toast.textContent = message;
  document.body.appendChild(toast);

  // Remove after animation
  setTimeout(() => {
    toast.remove();
  }, 2500);
}
