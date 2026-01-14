import { SearchResult } from "../api/ads";

export function formatInlineCitation(doc: SearchResult): string {
  const authors = doc.author || [];
  const year = doc.year || "n.d.";

  if (authors.length === 0) {
    return `(${year})`;
  }

  const firstAuthor = extractLastName(authors[0]);

  if (authors.length === 1) {
    return `${firstAuthor} (${year})`;
  } else if (authors.length === 2) {
    const secondAuthor = extractLastName(authors[1]);
    return `${firstAuthor} & ${secondAuthor} (${year})`;
  } else {
    return `${firstAuthor} et al. (${year})`;
  }
}

function extractLastName(authorName: string): string {
  // ADS format is "Last, First" or "Last, First Middle"
  const commaIndex = authorName.indexOf(",");
  if (commaIndex > 0) {
    return authorName.substring(0, commaIndex);
  }
  // Fallback: take first word
  return authorName.split(" ")[0];
}

export function formatResultDisplay(doc: SearchResult): {
  authors: string;
  title: string;
  year: string;
  publication: string;
} {
  const authors = doc.author || [];
  let authorStr: string;

  if (authors.length === 0) {
    authorStr = "Unknown";
  } else if (authors.length === 1) {
    authorStr = authors[0];
  } else if (authors.length === 2) {
    authorStr = `${authors[0]}; ${authors[1]}`;
  } else {
    authorStr = `${authors[0]} +${authors.length - 1} more`;
  }

  return {
    authors: authorStr,
    title: doc.title?.[0] || "Untitled",
    year: doc.year || "n.d.",
    publication: doc.pub || "",
  };
}
