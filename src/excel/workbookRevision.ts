export type WorkbookChangeKind = "content" | "structure" | "format";

export interface WorkbookRevision {
  sequence: number;
  content: number;
  structure: number;
  format: number;
}

export interface RevisionCoverage {
  content: boolean;
  structure: boolean;
  format: boolean;
  protection: boolean;
  sheetNames: boolean;
}

let revision: WorkbookRevision = { sequence: 0, content: 0, structure: 0, format: 0 };
let coverage: RevisionCoverage = {
  content: false,
  structure: false,
  format: false,
  protection: false,
  sheetNames: false
};

export function getWorkbookRevision(): WorkbookRevision {
  return { ...revision };
}

export function getRevisionCoverage(): RevisionCoverage {
  return { ...coverage };
}

export function setRevisionCoverage(next: Partial<RevisionCoverage>): RevisionCoverage {
  coverage = { ...coverage, ...next };
  return getRevisionCoverage();
}

export function bumpWorkbookRevision(kind: WorkbookChangeKind): WorkbookRevision {
  revision.sequence += 1;
  if (kind === "structure") {
    revision.structure += 1;
    revision.content += 1;
    revision.format += 1;
  } else {
    revision[kind] += 1;
  }
  return getWorkbookRevision();
}

/** Tests and a newly attached workbook start from an explicitly unknown era. */
export function resetWorkbookRevision(): void {
  revision = { sequence: 0, content: 0, structure: 0, format: 0 };
  coverage = { content: false, structure: false, format: false, protection: false, sheetNames: false };
}
