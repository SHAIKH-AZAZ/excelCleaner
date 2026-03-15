import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import * as XLSX from "xlsx";

const parserOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  preserveOrder: true,
  textNodeName: "#text",
  trimValues: false,
  parseTagValue: false,
};

/**
 * Extract worksheet names from an XLSX buffer.
 */
export function getSheetNames(buffer: Buffer): string[] {
  const workbook = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  return workbook.SheetNames;
}

/**
 * Count drawing objects (images, shapes, groups) per sheet.
 * Returns a map of sheet name → total object count.
 */
export async function getSheetDrawingCounts(
  buffer: Buffer
): Promise<Record<string, number>> {
  const zip = await JSZip.loadAsync(buffer);

  const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
  if (!workbookXml) return {};
  const sheets = parseWorkbook(workbookXml);

  const wbRelsXml = await zip
    .file("xl/_rels/workbook.xml.rels")
    ?.async("string");
  if (!wbRelsXml) return {};
  const wbRels = parseRels(wbRelsXml);

  const counts: Record<string, number> = {};

  for (const sheet of sheets) {
    const sheetRel = wbRels.find((r) => r.id === sheet.rId);
    if (!sheetRel) { counts[sheet.name] = 0; continue; }

    const sheetPath = `xl/${sheetRel.target.replace(/^\.?\//, "")}`;
    const sheetRelsPath = sheetPath.replace(
      /^(.*\/)([^/]+)$/,
      "$1_rels/$2.rels"
    );
    const sheetRelsXml = await zip.file(sheetRelsPath)?.async("string");
    if (!sheetRelsXml) { counts[sheet.name] = 0; continue; }

    const sheetRels = parseRels(sheetRelsXml);
    const drawingRels = sheetRels.filter(
      (r) =>
        r.type.includes("/drawing") ||
        r.type.includes("/vmlDrawing") ||
        r.type.includes("/image")
    );

    let total = 0;

    for (const rel of drawingRels) {
      let drawingPath: string;
      const target = rel.target;
      if (target.startsWith("../")) {
        drawingPath = `xl/${target.replace("../", "")}`;
      } else if (target.startsWith("/")) {
        drawingPath = target.slice(1);
      } else {
        const sheetDir = sheetPath.substring(0, sheetPath.lastIndexOf("/"));
        drawingPath = `${sheetDir}/${target}`;
      }

      const drawingXml = await zip.file(drawingPath)?.async("string");
      if (!drawingXml) continue;

      if (rel.type.includes("/vmlDrawing")) {
        // VML: count <v:shape> elements
        total += (drawingXml.match(/<v:shape\b/gi) || []).length;
      } else {
        // Modern drawing: count anchor elements (each = one object)
        total += (
          drawingXml.match(
            /<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b/gi
          ) || []
        ).length;
      }
    }

    counts[sheet.name] = total;
  }

  return counts;
}


/**
 * Helper: find the first element with a given tag in a parsed XML array (preserveOrder mode).
 */
function findElement(arr: any[], tagName: string): any | undefined {
  if (!Array.isArray(arr)) return undefined;
  return arr.find((el: any) => el[tagName] !== undefined);
}

/**
 * Helper: find all elements with a given tag.
 */
function findAllElements(arr: any[], tagName: string): any[] {
  if (!Array.isArray(arr)) return [];
  return arr.filter((el: any) => el[tagName] !== undefined);
}

/**
 * Parse the workbook.xml to get an ordered list of sheet names and their rId references.
 * READ-ONLY — does not modify the XML.
 */
function parseWorkbook(
  xml: string
): { name: string; sheetId: string; rId: string }[] {
  const parser = new XMLParser(parserOptions);
  const parsed = parser.parse(xml);

  const workbook = findElement(parsed, "workbook");
  if (!workbook) return [];

  const sheets = findElement(workbook["workbook"], "sheets");
  if (!sheets) return [];

  const sheetElements = findAllElements(sheets["sheets"], "sheet");
  return sheetElements.map((s: any) => ({
    name: s[":@"]["@_name"],
    sheetId: s[":@"]["@_sheetId"],
    rId: s[":@"]["@_r:id"],
  }));
}

/**
 * Parse a .rels file to get relationships. READ-ONLY.
 */
function parseRels(
  xml: string
): { id: string; type: string; target: string }[] {
  const parser = new XMLParser(parserOptions);
  const parsed = parser.parse(xml);

  const relationships = findElement(parsed, "Relationships");
  if (!relationships) return [];

  const rels = findAllElements(
    relationships["Relationships"],
    "Relationship"
  );
  return rels.map((r: any) => ({
    id: r[":@"]["@_Id"],
    type: r[":@"]["@_Type"],
    target: r[":@"]["@_Target"],
  }));
}

// ===== REGEX-BASED XML SURGERY (preserves original structure) =====

/**
 * Remove a <Relationship> element by its Id attribute from a .rels XML string.
 * Returns null if the file has no remaining Relationship elements.
 */
function removeRelById(xml: string, relId: string): string | null {
  // Match both self-closing and paired Relationship tags with the given Id
  const pattern = new RegExp(
    `\\s*<Relationship[^>]*\\bId\\s*=\\s*"${escapeRegex(relId)}"[^>]*\\/?>`,
    "g"
  );
  const result = xml.replace(pattern, "");

  // Check if any Relationship elements remain
  if (!/<Relationship\b/i.test(result)) {
    return null;
  }
  return result;
}

/**
 * Remove <drawing .../> and <legacyDrawing .../> elements from worksheet XML.
 */
function removeDrawingElements(xml: string): string {
  // Remove self-closing <drawing .../> 
  let result = xml.replace(/<drawing\b[^>]*\/>\s*/gi, "");
  // Remove paired <drawing ...>...</drawing>
  result = result.replace(/<drawing\b[^>]*>[\s\S]*?<\/drawing>\s*/gi, "");
  // Remove self-closing <legacyDrawing .../>
  result = result.replace(/<legacyDrawing\b[^>]*\/>\s*/gi, "");
  // Remove paired <legacyDrawing>...</legacyDrawing>
  result = result.replace(/<legacyDrawing\b[^>]*>[\s\S]*?<\/legacyDrawing>\s*/gi, "");
  return result;
}

/**
 * Remove an <Override> entry from [Content_Types].xml by PartName.
 */
function removeContentTypeOverride(xml: string, partName: string): string {
  const escaped = escapeRegex(partName);
  const pattern = new RegExp(
    `\\s*<Override[^>]*\\bPartName\\s*=\\s*"${escaped}"[^>]*\\/?>`,
    "g"
  );
  return xml.replace(pattern, "");
}

/**
 * Escape special regex characters in a string.
 */
function escapeRegex(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\\/]/g, "\\$&");
}

/**
 * Clean selected sheets by removing all drawing objects (images, shapes, grouped shapes).
 * Returns a cleaned XLSX buffer.
 */
export async function cleanSheets(
  buffer: Buffer,
  selectedSheets: string[]
): Promise<Buffer> {
  const zip = await JSZip.loadAsync(buffer);

  // Parse workbook to get sheet info (read-only)
  const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
  if (!workbookXml) throw new Error("Invalid XLSX: missing workbook.xml");

  const sheets = parseWorkbook(workbookXml);

  // Parse workbook rels to map rId → sheet file target (read-only)
  const wbRelsXml = await zip
    .file("xl/_rels/workbook.xml.rels")
    ?.async("string");
  if (!wbRelsXml) throw new Error("Invalid XLSX: missing workbook.xml.rels");

  const wbRels = parseRels(wbRelsXml);

  // Build a set of selected sheet names (lowercase for comparison)
  const selectedSet = new Set(selectedSheets.map((s) => s.toLowerCase()));

  // Track media files to potentially remove
  const mediaToRemove: Set<string> = new Set();

  // Read content types
  let contentTypesXml = await zip
    .file("[Content_Types].xml")
    ?.async("string");
  if (!contentTypesXml)
    throw new Error("Invalid XLSX: missing [Content_Types].xml");

  // Process each selected sheet
  for (const sheet of sheets) {
    if (!selectedSet.has(sheet.name.toLowerCase())) continue;

    // Find the sheet file from workbook rels
    const sheetRel = wbRels.find((r) => r.id === sheet.rId);
    if (!sheetRel) continue;

    const sheetPath = `xl/${sheetRel.target.replace(/^\.?\//, "")}`;

    // Parse the sheet's relationship file (read-only for data extraction)
    const sheetRelsPath = sheetPath.replace(
      /^(.*\/)([^/]+)$/,
      "$1_rels/$2.rels"
    );
    const sheetRelsXml = await zip.file(sheetRelsPath)?.async("string");

    if (!sheetRelsXml) {
      // No rels file means no drawings — skip
      continue;
    }

    const sheetRels = parseRels(sheetRelsXml);

    // Find drawing relationships
    const drawingRels = sheetRels.filter(
      (r) =>
        r.type.includes("/drawing") ||
        r.type.includes("/vmlDrawing") ||
        r.type.includes("/image")
    );

    if (drawingRels.length === 0) continue;

    // Surgically edit the rels XML (regex-based)
    let modifiedRelsXml: string | null = sheetRelsXml;

    for (const drawingRel of drawingRels) {
      // Resolve the drawing file path
      const drawingTarget = drawingRel.target;
      let drawingPath: string;

      if (drawingTarget.startsWith("../")) {
        drawingPath = `xl/${drawingTarget.replace("../", "")}`;
      } else if (drawingTarget.startsWith("/")) {
        drawingPath = drawingTarget.slice(1);
      } else {
        const sheetDir = sheetPath.substring(0, sheetPath.lastIndexOf("/"));
        drawingPath = `${sheetDir}/${drawingTarget}`;
      }

      // Collect media referenced by this drawing's rels (read-only parse)
      const drawingRelsPath = drawingPath.replace(
        /^(.*\/)([^/]+)$/,
        "$1_rels/$2.rels"
      );
      const drawingRelsFileXml = await zip
        .file(drawingRelsPath)
        ?.async("string");
      if (drawingRelsFileXml) {
        const drawingFileRels = parseRels(drawingRelsFileXml);
        for (const rel of drawingFileRels) {
          if (
            rel.target.includes("media/") ||
            rel.type.includes("/image")
          ) {
            const mediaName = rel.target.split("/").pop();
            if (mediaName) {
              mediaToRemove.add(mediaName);
            }
          }
        }
        // Remove the drawing's rels file
        zip.remove(drawingRelsPath);
      }

      // Remove the drawing file itself
      zip.remove(drawingPath);

      // Remove from content types (regex-based)
      contentTypesXml = removeContentTypeOverride(
        contentTypesXml,
        `/${drawingPath}`
      );

      // Remove the relationship from the sheet's rels (regex-based)
      if (modifiedRelsXml) {
        modifiedRelsXml = removeRelById(modifiedRelsXml, drawingRel.id);
      }
    }

    // Update or remove the sheet rels file
    if (modifiedRelsXml === null) {
      zip.remove(sheetRelsPath);
    } else {
      zip.file(sheetRelsPath, modifiedRelsXml);
    }

    // Remove <drawing> and <legacyDrawing> from the sheet XML (regex-based)
    const sheetXml = await zip.file(sheetPath)?.async("string");
    if (sheetXml) {
      const cleanedSheetXml = removeDrawingElements(sheetXml);
      zip.file(sheetPath, cleanedSheetXml);
    }
  }

  // Remove orphaned media files (only if not referenced by remaining drawings)
  const remainingMediaRefs: Set<string> = new Set();
  for (const entry of Object.keys(zip.files)) {
    if (entry.match(/^xl\/drawings\/_rels\/.*\.rels$/)) {
      const drawingRelsXml = await zip.file(entry)?.async("string");
      if (drawingRelsXml) {
        const drawingRels = parseRels(drawingRelsXml);
        for (const rel of drawingRels) {
          if (rel.target.startsWith("../media/")) {
            remainingMediaRefs.add(rel.target.replace("../media/", ""));
          }
        }
      }
    }
  }

  for (const media of mediaToRemove) {
    if (!remainingMediaRefs.has(media)) {
      const mediaPath = `xl/media/${media}`;
      zip.remove(mediaPath);
      contentTypesXml = removeContentTypeOverride(
        contentTypesXml,
        `/${mediaPath}`
      );
    }
  }

  // Update content types
  zip.file("[Content_Types].xml", contentTypesXml);

  // Generate the cleaned XLSX
  const cleaned = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

  return Buffer.from(cleaned);
}
