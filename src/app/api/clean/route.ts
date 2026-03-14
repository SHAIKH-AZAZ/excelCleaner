import { NextRequest, NextResponse } from "next/server";
import { readFile, unlink } from "fs/promises";
import path from "path";
import os from "os";
import { cleanSheets } from "@/lib/xlsxCleaner";

export const runtime = "nodejs";

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { fileId, selectedSheets, fileName } = body;

    if (!fileId || !selectedSheets || !Array.isArray(selectedSheets)) {
      return NextResponse.json(
        { error: "Missing fileId or selectedSheets" },
        { status: 400 }
      );
    }

    if (selectedSheets.length === 0) {
      return NextResponse.json(
        { error: "No sheets selected for cleaning" },
        { status: 400 }
      );
    }

    // Read the temp file
    const tmpDir = path.join(os.tmpdir(), "excel-cleaner");
    let tmpPath = path.join(tmpDir, `${fileId}.xlsx`);

    let buffer: Buffer;
    try {
      buffer = await readFile(tmpPath);
    } catch {
      // Try .xlsm extension
      try {
        tmpPath = path.join(tmpDir, `${fileId}.xlsm`);
        buffer = await readFile(tmpPath);
      } catch {
        return NextResponse.json(
          { error: "File not found. Please upload again." },
          { status: 404 }
        );
      }
    }

    // Clean the sheets
    const cleanedBuffer = await cleanSheets(buffer, selectedSheets);

    // Clean up temp file
    try {
      await unlink(tmpPath);
    } catch {
      // ignore cleanup errors
    }

    // Return the cleaned file
    const cleanedName = fileName
      ? fileName.replace(/\.(xlsx|xlsm)$/i, "_cleaned.$1")
      : "cleaned.xlsx";

    return new NextResponse(new Uint8Array(cleanedBuffer), {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${cleanedName}"`,
        "Content-Length": String(cleanedBuffer.length),
      },
    });
  } catch (error: any) {
    console.error("Error cleaning file:", error);
    return NextResponse.json(
      { error: error.message || "Failed to clean the file" },
      { status: 500 }
    );
  }
}
