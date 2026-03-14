import { NextRequest, NextResponse } from "next/server";
import { writeFile, mkdir } from "fs/promises";
import path from "path";
import { v4 as uuidv4 } from "uuid";
import { getSheetNames } from "@/lib/xlsxCleaner";
import os from "os";

export const runtime = "nodejs";

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get("file") as File | null;

    if (!file) {
      return NextResponse.json({ error: "No file uploaded" }, { status: 400 });
    }

    const ext = file.name.toLowerCase();
    if (!ext.endsWith(".xlsx") && !ext.endsWith(".xlsm")) {
      return NextResponse.json(
        { error: "Only .xlsx and .xlsm files are supported" },
        { status: 400 }
      );
    }

    // Read the file into a buffer
    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    // Extract sheet names
    const sheets = getSheetNames(buffer);

    if (sheets.length === 0) {
      return NextResponse.json(
        { error: "No sheets found in the workbook" },
        { status: 400 }
      );
    }

    // Save temp file for later processing
    const fileId = uuidv4();
    const fileExt = file.name.toLowerCase().endsWith(".xlsm") ? ".xlsm" : ".xlsx";
    const tmpDir = path.join(os.tmpdir(), "excel-cleaner");
    await mkdir(tmpDir, { recursive: true });
    const tmpPath = path.join(tmpDir, `${fileId}${fileExt}`);
    await writeFile(tmpPath, buffer);

    return NextResponse.json({
      sheets,
      fileId,
      fileName: file.name,
    });
  } catch (error: any) {
    console.error("Error processing upload:", error);
    return NextResponse.json(
      { error: error.message || "Failed to process the file" },
      { status: 500 }
    );
  }
}
