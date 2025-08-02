import "dotenv/config";
import Fastify from "fastify";
import { google } from "googleapis";
import { GoogleAuth } from "google-auth-library";
import fs from "fs";
import path from "path";

const server = Fastify({ logger: true });

server.get("/google-sheets", async () => {
  try {
    const data = await getGoogleSheetsData();
    return {
      message: "Google Sheets data retrieved successfully!",
      totalRecords: data.length,
      data,
    };
  } catch (error: any) {
    return {
      error: "Failed to read Google Sheets data",
      message: error.message,
    };
  }
});

// Google Sheets configuration
const SPREADSHEET_ID = process.env.GOOGLE_SPREADSHEET_ID;
const SHEET_NAME = process.env.GOOGLE_SHEET_NAME || "Página1";
const RANGE = process.env.GOOGLE_SHEET_RANGE || "A1:I169"; // Read all columns

const config = {
  spreadsheetId: SPREADSHEET_ID,
  sheetName: SHEET_NAME,
  range: `${SHEET_NAME}!${RANGE}`,
  credentialsPath: "credentials.json",
};

// Initialize Google Sheets API
const initializeGoogleSheets = async () => {
  try {
    console.log("🔐 Initializing Google Sheets authentication...");

    // Check if credentials file exists
    const credentialsPath = path.resolve(config.credentialsPath);
    if (!fs.existsSync(credentialsPath)) {
      throw new Error(`Credentials file not found at: ${credentialsPath}`);
    }

    // Read credentials
    const credentials = JSON.parse(fs.readFileSync(credentialsPath, "utf8"));
    console.log("📄 Credentials loaded successfully");

    // Initialize auth
    const auth = new GoogleAuth({
      keyFile: credentialsPath,
      scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
    });

    // Initialize Google Sheets API
    const sheets = google.sheets({ version: "v4", auth });

    console.log("✅ Google Sheets API initialized successfully");
    return sheets;
  } catch (error: any) {
    console.error("❌ Error initializing Google Sheets:", error.message);
    throw error;
  }
};

// Get spreadsheet information
const getSpreadsheetInfo = async (sheets: any, spreadsheetId: string) => {
  try {
    console.log("📋 Getting spreadsheet information...");

    const response = await sheets.spreadsheets.get({
      spreadsheetId,
      includeGridData: false,
    });

    const spreadsheet = response.data;
    console.log("📊 Spreadsheet Info:", {
      title: spreadsheet.properties?.title,
      sheets: spreadsheet.sheets?.map((sheet: any) => ({
        title: sheet.properties?.title,
        sheetId: sheet.properties?.sheetId,
        gridProperties: sheet.properties?.gridProperties,
      })),
    });

    return spreadsheet;
  } catch (error: any) {
    console.error("❌ Error getting spreadsheet info:", error.message);
    throw error;
  }
};

// Transform raw Google Sheets data into structured JSON
const transformGoogleSheetsData = (
  rawData: any[][],
  hasHeaders: boolean = true
) => {
  if (!Array.isArray(rawData) || rawData.length === 0) {
    console.log("⚠️ No data to transform");
    return [];
  }

  console.log(
    `🔄 Transforming ${rawData.length} rows of Google Sheets data...`
  );

  let headers: string[] = [];
  let dataRows: any[][] = [];

  if (hasHeaders && rawData.length > 0) {
    headers = rawData[0].map((header: any, index: number) =>
      String(header || `Column_${index + 1}`).trim()
    );
    dataRows = rawData.slice(1);
    console.log("📋 Headers found:", headers);
  } else {
    // Generate generic headers if no headers are present
    const maxColumns = Math.max(...rawData.map((row) => row.length));
    headers = Array.from({ length: maxColumns }, (_, i) => `Column_${i + 1}`);
    dataRows = rawData;
    console.log("📋 Generated headers:", headers);
  }

  const transformedData = dataRows
    .map((row: any[], rowIndex: number) => {
      // Skip empty rows
      if (!row || row.every((cell) => !cell || String(cell).trim() === "")) {
        return null;
      }

      // Clean and parse values
      const parseValue = (value: any) => {
        if (value === null || value === undefined) return "";

        const stringValue = String(value).trim();

        // Try to parse as number
        if (stringValue !== "" && !isNaN(Number(stringValue))) {
          // Handle Brazilian decimal format (comma as decimal separator)
          const normalizedValue = stringValue.replace(",", ".");
          const parsed = parseFloat(normalizedValue);
          if (!isNaN(parsed)) return parsed;
        }

        // Return as string
        return stringValue;
      };

      const transformedRow: any = {};

      headers.forEach((header, index) => {
        transformedRow[header] = parseValue(row[index]);
      });

      return transformedRow;
    })
    .filter((row) => row !== null); // Remove null rows

  console.log(
    `✅ Successfully transformed ${transformedData.length} valid rows`
  );

  // Show a sample of the transformed data
  if (transformedData.length > 0) {
    console.log("📊 Sample transformed data (first 2 rows):");
    console.log(JSON.stringify(transformedData.slice(0, 2), null, 2));
  }

  return transformedData;
};

// Main function to get data from Google Sheets
const getGoogleSheetsData = async () => {
  try {
    console.log("🚀 Starting Google Sheets data retrieval...");
    console.log("📋 Configuration:", config);

    if (!config.spreadsheetId) {
      throw new Error("GOOGLE_SPREADSHEET_ID environment variable is required");
    }

    // Initialize Google Sheets API
    const sheets = await initializeGoogleSheets();

    // Get spreadsheet information
    await getSpreadsheetInfo(sheets, config.spreadsheetId);

    // Read the data from the specified range
    console.log(`📖 Reading data from range: ${config.range}`);

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: config.spreadsheetId,
      range: config.range,
    });

    const rawData = response.data.values;

    if (!rawData || rawData.length === 0) {
      console.log("⚠️ No data found in the specified range");
      return [];
    }

    console.log(`📊 Raw data retrieved: ${rawData.length} rows`);
    console.log("🔍 First few rows:", rawData.slice(0, 3));

    // Transform the raw data into structured JSON
    const transformedData = transformGoogleSheetsData(rawData, true);

    console.log("📋 Final transformed data ready for use:", {
      totalRows: transformedData.length,
      sampleRow: transformedData[0] || null,
      headers:
        transformedData.length > 0 ? Object.keys(transformedData[0]) : [],
    });

    return transformedData;
  } catch (error: any) {
    console.error("❌ Error accessing Google Sheets:");

    if (error.response) {
      console.error("Status:", error.response.status);
      console.error("Status Text:", error.response.statusText);
      console.error(
        "API Response:",
        JSON.stringify(error.response.data, null, 2)
      );
    } else {
      console.error("Error message:", error.message);
    }

    throw error;
  }
};

// Start the server
const start = async () => {
  try {
    await server.listen({ port: 3001, host: "0.0.0.0" });
    console.log("🚀 Google Sheets server is running on http://localhost:3001");

    // Test the Google Sheets integration on startup
    console.log("\n🧪 Testing Google Sheets integration...");
    try {
      const sheetsData = await getGoogleSheetsData();
      console.log("📊 Google Sheets Data Retrieved:", {
        totalRecords: sheetsData.length,
        firstRecord: sheetsData[0] || null,
        lastRecord: sheetsData[sheetsData.length - 1] || null,
      });
    } catch (testError: any) {
      console.error("⚠️ Initial Google Sheets test failed:", testError.message);
    }
  } catch (err) {
    server.log.error(err);
    process.exit(1);
  }
};

// Export functions for use in other modules
export {
  getGoogleSheetsData,
  transformGoogleSheetsData,
  initializeGoogleSheets,
  getSpreadsheetInfo,
};

// Start the server if this file is run directly
if (require.main === module) {
  start();
}
