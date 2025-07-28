import "dotenv/config";
import Fastify from "fastify";
import axios from "axios";
import {
  ClientCredentialRequest,
  ConfidentialClientApplication,
} from "@azure/msal-node";

const server = Fastify({ logger: true });

server.get("/", async () => {
  return { message: "Hello Fastify with TypeScript!" };
});

const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const TENANT_ID = process.env.TENANT_ID;

// Extracted from
// https://excel.cloud.microsoft/open/onedrive/?docId=D1B4FC2F8C816AEE%21sa5cf69d6228847c78cf590c44bceed4d&driveId=D1B4FC2F8C816AEE
const EXCEL_DRIVE_ID = process.env.EXCEL_DRIVE_ID;
const EXCEL_ITEM_ID = process.env.EXCEL_ITEM_ID;
const WORK_SHEET_NAME = "historico";
const TABLE_NAME = "Table1";

const config = {
  clientId: CLIENT_ID,
  clientSecret: CLIENT_SECRET,
  tenantId: TENANT_ID,
  excelDriveId: EXCEL_DRIVE_ID,
  excelItemId: EXCEL_ITEM_ID,
  worksheetName: WORK_SHEET_NAME,
  tableName: TABLE_NAME,
};

const listAvailableWorksheets = async (
  accessToken: string,
  driveId: string,
  excelWorkbookId: string
) => {
  try {
    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${excelWorkbookId}/workbook/worksheets`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
      }
    );

    console.log("📋 Available worksheets:");
    response.data.value.forEach((sheet: any, index: number) => {
      console.log(`  ${index + 1}. "${sheet.name}"`);
    });

    return response.data.value;
  } catch (err: any) {
    console.error(
      "❌ Error listing worksheets:",
      err.response?.data || err.message
    );
    return [];
  }
};

const getExcelDataFromSharepoint = async () => {
  try {
    console.log("config", config);
    console.log("Starting authentication...");

    // 1. Configure the MSAL client for Client Credentials Flow
    const clientApp = new ConfidentialClientApplication({
      auth: {
        clientId: config.clientId!,
        clientSecret: config.clientSecret!,
        authority: `https://login.microsoftonline.com/${config.tenantId!}`,
      },
    });

    const clientCredentialRequest: ClientCredentialRequest = {
      scopes: ["https://graph.microsoft.com/.default"],
    };

    console.log("Acquiring token...");
    const authResponse = await clientApp.acquireTokenByClientCredential(
      clientCredentialRequest
    );
    const accessToken = authResponse?.accessToken;

    if (!accessToken) {
      throw new Error("Failed to acquire access token");
    }

    console.log("Token acquired successfully");

    const excelWorkbookId = config.excelItemId!; // In this context, the itemId is the workbook ID
    const driveId = config.excelDriveId!;

    // List available worksheets first
    console.log("\n📋 Checking available worksheets...");
    await listAvailableWorksheets(accessToken, driveId, excelWorkbookId);

    // 2. Make direct axios call instead of using Graph client
    // Example for reading a named table:
    // const excelApiEndpoint = `/drives/${driveId}/items/${excelWorkbookId}/workbook/worksheets/${config.worksheetName}/tables/${config.tableName}/rows`;

    // Alternative for reading the entire used range of a sheet:
    const excelApiEndpoint = `/drives/${driveId}/items/${excelWorkbookId}/workbook/worksheets/${config.worksheetName}/usedRange`;

    // Alternative for reading a specific range:
    // const excelApiEndpoint = `/drives/${driveId}/items/${excelWorkbookId}/workbook/worksheets/${config.worksheetName}/range(address='A1:Z100')`;

    console.log(
      `\n🔍 Attempting to access worksheet: "${config.worksheetName}"`
    );
    console.log("Making API call to:", excelApiEndpoint);

    const response = await axios.get(
      `https://graph.microsoft.com/v1.0${excelApiEndpoint}`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        timeout: 30000, // 30 second timeout
      }
    );

    console.log("response status:", response.status);
    console.log("response data:", response.data);

    const excelData = response.data.value || response.data.values; // For tables, it's .value; for usedRange, it's .values
    console.log("Successfully read Excel Data for Cron Job:", excelData);

    return excelData;
  } catch (err: any) {
    console.error("❌ Error accessing Excel file:");

    if (err.response) {
      // The request was made and the server responded with a status code
      console.error("Status:", err.response.status);
      console.error("Status Text:", err.response.statusText);
      console.error(
        "API Response:",
        JSON.stringify(err.response.data, null, 2)
      );
      console.error("Request URL:", err.config?.url);
    } else if (err.request) {
      // The request was made but no response was received
      console.error("No response received:", err.message);
    } else {
      // Something happened in setting up the request
      console.error("Request setup error:", err.message);
    }

    throw err;
  }
};

const start = async () => {
  try {
    await server.listen({ port: 3000, host: "0.0.0.0" });
    console.log("🚀 Server is running on http://localhost:3000");

    const excelData = await getExcelDataFromSharepoint();
    console.log("Excel Data:", excelData);
  } catch (err) {
    server.log.error(err);
    process.exit(1);
  }
};

start();
