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

// SharePoint configuration - extracted from SharePoint URL
const SHAREPOINT_SITE_URL = "gadocerto902.sharepoint.com";
const DOCUMENT_ID = "ab9a6aeb-9387-455f-9be5-9e8be5d4c749"; // From sourcedoc parameter
const WORK_SHEET_NAME = "historico";
const TABLE_NAME = "Table1";

const config = {
  clientId: CLIENT_ID,
  clientSecret: CLIENT_SECRET,
  tenantId: TENANT_ID,
  sharepointSiteUrl: SHAREPOINT_SITE_URL,
  documentId: DOCUMENT_ID,
  worksheetName: WORK_SHEET_NAME,
  tableName: TABLE_NAME,
};

// Get SharePoint site information
const getSharePointSite = async (accessToken: string, siteUrl: string) => {
  try {
    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${siteUrl}`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
      }
    );

    console.log("📍 SharePoint Site Info:", {
      id: response.data.id,
      name: response.data.displayName,
      webUrl: response.data.webUrl,
    });

    return response.data;
  } catch (err: any) {
    console.error(
      "❌ Error getting SharePoint site:",
      err.response?.data || err.message
    );
    throw err;
  }
};

// Find the document by ID in SharePoint
const findDocumentInSharePoint = async (
  accessToken: string,
  siteId: string,
  documentId: string
) => {
  try {
    console.log("🔍 Searching for document with ID:", documentId);

    // Search across all drives in the site
    const drivesResponse = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
      }
    );

    console.log(
      "📁 Available drives:",
      drivesResponse.data.value.map((d: any) => ({
        id: d.id,
        name: d.name,
        driveType: d.driveType,
      }))
    );

    // Try to find the document in each drive
    for (const drive of drivesResponse.data.value) {
      try {
        console.log(`\n🔍 Searching in drive: "${drive.name}"`);

        // Method 1: Try to list all files in the root first
        try {
          const rootFilesResponse = await axios.get(
            `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${drive.id}/root/children`,
            {
              headers: {
                Authorization: `Bearer ${accessToken}`,
                "Content-Type": "application/json",
              },
            }
          );

          console.log(
            `📄 Files in root of "${drive.name}":`,
            rootFilesResponse.data.value.map((f: any) => ({
              id: f.id,
              name: f.name,
              size: f.size,
              webUrl: f.webUrl,
            }))
          );

          // Look for Excel files
          const excelFiles = rootFilesResponse.data.value.filter(
            (item: any) =>
              item.name?.toLowerCase().endsWith(".xlsx") ||
              item.name?.toLowerCase().endsWith(".xls")
          );

          if (excelFiles.length > 0) {
            console.log(
              "📊 Found Excel files:",
              excelFiles.map((f: any) => f.name)
            );

            // Try to find our target document
            const targetDoc = excelFiles.find(
              (item: any) =>
                item.id.includes(documentId) ||
                item.name.toLowerCase().includes("historico") ||
                item.webUrl?.includes(documentId)
            );

            if (targetDoc) {
              console.log("✅ Found target document in root:", {
                id: targetDoc.id,
                name: targetDoc.name,
                driveId: drive.id,
                webUrl: targetDoc.webUrl,
              });

              return {
                driveId: drive.id,
                itemId: targetDoc.id,
                document: targetDoc,
              };
            }
          }
        } catch (rootErr: any) {
          console.log(
            `⚠️  Could not list root files in "${drive.name}":`,
            rootErr.response?.status
          );
        }

        // Method 2: Try a simpler search without wildcards
        try {
          const searchResponse = await axios.get(
            `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${drive.id}/root/search(q='historico')`,
            {
              headers: {
                Authorization: `Bearer ${accessToken}`,
                "Content-Type": "application/json",
              },
            }
          );

          console.log(
            `📄 Search results for "historico" in "${drive.name}":`,
            searchResponse.data.value.map((f: any) => ({
              id: f.id,
              name: f.name,
              webUrl: f.webUrl,
            }))
          );

          // Look for the document by ID or name pattern
          const targetDoc = searchResponse.data.value.find(
            (item: any) =>
              item.id.includes(documentId) ||
              item.name.toLowerCase().includes("historico") ||
              item.webUrl?.includes(documentId)
          );

          if (targetDoc) {
            console.log("✅ Found target document via search:", {
              id: targetDoc.id,
              name: targetDoc.name,
              driveId: drive.id,
              webUrl: targetDoc.webUrl,
            });

            return {
              driveId: drive.id,
              itemId: targetDoc.id,
              document: targetDoc,
            };
          }
        } catch (searchErr: any) {
          console.log(
            `⚠️  Search failed in drive "${drive.name}":`,
            searchErr.response?.status,
            searchErr.response?.data?.error?.message
          );
        }

        // Method 3: Try to access the document directly using the GUID
        try {
          console.log(
            `🎯 Trying direct access with document ID: ${documentId}`
          );
          const directResponse = await axios.get(
            `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${drive.id}/items/${documentId}`,
            {
              headers: {
                Authorization: `Bearer ${accessToken}`,
                "Content-Type": "application/json",
              },
            }
          );

          console.log("✅ Found document via direct access:", {
            id: directResponse.data.id,
            name: directResponse.data.name,
            driveId: drive.id,
            webUrl: directResponse.data.webUrl,
          });

          return {
            driveId: drive.id,
            itemId: directResponse.data.id,
            document: directResponse.data,
          };
        } catch (directErr: any) {
          console.log(
            `⚠️  Direct access failed in "${drive.name}":`,
            directErr.response?.status
          );
        }
      } catch (driveErr: any) {
        console.log(
          `⚠️  Could not access drive "${drive.name}":`,
          driveErr.response?.status
        );
      }
    }

    throw new Error("Document not found in any drive");
  } catch (err: any) {
    console.error(
      "❌ Error finding document:",
      err.response?.data || err.message
    );
    throw err;
  }
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

    // 2. Get SharePoint site information
    console.log("\n🌐 Getting SharePoint site information...");
    const siteInfo = await getSharePointSite(
      accessToken,
      config.sharepointSiteUrl!
    );

    // 3. Find the document in SharePoint
    console.log("\n🔍 Finding the Excel document...");
    const documentInfo = await findDocumentInSharePoint(
      accessToken,
      siteInfo.id,
      config.documentId!
    );

    const { driveId, itemId } = documentInfo;

    // 4. List available worksheets first
    console.log("\n📋 Checking available worksheets...");
    await listAvailableWorksheets(accessToken, driveId, itemId);

    // 5. Read the Excel data
    const excelApiEndpoint = `/drives/${driveId}/items/${itemId}/workbook/worksheets/${config.worksheetName}/usedRange`;

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
