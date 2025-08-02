import "dotenv/config";
import Fastify from "fastify";
import { getGoogleSheetsData } from "./google-sheets-reader";
// Note: You would need to export the function from server.ts to import it here
// import { getExcelDataFromSharepoint } from "./server";

const server = Fastify({ logger: true });

// Home route
server.get("/", async () => {
  return {
    message: "Integrated Server - Microsoft Excel & Google Sheets Reader",
    endpoints: [
      "GET /microsoft-excel - Read data from Microsoft SharePoint Excel",
      "GET /google-sheets - Read data from Google Sheets",
      "GET /compare - Compare data from both sources",
    ],
  };
});

// Microsoft Excel endpoint
server.get("/microsoft-excel", async (request, reply) => {
  try {
    // Note: This would require exporting the function from server.ts
    // const data = await getExcelDataFromSharepoint();

    // For now, return a placeholder
    return {
      error: "Microsoft Excel integration not available",
      message:
        "Please export getExcelDataFromSharepoint function from server.ts to use this endpoint",
    };
  } catch (error: any) {
    reply.status(500);
    return {
      error: "Failed to read Microsoft Excel data",
      message: error.message,
    };
  }
});

// Google Sheets endpoint
server.get("/google-sheets", async (request, reply) => {
  try {
    const data = await getGoogleSheetsData();
    return {
      source: "google-sheets",
      message: "Google Sheets data retrieved successfully!",
      totalRecords: data.length,
      data,
    };
  } catch (error: any) {
    reply.status(500);
    return {
      error: "Failed to read Google Sheets data",
      message: error.message,
    };
  }
});

// Compare data from both sources
server.get("/compare", async (request, reply) => {
  try {
    const results: any = {
      comparison: "Data comparison between Microsoft Excel and Google Sheets",
      timestamp: new Date().toISOString(),
    };

    // Get Google Sheets data
    try {
      const googleData = await getGoogleSheetsData();
      results.googleSheets = {
        status: "success",
        totalRecords: googleData.length,
        sampleRecord: googleData[0] || null,
        headers: googleData.length > 0 ? Object.keys(googleData[0]) : [],
      };
    } catch (googleError: any) {
      results.googleSheets = {
        status: "error",
        message: googleError.message,
      };
    }

    // Get Microsoft Excel data
    try {
      // Note: This would require exporting the function from server.ts
      // const excelData = await getExcelDataFromSharepoint();
      results.microsoftExcel = {
        status: "not_available",
        message: "Microsoft Excel integration not configured for this endpoint",
      };
    } catch (excelError: any) {
      results.microsoftExcel = {
        status: "error",
        message: excelError.message,
      };
    }

    return results;
  } catch (error: any) {
    reply.status(500);
    return {
      error: "Failed to compare data sources",
      message: error.message,
    };
  }
});

// Health check endpoint
server.get("/health", async () => {
  return {
    status: "healthy",
    timestamp: new Date().toISOString(),
    services: {
      "google-sheets": "available",
      "microsoft-excel": "requires configuration",
    },
  };
});

// Start the server
const start = async () => {
  try {
    await server.listen({ port: 3002, host: "0.0.0.0" });
    console.log("🚀 Integrated server is running on http://localhost:3002");
    console.log("📋 Available endpoints:");
    console.log("  - GET / (home)");
    console.log("  - GET /google-sheets (Google Sheets data)");
    console.log("  - GET /microsoft-excel (Microsoft Excel data)");
    console.log("  - GET /compare (Compare both sources)");
    console.log("  - GET /health (Health check)");
  } catch (err) {
    server.log.error(err);
    process.exit(1);
  }
};

// Start the server if this file is run directly
if (require.main === module) {
  start();
}

export { server };
