const express = require("express");
const bodyParser = require("body-parser");
const axios = require("axios");
const getAccessToken = require("./auth");

// Load .env only if NOT in production
if (process.env.NODE_ENV !== "production") {
  require("dotenv").config();
}

const app = express();
app.use(bodyParser.json());

// Load environment variables
const tenantId = process.env.TENANT_ID;
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const siteId = process.env.SITE_ID;
const listId = process.env.LIST_ID;

// Validate required environment variables
if (!tenantId || !clientId || !clientSecret || !siteId || !listId) {
  console.error("❌ Missing one or more required environment variables.");
  process.exit(1);
}

// Root endpoint (health check)
app.get("/", (req, res) => {
  res.send("✅ Outlook Adaptive Card handler is running!");
});

// Main endpoint to receive responses
app.post("/response", async (req, res) => {
  const { ApprovalStatus, Comments, ID } = req.body;

  console.log("🔄 Received Adaptive Card response payload:", req.body);

  if (!ApprovalStatus || !ID) {
    console.error("❌ Missing required fields: ApprovalStatus or ID");
    return res.status(400).send({
      type: "MessageCard",
      text: `❌ Missing required fields: ApprovalStatus or ID`
    });
  }

  try {
    const token = await getAccessToken(tenantId, clientId, clientSecret);
    const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${ID}/fields`;

    const updatePayload = {
      ApprovalStatus: ApprovalStatus,
      Comments: Comments || ""
    };

    console.log("📡 PATCH request to:", url);
    console.log("📝 Payload:", updatePayload);

    const response = await axios.patch(url, updatePayload, {
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      }
    });

    console.log("✅ SharePoint item updated:", response.status);

    res.status(200).send({
      type: "MessageCard",
      text: `✅ SharePoint item updated with status: ${ApprovalStatus}`
    });

  } catch (error) {
    // Log detailed error info
    console.error("❌ Error updating SharePoint:", error.message);
    if (error.response) {
      console.error("Response data:", JSON.stringify(error.response.data, null, 2));
      console.error("Response status:", error.response.status);
      console.error("Response headers:", error.response.headers);
    }
    console.error("Stack trace:", error.stack);

    res.status(500).send({
      type: "MessageCard",
      text: `❌ Error updating SharePoint: ${error.response?.data?.error?.message || error.message}`
    });
  }
});
