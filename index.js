const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const axios = require("axios");
const getAccessToken = require("./auth");
require("dotenv").config();

const app = express();

// Enable CORS for your SharePoint domain
app.use(cors({
  origin: "https://sachagroup.sharepoint.com"
}));

app.use(bodyParser.json());

// Load environment variables
const tenantId = process.env.TENANT_ID;
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const siteId = process.env.SITE_ID;
const listId = process.env.LIST_ID;

// Health check endpoint
app.get("/", (req, res) => {
  res.send("✅ Outlook Adaptive Card handler is running!");
});

// Endpoint to receive Adaptive Card responses (from Outlook emails)
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
    console.log("TENANT_ID:", process.env.TENANT_ID);
    console.log("CLIENT_ID:", process.env.CLIENT_ID);
    console.log("CLIENT_SECRET:", process.env.CLIENT_SECRET ? "set" : "MISSING");

    const token = await getAccessToken(
  tenantId,
  clientId,
  clientSecret
);
    // Get access token from your auth module
    //const token = await getAccessToken(); // Make sure this function works as expected
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
    console.error("❌ Error updating SharePoint:", error.response?.data || error.message);
    res.status(500).send({
      type: "MessageCard",
      text: `❌ Error updating SharePoint: ${error.response?.data?.error?.message || error.message}`
    });
  }
});

// Endpoint to send Adaptive Card to Teams via webhook
app.post("/send-teams-card", async (req, res) => {
  const webhookUrl = process.env.TEAMS_WEBHOOK_URL; // Store in your Render env vars!
  const { card } = req.body;

  if (!card) {
    return res.status(400).send("Missing card");
  }

  try {
    const teamsPayload = {
      type: "message",
      attachments: [
        {
          contentType: "application/vnd.microsoft.card.adaptive",
          content: card
        }
      ]
    };

    const response = await axios.post(webhookUrl, teamsPayload, {
      headers: { "Content-Type": "application/json" }
    });

    console.log("✅ Card sent to Teams:", response.status);
    res.status(200).send("Card sent to Teams!");
  } catch (error) {
    console.error("❌ Error sending card to Teams:", error.response?.data || error.message);
    res.status(500).send(
      error.response?.data?.error?.message || error.message || "Failed to send to Teams"
    );
  }
});

// Start the server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`✅ Server is running on port ${PORT}`);
});