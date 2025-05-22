const express = require("express");
const bodyParser = require("body-parser");
const axios = require("axios");
const getAccessToken = require("./auth");
require("dotenv").config();

const app = express();
app.use(bodyParser.json());

// Load environment variables (Render sets these automatically)
const tenantId = process.env.TENANT_ID;
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const siteId = process.env.SITE_ID;
const listId = process.env.LIST_ID;

// Health check endpoint
app.get("/", (req, res) => {
  res.send("✅ Outlook Adaptive Card handler is running!");
});

// Endpoint to receive Adaptive Card responses
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
    // Get access token from your auth module
    const token = await getAccessToken(tenantId, clientId, clientSecret);
console.log(token);
    if (!token) {
      console.error("❌ Failed to get access token");
      return res.status(500).send({
        type: "MessageCard",
        text: `❌ Failed to get access token`
      });
    }

    const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${ID}/fields`;

    const updatePayload = {
      ApprovalStatus: "ApprovalStatus",
      Comments: "Comments" || ""
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

// Listen on the port provided by Render or fallback to 3000 locally
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`✅ Server is running on port ${PORT}`);
});
