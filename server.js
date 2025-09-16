// server.js
const express = require("express");
const axios = require("axios");
const app = express();
app.use(express.json());

const tenantId = 06f0ceb1-3e11-4804-b0b7-2b70b7fb6d69;
const clientId = b8d1788a-2d6b-4ee8-b1da-0955df3839a9;
const clientSecret = 9tC8Q~56SJ0dgF73~2bERsCrpFXUyuwBMZ4d2aE1;

app.post("/convert-docx-to-pdf", async (req, res) => {
  const { itemId } = req.body;
  if (!itemId) return res.status(400).send("Missing itemId");

  try {
    // 1. Get access token
    const tokenResponse = await axios.post(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        scope: "https://graph.microsoft.com/.default",
        grant_type: "client_credentials",
      }),
      { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
    );

    const accessToken = tokenResponse.data.access_token;

    // 2. Get the file as PDF
    const pdfResponse = await axios.get(
      `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/content`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/pdf",
        },
        responseType: "arraybuffer", // so we get the binary PDF
      }
    );

    res.setHeader("Content-Type", "application/pdf");
    res.send(pdfResponse.data);
  } catch (err) {
    console.error(err.response?.data || err.message);
    res.status(500).send("Something went wrong");
  }
});

app.listen(3000, () => console.log("Server running on port 3000"));
