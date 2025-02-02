const { google } = require("googleapis");

/**
 * Lists the labels in the user's account.
 *
 * @param {google.auth.OAuth2} auth An authorized OAuth2 client.
 */
async function listLabels(auth) {
  const gmail = google.gmail({ version: "v1", auth });
  const res = await gmail.users.labels.list({
    userId: "me",
  });
  const labels = res.data.labels;
  if (!labels || labels.length === 0) {
    console.log("No labels found.");
    return;
  }

  return labels;
}

async function listMessages(auth) {
  const gmail = google.gmail({ version: "v1", auth });
  const res = await gmail.users.messages.list({
    userId: "me",
    maxResults: 1,
  });

  let latestMessageId = res.data.messages[0].id;

  const message = await gmail.users.messages.get({
    userId: "me",
    id: latestMessageId,
  });

  const body = JSON.stringify(message.data.payload.body.data);

  let mailBody = new Buffer.from(body, "base64").toString();

  return mailBody;
}

async function sendEmail(auth, content) {
  const gmail = google.gmail({ version: "v1", auth });
  const encodedMessage = Buffer.from(content)
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
  const res = await gmail.users.messages.send({
    userId: "me",
    requestBody: {
      raw: encodedMessage,
    },
  });

  return res.data;
}

module.exports = {
  listLabels: listLabels,
  listMessages: listMessages,
  sendEmail: sendEmail,
};
