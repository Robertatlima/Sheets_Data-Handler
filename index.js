const fs = require("fs").promises;
const path = require("path");
const process = require("process");
const { authenticate } = require("@google-cloud/local-auth");
const { google } = require("googleapis");

/**Scopes are the access levels that your program will have,
if you want reading level only, add: ...sheets.readonly
*/
const SCOPES = ["https://www.googleapis.com/auth/spreadsheets"];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = path.join(process.cwd(), "token.json");
const CREDENTIALS_PATH = path.join(process.cwd(), "credentials.json");

/**
 * Reads previously authorized credentials from the save file.
 *
 * @return {Promise<OAuth2Client|null>}
 */
async function loadSavedCredentialsIfExist() {
  try {
    const content = await fs.readFile(TOKEN_PATH);
    const credentials = JSON.parse(content);
    return google.auth.fromJSON(credentials);
  } catch (err) {
    return null;
  }
}

/**
 * Serializes credentials to a file comptible with GoogleAUth.fromJSON.
 *
 * @param {OAuth2Client} client
 * @return {Promise<void>}
 */
async function saveCredentials(client) {
  const content = await fs.readFile(CREDENTIALS_PATH);
  const keys = JSON.parse(content);
  const key = keys.installed || keys.web;
  const payload = JSON.stringify({
    type: "authorized_user",
    client_id: key.client_id,
    client_secret: key.client_secret,
    refresh_token: client.credentials.refresh_token,
  });
  await fs.writeFile(TOKEN_PATH, payload);
}

/**
 * Load or request or authorization to call APIs.
 *
 */
async function authorize() {
  let client = await loadSavedCredentialsIfExist();
  if (client) {
    return client;
  }
  client = await authenticate({
    scopes: SCOPES,
    keyfilePath: CREDENTIALS_PATH,
  });
  if (client.credentials) {
    await saveCredentials(client);
  }
  return client;
}

/**
 * Prints the names and majors of students in a sample spreadsheet:
 * @see https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 */

async function calculateSituation(auth) {
  try {
    const sheets = google.sheets({ version: "v4", auth });
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: "1_ingb7t4NRT3T1TqNsjO8Jpwb5KakfMiHGs9AaWQpVQ",
      range: "engenharia_de_software!A2:H27",
    });
    const rows = res.data.values;
    const situationData = [];

    let totalClasses = 0;
    const checkTotalClasses = rows[0].join(",");
    const match = checkTotalClasses.match(/\d+/);

    if (match) {
      totalClasses = parseInt(match[0], 10);
    }

    let absenceThreshold = totalClasses * 0.25;
    const notorFinalApprovalData = [];

    rows.forEach((row, index) => {
      if (index > 1) {
        const calculateNoteForFinalApproval = (media) => {
          return Math.ceil(10 - media);
        };
        let P1 = parseFloat(row[3]) / 10;
        let P2 = parseFloat(row[4]) / 10;
        let P3 = parseFloat(row[5]) / 10;
        const media = (parseFloat(P1 + P2 + P3) / 3).toFixed(2);
        let absenceCount = parseInt(row[2]);
        if (absenceCount > absenceThreshold) {
          situationData.push(["Reprovado por Falta"]);
          notorFinalApprovalData.push([0]);
        } else if (media < 5.0) {
          situationData.push([`Reprovado por Nota`]);
          notorFinalApprovalData.push([0]);
        } else if (media >= 5.0 && media < 7.0) {
          situationData.push([`Exame Final`]);
          notorFinalApprovalData.push([calculateNoteForFinalApproval(media)]);
        } else {
          situationData.push([`Aprovado`]);
          notorFinalApprovalData.push([0]);
        }
      }
    });

    await sheets.spreadsheets.values.update({
      spreadsheetId: "1_ingb7t4NRT3T1TqNsjO8Jpwb5KakfMiHGs9AaWQpVQ",
      range: "G4",
      valueInputOption: "USER_ENTERED",
      resource: { values: situationData },
    });

    await sheets.spreadsheets.values.update({
      spreadsheetId: "1_ingb7t4NRT3T1TqNsjO8Jpwb5KakfMiHGs9AaWQpVQ",
      range: "H4",
      valueInputOption: "USER_ENTERED",
      resource: { values: notorFinalApprovalData },
    });

    console.log("Planilha atualizada com sucesso.");
  } catch (error) {
    console.error("Erro ao atualizar a planilha:", error.message);
  }
}

authorize().then(calculateSituation).catch(console.error);
