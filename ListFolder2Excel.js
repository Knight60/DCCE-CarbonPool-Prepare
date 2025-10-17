const fs = require('fs').promises;
const readline = require('readline');
const path = require('path');
const { google } = require('googleapis');
const xl = require('excel4node');

// --- CONFIGURATION ---
const outName = 'DCCE-SpeciesList.xlsx'; // Output Excel file name
const ROOT_FOLDER_ID = '17DsEtWg2aYlGxo2v8xp9HMhJtXx1E7xx'; // The starting folder ID

// --- GOOGLE API SETUP ---
const SCOPES = ['https://www.googleapis.com/auth/drive.metadata.readonly'];
const CREDENTIALS_PATH = path.join(__dirname, 'DCCE-CarbonPool-Credential.json');
const TOKEN_PATH = path.join(__dirname, 'DCCE-CarbonPool-Token.json');

// --- SCRIPT START ---
main();

async function main() {
    try {
        const credentials = JSON.parse(await fs.readFile(CREDENTIALS_PATH));
        const auth = await authorize(credentials);
        await startProcessing(auth);
    } catch (err) {
        console.error('Error during script execution:', err);
    }
}

async function authorize(credentials) {
    const { client_secret, client_id, redirect_uris } = credentials.installed;
    const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

    try {
        const token = await fs.readFile(TOKEN_PATH);
        oAuth2Client.setCredentials(JSON.parse(token));
        return oAuth2Client;
    } catch (err) {
        console.log('Token not found, generating a new one...');
        return await getNewToken(oAuth2Client);
    }
}

function getNewToken(oAuth2Client) {
    return new Promise((resolve, reject) => {
        const authUrl = oAuth2Client.generateAuthUrl({ access_type: 'offline', scope: SCOPES });
        console.log('Authorize this app by visiting this url:', authUrl);
        const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

        rl.question('Enter the code from that page here: ', (code) => {
            rl.close();
            oAuth2Client.getToken(code, (err, token) => {
                if (err) return reject(new Error('Error retrieving access token', err));
                oAuth2Client.setCredentials(token);
                fs.writeFile(TOKEN_PATH, JSON.stringify(token));
                console.log('Token stored to', TOKEN_PATH);
                resolve(oAuth2Client);
            });
        });
    });
}

/**
 * Main function to start scanning Drive and creating the Excel file.
 * @param {google.auth.OAuth2} auth An authorized OAuth2 client.
 */
async function startProcessing(auth) {
    const drive = google.drive({ version: 'v3', auth });
    const wb = new xl.Workbook();
    const ws = wb.addWorksheet('FileList');

    // Setup Excel Headers
    const outColumns = ['Folder', 'ID'];
    outColumns.forEach((header, index) => {
        ws.cell(1, index + 1).string(header).style({ font: { bold: true } });
    });

    let currentRow = 2; // Start writing data from row 2

    console.log(`🚀 Starting to scan Google Drive folder: ${ROOT_FOLDER_ID}`);

    const context = { ws, currentRow };

    await traverseDriveFolder(drive, context, ROOT_FOLDER_ID, '');

    await wb.write(outName);
    console.log(`✅ Success! Excel file created: ${outName}`);
}

/**
 * Recursively traverses a Google Drive folder and writes image file info to the worksheet.
 * @param {google.drive_v3.Drive} drive The authenticated Drive API client.
 * @param {object} context Contains worksheet (ws) and currentRow.
 * @param {string} folderId The ID of the folder to scan.
 * @param {string} currentPath The file path built so far.
 */
async function traverseDriveFolder(drive, context, folderId, currentPath) {
    let pageToken = null;
    do {
        const res = await drive.files.list({
            q: `'${folderId}' in parents and trashed = false`,
            fields: 'nextPageToken, files(id, name, mimeType, size)',
            orderBy: 'name', // Sorts files and folders alphabetically
            pageToken: pageToken,
            pageSize: 1000,
        });

        const files = res.data.files;
        if (files && files.length > 0) {
            // Separate folders and files to process folders first
            const folders = files.filter(file => file.mimeType === 'application/vnd.google-apps.folder');
            //const imageFiles = files.filter(file => isImageFile(file.mimeType));

            // 1. Recursively process all subfolders first
            for (const folder of folders) {
                context.ws.cell(context.currentRow, 1).string(folder.name);
                context.ws.cell(context.currentRow, 2).string(folder.id);
                context.currentRow++;

                const newPath = currentPath ? path.join(currentPath, folder.name).replace(/\\/g, '/') : folder.name;
                console.log(`📂 Entering folder: ${newPath}`);
                await traverseDriveFolder(drive, context, folder.id, newPath);
            }

            /*
            // 2. Process all files in the current folder
            for (const file of imageFiles) {
                // Correctly access 'ws' via the 'context' object
                context.ws.cell(context.currentRow, 1).string(currentPath || '/');
                context.ws.cell(context.currentRow, 2).string(file.name);
                context.ws.cell(context.currentRow, 3).string(file.size);
                context.ws.cell(context.currentRow, 4).string(file.id);
                context.currentRow++;
            
            }
            */
        }
        pageToken = res.data.nextPageToken;
    }
    while (pageToken);
}

/**
 * Checks if a MIME type corresponds to a common image format.
 * @param {string} mimeType The MIME type of the file.
 * @returns {boolean}
 */
function isImageFile(mimeType) {
    const imageTypes = [
        'image/jpeg', 'image/png', 'image/gif', 'image/bmp', 'image/svg+xml',
        'image/webp', 'image/tiff', 'image/heif', 'image/heic'
    ];
    return imageTypes.includes(mimeType);
}