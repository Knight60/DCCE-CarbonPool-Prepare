// --- เพิ่ม Library ที่จำเป็น ---
const fs = require('fs').promises;
const path = require('path');
const readline = require('readline');
const { google } = require('googleapis');
const XLSX = require('xlsx');
const pLimit = require('p-limit').default; // << FIX: แก้ไขการ import
const retry = require('async-retry');

// --- CONFIGURATION ---
const CONFIG = {
    EXCEL_FILE_PATH: './DCCE2025-AiTaxonomy-Sp.xlsx',
    SHEET_NAME: 'AiTaxonomy-2025',
    FILE_ID_COLUMN: 'File ID',
    FOLDER_ID_COLUMN: 'Folder ID',
    FILE_NAME_COLUMN: 'File Name'
};

// --- GOOGLE API SETUP ---
const SCOPES = ['https://www.googleapis.com/auth/drive'];
const CREDENTIALS_PATH = path.join(__dirname, 'DCCE-CarbonPool-Credential.json');
const TOKEN_PATH = path.join(__dirname, 'DCCE-CarbonPool-Token.json');

// --- SCRIPT START ---
main();

async function main() {
    try {
        console.log('🔄 กำลังยืนยันตัวตนกับ Google...');
        const credentials = JSON.parse(await fs.readFile(CREDENTIALS_PATH));
        const auth = await authorize(credentials);
        console.log('✅ ยืนยันตัวตนสำเร็จ');
        await processShortcutsWithLimitAndRetry(auth);
    } catch (err) {
        console.error('❌ เกิดข้อผิดพลาดร้ายแรงระหว่างการทำงาน:', err.message);
    }
}

// ==============================================================================
//  ส่วนของการยืนยันตัวตน (Authentication)
// ==============================================================================
async function authorize(credentials) {
    const { client_secret, client_id, redirect_uris } = credentials.installed;
    const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
    try {
        const token = await fs.readFile(TOKEN_PATH);
        oAuth2Client.setCredentials(JSON.parse(token));
        return oAuth2Client;
    } catch (err) {
        console.log('⚠️ ไม่พบ Token, กำลังสร้าง Token ใหม่...');
        return await getNewToken(oAuth2Client);
    }
}

function getNewToken(oAuth2Client) {
    return new Promise((resolve, reject) => {
        const authUrl = oAuth2Client.generateAuthUrl({ access_type: 'offline', scope: SCOPES });
        console.log('กรุณาไปที่ URL นี้เพื่อทำการยืนยันตัวตน:', authUrl);
        const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
        rl.question('นำโค้ดที่ได้จากหน้านั้นมาวางที่นี่: ', (code) => {
            rl.close();
            oAuth2Client.getToken(code, (err, token) => {
                if (err) return reject(new Error('เกิดข้อผิดพลาดในการดึง Access Token', err));
                oAuth2Client.setCredentials(token);
                fs.writeFile(TOKEN_PATH, JSON.stringify(token));
                console.log('Token ถูกบันทึกไว้ที่', TOKEN_PATH);
                resolve(oAuth2Client);
            });
        });
    });
}

// ==============================================================================
//  ส่วนหลักของโปรแกรม (Core Logic)
// ==============================================================================

/**
 * สร้าง Shortcut หนึ่งรายการ พร้อมกลไก Retry อัตโนมัติ
 */
async function createSingleShortcutWithRetry(drive, fileId, folderId, fileName, rowIndexForLog) {
    return retry(
        async () => {
            const shortcutMetadata = {
                name: fileName,
                mimeType: 'application/vnd.google-apps.shortcut',
                shortcutDetails: { targetId: fileId },
                parents: [folderId]
            };
            return drive.files.create({
                resource: shortcutMetadata,
                fields: 'id, name'
            });
        }, {
        retries: 5,
        factor: 2,
        minTimeout: 1000,
        onRetry: (error, attempt) => {
            console.warn(`[แถวที่ ${rowIndexForLog}] ⚠️ (ครั้งที่ ${attempt}) เกิดข้อผิดพลาดในการสร้าง '${fileName}', กำลังลองใหม่...`);
        }
    }
    );
}

/**
 * อ่าน Excel, ตรวจสอบไฟล์ซ้ำ, และสร้าง Shortcuts โดยใช้ p-limit และ async-retry
 */
async function processShortcutsWithLimitAndRetry(auth) {
    // << FIX: ตั้งค่า auth client ให้เป็น default สำหรับทุก request
    google.options({ auth: auth });

    // สร้าง drive client โดยไม่ต้องส่ง auth ซ้ำ
    const drive = google.drive({ version: 'v3' });

    const limit = pLimit(100);

    let successCount = 0;
    let errorCount = 0;
    let missingDataCount = 0;
    let alreadyExistsCount = 0;

    try {
        console.log(`\n📖 กำลังอ่านไฟล์ Excel: ${CONFIG.EXCEL_FILE_PATH}`);
        const workbook = XLSX.readFile(CONFIG.EXCEL_FILE_PATH);
        const sheet = workbook.Sheets[CONFIG.SHEET_NAME];
        if (!sheet) throw new Error(`ไม่พบ Sheet '${CONFIG.SHEET_NAME}'`);

        const data = XLSX.utils.sheet_to_json(sheet);
        console.log(`🔍 พบข้อมูลทั้งหมด ${data.length} แถว`);

        const promises = data.map((row, index) => {
            const rowIndexForLog = index + 2;
            const fileId = row[CONFIG.FILE_ID_COLUMN];
            const folderId = row[CONFIG.FOLDER_ID_COLUMN];
            const fileName = row[CONFIG.FILE_NAME_COLUMN];

            if (!fileId || !folderId || !fileName) {
                missingDataCount++;
                return Promise.resolve({ status: 'skipped_missing_data' });
            }

            return limit(async () => {
                try {
                    // ตรวจสอบไฟล์ซ้ำ
                    const query = `'${folderId}' in parents and name = '${fileName.replace(/'/g, "\\'")}' and trashed = false`;
                    const existingFiles = await drive.files.list({
                        q: query,
                        fields: 'files(id)',
                        pageSize: 1
                    });

                    if (existingFiles.data.files && existingFiles.data.files.length > 0) {
                        alreadyExistsCount++;
                        return { status: 'skipped_exists' };
                    }

                    // ถ้าไม่พบไฟล์ซ้ำ ให้สร้าง shortcut
                    return createSingleShortcutWithRetry(drive, fileId, folderId, fileName, rowIndexForLog);

                } catch (error) {
                    throw new Error(`[แถวที่ ${rowIndexForLog}] Error during existence check for '${fileName}': ${error.message}`);
                }
            });
        });

        console.log(`\n🔄 เริ่มประมวลผล ${promises.length} รายการ (จำกัดการทำงานพร้อมกัน ${limit.concurrency} รายการ)...`);

        const results = await Promise.allSettled(promises);

        results.forEach((result, index) => {
            if (result.status === 'fulfilled') {
                if (result.value.status !== 'skipped_missing_data' && result.value.status !== 'skipped_exists') {
                    successCount++;
                }
            } else {
                errorCount++;
                const rowIndexForLog = index + 2;
                console.error(`[แถวที่ ${rowIndexForLog}] ❌ ล้มเหลวถาวร: ${result.reason.message}`);
            }
        });

        console.log('\n========================================');
        console.log('✨ การทำงานเสร็จสิ้น ✨');
        console.log(`- ✅ สร้าง Shortcut สำเร็จ: ${successCount} รายการ`);
        console.log(`- ⏩ ข้าม (มีไฟล์อยู่แล้ว): ${alreadyExistsCount} รายการ`);
        console.log(`- ⚠️ ข้าม (ข้อมูลไม่ครบ): ${missingDataCount} รายการ`);
        console.log(`- ❌ เกิดข้อผิดพลาด: ${errorCount} รายการ`);
        console.log('========================================');

    } catch (error) {
        console.error('❌ เกิดข้อผิดพลาดในฟังก์ชัน processShortcutsWithLimitAndRetry:', error.message);
    }
}