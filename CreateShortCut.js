const fs = require('fs').promises;
const path = require('path');
const readline = require('readline');
const { google } = require('googleapis');
const XLSX = require('xlsx');

// --- CONFIGURATION ---
const CONFIG = {
    EXCEL_FILE_PATH: './DCCE2025-AiTaxonomy-Sp.xlsx',
    SHEET_NAME: 'AiTaxonomy-2025',
    FILE_ID_COLUMN: 'File ID',
    FOLDER_ID_COLUMN: 'Folder ID',
    FAILED_ITEMS_OUTPUT_FILE: 'failed_shortcuts_summary.xlsx',
    MAX_RETRIES: 3, // จำนวนครั้งสูงสุดที่จะลองใหม่
    RETRY_DELAY_MS: 1000 // เวลาที่รอ (มิลลิวินาที) ก่อนลองครั้งถัดไป
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
        await processShortcutsConcurrently(auth);
    } catch (err) {
        console.error('❌ เกิดข้อผิดพลาดร้ายแรงระหว่างการทำงาน:', err.message);
    }
}

// ==============================================================================
//  ส่วนของการยืนยันตัวตน (Authentication) - ไม่มีการเปลี่ยนแปลง
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
 * สร้าง Shortcut หนึ่งรายการ (*** แก้ไขแล้ว ***)
 */
async function createSingleShortcut(drive, fileId, folderId) {
    let originalFileName = 'N/A (Failed to fetch name)';
    try {
        // 1. ดึงชื่อของไฟล์ต้นทาง (เพิ่ม supportsAllDrives)
        const fileInfo = await drive.files.get({
            fileId: fileId,
            fields: 'name',
            supportsAllDrives: true // <== เพิ่มบรรทัดนี้
        });
        originalFileName = fileInfo.data.name;

        // 2. Metadata สำหรับสร้าง Shortcut
        const shortcutMetadata = {
            name: originalFileName,
            mimeType: 'application/vnd.google-apps.shortcut',
            shortcutDetails: { targetId: fileId },
            parents: [folderId]
        };

        // 3. เรียก API เพื่อสร้าง Shortcut (เพิ่ม supportsAllDrives)
        return await drive.files.create({
            resource: shortcutMetadata,
            fields: 'id',
            supportsAllDrives: true // <== เพิ่มบรรทัดนี้
        });
    } catch (error) {
        const enhancedError = new Error(error.message);
        enhancedError.fileName = originalFileName;
        enhancedError.fileId = fileId;
        throw enhancedError;
    }
}

/**
 * พยายามสร้าง Shortcut พร้อม Retry Logic
 */
async function createShortcutWithRetries(drive, fileId, folderId, rowIndex) {
    for (let attempt = 1; attempt <= CONFIG.MAX_RETRIES; attempt++) {
        try {
            return await createSingleShortcut(drive, fileId, folderId);
        } catch (error) {
            console.warn(`[แถวที่ ${rowIndex}] ⚠️ พยายามครั้งที่ ${attempt}/${CONFIG.MAX_RETRIES} ล้มเหลว: ${error.message}`);
            if (attempt === CONFIG.MAX_RETRIES) {
                console.error(`[แถวที่ ${rowIndex}] ❌ ล้มเหลวถาวรหลังลองครบ ${CONFIG.MAX_RETRIES} ครั้ง`);
                throw error;
            }
            await new Promise(resolve => setTimeout(resolve, CONFIG.RETRY_DELAY_MS));
        }
    }
}


/**
 * บันทึกรายการที่ล้มเหลวลงในไฟล์ Excel
 */
async function saveFailedItemsToFile(items) {
    if (items.length === 0) {
        console.log('🎉 ไม่มีรายการที่ล้มเหลว ไม่จำเป็นต้องสร้างไฟล์สรุป');
        return;
    }
    const filePath = path.join(__dirname, CONFIG.FAILED_ITEMS_OUTPUT_FILE);
    try {
        console.log(`\n💾 กำลังบันทึก ${items.length} รายการที่ล้มเหลว...`);
        const worksheet = XLSX.utils.json_to_sheet(items);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'FailedItems');
        XLSX.writeFile(workbook, filePath);
        console.log(`✅ รายการที่ล้มเหลวถูกบันทึกไว้ที่: ${filePath}`);
    } catch (error) {
        console.error('❌ เกิดข้อผิดพลาดร้ายแรงขณะบันทึกไฟล์สรุป:', error.message);
    }
}


/**
 * อ่านไฟล์ Excel และสร้าง Google Drive shortcuts
 */
async function processShortcutsConcurrently(auth) {
    const drive = google.drive({ version: 'v3', auth });
    const CHUNK_SIZE = 100;
    let successCount = 0;
    const failedItems = [];

    try {
        console.log(`\n📖 กำลังอ่านไฟล์ Excel: ${CONFIG.EXCEL_FILE_PATH}`);
        const workbook = XLSX.readFile(CONFIG.EXCEL_FILE_PATH);
        const sheet = workbook.Sheets[CONFIG.SHEET_NAME];
        if (!sheet) throw new Error(`ไม่พบ Sheet '${CONFIG.SHEET_NAME}'`);

        const data = XLSX.utils.sheet_to_json(sheet);
        console.log(`🔍 พบข้อมูลทั้งหมด ${data.length} แถว`);

        for (let i = 3550; i < data.length; i += CHUNK_SIZE) {
            const chunk = data.slice(i, i + CHUNK_SIZE);
            const currentChunkNum = Math.floor(i / CHUNK_SIZE) + 1;
            const totalChunks = Math.ceil((data.length - 3550) / CHUNK_SIZE);

            console.log(`\n🔄 กำลังประมวลผลกลุ่มที่ ${currentChunkNum} / ${totalChunks} (แถวที่ ${i + 2} ถึง ${i + 1 + chunk.length})`);

            const promises = chunk.map((row, index) => {
                const rowIndexForLog = i + index + 2;
                const fileId = row[CONFIG.FILE_ID_COLUMN];
                const folderId = row[CONFIG.FOLDER_ID_COLUMN];

                if (!fileId || !folderId) {
                    const reason = `ข้อมูล File ID หรือ Folder ID ไม่ครบ`;
                    return Promise.reject({
                        message: reason,
                        fileName: 'N/A (Missing data)',
                        fileId: fileId || 'N/A',
                        rowIndex: rowIndexForLog
                    });
                }

                return createShortcutWithRetries(drive, fileId, folderId, rowIndexForLog);
            });

            const results = await Promise.allSettled(promises);

            results.forEach((result, index) => {
                const rowIndexForLog = i + index + 2;
                const row = chunk[index];
                if (result.status === 'fulfilled') {
                    successCount++;
                } else {
                    const reason = result.reason;
                    failedItems.push({
                        'File Name': reason.fileName || 'N/A',
                        'File ID': reason.fileId || row[CONFIG.FILE_ID_COLUMN] || 'N/A',
                        'Row in Excel': rowIndexForLog,
                        'Error Message': reason.message
                    });
                }
            });
            const chunkSuccess = results.filter(r => r.status === 'fulfilled').length;
            const chunkFailed = results.filter(r => r.status === 'rejected').length;
            console.log(`👍 กลุ่มที่ ${currentChunkNum} เสร็จสิ้น (สำเร็จ: ${chunkSuccess}, ผิดพลาด: ${chunkFailed})`);
        }

        console.log('\n========================================');
        console.log('✨ การทำงานทั้งหมดเสร็จสิ้น ✨');
        console.log(`- สร้าง Shortcut สำเร็จ: ${successCount} รายการ`);
        console.log(`- เกิดข้อผิดพลาด/ข้ามไป: ${failedItems.length} รายการ`);
        console.log('========================================');

        await saveFailedItemsToFile(failedItems);

    } catch (error) {
        console.error('❌ เกิดข้อผิดพลาดในฟังก์ชัน processShortcutsConcurrently:', error.message);
        await saveFailedItemsToFile(failedItems);
    }
}