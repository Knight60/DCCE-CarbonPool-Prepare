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
    FOLDER_ID_COLUMN: 'Folder ID'
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
        await processShortcutsConcurrently(auth); // <== เรียกใช้ฟังก์ชันใหม่
    } catch (err) {
        console.error('❌ เกิดข้อผิดพลาดร้ายแรงระหว่างการทำงาน:', err.message);
    }
}

// ==============================================================================
//  ส่วนของการยืนยันตัวตน (Authentication) - คงไว้เหมือนเดิม
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
//  ส่วนหลักของโปรแกรม (Core Logic) - แก้ไขใหม่เพื่อใช้ Concurrent Requests
// ==============================================================================

/**
 * สร้าง Shortcut หนึ่งรายการ (เป็นฟังก์ชัน helper)
 * @param {google.drive_v3.Drive} drive The authenticated Drive API client.
 * @param {string} fileId ID ของไฟล์ต้นทาง
 * @param {string} folderId ID ของโฟลเดอร์ปลายทาง
 */
async function createSingleShortcut(drive, fileId, folderId) {
    // 1. ดึงชื่อของไฟล์ต้นทาง
    const fileInfo = await drive.files.get({
        fileId: fileId,
        fields: 'name'
    });
    const originalFileName = fileInfo.data.name;

    // 2. Metadata สำหรับสร้าง Shortcut
    const shortcutMetadata = {
        name: originalFileName,
        mimeType: 'application/vnd.google-apps.shortcut',
        shortcutDetails: {
            targetId: fileId
        },
        parents: [folderId]
    };

    // 3. เรียก API เพื่อสร้าง Shortcut
    return drive.files.create({
        resource: shortcutMetadata,
        fields: 'id'
    });
}


/**
 * อ่านไฟล์ Excel และสร้าง Google Drive shortcuts แบบ Concurrent (พร้อมกันทีละกลุ่ม)
 * @param {google.auth.OAuth2} auth An authorized OAuth2 client.
 */
async function processShortcutsConcurrently(auth) {
    const drive = google.drive({ version: 'v3', auth });
    // การส่ง request พร้อมกันเยอะเกินไปอาจเจอปัญหา Rate Limit ของ Google
    // แนะนำให้เริ่มจากค่าน้อยๆ เช่น 20-50 แล้วค่อยๆ ปรับเพิ่ม
    const CHUNK_SIZE = 100;
    let successCount = 0;
    let errorCount = 0;

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
            const totalChunks = Math.ceil(data.length / CHUNK_SIZE);

            console.log(`\n🔄 กำลังประมวลผลกลุ่มที่ ${currentChunkNum} / ${totalChunks} (แถวที่ ${i + 1} ถึง ${i + chunk.length}) จาก ${data.length}`);

            // สร้าง Array ของ Promises สำหรับทุกรายการใน chunk นี้
            const promises = chunk.map((row, index) => {
                const rowIndexForLog = i + index + 2; // +2 เพื่อให้ตรงกับเลขแถวใน Excel
                const fileId = row[CONFIG.FILE_ID_COLUMN];
                const folderId = row[CONFIG.FOLDER_ID_COLUMN];

                if (!fileId || !folderId) {
                    console.warn(`[แถวที่ ${rowIndexForLog}] ⚠️ ข้อมูลไม่ครบ ข้าม...`);
                    // คืนค่า Promise ที่ reject ทันทีเพื่อให้นับเป็น error
                    return Promise.reject(new Error(`Missing File ID or Folder ID at row ${rowIndexForLog}`));
                }

                return createSingleShortcut(drive, fileId, folderId);
            });

            // รอให้ทุก Promises ใน chunk นี้ทำงานเสร็จสิ้น (ไม่ว่าจะสำเร็จหรือล้มเหลว)
            const results = await Promise.allSettled(promises);

            // ตรวจสอบผลลัพธ์
            results.forEach((result, index) => {
                const rowIndexForLog = i + index + 2;
                if (result.status === 'fulfilled') {
                    successCount++;
                    // console.log(`[แถวที่ ${rowIndexForLog}] ✅ สำเร็จ`);
                } else {
                    errorCount++;
                    // แสดงเฉพาะ error ที่เกิดขึ้นจริงจากการเรียก API
                    if (result.reason.message.includes('Missing') === false) {
                        console.error(`[แถวที่ ${rowIndexForLog}] ❌ ล้มเหลว: ${result.reason.message}`);
                    }
                }
            });
            console.log(`👍 กลุ่มที่ ${currentChunkNum} ประมวลผลเสร็จสิ้น (สำเร็จ: ${results.filter(r => r.status === 'fulfilled').length}, ผิดพลาด: ${results.filter(r => r.status === 'rejected').length})`);
        }

        console.log('\n========================================');
        console.log('✨ การทำงานเสร็จสิ้น ✨');
        console.log(`- สร้าง Shortcut สำเร็จ: ${successCount} รายการ`);
        console.log(`- เกิดข้อผิดพลาด/ข้ามไป: ${errorCount} รายการ`);
        console.log('========================================');

    } catch (error) {
        console.error('❌ เกิดข้อผิดพลาดในฟังก์ชัน processShortcutsConcurrently:', error.message);
    }
}