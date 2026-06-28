const { Client, LocalAuth } = require('whatsapp-web.js');
const cron = require('node-cron');
const qrcode = require('qrcode-terminal');
const XLSX = require('xlsx');
const fs = require('fs');

const client = new Client({
    authStrategy: new LocalAuth()
});

const LOG_FILE = 'log_terkirim.txt';

client.on('qr', (qr) => {
    qrcode.generate(qr, { small: true });
    console.log('=== SILAKAN SCAN QR CODE DI ATAS ===');
});

// Mengecek jadwal begitu bot menyala (misal: saat laptop baru dihidupkan)
client.on('ready', () => {
    console.log('Bot WhatsApp sudah SIAP dan AKTIF!');
    console.log('Menjalankan pengecekan instan saat startup...');
    cekDanKirimPengingat();
});

// FITUR PENCARI ID GRUP & TES BOT
client.on('message_create', async (msg) => {
    // RADAR PELACAK: Menampilkan semua pesan yang terbaca ke terminal
    console.log(`[RADAR] Pesan terdeteksi: "${msg.body}"`);

    const chat = await msg.getChat();
    if (chat.isGroup) {
        console.log(`[INFO GRUP] Nama: "${chat.name}" | ID: ${chat.id._serialized}`);
    }

    // Fitur untuk ngetes bot hidup atau tidak
    if (msg.body === '!tes') {
        msg.reply('Halo! Bot pengingat jadwal Jumat dalam keadaan AKTIF dan siap bertugas. 🤖✅');
    }

    // FITUR BARU: Uji coba tembak pesan jadwal langsung ke grup
    if (msg.body === '!tes-jadwal') {
        msg.reply('⏳ Sedang memproses bacaan dari Excel dan mengirim ke grup...');
        await kirimPengingatJumat();
    }
});

function getTodayDateString() {
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const dd = String(today.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
}

function getNextFridayDate() {
    const today = new Date();
    const dayOfWeek = today.getDay();
    const daysUntilFriday = (5 - dayOfWeek + 7) % 7 === 0 ? 7 : (5 - dayOfWeek + 7) % 7;
    const nextFriday = new Date(today);
    nextFriday.setDate(today.getDate() + daysUntilFriday);

    const yyyy = nextFriday.getFullYear();
    const mm = String(nextFriday.getMonth() + 1).padStart(2, '0');
    const dd = String(nextFriday.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
}

async function kirimPengingatJumat(targetGroupId = null) {
    const groupId = targetGroupId || '120363420334346505@g.us';

    const tanggalJumatTarget = getNextFridayDate();

    try {
        if (!fs.existsSync('jadwal.xlsx')) {
            console.error('Error: File "jadwal.xlsx" tidak ditemukan!');
            return false;
        }

        const workbook = XLSX.readFile('jadwal.xlsx', { cellDates: true });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet);

        const dataBersih = data.map(row => {
            const newRow = {};
            for (let key in row) {
                newRow[key.trim()] = row[key];
            }
            return newRow;
        });

        const konversiTanggal = (input) => {
            if (!input) return '';

            if (input instanceof Date) {
                // [TRIK 12 JAM]: Menambahkan 12 jam untuk mengatasi tanggal mundur akibat zona waktu
                const waktuAman = new Date(input.getTime() + (12 * 60 * 60 * 1000));

                const yyyy = waktuAman.getFullYear();
                const mm = String(waktuAman.getMonth() + 1).padStart(2, '0');
                const dd = String(waktuAman.getDate()).padStart(2, '0');
                return `${yyyy}-${mm}-${dd}`;
            }
            return String(input).trim();
        };

        // Radar untuk memastikan tanggal yang dibaca sudah akurat
        console.log(`\n🔍 MENCARI JADWAL: ${tanggalJumatTarget}`);
        if (dataBersih.length > 0) {
            console.log(`Format tgl baris pertama TERBACA sbg: "${konversiTanggal(dataBersih[0].Tanggal)}"`);
        }

        const jadwalDitemukan = dataBersih.find(row => konversiTanggal(row.Tanggal) === tanggalJumatTarget);

        let teksPesan = '';

        if (jadwalDitemukan) {
            const keys = Object.keys(jadwalDitemukan);
            // Mencari key untuk No. WA (bisa "No. WA", "No WA", atau fallback ke kolom ke-3)
            const waKey = keys.find(k => k.toLowerCase().includes('wa')) || keys[2];
            const noWA = waKey ? jadwalDitemukan[waKey] : undefined;

            if (noWA) {
                teksPesan = `*PENGINGAT JADWAL JUMAT MINGGUAN*\n\n` +
                    `Assalamu'alaikum Wr. Wb.\n` +
                    `Mengingatkan kepada marbot untuk melakukan konfirmasi kembali kepada Ustadz yang bertugas besok:\n\n` +
                    `📅 *Hari/Tanggal:* Jumat, ${tanggalJumatTarget}\n` +
                    `👤 *Imam & Khotib:* ${jadwalDitemukan.Nama}\n` +
                    `📱 *No. WA:* ${noWA}\n\n` +
                    `Mohon segera menghubungi beliau. Terima kasih.\n\n` +
                    `Wassalamu'alaikum Wr. Wb.\n\n` +
                    `-fahryan`;
            } else {
                teksPesan = `*PENGINGAT JADWAL JUMAT MINGGUAN*\n\n` +
                    `Assalamu'alaikum Wr. Wb.\n` +
                    `Mengingatkan kepada marbot untuk melakukan konfirmasi kembali kepada Ustadz yang bertugas besok:\n\n` +
                    `📅 *Hari/Tanggal:* Jumat, ${tanggalJumatTarget}\n` +
                    `👤 *Imam & Khotib:* ${jadwalDitemukan.Nama}\n\n` +
                    `⚠️ _Catatan: Nomor WA tidak ditemukan di dalam data Excel. Mohon untuk mencari kontak beliau secara mandiri._\n\n` +
                    `Mohon segera menghubungi beliau. Terima kasih.\n\n` +
                    `Wassalamu'alaikum Wr. Wb.\n\n` +
                    `-fahryan`;
            }
        } else {
            teksPesan = `*PENGINGAT JADWAL JUMAT MINGGUAN*\n\n` +
                `Assalamu'alaikum Wr. Wb.\n` +
                `Mengingatkan untuk konfirmasi jadwal Imam & Khotib Jumat besok (*${tanggalJumatTarget}*).\n\n` +
                `⚠️ _Tidak ditemukan nama penceramah pada data jadwal untuk tanggal tersebut. Mohon dilakukan pengecekan lebih lanjut._\n\n` +
                `Wassalamu'alaikum Wr. Wb.\n\n` +
                `-fahryan`;
        }

        await client.sendMessage(groupId, teksPesan);
        console.log(`[SUKSES] Pesan berhasil dikirim pada ${new Date().toLocaleTimeString()}!\n`);
        return true;

    } catch (error) {
        console.error('Terjadi kesalahan:', error);
        return false;
    }
}

// Fungsi Pengecekan Utama
async function cekDanKirimPengingat() {
    const now = new Date();
    const dayOfWeek = now.getDay(); // 4 = Kamis

    // Cek apakah hari ini hari Kamis
    if (dayOfWeek === 4) {
        const hariIni = getTodayDateString();
        let sudahTerkirim = false;

        // Cek log memori
        if (fs.existsSync(LOG_FILE)) {
            const lastSentDate = fs.readFileSync(LOG_FILE, 'utf8').trim();
            if (lastSentDate === hariIni) {
                sudahTerkirim = true;
            }
        }

        // Eksekusi pengiriman jika belum
        if (!sudahTerkirim) {
            console.log('Hari Kamis terdeteksi. Mencoba mengirim pesan...');
            const berhasil = await kirimPengingatJumat();

            if (berhasil) {
                fs.writeFileSync(LOG_FILE, hariIni);
                console.log(`Pengiriman tercatat untuk tanggal ${hariIni}.`);
            }
        } else {
            console.log('Pesan untuk hari Kamis ini sudah terkirim sebelumnya. Aman!');
        }
    } else {
        console.log('Hari ini bukan Kamis. Santai dulu.');
    }
}

// Jika laptop terus menyala berhari-hari, cron ini yang akan mengecek setiap 5 menit
cron.schedule('*/5 * * * *', () => {
    cekDanKirimPengingat();
});

client.initialize();