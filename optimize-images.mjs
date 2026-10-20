// ============================================================
//  optimize-images.mjs — ย่อขนาดรูปที่ใหญ่เกินจริง (รันครั้งเดียว)
//  วิธีใช้:  npm i sharp   แล้ว   node optimize-images.mjs
//
//  ปลอดภัย:
//   • สำรองต้นฉบับไว้ใน ./images-backup/ ก่อน (ครั้งแรกครั้งเดียว)
//   • ย่อเฉพาะรูปที่กว้างเกินเป้าหมาย (ไม่ขยายรูปเล็ก)
//   • เขียนทับไฟล์เดิมเฉพาะเมื่อได้ไฟล์ "เล็กลงจริง" (คงชื่อ/นามสกุลเดิม → ไม่ต้องแก้ HTML)
//
//  กู้คืน: ก๊อปไฟล์จาก ./images-backup/ กลับมาทับได้ตลอด
// ============================================================
import sharp from 'sharp';
import { promises as fs } from 'fs';
import path from 'path';

const BACKUP = 'images-backup';

// [ ไฟล์, ความกว้างเป้าหมาย(px) ]  — เป้าหมาย ~2 เท่าของขนาดที่แสดงจริง (เผื่อจอ retina)
const JOBS = [
  // Love Tree (ตัวหลัก ~7.8MB)
  ['Love Tree/ต้น.PNG', 900],
  ['Love Tree/รูปคู่รัก.PNG', 460],
  ['Love Tree/ฟุ่ม1.PNG', 480],
  ['Love Tree/ฟุ่ม2.PNG', 480],
  ['Love Tree/ฟุ่ม3.PNG', 480],
  ['Love Tree/canva-watercolor-rose-flower-MAHABloemX4.png', 140],
  // โลโก้ (4500×4500 → โชว์แค่ ~110px)
  ['logo.png', 320],
  // รูปปก / นับถอยหลัง (เต็มจอ)
  ['รูปประกอบการ์ดเชิญ/cover.jpg', 1200],
  ['รูปประกอบการ์ดเชิญ/countdown.jpg', 1200],
  // รูปคู่บ่าวสาว (วงกลมเล็ก)
  ['รูปประกอบการ์ดเชิญ/groom-em-crop.jpg', 560],
  ['รูปประกอบการ์ดเชิญ/bride-paeng.jpg', 560],
  // รูปสถานที่
  ['รูปสถานที่/710298291_1707493166931898_4535512675905232597_n.jpg', 1100],
  ['รูปสถานที่/709134475_2413048042505452_8347293504824554399_n.jpg', 760],
  ['รูปสถานที่/711708635_1712209976454956_5419472830277624580_n.jpg', 760],
];

const kb = b => (b / 1024).toFixed(0) + ' KB';

async function backupOnce(file, buf) {
  const bak = path.join(BACKUP, file);
  await fs.mkdir(path.dirname(bak), { recursive: true });
  try { await fs.access(bak); }              // มีสำรองแล้ว → ไม่ทับ (กันเขียนทับสำรองด้วยไฟล์ที่ย่อแล้ว)
  catch { await fs.writeFile(bak, buf); }
}

async function run() {
  let totalBefore = 0, totalAfter = 0, changed = 0;
  console.log('เริ่มย่อรูป... (ต้นฉบับสำรองไว้ใน ./' + BACKUP + '/)\n');

  for (const [file, width] of JOBS) {
    try {
      let buf;
      try { buf = await fs.readFile(file); }
      catch { console.log('– ข้าม (ไม่พบไฟล์):', file); continue; }

      const meta = await sharp(buf).metadata();
      if (!meta.width) { console.log('– ข้าม (อ่าน metadata ไม่ได้):', file); continue; }
      if (meta.width <= width) { console.log(`– ข้าม (เล็กอยู่แล้ว ${meta.width}px): ${file}`); continue; }

      const ext = path.extname(file).toLowerCase();
      let p = sharp(buf).resize({ width, withoutEnlargement: true });
      if (ext === '.jpg' || ext === '.jpeg') p = p.jpeg({ quality: 82, mozjpeg: true });
      else p = p.png({ compressionLevel: 9, quality: 90 });   // คง alpha (ความโปร่งใส) ไว้
      const out = await p.toBuffer();

      if (out.length < buf.length * 0.95) {          // เล็กลงอย่างน้อย 5% จึงเขียนทับ
        await backupOnce(file, buf);
        await fs.writeFile(file, out);
        totalBefore += buf.length; totalAfter += out.length; changed++;
        console.log(`✓ ${file}\n    ${meta.width}px → ${width}px   ${kb(buf.length)} → ${kb(out.length)}  (ลด ${((1 - out.length / buf.length) * 100).toFixed(0)}%)`);
      } else {
        console.log(`– คงเดิม (ย่อแล้วไม่เล็กลงพอ): ${file}`);
      }
    } catch (e) {
      console.log('! ERROR', file, '-', e.message);
    }
  }

  console.log(`\n─────────────────────────────`);
  console.log(`ย่อสำเร็จ ${changed} ไฟล์`);
  console.log(`รวม ${kb(totalBefore)} → ${kb(totalAfter)}   (ประหยัด ${kb(totalBefore - totalAfter)})`);
  console.log(`\nขั้นต่อไป:`);
  console.log(`  1) ตรวจรูปในการ์ดว่ายังคมชัดดี`);
  console.log(`  2) bump เลข cache ใน sw.js (wedding-planner-vNN) แล้ว git add -A && commit && push`);
  console.log(`  (ถ้าไม่พอใจ: ก๊อปไฟล์จาก ./${BACKUP}/ กลับมาทับ)`);
}

run();
