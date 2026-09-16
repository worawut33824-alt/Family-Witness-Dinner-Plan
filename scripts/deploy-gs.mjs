#!/usr/bin/env node
/**
 * Deploy code.gs -> Google Apps Script ด้วยคำสั่งเดียว
 *
 *   npm run deploy:gs
 *
 * ทำอะไรบ้าง:
 *   1. copy code.gs (source จริงใน git) -> .appsscript/รหัส.js (ชื่อไฟล์ที่โปรเจกต์ Apps Script ใช้)
 *   2. clasp push -f  — อัปโหลด source ขึ้นโปรเจกต์
 *   3. clasp deploy -i <DEPLOYMENT_ID>  — อัปเดต deployment เดิม (คง URL /exec เดิม
 *      ที่ invite.html เรียกใช้ ไม่ต้องแก้ SCRIPT_URL)
 *
 * ต้อง login ครั้งเดียวก่อน:  npx clasp login
 * (credential อยู่ใน ~/.clasprc.json — gitignore ไว้แล้ว ห้าม push ขึ้น public repo)
 */
import { execFileSync } from 'node:child_process';
import { copyFileSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');

// deployment เดิมที่ invite.html เรียก (ตรงกับ /exec URL) — อัปเดตตัวนี้เพื่อคง URL เดิม
const DEPLOYMENT_ID =
  'AKfycbyGS50jFbBmz6nU-o6zPBGZrSokuyxD1TPQMhpfzwbpmzOJSg4emYaXKHv6QB_8Q9wAbg';

const SRC  = join(root, 'code.gs');
const DEST = join(root, '.appsscript', 'รหัส.js'); // ชื่อไฟล์ที่โปรเจกต์ใช้ (locale ไทย = "รหัส")

const npx = process.platform === 'win32' ? 'npx.cmd' : 'npx';
const run = (args) =>
  execFileSync(npx, args, { cwd: root, stdio: 'inherit', shell: false });

if (!existsSync(join(root, '.appsscript'))) {
  console.error('❌ ยังไม่มีโฟลเดอร์ .appsscript — รัน `npx clasp pull` ก่อนหนึ่งครั้ง');
  process.exit(1);
}

const label = process.argv.slice(2).join(' ') || `deploy ${new Date().toISOString()}`;

console.log('1/3  copy code.gs -> .appsscript/รหัส.js');
copyFileSync(SRC, DEST);

console.log('2/3  clasp push');
run(['clasp', 'push', '-f']);

console.log(`3/3  clasp deploy (คง URL เดิม) — "${label}"`);
run(['clasp', 'deploy', '-i', DEPLOYMENT_ID, '-d', label]);

console.log('\n✅ deploy เสร็จ — /exec URL เดิมรันโค้ดล่าสุดแล้ว');
console.log('   ⚠️  ถ้าแก้ข้อความในชีต (เช่น หัวตาราง) ต้องยิง action format ซ้ำ');
console.log('       เช่น: curl -sL "<SCRIPT_URL>?action=format_all"');
