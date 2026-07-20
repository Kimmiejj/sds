# เข็มทิศอาชีพ RIASEC Thailand

เว็บแบบประเมิน S.D.S. / Holland Code สำหรับนักเรียนไทย ใช้คำถามจากเอกสารต้นฉบับครบ 216 ข้อ (กิจกรรม 66, ความสามารถ 66, อาชีพ 84) พร้อมข้อมูลคณะ อาชีพ MyTCAS และเว็บไซต์ตลาดแรงงานไทย

## สิ่งที่ระบบบันทึก

- ชื่อ, นามสกุล, ชื่อเล่น, ระดับชั้น, ห้อง, เลขที่ และการยินยอม
- คำตอบทั้ง 216 ข้อ, คะแนน R/I/A/S/E/C และ Holland Code
- ข้อมูลเก็บใน Google Sheet: `SDS-data` โดยครูเป็นผู้ดูแลสิทธิ์

## เชื่อม Google Sheet ครั้งเดียว

GitHub Pages เป็นเว็บสาธารณะแบบ static จึงไม่ควรใส่สิทธิ์แก้ไข Google Sheet ไว้บนหน้าเว็บโดยตรง ระบบจึงใช้ Google Apps Script เป็นช่องรับข้อมูลที่ปลอดภัยกว่า

1. เปิด [Google Sheet SDS-data](https://docs.google.com/spreadsheets/d/1955ghvAH8XdYs2Q356-YQhFdt4_0LoS0jTbnfaXswF8/edit)
2. เลือก **ส่วนขยาย → Apps Script**
3. แทนที่เนื้อหาใน `Code.gs` ด้วยไฟล์ [apps-script/Code.gs](apps-script/Code.gs) และเพิ่มไฟล์ `appsscript.json` ตามในโฟลเดอร์เดียวกัน
4. กด **Deploy → New deployment → Web app**
   - Execute as: **Me**
   - Who has access: **Anyone** (หรือผู้ใช้ Google ในโดเมนโรงเรียน หากนักเรียนทุกคนมีบัญชีโดเมน)
5. คัดลอก URL ที่ลงท้ายด้วย `/exec` แล้วใส่ใน `config.js`:

```js
window.SDS_CONFIG = { apiUrl: "วาง_URL_ที่นี่" };
```

6. Commit และ push `config.js` ไปที่ `main` อีกครั้ง เว็บจะเริ่มรับคำตอบได้ทันที

## เผยแพร่เว็บไซต์

ไฟล์ `.github/workflows/pages.yml` ตั้งค่าให้ GitHub Pages เผยแพร่เว็บอัตโนมัติทุกครั้งที่ push branch `main` แล้ว

## ข้อควรทราบ

ผล RIASEC ใช้เพื่อสำรวจความสนใจทางการศึกษาและอาชีพ ไม่ใช่การตัดสินความสามารถหรืออนาคตของนักเรียน


https://kimmiejj.github.io/sds/
