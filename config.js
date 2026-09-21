// thai-ur office — ตั้งค่าหน้าเว็บ
// แก้ 2 ค่านี้แล้ว push ขึ้น GitHub ได้เลย
//
// GOOGLE_CLIENT_ID : จาก Google Cloud Console → Credentials → OAuth 2.0 Client ID (ชนิด Web application)
// API_URL          : จาก Apps Script → Deploy → Web app (หรือเมนูในชีต "ดู URL ของ API")
//
// ค่าทั้งสองนี้ไม่ใช่ความลับ เปิดเผยในหน้าเว็บได้ตามปกติของ Google Sign-In
// ความปลอดภัยอยู่ที่ Authorized JavaScript origins และการตรวจ token ฝั่งเซิร์ฟเวอร์

window.CHECKIN_CONFIG = {
  GOOGLE_CLIENT_ID: '460755476572-cluketgfop5b5otiqkm9vhoftnm4qtgm.apps.googleusercontent.com',
  API_URL: 'https://script.google.com/macros/s/AKfycbz4SWsHbp68Q6p20i2iTMUJAJjPRi6yJmPPVW2uDb-ojqLYK0xKpysOOL1V5JDUb_fQ/exec'
};
