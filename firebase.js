import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";

// ตั้งค่าการเชื่อมต่อ Firebase
const firebaseConfig = {
  apiKey: "AIzaSyDaPYyH9hz_s90snUMBD_fPkjFLKYjKUoc",
  authDomain: "student-attendance-db.firebaseapp.com",
  projectId: "student-attendance-db",
  storageBucket: "student-attendance-db.firebasestorage.app",
  messagingSenderId: "862325141839",
  appId: "1:862392518439:web:99a0d21acb5661bb830b02"
};

// เริ่มต้นใช้งาน Firebase
const app = initializeApp(firebaseConfig);

// ส่งออก db สำหรับใช้งาน Firestore ในไฟล์อื่นๆ (เช่น App.jsx)
export const db = getFirestore(app);

export default app;