import React, { useState, useEffect, useRef } from 'react';
import { QRCodeSVG } from 'qrcode.react';
import * as XLSX from 'xlsx';
import { db } from './firebase';
import { collection, onSnapshot, doc, setDoc, updateDoc, deleteDoc, getDocs, query, where } from 'firebase/firestore';

function App() {
  const [studentModeSubject, setStudentModeSubject] = useState(null);
  const [activeTab, setActiveTab] = useState('checkin');
  const [students, setStudents] = useState([]);
  const [subjects, setSubjects] = useState([]); 
  const [schedules, setSchedules] = useState([]); 
  const [selectedSubject, setSelectedSubject] = useState(''); 
  
  const getTodayISO = () => {
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  };

  const [selectedDate, setSelectedDate] = useState(getTodayISO());
  const [studentInputId, setStudentInputId] = useState('');
  const [studentCheckinStatus, setStudentCheckinStatus] = useState(null);

  // Subject & Schedule Form States
  const [editingSubjectId, setEditingSubjectId] = useState(null);
  const [inputSubjectCode, setInputSubjectCode] = useState('');
  const [inputSubjectName, setInputSubjectName] = useState('');
  const [inputSec, setInputSec] = useState('');
  const [scheduleDay, setScheduleDay] = useState('วันจันทร์');
  const [scheduleTimeStart, setScheduleTimeStart] = useState('08:30'); 
  const [scheduleTimeEnd, setScheduleTimeEnd] = useState('11:30');     
  const [scheduleRoom, setScheduleRoom] = useState('');         
  
  const [currentAlerts, setCurrentAlerts] = useState([]);
  const [audioEnabled, setAudioEnabled] = useState(false);

  // Modal States
  const [deleteSubjectModal, setDeleteSubjectModal] = useState({ isOpen: false, subjectId: null });
  const [customAlert, setCustomAlert] = useState({ isOpen: false, title: '', message: '', type: 'info' });

  // 💡 [แก้ไข] สถานะสำหรับกำหนด IP ของคอมพิวเตอร์เพื่อแก้ปัญหา QR สแกนไม่ได้
  const [customIp, setCustomIp] = useState(window.location.hostname === 'localhost' ? '192.168.1.50' : window.location.hostname);

  const POINTS_PER_PRESENT = 1; 
  const POINTS_PER_LATE = 0.5;  
  const MAX_ATTENDANCE_SCORE = 10; 

  const playedAlertsRef = useRef(new Set()); 
  const isFirstStudentLoad = useRef(true);

  const showAlert = (message, title = 'แจ้งเตือน', type = 'info') => {
    setCustomAlert({ isOpen: true, title, message, type });
  };

  const closeAlert = () => {
    setCustomAlert({ isOpen: false, title: '', message: '', type: 'info' });
  };

  const timeSlots = [
    { label: '7:00-8:00', startH: 7, endH: 8 },
    { label: '8:00-9:00', startH: 8, endH: 9 },
    { label: '9:00-10:00', startH: 9, endH: 10 },
    { label: '10:00-11:00', startH: 10, endH: 11 },
    { label: '11:00-12:00', startH: 11, endH: 12 },
    { label: '12:00-13:00', startH: 12, endH: 13 },
    { label: '13:00-14:00', startH: 13, endH: 15 },
    { label: '14:00-15:00', startH: 14, endH: 15 },
    { label: '15:00-16:00', startH: 15, endH: 16 },
    { label: '16:00-17:00', startH: 17, endH: 17 },
    { label: '17:00-18:00', startH: 18, endH: 18 },
    { label: '18:00-19:00', startH: 18, endH: 19 },
    { label: '19:00-20:00', startH: 19, endH: 20 }
  ];

  const daysOfWeek = [
    { name: 'วันจันทร์', short: 'จ.', bg: '#fef08a', boxBg: '#fef3c7', textColor: '#854d0e', border: '#fde047' },
    { name: 'วันอังคาร', short: 'อ.', bg: '#fbcfe8', boxBg: '#fce7f3', textColor: '#9d174d', border: '#f9a8d4' },
    { name: 'วันพุธ', short: 'พ.', bg: '#a7f3d0', boxBg: '#d1fae5', textColor: '#065f46', border: '#6ee7b7' },
    { name: 'วันพฤหัสบดี', short: 'พฤ.', bg: '#fed7aa', boxBg: '#ffedd5', textColor: '#9a3412', border: '#fdba74' },
    { name: 'วันศุกร์', short: 'ศ.', bg: '#bae6fd', boxBg: '#e0f2fe', textColor: '#075985', border: '#7dd3fc' },
    { name: 'วันเสาร์', short: 'ส.', bg: '#e9d5ff', boxBg: '#f3e8ff', textColor: '#6b21a8', border: '#d8b4fe' },
    { name: 'วันอาทิตย์', short: 'อา.', bg: '#fca5a5', boxBg: '#fee2e2', textColor: '#991b1b', border: '#f87171' }
  ];

  const parseTimeToMinutes = (timeStr) => {
    if (!timeStr) return 0;
    const cleanStr = String(timeStr).trim().replace('.', ':');
    const parts = cleanStr.split(':');
    const h = parseInt(parts[0], 10) || 0;
    const m = parseInt(parts[1], 10) || 0;
    return h * 60 + m;
  };

  const getGridColumnRange = (timeStart, timeEnd) => {
    const startMins = parseTimeToMinutes(timeStart);
    const endMins = parseTimeToMinutes(timeEnd);
    const baseMins = 7 * 60;
    const colStart = Math.max(2, Math.floor((startMins - baseMins) / 60) + 2);
    const colEnd = Math.min(15, Math.ceil((endMins - baseMins) / 60) + 2);
    return { colStart, colEnd: colEnd <= colStart ? colStart + 1 : colEnd };
  };

  const playDingSound1Sec = () => {
    try {
      const AudioContext = window.AudioContext || window.webkitAudioContext;
      if (!AudioContext) return;
      const ctx = new AudioContext();
      for (let i = 0; i < 2; i++) {
        const startTime = i * 0.5;
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.type = 'sine';
        osc.frequency.setValueAtTime(880, ctx.currentTime + startTime);
        osc.frequency.exponentialRampToValueAtTime(1760, ctx.currentTime + startTime + 0.2);
        gain.gain.setValueAtTime(0.3, ctx.currentTime + startTime);
        gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + startTime + 0.4);
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.start(ctx.currentTime + startTime);
        osc.stop(ctx.currentTime + startTime + 0.4);
      }
    } catch (e) {
      console.error("Audio error:", e);
    }
  };

  const playScheduleAlertSound5Sec = () => {
    try {
      const AudioContext = window.AudioContext || window.webkitAudioContext;
      if (!AudioContext) return;
      const ctx = new AudioContext();
      for (let i = 0; i < 10; i++) {
        const startTime = i * 0.5;
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.type = 'sine';
        osc.frequency.setValueAtTime(i % 2 === 0 ? 987.77 : 783.99, ctx.currentTime + startTime);
        gain.gain.setValueAtTime(0, ctx.currentTime + startTime);
        gain.gain.linearRampToValueAtTime(0.3, ctx.currentTime + startTime + 0.05);
        gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + startTime + 0.4);
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.start(ctx.currentTime + startTime);
        osc.stop(ctx.currentTime + startTime + 0.4);
      }
    } catch (e) {
      console.error("Audio error:", e);
    }
  };

  const enableAudioPermission = () => {
    playScheduleAlertSound5Sec(); 
    setAudioEnabled(true);
  };

  useEffect(() => {
    const searchParams = new URLSearchParams(window.location.search);
    let subjectParam = searchParams.get('subject');

    if (!subjectParam && window.location.hash.includes('?')) {
      const hashQuery = window.location.hash.split('?')[1];
      const hashParams = new URLSearchParams(hashQuery);
      subjectParam = hashParams.get('subject');
    }

    if (subjectParam) {
      setStudentModeSubject(decodeURIComponent(subjectParam));
    }
  }, []);

  const calculateAutoGrade = (totalScore) => {
    if (totalScore >= 80) return 'A';
    if (totalScore >= 75) return 'B+';
    if (totalScore >= 70) return 'B';
    if (totalScore >= 65) return 'C+';
    if (totalScore >= 60) return 'C';
    if (totalScore >= 55) return 'D+';
    if (totalScore >= 50) return 'D';
    return 'F';
  };

  useEffect(() => {
    const unsubscribe = onSnapshot(collection(db, 'students'), (snapshot) => {
      const studentList = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      }));
      studentList.sort((a, b) => String(a.id).localeCompare(String(b.id)));

      if (!isFirstStudentLoad.current) {
        snapshot.docChanges().forEach((change) => {
          if (change.type === 'modified') {
            playDingSound1Sec(); 
          }
        });
      } else {
        isFirstStudentLoad.current = false;
      }

      setStudents(studentList);
    });
    return () => unsubscribe();
  }, []);

  useEffect(() => {
    const unsubscribe = onSnapshot(collection(db, 'subjects'), (snapshot) => {
      const subjectList = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      }));
      setSubjects(subjectList);
      if (subjectList.length > 0 && !selectedSubject) {
        setSelectedSubject(subjectList[0].id);
      }
    });
    return () => unsubscribe();
  }, [selectedSubject]);

  useEffect(() => {
    const unsubscribe = onSnapshot(collection(db, 'schedules'), (snapshot) => {
      const scheduleList = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      }));
      setSchedules(scheduleList);
    });
    return () => unsubscribe();
  }, []);

  useEffect(() => {
    const checkScheduleTime = () => {
      const now = new Date();
      const options = { timeZone: 'Asia/Bangkok', hour: '2-digit', minute: '2-digit', hour12: false };
      const currentTimeStr = now.toLocaleTimeString('th-TH', options); 
      
      const dayIndex = now.getDay();
      const daysArr = ['วันอาทิตย์', 'วันจันทร์', 'วันอังคาร', 'วันพุธ', 'วันพฤหัสบดี', 'วันศุกร์', 'วันเสาร์'];
      const currentDayName = daysArr[dayIndex];

      const activeAlerts = [];

      schedules.forEach((sch) => {
        if (sch.day === currentDayName) {
          if (currentTimeStr >= sch.timeStart && currentTimeStr <= sch.timeEnd) {
            const sub = subjects.find(s => s.id === sch.subjectId);
            activeAlerts.push({
              ...sch,
              subjectName: sub ? sub.name : 'ไม่ระบุชื่อวิชา'
            });

            if (!playedAlertsRef.current.has(sch.id)) {
              playScheduleAlertSound5Sec(); 
              playedAlertsRef.current.add(sch.id);
            }
          }
        }
      });

      setCurrentAlerts(activeAlerts);
    };

    const timer = setInterval(checkScheduleTime, 1000); 
    return () => clearInterval(timer);
  }, [schedules, subjects]);

  const checkIfLateBySchedule = (subjectId) => {
    const now = new Date();
    const daysArr = ['วันอาทิตย์', 'วันจันทร์', 'วันอังคาร', 'วันพุธ', 'วันพฤหัสบดี', 'วันศุกร์', 'วันเสาร์'];
    const currentDayName = daysArr[now.getDay()];

    const todaySchedule = schedules.find(s => s.subjectId === subjectId && s.day === currentDayName);

    if (todaySchedule && todaySchedule.timeStart) {
      const startTotalMinutes = parseTimeToMinutes(todaySchedule.timeStart);
      const currentTotalMinutes = now.getHours() * 60 + now.getMinutes();

      if (currentTotalMinutes > startTotalMinutes + 10) {
        return { isLate: true, reason: `(สายเกิน 10 นาที จากเวลาเริ่ม ${todaySchedule.timeStart} น.)` };
      }
    }

    return { isLate: false, reason: '' };
  };

  const filteredStudents = students.filter(std => {
    if (!selectedSubject) return true;
    return !std.subjects || std.subjects.includes(selectedSubject);
  });

  const handleStudentSelfCheckin = async (e) => {
    e.preventDefault();
    if (!studentInputId.trim()) return;

    const stdId = studentInputId.trim();
    const student = students.find(s => String(s.id).trim() === stdId);

    if (!student) {
      setStudentCheckinStatus({ type: 'error', message: `❌ ไม่พบรหัสนักศึกษา ${stdId} ในระบบ` });
      return;
    }

    const currentSubject = studentModeSubject;
    const now = new Date();
    const todayKey = getTodayISO();
    const dateStr = now.toLocaleDateString('th-TH');
    const timeStr = now.toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    const fullTimeLog = `${dateStr} ${timeStr}`;

    const lateCheck = checkIfLateBySchedule(currentSubject);
    const finalStatus = lateCheck.isLate ? 'สาย' : 'มาเรียน';

    const statusKey = `status_${currentSubject}_${todayKey}`;
    const timeKey = `time_${currentSubject}_${todayKey}`;

    const oldAttScore = student[`scoreAtt_${currentSubject}`] || 0;
    const oldStatus = student[statusKey] || 'ขาด';

    let baseScore = oldAttScore;
    if (oldStatus === 'มาเรียน') baseScore -= POINTS_PER_PRESENT;
    if (oldStatus === 'สาย') baseScore -= POINTS_PER_LATE;

    const scoreToAdd = finalStatus === 'มาเรียน' ? POINTS_PER_PRESENT : POINTS_PER_LATE;
    let newAttScore = baseScore + scoreToAdd;
    if (newAttScore > MAX_ATTENDANCE_SCORE) newAttScore = MAX_ATTENDANCE_SCORE;

    const tasks = student[`scoreTasks_${currentSubject}`] || 0;
    const midterm = student[`scoreMidterm_${currentSubject}`] || 0;
    const final = student[`scoreFinal_${currentSubject}`] || 0;
    
    const newTotal = newAttScore + tasks + midterm + final;
    const autoGrade = calculateAutoGrade(newTotal);

    try {
      const studentRefDoc = doc(db, 'students', String(student.id).trim());
      await updateDoc(studentRefDoc, {
        [statusKey]: finalStatus,
        [timeKey]: fullTimeLog,
        [`scoreAtt_${currentSubject}`]: newAttScore, 
        [`grade_${currentSubject}`]: autoGrade      
      });

      playDingSound1Sec(); 

      const statusIcon = finalStatus === 'มาเรียน' ? '✅' : '⚠️';
      setStudentCheckinStatus({ 
        type: 'success', 
        message: `🎉 บันทึกสำเร็จ!\nชื่อ: ${student.name}\nสถานะ: ${statusIcon} ${finalStatus} ${lateCheck.reason}\nเวลา: ${fullTimeLog}` 
      });
      setStudentInputId('');
    } catch (error) {
      console.error(error);
      setStudentCheckinStatus({ type: 'error', message: 'เกิดข้อผิดพลาดในการบันทึกข้อมูล' });
    }
  };

  const handleEditSubjectClick = (sub) => {
    setEditingSubjectId(sub.id);
    setInputSubjectCode(sub.id);
    setInputSubjectName(sub.name || '');
    setInputSec(sub.sec || '');

    const existingSchedule = schedules.find(s => s.subjectId === sub.id);
    if (existingSchedule) {
      setScheduleDay(existingSchedule.day || 'วันจันทร์');
      setScheduleTimeStart(existingSchedule.timeStart || '08:30');
      setScheduleTimeEnd(existingSchedule.timeEnd || '11:30');
      setScheduleRoom(existingSchedule.room || '');
    }
  };

  const handleCancelEdit = () => {
    setEditingSubjectId(null);
    setInputSubjectCode('');
    setInputSubjectName('');
    setInputSec('');
    setScheduleTimeStart('08:30');
    setScheduleTimeEnd('11:30');
    setScheduleRoom('');
  };

  const handleAddOrUpdateFullCourse = async (e) => {
    e.preventDefault();
    if (!inputSubjectCode.trim() || !inputSubjectName.trim() || !scheduleTimeStart || !scheduleTimeEnd || !scheduleRoom.trim()) {
      showAlert('กรุณากรอกข้อมูลให้ครบถ้วนครับ', 'ข้อผิดพลาด', 'error');
      return;
    }

    const code = inputSubjectCode.trim();
    const name = inputSubjectName.trim();
    const sec = inputSec.trim() || '1';

    try {
      if (editingSubjectId) {
        if (editingSubjectId !== code) {
          await deleteDoc(doc(db, 'subjects', editingSubjectId));
          const qSchedules = query(collection(db, 'schedules'), where('subjectId', '==', editingSubjectId));
          const schSnap = await getDocs(qSchedules);
          schSnap.forEach(async (d) => {
            await deleteDoc(doc(db, 'schedules', d.id));
          });
        }

        await setDoc(doc(db, 'subjects', code), { name, sec }, { merge: true });

        const qSchedules = query(collection(db, 'schedules'), where('subjectId', '==', editingSubjectId));
        const schSnap = await getDocs(qSchedules);
        
        if (!schSnap.empty) {
          schSnap.forEach(async (schDoc) => {
            await updateDoc(doc(db, 'schedules', schDoc.id), {
              subjectId: code,
              subjectName: name,
              sec: sec,
              day: scheduleDay,
              timeStart: scheduleTimeStart.trim(),
              timeEnd: scheduleTimeEnd.trim(),
              room: scheduleRoom.trim()
            });
          });
        } else {
          const scheduleId = `${scheduleDay}_${code}_${Date.now()}`;
          await setDoc(doc(db, 'schedules', scheduleId), {
            day: scheduleDay,
            subjectId: code,
            subjectName: name,
            sec,
            timeStart: scheduleTimeStart.trim(),
            timeEnd: scheduleTimeEnd.trim(),
            room: scheduleRoom.trim()
          });
        }

        setSelectedSubject(code);
        handleCancelEdit();
        showAlert(`อัปเดตข้อมูลวิชา [${code}] เรียบร้อยแล้ว!`, 'แก้ไขสำเร็จ', 'success');
      } else {
        await setDoc(doc(db, 'subjects', code), { name, sec }, { merge: true });
        
        const scheduleId = `${scheduleDay}_${code}_${Date.now()}`;
        await setDoc(doc(db, 'schedules', scheduleId), {
          day: scheduleDay,
          subjectId: code,
          subjectName: name,
          sec,
          timeStart: scheduleTimeStart.trim(),
          timeEnd: scheduleTimeEnd.trim(),
          room: scheduleRoom.trim()
        });

        setSelectedSubject(code);
        setInputSubjectCode(''); setInputSubjectName(''); setInputSec(''); setScheduleTimeStart('08:30'); setScheduleTimeEnd('11:30'); setScheduleRoom('');
        showAlert('บันทึกเพิ่มวิชาเข้าตารางสำเร็จ!', 'สำเร็จ', 'success');
      }
    } catch (error) { 
      console.error(error); 
      showAlert('เกิดข้อผิดพลาดในการบันทึกข้อมูล', 'ข้อผิดพลาด', 'error');
    }
  };

  const confirmDeleteSubject = async () => {
    if (!deleteSubjectModal.subjectId) return;
    const subId = deleteSubjectModal.subjectId;

    try {
      await deleteDoc(doc(db, 'subjects', subId));
      const q = query(collection(db, 'schedules'), where('subjectId', '==', subId));
      const querySnapshot = await getDocs(q);
      querySnapshot.forEach(async (d) => {
        await deleteDoc(doc(db, 'schedules', d.id));
      });

      if (selectedSubject === subId) setSelectedSubject('');
      if (editingSubjectId === subId) handleCancelEdit();

      setDeleteSubjectModal({ isOpen: false, subjectId: null });
      showAlert(`ถอน/ลบรายวิชา [${subId}] ออกจากระบบเรียบร้อยแล้ว`, 'ลบสำเร็จ', 'success');
    } catch (error) {
      console.error('Delete Subject Error:', error);
      showAlert('เกิดข้อผิดพลาดในการลบวิชา', 'ข้อผิดพลาด', 'error');
    }
  };

  const handleManualClick = async (student, currentStatus) => {
    if (!selectedSubject) return showAlert('กรุณาเลือกรายวิชาก่อนครับ', 'ข้อผิดพลาด', 'error');
    let nextStatus = currentStatus === 'มาเรียน' ? 'สาย' : currentStatus === 'สาย' ? 'ลา' : currentStatus === 'ลา' ? 'ขาด' : 'มาเรียน';
    
    const now = new Date();
    const fullTimeLog = nextStatus === 'ขาด' ? '-' : `${now.toLocaleDateString('th-TH')} ${now.toLocaleTimeString('th-TH')}`;

    const statusKey = `status_${selectedSubject}_${selectedDate}`;
    const timeKey = `time_${selectedSubject}_${selectedDate}`;

    await updateDoc(doc(db, 'students', String(student.id).trim()), {
      [statusKey]: nextStatus,
      [timeKey]: fullTimeLog
    });
  };

  const handleInlineScoreChange = async (student, type, value) => {
    if (!selectedSubject) return;
    const numValue = value.trim() === '' ? 0 : Number(value);

    let att = type === 'att' ? numValue : (student[`scoreAtt_${selectedSubject}`] || 0);
    let tasks = type === 'tasks' ? numValue : (student[`scoreTasks_${selectedSubject}`] || 0);
    let midterm = type === 'midterm' ? numValue : (student[`scoreMidterm_${selectedSubject}`] || 0);
    let final = type === 'final' ? numValue : (student[`scoreFinal_${selectedSubject}`] || 0);
    
    const currentTotal = att + tasks + midterm + final;
    const autoGrade = calculateAutoGrade(currentTotal);

    await updateDoc(doc(db, 'students', String(student.id).trim()), {
      [`scoreAtt_${selectedSubject}`]: att,
      [`scoreTasks_${selectedSubject}`]: tasks,
      [`scoreMidterm_${selectedSubject}`]: midterm,
      [`scoreFinal_${selectedSubject}`]: final,
      [`grade_${selectedSubject}`]: autoGrade
    });
  };

  const downloadSampleExcel = () => {
    const sampleData = [
      {
        'id': '65121802',
        'Email': '65121802@g.cmru.ac.th',
        'name': 'นางสาวจริญญา มณีรัตน์',
        'Section / หมู่เรียน': 'Sect 51 หมู่เรียน คพ 65.ค.บ.4.1 : คอมพิวเตอร์ศึกษา'
      },
      {
        'id': '65121803',
        'Email': '65121803@g.cmru.ac.th',
        'name': 'นางสาวปวริศา ใจเอิบ',
        'Section / หมู่เรียน': 'Sect 67 หมู่เรียน คพ 65.ค.บ.4.2 : คอมพิวเตอร์ศึกษา'
      }
    ];

    const ws = XLSX.utils.json_to_sheet(sampleData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    XLSX.writeFile(wb, "รายชื่อนศ65_ตัวอย่าง.xlsx");
  };

  const handleExcelUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const bstr = evt.target.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsName = wb.SheetNames[0];
        const ws = wb.Sheets[wsName];

        const rawData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

        if (rawData.length === 0) {
          showAlert('ไฟล์ไม่มีข้อมูล หรือรูปแบบไม่ถูกต้องครับ', 'ข้อผิดพลาด', 'error');
          return;
        }

        let headerRowIndex = -1;
        let idColIndex = -1;
        let nameColIndex = -1;
        let sectColIndex = -1;

        for (let i = 0; i < Math.min(rawData.length, 15); i++) {
          const row = rawData[i];
          for (let j = 0; j < row.length; j++) {
            const cellText = String(row[j]).trim();
            if (idColIndex === -1 && /^id$|รหัส/i.test(cellText)) {
              idColIndex = j;
              headerRowIndex = i;
            }
            if (nameColIndex === -1 && /^name$|ชื่อ/i.test(cellText)) {
              nameColIndex = j;
            }
            if (sectColIndex === -1 && /sect|หมู่/i.test(cellText)) {
              sectColIndex = j;
            }
          }
          if (idColIndex !== -1) break; 
        }

        if (idColIndex === -1) {
          idColIndex = 0;
          nameColIndex = 4;
          sectColIndex = 5;
          headerRowIndex = 0;
        }

        let count = 0;
        let targetSubject = selectedSubject;

        for (let i = headerRowIndex + 1; i < rawData.length; i++) {
          const row = rawData[i];
          if (!row || row.length === 0) continue;

          const rawId = row[idColIndex];
          if (!rawId) continue;

          const studentId = String(rawId).trim();
          if (studentId.length < 4 || !/^\d+$/.test(studentId)) continue; 

          const studentName = nameColIndex < row.length && row[nameColIndex] ? String(row[nameColIndex]).trim() : 'ไม่ระบุชื่อ';
          const sectDetail = sectColIndex < row.length && row[sectColIndex] ? String(row[sectColIndex]).trim() : '';

          let extractedSec = '1';
          const secMatch = sectDetail.match(/Sect\s*(\d+)/i);
          if (secMatch) {
            extractedSec = secMatch[1];
          }

          if (!targetSubject) {
            targetSubject = 'COM3504';
            await setDoc(doc(db, 'subjects', targetSubject), { name: 'วิชาคอมพิวเตอร์ศึกษา', sec: extractedSec }, { merge: true });
            setSelectedSubject(targetSubject);
          }

          const existingStudent = students.find(s => s.id === studentId);
          let updatedSubjects = existingStudent?.subjects || [];
          
          if (!updatedSubjects.includes(targetSubject)) {
            updatedSubjects.push(targetSubject);
          }

          await setDoc(doc(db, 'students', studentId), { 
            name: studentName,
            sec: extractedSec,
            subjects: updatedSubjects
          }, { merge: true });

          count++;
        }

        showAlert(`นำเข้าข้อมูลนักศึกษาจำนวน ${count} คน สำเร็จเรียบร้อยแล้ว!`, 'สำเร็จ', 'success');
      } catch (err) {
        console.error("Excel Import Error:", err);
        showAlert('เกิดข้อผิดพลาดในการอ่านไฟล์ Excel กรุณาตรวจสอบไฟล์อีกครั้ง', 'ข้อผิดพลาด', 'error');
      }
    };

    reader.readAsBinaryString(file);
    e.target.value = null;
  };

  const exportToExcel = () => {
    if (!selectedSubject) return showAlert('กรุณาเลือกรายวิชาครับ', 'ข้อผิดพลาด', 'error');
    const formattedData = filteredStudents.map(std => ({
      'รหัสนักศึกษา (id)': std.id,
      'ชื่อ-นามสกุล (name)': std.name,
      [`สถานะเข้าเรียนวันที่ ${selectedDate}`]: std[`status_${selectedSubject}_${selectedDate}`] || 'ขาด',
      [`เวลาบันทึกวันที่ ${selectedDate}`]: std[`time_${selectedSubject}_${selectedDate}`] || '-',
      'คะแนนมาเรียน/จิตพิสัย (10)': std[`scoreAtt_${selectedSubject}`] || 0,
      'คะแนนงาน/รายงาน/โปรเจกต์ (50)': std[`scoreTasks_${selectedSubject}`] || 0,
      'คะแนนสอบกลางภาค (20)': std[`scoreMidterm_${selectedSubject}`] || 0,
      'คะแนนสอบปลายภาค (20)': std[`scoreFinal_${selectedSubject}`] || 0,
      'คะแนนรวมทั้งหมด (100)': (std[`scoreAtt_${selectedSubject}`]||0)+(std[`scoreTasks_${selectedSubject}`]||0)+(std[`scoreMidterm_${selectedSubject}`]||0)+(std[`scoreFinal_${selectedSubject}`]||0),
      'เกรดรวมอัตโนมัติ': std[`grade_${selectedSubject}`] || 'F'
    }));

    const ws = XLSX.utils.json_to_sheet(formattedData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "รายงานสมุดเกรด");
    XLSX.writeFile(wb, `สมุดเกรด_วิชา_${selectedSubject}_วันที่_${selectedDate}.xlsx`);
  };

  // 💡 [แก้ไขจุดสำคัญ] ปรับแก้ฟังก์ชันคำนวณ URL ให้เปลี่ยน localhost เป็น IP หรือ Domain เพื่อให้มือถือสแกนติด
  const getFullQrUrl = () => {
    if (!selectedSubject) return '';
    
    let { protocol, hostname, port, pathname } = window.location;
    const portStr = port ? `:${port}` : '';

    // ถ้าเปิดจาก localhost ในคอมพิวเตอร์ จะสลับไปใช้ IP address ในวงแลนแทน
    if (hostname === 'localhost' || hostname === '127.0.0.1') {
      hostname = customIp.trim() || '192.168.1.50';
    }

    const baseOrigin = `${protocol}//${hostname}${portStr}`;
    
    if (window.location.hash) {
      const baseHash = window.location.hash.split('?')[0];
      return `${baseOrigin}${pathname}${baseHash}?subject=${encodeURIComponent(selectedSubject)}`;
    }

    return `${baseOrigin}${pathname}?subject=${encodeURIComponent(selectedSubject)}`;
  };

  const qrUrl = getFullQrUrl();

  // ฝั่งนักศึกษาเปิดผ่านมือถือ
  if (studentModeSubject) {
    const currentSub = subjects.find(s => s.id === studentModeSubject);
    return (
      <div style={{ maxWidth: '450px', margin: '0 auto', padding: '15px', fontFamily: 'Tahoma, sans-serif', color: '#333', backgroundColor: '#e2e8f0', minHeight: '100vh', boxSizing: 'border-box' }}>
        <div style={{ textAlign: 'center', background: '#fff', padding: '20px 15px', borderRadius: '8px', border: '1px solid #cbd5e1', boxShadow: '0 4px 12px rgba(0,0,0,0.1)' }}>
          <h2 style={{ color: '#005f73', margin: '0 0 10px 0', fontSize: '1.3rem' }}>📲 เช็คชื่อเข้าเรียน (CMRU)</h2>
          <div style={{ background: '#f8fafc', padding: '12px', borderRadius: '6px', border: '1px solid #cbd5e1', marginBottom: '15px' }}>
            <div style={{ fontSize: '12px', color: '#64748b' }}>รายวิชา:</div>
            <div style={{ fontSize: '15px', fontWeight: 'bold', color: '#0284c7', wordBreak: 'break-word' }}>[{studentModeSubject}] {currentSub?.name}</div>
            {currentSub?.sec && <div style={{ fontSize: '12px', color: '#475569', marginTop: '4px' }}>Sec: {currentSub.sec}</div>}
            <div style={{ fontSize: '11px', color: '#64748b', marginTop: '4px' }}>ประจำวันที่: {new Date().toLocaleDateString('th-TH')}</div>
          </div>

          {studentCheckinStatus ? (
            <div style={{ padding: '15px', borderRadius: '6px', background: studentCheckinStatus.type === 'success' ? '#d1fae5' : '#fee2e2', border: `1px solid ${studentCheckinStatus.type === 'success' ? '#10b981' : '#ef4444'}`, color: '#000', whiteSpace: 'pre-line', marginBottom: '15px' }}>
              {studentCheckinStatus.message}
              <div style={{ marginTop: '15px' }}>
                <button onClick={() => setStudentCheckinStatus(null)} style={{ background: '#0284c7', color: '#fff', border: 'none', padding: '8px 16px', borderRadius: '4px', cursor: 'pointer' }}>กรอกรหัสใหม่</button>
              </div>
            </div>
          ) : (
            <form onSubmit={handleStudentSelfCheckin} style={{ display: 'flex', flexDirection: 'column', gap: '15px' }}>
              <div>
                <label style={{ display: 'block', textAlign: 'left', marginBottom: '8px', fontSize: '13px', color: '#333' }}>กรอกรหัสนักศึกษาของคุณ:</label>
                <input 
                  type="text" 
                  placeholder="เช่น 65121802" 
                  value={studentInputId} 
                  onChange={(e) => setStudentInputId(e.target.value)} 
                  style={{ width: '100%', padding: '12px', borderRadius: '4px', border: '2px solid #0284c7', background: '#ffffff', color: '#000000', fontSize: '18px', textAlign: 'center', boxSizing: 'border-box' }} 
                  required
                />
              </div>

              <div style={{ fontSize: '12px', color: '#b45309', background: '#fef3c7', padding: '8px', borderRadius: '4px', border: '1px solid #fde68a' }}>
                 หากเช็คชื่อเกินเวลาเริ่มเรียน 10 นาที ระบบจะบันทึกเป็น <strong>"สาย"</strong> อัตโนมัติ
              </div>

              <button type="submit" style={{ background: '#0284c7', color: '#fff', border: 'none', padding: '12px', borderRadius: '4px', fontSize: '15px', fontWeight: 'bold', cursor: 'pointer' }}>
                 ยืนยันกดเช็คชื่อ
              </button>
            </form>
          )}
        </div>
      </div>
    );
  }

  return (
    <div style={{ backgroundColor: '#ffffff', minHeight: '100vh', fontFamily: 'Tahoma, sans-serif', color: '#333', fontSize: '13px' }}>
      
      {/* CMRU Top Header */}
      <div style={{ background: 'linear-gradient(180deg, #facc15 0%, #eab308 100%)', borderBottom: '2px solid #ca8a04', padding: '8px 20px', display: 'flex', flexWrap: 'wrap', justifyContent: 'space-between', alignItems: 'center', gap: '10px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
          <div style={{ background: '#0284c7', color: '#fff', borderRadius: '50%', width: '42px', height: '42px', display: 'flex', justifyContent: 'center', alignItems: 'center', fontWeight: 'bold', fontSize: '14px', border: '2px solid #fff', flexShrink: 0 }}>
            CMRU
          </div>
          <div>
            <div style={{ fontSize: '18px', fontWeight: 'bold', color: '#1e3a8a' }}>มหาวิทยาลัยราชภัฏเชียงใหม่</div>
            <div style={{ fontSize: '12px', color: '#1e40af', fontWeight: '600' }}>Chiang Mai Rajabhat University</div>
          </div>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: '15px' }}>
          <button 
            onClick={enableAudioPermission} 
            style={{ 
              background: audioEnabled ? '#10b981' : '#f97316', 
              color: '#fff', 
              border: 'none', 
              padding: '6px 14px', 
              borderRadius: '4px', 
              fontWeight: 'bold', 
              fontSize: '12px',
              cursor: 'pointer' 
            }}
          >
            {audioEnabled ? '🔔 เปิดระบบเสียงแล้ว' : '🔔 เปิดระบบเสียงเตือน'}
          </button>
        </div>
      </div>

      {/* Alert Bar */}
      {currentAlerts.length > 0 && (
        <div style={{ background: '#ef4444', color: '#fff', padding: '10px 20px', fontWeight: 'bold', display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: '10px' }}>
          <span>⏰ ขณะนี้ถึงเวลาเรียนวิชา:</span>
          {currentAlerts.map(al => (
            <span key={al.id} style={{ background: '#fff', color: '#b91c1c', padding: '2px 8px', borderRadius: '4px' }}>
              [{al.subjectId}] {al.subjectName} ({al.timeStart} - {al.timeEnd} น.) ห้อง {al.room}
            </span>
          ))}
        </div>
      )}

      {/* Main Container Layout */}
      <div style={{ display: 'flex', flexWrap: 'wrap', minHeight: 'calc(100vh - 65px)' }}>
        
        {/* Left Sidebar Menu */}
        <div style={{ width: '220px', background: '#f1f5f9', borderRight: '1px solid #cbd5e1', padding: '10px', flexShrink: 0, boxSizing: 'border-box' }}>
          <div style={{ background: '#334155', color: '#fff', padding: '6px 10px', fontWeight: 'bold', fontSize: '12px', borderRadius: '2px', marginBottom: '8px' }}>
            เมนูหลัก
          </div>
          
          <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
            <button 
              onClick={() => setActiveTab('checkin')} 
              style={{ textAlign: 'left', padding: '8px 10px', background: activeTab === 'checkin' ? '#0284c7' : 'transparent', color: activeTab === 'checkin' ? '#fff' : '#334155', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: activeTab === 'checkin' ? 'bold' : 'normal' }}
            >
              📋 เช็คชื่อ & สแกน QR Code
            </button>
            <button 
              onClick={() => setActiveTab('grades')} 
              style={{ textAlign: 'left', padding: '8px 10px', background: activeTab === 'grades' ? '#0284c7' : 'transparent', color: activeTab === 'grades' ? '#fff' : '#334155', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: activeTab === 'grades' ? 'bold' : 'normal' }}
            >
              💯 สมุดบันทึกคะแนน & ตัดเกรด
            </button>
            <button 
              onClick={() => setActiveTab('schedule')} 
              style={{ textAlign: 'left', padding: '8px 10px', background: activeTab === 'schedule' ? '#0284c7' : 'transparent', color: activeTab === 'schedule' ? '#fff' : '#334155', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: activeTab === 'schedule' ? 'bold' : 'normal' }}
            >
              🗓️ ตารางเรียน / ตารางสอน
            </button>
            <button 
              onClick={() => setActiveTab('manage')} 
              style={{ textAlign: 'left', padding: '8px 10px', background: activeTab === 'manage' ? '#0284c7' : 'transparent', color: activeTab === 'manage' ? '#fff' : '#334155', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: activeTab === 'manage' ? 'bold' : 'normal' }}
            >
              ⚙️ จัดการรายวิชา / ข้อมูล
            </button>
          </div>

          <div style={{ marginTop: '20px', background: '#fff', padding: '10px', borderRadius: '4px', border: '1px solid #cbd5e1' }}>
            <div style={{ fontWeight: 'bold', fontSize: '11px', color: '#64748b', marginBottom: '6px' }}>เลือกรายวิชาทำงาน:</div>
            <select 
              value={selectedSubject} 
              onChange={(e) => setSelectedSubject(e.target.value)} 
              style={{ width: '100%', padding: '6px', borderRadius: '4px', border: '1px solid #94a3b8', fontSize: '12px' }}
            >
              {subjects.map(s => (
                <option key={s.id} value={s.id}>[{s.id}] {s.name}</option>
              ))}
            </select>
          </div>
        </div>

        {/* Main Content Area */}
        <div style={{ flex: 1, padding: '15px', overflowX: 'auto', boxSizing: 'border-box' }}>
          
          {/* TAB 1: CHECK-IN & QR CODE COMBINED */}
          {activeTab === 'checkin' && (
            <div>
              {/* QR Code Top Section */}
              <div style={{ background: '#f8fafc', border: '1px solid #cbd5e1', borderRadius: '8px', padding: '15px', marginBottom: '20px', textAlign: 'center' }}>
                <h3 style={{ margin: '0 0 10px 0', color: '#0284c7' }}>📱 QR Code สแกนเช็คชื่อเข้าเรียน</h3>
                {selectedSubject ? (
                  <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
                    <div style={{ background: '#fff', padding: '12px', borderRadius: '8px', border: '1px solid #cbd5e1', boxShadow: '0 2px 6px rgba(0,0,0,0.05)' }}>
                      <QRCodeSVG value={qrUrl} size={160} level="M" includeMargin={true} />
                    </div>
                    <div style={{ fontSize: '13px', fontWeight: 'bold', marginTop: '8px', color: '#0f172a' }}>
                      วิชา: [{selectedSubject}] {subjects.find(s => s.id === selectedSubject)?.name}
                    </div>
                    <div style={{ fontSize: '11px', color: '#64748b', marginTop: '4px' }}>
                      URL สแกน: <a href={qrUrl} target="_blank" rel="noreferrer" style={{ color: '#0284c7', wordBreak: 'break-all' }}>{qrUrl}</a>
                    </div>

                    {/* 💡 [ส่วนปรับแต่ง IP เครื่องคอมพิวเตอร์] */}
                    <div style={{ marginTop: '12px', background: '#f1f5f9', padding: '8px 12px', borderRadius: '6px', border: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', gap: '8px', fontSize: '11px' }}>
                      <span>⚙️ หากสแกนไม่ติด ให้เปลี่ยนเป็น IP เครื่องคอมฯ คุณ:</span>
                      <input 
                        type="text" 
                        value={customIp} 
                        onChange={(e) => setCustomIp(e.target.value)} 
                        placeholder="เช่น 192.168.1.50"
                        style={{ padding: '3px 6px', border: '1px solid #cbd5e1', borderRadius: '4px', width: '120px', fontSize: '11px', textAlign: 'center' }}
                      />
                    </div>
                  </div>
                ) : (
                  <div style={{ color: '#ef4444' }}>กรุณาเลือกรายวิชาทางเมนูด้านซ้ายก่อนครับ</div>
                )}
              </div>

              {/* Attendance Table Section */}
              <div style={{ display: 'flex', flexWrap: 'wrap', justifyContent: 'space-between', alignItems: 'center', gap: '10px', marginBottom: '10px', background: '#f8fafc', padding: '10px', borderRadius: '6px', border: '1px solid #e2e8f0' }}>
                <div>
                  <h3 style={{ margin: 0, color: '#0f172a' }}>📋 ตารางเช็คชื่อเข้าเรียนประจำวัน</h3>
                  <div style={{ color: '#64748b', fontSize: '12px', marginTop: '2px' }}>
                    วิชาที่เลือก: <strong style={{ color: '#0284c7' }}>{selectedSubject ? `[${selectedSubject}] ${subjects.find(s => s.id === selectedSubject)?.name || ''}` : 'ยังไม่ได้เลือกวิชา'}</strong>
                  </div>
                </div>

                <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                  <label style={{ fontSize: '12px', fontWeight: 'bold' }}>เลือกวันที่:</label>
                  <input 
                    type="date" 
                    value={selectedDate} 
                    onChange={(e) => setSelectedDate(e.target.value)} 
                    style={{ padding: '5px 8px', borderRadius: '4px', border: '1px solid #cbd5e1' }}
                  />
                </div>
              </div>

              <div style={{ overflowX: 'auto', background: '#fff', border: '1px solid #cbd5e1', borderRadius: '4px' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', textAlign: 'center' }}>
                  <thead>
                    <tr style={{ background: '#f1f5f9', borderBottom: '2px solid #cbd5e1', color: '#334155' }}>
                      <th style={{ padding: '10px', borderRight: '1px solid #cbd5e1', width: '120px' }}>รหัสนักศึกษา</th>
                      <th style={{ padding: '10px', borderRight: '1px solid #cbd5e1', textAlign: 'left' }}>ชื่อ-นามสกุล</th>
                      <th style={{ padding: '10px', borderRight: '1px solid #cbd5e1', width: '150px' }}>สถานะเข้าเรียน ({selectedDate})</th>
                      <th style={{ padding: '10px', width: '200px' }}>เวลาที่บันทึกสแกน</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredStudents.length === 0 ? (
                      <tr><td colSpan="4" style={{ padding: '20px', color: '#94a3b8' }}>ไม่มีข้อมูลนักศึกษาในรายวิชานี้</td></tr>
                    ) : (
                      filteredStudents.map(std => {
                        const statusKey = `status_${selectedSubject}_${selectedDate}`;
                        const timeKey = `time_${selectedSubject}_${selectedDate}`;
                        const currentStatus = std[statusKey] || 'ขาด';
                        const timeLog = std[timeKey] || '-';

                        let statusColor = '#ef4444';
                        if (currentStatus === 'มาเรียน') statusColor = '#10b981';
                        if (currentStatus === 'สาย') statusColor = '#f59e0b';
                        if (currentStatus === 'ลา') statusColor = '#6b7280';

                        return (
                          <tr key={std.id} style={{ borderBottom: '1px solid #e2e8f0' }}>
                            <td style={{ padding: '8px', borderRight: '1px solid #e2e8f0', fontWeight: 'bold' }}>{std.id}</td>
                            <td style={{ padding: '8px', borderRight: '1px solid #e2e8f0', textAlign: 'left' }}>{std.name}</td>
                            <td style={{ padding: '8px', borderRight: '1px solid #e2e8f0' }}>
                              <button 
                                onClick={() => handleManualClick(std, currentStatus)} 
                                style={{ background: statusColor, color: '#fff', border: 'none', padding: '5px 12px', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', fontWeight: 'bold', width: '80px' }}
                              >
                                {currentStatus}
                              </button>
                            </td>
                            <td style={{ padding: '8px', fontSize: '11px', color: '#64748b' }}>{timeLog}</td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* TAB 2: GRADES & SCORE RECORD */}
          {activeTab === 'grades' && (
            <div>
              <div style={{ display: 'flex', flexWrap: 'wrap', justifyContent: 'space-between', alignItems: 'center', gap: '10px', marginBottom: '15px', background: '#f8fafc', padding: '10px', borderRadius: '6px', border: '1px solid #e2e8f0' }}>
                <div>
                  <h3 style={{ margin: 0, color: '#0f172a' }}>💯 สมุดบันทึกคะแนนและตัดเกรด</h3>
                  <div style={{ color: '#64748b', fontSize: '12px', marginTop: '2px' }}>
                    วิชาที่เลือก: <strong style={{ color: '#0284c7' }}>{selectedSubject ? `[${selectedSubject}] ${subjects.find(s => s.id === selectedSubject)?.name || ''}` : 'ยังไม่ได้เลือกวิชา'}</strong>
                  </div>
                </div>

                <button onClick={exportToExcel} style={{ background: '#16a34a', color: '#fff', border: 'none', padding: '6px 14px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' }}>
                  📊 ส่งออกไฟล์ Excel
                </button>
              </div>

              <div style={{ overflowX: 'auto', background: '#fff', border: '1px solid #cbd5e1', borderRadius: '4px' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', textAlign: 'center' }}>
                  <thead>
                    <tr style={{ background: '#f1f5f9', borderBottom: '2px solid #cbd5e1', color: '#334155' }}>
                      <th style={{ padding: '8px', borderRight: '1px solid #cbd5e1' }}>รหัสนักศึกษา</th>
                      <th style={{ padding: '8px', borderRight: '1px solid #cbd5e1', textAlign: 'left' }}>ชื่อ-นามสกุล</th>
                      <th style={{ padding: '8px', borderRight: '1px solid #cbd5e1', background: '#fef3c7' }}>เข้าเรียน (10)</th>
                      <th style={{ padding: '8px', borderRight: '1px solid #cbd5e1', background: '#e0f2fe' }}>งาน (50)</th>
                      <th style={{ padding: '8px', borderRight: '1px solid #cbd5e1', background: '#fce7f3' }}>กลางภาค (20)</th>
                      <th style={{ padding: '8px', borderRight: '1px solid #cbd5e1', background: '#d1fae5' }}>ปลายภาค (20)</th>
                      <th style={{ padding: '8px', borderRight: '1px solid #cbd5e1' }}>รวม (100)</th>
                      <th style={{ padding: '8px' }}>เกรด</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredStudents.length === 0 ? (
                      <tr><td colSpan="8" style={{ padding: '20px', color: '#94a3b8' }}>ไม่มีข้อมูลนักศึกษาในรายวิชานี้</td></tr>
                    ) : (
                      filteredStudents.map(std => {
                        const attScore = std[`scoreAtt_${selectedSubject}`] || 0;
                        const tasksScore = std[`scoreTasks_${selectedSubject}`] || 0;
                        const midtermScore = std[`scoreMidterm_${selectedSubject}`] || 0;
                        const finalScore = std[`scoreFinal_${selectedSubject}`] || 0;
                        const totalScore = attScore + tasksScore + midtermScore + finalScore;

                        return (
                          <tr key={std.id} style={{ borderBottom: '1px solid #e2e8f0' }}>
                            <td style={{ padding: '6px', borderRight: '1px solid #e2e8f0', fontWeight: 'bold' }}>{std.id}</td>
                            <td style={{ padding: '6px', borderRight: '1px solid #e2e8f0', textAlign: 'left' }}>{std.name}</td>
                            <td style={{ padding: '6px', borderRight: '1px solid #e2e8f0', background: '#fffbeb' }}>
                              <input type="number" value={attScore} onChange={(e) => handleInlineScoreChange(std, 'att', e.target.value)} style={{ width: '50px', textAlign: 'center' }} />
                            </td>
                            <td style={{ padding: '6px', borderRight: '1px solid #e2e8f0', background: '#f0f9ff' }}>
                              <input type="number" value={tasksScore} onChange={(e) => handleInlineScoreChange(std, 'tasks', e.target.value)} style={{ width: '50px', textAlign: 'center' }} />
                            </td>
                            <td style={{ padding: '6px', borderRight: '1px solid #e2e8f0', background: '#fdf2f8' }}>
                              <input type="number" value={midtermScore} onChange={(e) => handleInlineScoreChange(std, 'midterm', e.target.value)} style={{ width: '50px', textAlign: 'center' }} />
                            </td>
                            <td style={{ padding: '6px', borderRight: '1px solid #e2e8f0', background: '#ecfdf5' }}>
                              <input type="number" value={finalScore} onChange={(e) => handleInlineScoreChange(std, 'final', e.target.value)} style={{ width: '50px', textAlign: 'center' }} />
                            </td>
                            <td style={{ padding: '6px', borderRight: '1px solid #e2e8f0', fontWeight: 'bold', color: '#0369a1' }}>{totalScore}</td>
                            <td style={{ padding: '6px', fontWeight: 'bold', color: totalScore >= 50 ? '#15803d' : '#b91c1c' }}>{std[`grade_${selectedSubject}`] || calculateAutoGrade(totalScore)}</td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* TAB 3: SCHEDULE */}
          {activeTab === 'schedule' && (
            <div>
              <h3 style={{ marginTop: 0, color: '#0f172a' }}>🗓️ ตารางเรียน / ตารางสอนประจำสัปดาห์</h3>
              <div style={{ overflowX: 'auto', background: '#fff', border: '1px solid #cbd5e1', borderRadius: '6px', padding: '10px' }}>
                <div style={{ display: 'grid', gridTemplateColumns: '80px repeat(13, minmax(60px, 1fr))', gap: '4px', fontSize: '11px', textAlign: 'center' }}>
                  <div style={{ background: '#334155', color: '#fff', padding: '8px 2px', fontWeight: 'bold' }}>วัน / เวลา</div>
                  {timeSlots.map((ts, idx) => (
                    <div key={idx} style={{ background: '#f1f5f9', padding: '8px 2px', fontWeight: 'bold', borderBottom: '2px solid #cbd5e1' }}>{ts.label}</div>
                  ))}

                  {daysOfWeek.map((dayObj) => (
                    <React.Fragment key={dayObj.name}>
                      <div style={{ background: dayObj.bg, color: dayObj.textColor, padding: '15px 2px', fontWeight: 'bold', display: 'flex', alignItems: 'center', justifyContent: 'center', borderRadius: '4px' }}>
                        {dayObj.name}
                      </div>

                      <div style={{ gridColumn: '2 / 15', display: 'grid', gridTemplateColumns: 'repeat(13, 1fr)', gap: '4px', position: 'relative' }}>
                        {schedules.filter(s => s.day === dayObj.name).map((sch) => {
                          const { colStart, colEnd } = getGridColumnRange(sch.timeStart, sch.timeEnd);
                          return (
                            <div 
                              key={sch.id} 
                              style={{ 
                                gridColumn: `${colStart - 1} / ${colEnd - 1}`, 
                                background: dayObj.boxBg, 
                                border: `1px solid ${dayObj.border}`, 
                                borderRadius: '4px', 
                                padding: '4px', 
                                color: dayObj.textColor, 
                                textAlign: 'left',
                                fontSize: '11px',
                                overflow: 'hidden'
                              }}
                            >
                              <div style={{ fontWeight: 'bold' }}>[{sch.subjectId}] {sch.subjectName}</div>
                              <div>เวลา: {sch.timeStart} - {sch.timeEnd} น.</div>
                              <div>ห้อง: {sch.room} (Sec {sch.sec || '1'})</div>
                            </div>
                          );
                        })}
                      </div>
                    </React.Fragment>
                  ))}
                </div>
              </div>
            </div>
          )}

          {/* TAB 4: MANAGE & IMPORT */}
          {activeTab === 'manage' && (
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '20px' }}>
              <div style={{ flex: '1 1 320px', background: '#f8fafc', padding: '15px', borderRadius: '8px', border: '1px solid #cbd5e1' }}>
                <h4 style={{ marginTop: 0, color: '#0284c7' }}>{editingSubjectId ? '✏️ แก้ไขข้อมูลรายวิชา' : '➕ เพิ่มรายวิชาและตารางเรียน'}</h4>
                <form onSubmit={handleAddOrUpdateFullCourse} style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                  <div>
                    <label style={{ fontSize: '11px', fontWeight: 'bold' }}>รหัสวิชา:</label>
                    <input type="text" value={inputSubjectCode} onChange={(e) => setInputSubjectCode(e.target.value)} style={{ width: '100%', padding: '6px', borderRadius: '4px', border: '1px solid #cbd5e1' }} required />
                  </div>
                  <div>
                    <label style={{ fontSize: '11px', fontWeight: 'bold' }}>ชื่อรายวิชา:</label>
                    <input type="text" value={inputSubjectName} onChange={(e) => setInputSubjectName(e.target.value)} style={{ width: '100%', padding: '6px', borderRadius: '4px', border: '1px solid #cbd5e1' }} required />
                  </div>
                  <div>
                    <label style={{ fontSize: '11px', fontWeight: 'bold' }}>Section / หมู่เรียน:</label>
                    <input type="text" value={inputSec} onChange={(e) => setInputSec(e.target.value)} placeholder="เช่น 1 หรือ 51" style={{ width: '100%', padding: '6px', borderRadius: '4px', border: '1px solid #cbd5e1' }} />
                  </div>
                  <div>
                    <label style={{ fontSize: '11px', fontWeight: 'bold' }}>วันที่มีเรียน:</label>
                    <select value={scheduleDay} onChange={(e) => setScheduleDay(e.target.value)} style={{ width: '100%', padding: '6px', borderRadius: '4px', border: '1px solid #cbd5e1' }}>
                      {daysOfWeek.map(d => <option key={d.name} value={d.name}>{d.name}</option>)}
                    </select>
                  </div>
                  <div style={{ display: 'flex', gap: '10px' }}>
                    <div style={{ flex: 1 }}>
                      <label style={{ fontSize: '11px', fontWeight: 'bold' }}>เวลาเริ่ม:</label>
                      <input type="time" value={scheduleTimeStart} onChange={(e) => setScheduleTimeStart(e.target.value)} style={{ width: '100%', padding: '6px', borderRadius: '4px', border: '1px solid #cbd5e1' }} required />
                    </div>
                    <div style={{ flex: 1 }}>
                      <label style={{ fontSize: '11px', fontWeight: 'bold' }}>เวลาสิ้นสุด:</label>
                      <input type="time" value={scheduleTimeEnd} onChange={(e) => setScheduleTimeEnd(e.target.value)} style={{ width: '100%', padding: '6px', borderRadius: '4px', border: '1px solid #cbd5e1' }} required />
                    </div>
                  </div>
                  <div>
                    <label style={{ fontSize: '11px', fontWeight: 'bold' }}>ห้องเรียน:</label>
                    <input type="text" value={scheduleRoom} onChange={(e) => setScheduleRoom(e.target.value)} placeholder="เช่น 2804" style={{ width: '100%', padding: '6px', borderRadius: '4px', border: '1px solid #cbd5e1' }} required />
                  </div>

                  <div style={{ display: 'flex', gap: '10px', marginTop: '10px' }}>
                    <button type="submit" style={{ flex: 1, background: '#0284c7', color: '#fff', border: 'none', padding: '8px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' }}>
                      {editingSubjectId ? 'บันทึกการแก้ไข' : 'บันทึกรายวิชา'}
                    </button>
                    {editingSubjectId && (
                      <button type="button" onClick={handleCancelEdit} style={{ background: '#64748b', color: '#fff', border: 'none', padding: '8px', borderRadius: '4px', cursor: 'pointer' }}>
                        ยกเลิก
                      </button>
                    )}
                  </div>
                </form>
              </div>

              <div style={{ flex: '1 1 350px', display: 'flex', flexDirection: 'column', gap: '20px' }}>
                <div style={{ background: '#f8fafc', padding: '15px', borderRadius: '8px', border: '1px solid #cbd5e1' }}>
                  <h4 style={{ marginTop: 0, color: '#16a34a' }}>📥 นำเข้าข้อมูลรายชื่อนักศึกษา (Excel)</h4>
                  <div style={{ display: 'flex', gap: '10px', alignItems: 'center', marginBottom: '10px' }}>
                    <button onClick={downloadSampleExcel} style={{ background: '#64748b', color: '#fff', border: 'none', padding: '6px 12px', borderRadius: '4px', cursor: 'pointer', fontSize: '11px' }}>
                      📄 โหลดไฟล์ตัวอย่าง (.xlsx)
                    </button>
                  </div>
                  <input type="file" accept=".xlsx, .xls" onChange={handleExcelUpload} style={{ fontSize: '12px' }} />
                </div>

                <div style={{ background: '#f8fafc', padding: '15px', borderRadius: '8px', border: '1px solid #cbd5e1' }}>
                  <h4 style={{ marginTop: 0, color: '#334155' }}>📚 รายวิชาทั้งหมดในระบบ</h4>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', maxHeight: '250px', overflowY: 'auto' }}>
                    {subjects.map(sub => (
                      <div key={sub.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: '#fff', padding: '8px', borderRadius: '4px', border: '1px solid #e2e8f0' }}>
                        <div>
                          <div style={{ fontWeight: 'bold', color: '#0284c7' }}>[{sub.id}] {sub.name}</div>
                          <div style={{ fontSize: '11px', color: '#64748b' }}>Sec: {sub.sec || '1'}</div>
                        </div>
                        <div style={{ display: 'flex', gap: '6px' }}>
                          <button onClick={() => handleEditSubjectClick(sub)} style={{ background: '#f59e0b', color: '#fff', border: 'none', padding: '4px 8px', borderRadius: '4px', cursor: 'pointer', fontSize: '11px' }}>แก้ไข</button>
                          <button onClick={() => setDeleteSubjectModal({ isOpen: true, subjectId: sub.id })} style={{ background: '#ef4444', color: '#fff', border: 'none', padding: '4px 8px', borderRadius: '4px', cursor: 'pointer', fontSize: '11px' }}>ถอนวิชา</button>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            </div>
          )}

        </div>
      </div>

      {/* Modal Custom Alert */}
      {customAlert.isOpen && (
        <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.5)', display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 1000 }}>
          <div style={{ background: '#fff', padding: '20px', borderRadius: '8px', maxWidth: '400px', width: '90%', textAlign: 'center' }}>
            <h4 style={{ marginTop: 0, color: customAlert.type === 'error' ? '#ef4444' : '#0284c7' }}>{customAlert.title}</h4>
            <p style={{ fontSize: '13px', color: '#334155' }}>{customAlert.message}</p>
            <button onClick={closeAlert} style={{ background: '#0284c7', color: '#fff', border: 'none', padding: '6px 16px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold', marginTop: '10px' }}>ตกลง</button>
          </div>
        </div>
      )}

      {/* Modal Confirm Delete Subject */}
      {deleteSubjectModal.isOpen && (
        <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.5)', display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 1000 }}>
          <div style={{ background: '#fff', padding: '20px', borderRadius: '8px', maxWidth: '400px', width: '90%', textAlign: 'center' }}>
            <h4 style={{ marginTop: 0, color: '#ef4444' }}>⚠️ ยืนยันการถอน/ลบรายวิชา</h4>
            <p style={{ fontSize: '13px', color: '#334155' }}>ต้องการลบวิชา [{deleteSubjectModal.subjectId}] ใช่หรือไม่? ข้อมูลตารางสอนที่เกี่ยวข้องจะถูกลบด้วย</p>
            <div style={{ display: 'flex', justifyContent: 'center', gap: '10px', marginTop: '15px' }}>
              <button onClick={confirmDeleteSubject} style={{ background: '#ef4444', color: '#fff', border: 'none', padding: '6px 16px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' }}>ยืนยันลบ</button>
              <button onClick={() => setDeleteSubjectModal({ isOpen: false, subjectId: null })} style={{ background: '#64748b', color: '#fff', border: 'none', padding: '6px 16px', borderRadius: '4px', cursor: 'pointer' }}>ยกเลิก</button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}

export default App;