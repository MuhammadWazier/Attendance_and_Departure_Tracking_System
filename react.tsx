import React, { useState, useEffect, useRef, useMemo } from 'react';
import Chart from 'chart.js/auto';
import * as XLSX from 'xlsx';
import syrianEagleLogo from './SyrianEagleLogo.png';

const BACKUP_FORMAT = 'xface-attendance-backup';
const DATABASE_NAME = 'xface-attendance-database';
const DATABASE_VERSION = 1;
const DATA_STORE_NAME = 'application-data';
const APP_DATA_KEY = 'current';
let databasePromise;

const openAppDatabase = () => {
  if (!('indexedDB' in window)) {
    return Promise.reject(new Error('هذا المتصفح لا يدعم قاعدة البيانات المحلية IndexedDB.'));
  }
  if (!databasePromise) {
    databasePromise = new Promise((resolve, reject) => {
      const request = window.indexedDB.open(DATABASE_NAME, DATABASE_VERSION);
      request.onupgradeneeded = () => {
        if (!request.result.objectStoreNames.contains(DATA_STORE_NAME)) {
          request.result.createObjectStore(DATA_STORE_NAME, { keyPath: 'key' });
        }
      };
      request.onsuccess = () => {
        request.result.onversionchange = () => {
          request.result.close();
          databasePromise = undefined;
        };
        resolve(request.result);
      };
      request.onerror = () => reject(request.error || new Error('تعذر فتح قاعدة البيانات المحلية.'));
      request.onblocked = () => reject(new Error('قاعدة البيانات المحلية مشغولة في نافذة أخرى.'));
    }).catch(error => {
      databasePromise = undefined;
      throw error;
    });
  }
  return databasePromise;
};

const readStoredAppData = async () => {
  const database = await openAppDatabase();
  return new Promise((resolve, reject) => {
    const transaction = database.transaction(DATA_STORE_NAME, 'readonly');
    const request = transaction.objectStore(DATA_STORE_NAME).get(APP_DATA_KEY);
    request.onsuccess = () => resolve(request.result?.data);
    request.onerror = () => reject(request.error || new Error('تعذر قراءة البيانات المحلية.'));
    transaction.onabort = () => reject(transaction.error || new Error('تعذرت قراءة البيانات المحلية.'));
  });
};

const writeStoredAppData = async data => {
  const database = await openAppDatabase();
  return new Promise((resolve, reject) => {
    const transaction = database.transaction(DATA_STORE_NAME, 'readwrite');
    transaction.objectStore(DATA_STORE_NAME).put({ key: APP_DATA_KEY, data });
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error || new Error('تعذر حفظ البيانات المحلية.'));
    transaction.onabort = () => reject(transaction.error || new Error('أُلغيت عملية حفظ البيانات المحلية.'));
  });
};

const createEmptyAppData = () => ({
  settings: {
    startTime: '08:00',
    endTime: '16:00',
    graceLate: 15,
    graceEarly: 15,
    workDays: ['الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس'],
    holidays: []
  },
  attendance: [],
  leaves: [],
  employeesStatus: {}
});

const isObjectRecord = value => value !== null && typeof value === 'object' && !Array.isArray(value);

export default function App() {
  // Main Application State
  const [activeTab, setActiveTab] = useState('dashboard');
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [selectedDirectoryEmployeeId, setSelectedDirectoryEmployeeId] = useState('');
  const [toast, setToast] = useState({ show: false, message: '', type: 'success' });
  const [isDataReady, setIsDataReady] = useState(false);
  
  const [appData, setAppData] = useState({
    settings: {
      startTime: "08:00",
      endTime: "16:00",
      graceLate: 15,
      graceEarly: 15,
      workDays: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس"],
      holidays: [
        { id: "H1", title: "عيد الفطر السعيد", fromDate: "2026-03-20", toDate: "2026-03-23" },
        { id: "H2", title: "عيد الجلاء", fromDate: "2026-04-17", toDate: "2026-04-17" },
        { id: "H3", title: "بداية العام الدراسي", fromDate: "2026-10-06", toDate: "2026-10-06" }
      ]
    },
    attendance: [],
    leaves: [],
    employeesStatus: {}
  });

  const [tempUploadedAttendance, setTempUploadedAttendance] = useState([]);
  
  // Modal State for Employee Status
  const [empModal, setEmpModal] = useState({ isOpen: false, empId: '', status: 'نشط', date: '', directorate: '', department: '' });

  // Leave Form State
  const [leaveForm, setLeaveForm] = useState({
    empName: '',
    empId: '',
    type: 'سنوية',
    leaveUnit: 'أيام',
    fromDate: '',
    toDate: '',
    fromTime: '',
    toTime: '',
    documentNumber: '',
    status: 'معتمدة',
    notes: ''
  });
  const [leaveEmployeeSearch, setLeaveEmployeeSearch] = useState('');
  const [showLeaveEmployeeSuggestions, setShowLeaveEmployeeSuggestions] = useState(false);
  const [activeLeaveEmployeeSuggestion, setActiveLeaveEmployeeSuggestion] = useState(-1);

  // Settings Form State
  const [settingsForm, setSettingsForm] = useState({
    startTime: "08:00",
    endTime: "16:00",
    graceLate: 15,
    graceEarly: 15,
    workDays: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس"]
  });

  const [holidayForm, setHolidayForm] = useState({ title: '', fromDate: '', toDate: '' });

  // Reports Filter State
  const [reportFilters, setReportFilters] = useState({
    search: '',
    directorate: '',
    department: '',
    status: '',
    startDate: '',
    endDate: ''
  });

  // Employees view filter
  const [showInactiveToggle, setShowInactiveToggle] = useState(false);
  const [employeeDirectorySearch, setEmployeeDirectorySearch] = useState('');

  // Chart references
  const pieChartRef = useRef(null);
  const barChartRef = useRef(null);
  const pieInstance = useRef(null);
  const barInstance = useRef(null);
  const toastTimeout = useRef(null);
  const backupFileInputRef = useRef(null);

  useEffect(() => {
    let cancelled = false;
    const initializeDatabase = async () => {
      try {
        const storedData = await readStoredAppData();
        if (storedData) {
          const normalizedData = normalizeAppData(storedData);
          if (!cancelled) {
            setAppData(normalizedData);
            setSettingsForm(normalizedData.settings);
          }
          return;
        }

        const legacyData = localStorage.getItem('xface_react_app_data');
        if (legacyData) {
          const migratedData = normalizeAppData(JSON.parse(legacyData));
          await writeStoredAppData(migratedData);
          localStorage.removeItem('xface_react_app_data');
          if (!cancelled) {
            setAppData(migratedData);
            setSettingsForm(migratedData.settings);
            showToastMsg('تم نقل بياناتك المحفوظة إلى قاعدة البيانات المحلية الجديدة.');
          }
        } else {
          await loadDemoData(false);
        }
      } catch (error) {
        console.error(error);
        if (!cancelled) {
          try {
            const legacyData = localStorage.getItem('xface_react_app_data');
            if (legacyData) {
              const fallbackData = normalizeAppData(JSON.parse(legacyData));
              setAppData(fallbackData);
              setSettingsForm(fallbackData.settings);
            }
          } catch (fallbackError) {
            console.error(fallbackError);
          }
          showToastMsg('تعذر الوصول إلى قاعدة البيانات المحلية. تحقق من إعدادات المتصفح؛ قد لا تُحفظ التعديلات.', 'error');
        }
      } finally {
        if (!cancelled) setIsDataReady(true);
      }
    };
    initializeDatabase();

    return () => {
      cancelled = true;
      if (toastTimeout.current) clearTimeout(toastTimeout.current);
      pieInstance.current?.destroy();
      barInstance.current?.destroy();
    };
  }, []);

  useEffect(() => {
    if (activeTab === 'settings') {
      setSettingsForm(appData.settings);
    }
  }, [activeTab]);

  const normalizeAppData = data => {
    if (!data || typeof data !== 'object' || Array.isArray(data)) {
      throw new Error('بيانات قاعدة البيانات المحلية غير صالحة.');
    }
    const normalizedData = {
      ...data,
      settings: {
        ...appData.settings,
        ...(data.settings || {}),
        workDays: Array.isArray(data.settings?.workDays) ? data.settings.workDays : appData.settings.workDays,
        holidays: Array.isArray(data.settings?.holidays) ? data.settings.holidays : []
      },
      attendance: (Array.isArray(data.attendance) ? data.attendance : [])
        .filter(rec => rec && typeof rec === 'object')
        .map(rec => ({
          ...rec,
          directorate: rec.directorate || rec.department || 'عامة',
          department: rec.directorate ? (rec.department || '') : '',
          day: getArabicDay(rec.date) || rec.day
        })),
      leaves: Array.isArray(data.leaves) ? data.leaves : [],
      employeesStatus: Object.fromEntries(
        Object.entries(data.employeesStatus && typeof data.employeesStatus === 'object' ? data.employeesStatus : {}).map(([empId, status]) => [
          empId,
          status && typeof status === 'object' && status.directorate
            ? { ...status, department: status.department || '' }
            : { ...(status && typeof status === 'object' ? status : {}), directorate: status?.department || '', department: '' }
        ])
      )
    };
    return normalizedData;
  };

  const saveToDatabase = async newData => {
    try {
      await writeStoredAppData(newData);
      return true;
    } catch (e) {
      console.error(e);
      showToastMsg("تعذر حفظ التغييرات في قاعدة البيانات المحلية. تحقق من إعدادات المتصفح ومساحة التخزين.", "error");
      return false;
    }
  };

  const createBackup = () => {
    try {
      const backup = {
        format: BACKUP_FORMAT,
        version: 1,
        exportedAt: new Date().toISOString(),
        data: appData
      };
      const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `xface-attendance-backup-${new Date().toISOString().replace(/[:.]/g, '-')}.json`;
      link.click();
      setTimeout(() => URL.revokeObjectURL(url), 0);
      showToastMsg('تم إنشاء ملف النسخة الاحتياطية بنجاح');
    } catch (error) {
      console.error(error);
      showToastMsg('تعذر إنشاء النسخة الاحتياطية.', 'error');
    }
  };

  const importBackup = async (event) => {
    const input = event.target;
    const file = input.files?.[0];
    if (!file) return;
    input.value = '';

    try {
      const backup = JSON.parse(await file.text());
      const data = backup?.data;
      const validSettings = isObjectRecord(data?.settings) &&
        typeof data.settings.startTime === 'string' &&
        typeof data.settings.endTime === 'string' &&
        Number.isFinite(data.settings.graceLate) &&
        Number.isFinite(data.settings.graceEarly) &&
        Array.isArray(data.settings.workDays) &&
        data.settings.workDays.every(day => typeof day === 'string') &&
        Array.isArray(data.settings.holidays) &&
        data.settings.holidays.every(holiday =>
          isObjectRecord(holiday) &&
          typeof holiday.id === 'string' &&
          typeof holiday.title === 'string' &&
          typeof holiday.fromDate === 'string' &&
          typeof holiday.toDate === 'string'
        );
      const validData = isObjectRecord(data) &&
        validSettings &&
        Array.isArray(data.attendance) &&
        data.attendance.every(isObjectRecord) &&
        Array.isArray(data.leaves) &&
        data.leaves.every(isObjectRecord) &&
        isObjectRecord(data.employeesStatus) &&
        Object.values(data.employeesStatus).every(isObjectRecord);

      if (backup?.format !== BACKUP_FORMAT || backup.version !== 1 || !validData) {
        throw new Error('ملف النسخة الاحتياطية غير صالح أو غير مدعوم.');
      }
      if (!await saveToDatabase(data)) return;

      setAppData(data);
      setSettingsForm(data.settings);
      setTempUploadedAttendance([]);
      setSelectedDirectoryEmployeeId('');
      showToastMsg('تم استيراد النسخة الاحتياطية واستعادة البيانات بنجاح');
    } catch (error) {
      console.error(error);
      showToastMsg(error instanceof Error ? error.message : 'تعذر استيراد النسخة الاحتياطية.', 'error');
    }
  };

  const deleteAllData = async () => {
    const confirmed = window.confirm('هل أنت متأكد من حذف جميع بيانات الحضور والإجازات والموظفين والعطل؟ لا يمكن التراجع عن هذه العملية. يُنصح بإنشاء نسخة احتياطية قبل المتابعة.');
    if (!confirmed) return;

    const emptyData = createEmptyAppData();
    if (!await saveToDatabase(emptyData)) return;

    setAppData(emptyData);
    setSettingsForm(emptyData.settings);
    setTempUploadedAttendance([]);
    setSelectedDirectoryEmployeeId('');
    setLeaveForm({ empName: '', empId: '', type: 'سنوية', leaveUnit: 'أيام', fromDate: '', toDate: '', fromTime: '', toTime: '', documentNumber: '', status: 'معتمدة', notes: '' });
    setLeaveEmployeeSearch('');
    setShowLeaveEmployeeSuggestions(false);
    setEmpModal({ isOpen: false, empId: '', status: 'نشط', date: '', directorate: '', department: '' });
    showToastMsg('تم حذف جميع البيانات بنجاح', 'warning');
  };

  const showToastMsg = (message, type = 'success') => {
    if (toastTimeout.current) clearTimeout(toastTimeout.current);
    setToast({ show: true, message, type });
    toastTimeout.current = setTimeout(() => {
      setToast(prev => ({ ...prev, show: false }));
    }, 3500);
  };

  const loadDemoData = async (notify = true) => {
    const demo = {
      settings: {
        startTime: "08:00",
        endTime: "16:00",
        graceLate: 15,
        graceEarly: 15,
        workDays: ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس"],
        holidays: [
          { id: "H1", title: "عيد الفطر السعيد", fromDate: "2026-03-20", toDate: "2026-03-23" },
          { id: "H2", title: "عيد الجلاء", fromDate: "2026-04-17", toDate: "2026-04-17" },
          { id: "H3", title: "بداية العام الدراسي", fromDate: "2026-10-06", toDate: "2026-10-06" }
        ]
      },
      attendance: [
        { empId: "1001", name: "أحمد محمد العلي", directorate: "مديرية المالية والإدارية", department: "تقنية المعلومات", date: "2026-10-07", day: "الأربعاء", firstIn: "08:10", lastOut: "16:00" },
        { empId: "1002", name: "سارة خالد الأحمد", directorate: "مديرية الموارد البشرية", department: "شؤون العاملين", date: "2026-10-07", day: "الأربعاء", firstIn: "08:35", lastOut: "15:45" },
        { empId: "1003", name: "محمود حسن إبراهيم", directorate: "مديرية المالية والإدارية", department: "المالية", date: "2026-10-07", day: "الأربعاء", firstIn: "07:55", lastOut: "16:05" },
        { empId: "1004", name: "فاطمة عمر الصالح", directorate: "مديرية التخطيط والتعاون", department: "التخطيط", date: "2026-10-07", day: "الأربعاء", firstIn: "09:15", lastOut: "15:30" },
        { empId: "1005", name: "خالد سعيد النجار", directorate: "مديرية الإدارة العامة", department: "الأمانة العامة", date: "2026-10-07", day: "الأربعاء", firstIn: "", lastOut: "" },
        { empId: "1006", name: "ريما باسم الشام", directorate: "مديرية التخطيط والتعاون", department: "التعاون الدولي", date: "2026-10-07", day: "الأربعاء", firstIn: "07:15", lastOut: "17:20" },
        { empId: "1008", name: "ليان سامي الخطيب", directorate: "مديرية الموارد البشرية", department: "التدريب والتطوير", date: "2026-10-07", day: "الأربعاء", firstIn: "08:00", lastOut: "16:00" },
        { empId: "1009", name: "عمر زهير الحسن", directorate: "مديرية الخدمات", department: "النظافة والصيانة", date: "2026-10-07", day: "الأربعاء", firstIn: "", lastOut: "" },
        { empId: "1010", name: "نور الدين مراد", directorate: "مديرية الرقابة والتدقيق", department: "التدقيق الداخلي", date: "2026-10-09", day: "الجمعة", firstIn: "09:00", lastOut: "13:00" },
        { empId: "1003", name: "محمود حسن إبراهيم", directorate: "مديرية المالية والإدارية", department: "المالية", date: "2026-10-06", day: "الثلاثاء", firstIn: "08:00", lastOut: "16:00" },
        { empId: "1011", name: "شهد مازن القاسم", directorate: "مديرية الشؤون القانونية", department: "العقود والمعاملات", date: "2026-10-08", day: "الخميس", firstIn: "", lastOut: "16:00" },
        { empId: "1001", name: "أحمد محمد العلي", directorate: "مديرية المالية والإدارية", department: "تقنية المعلومات", date: "2026-10-08", day: "الخميس", firstIn: "08:05", lastOut: "16:00" },
        { empId: "1002", name: "سارة خالد الأحمد", directorate: "مديرية الموارد البشرية", department: "شؤون العاملين", date: "2026-10-08", day: "الخميس", firstIn: "08:12", lastOut: "16:00" },
        { empId: "1003", name: "محمود حسن إبراهيم", directorate: "مديرية المالية والإدارية", department: "المالية", date: "2026-10-08", day: "الخميس", firstIn: "08:00", lastOut: "16:00" },
        { empId: "1004", name: "فاطمة عمر الصالح", directorate: "مديرية التخطيط والتعاون", department: "التخطيط", date: "2026-10-08", day: "الخميس", firstIn: "09:00", lastOut: "15:30" },
        { empId: "1007", name: "عباس جمال الشريف", directorate: "مديرية الإدارة العامة", department: "المكاتب", date: "2026-10-08", day: "الخميس", firstIn: "08:00", lastOut: "16:00" }
      ],
      leaves: [
        { id: "L1", empId: "1005", empName: "خالد سعيد النجار", type: "مرضية", leaveUnit: "أيام", fromDate: "2026-10-07", toDate: "2026-10-07", status: "معتمدة", documentNumber: "طبي-2026-14", notes: "تقرير طبي رسمي" },
        { id: "L2", empId: "1008", empName: "ليان سامي الخطيب", type: "مهمة / مغادرة رسمية", leaveUnit: "ساعية", fromDate: "2026-10-07", toDate: "2026-10-07", fromTime: "10:00", toTime: "12:00", status: "معتمدة", documentNumber: "م-2026-35", notes: "مهمة رسمية خارج المديرية" },
        { id: "L3", empId: "1009", empName: "عمر زهير الحسن", type: "سنوية", leaveUnit: "أيام", fromDate: "2026-10-07", toDate: "2026-10-08", status: "معتمدة", documentNumber: "سنوي-2026-09", notes: "إجازة سنوية اعتيادية" },
        { id: "L4", empId: "1011", empName: "شهد مازن القاسم", type: "أخرى", leaveUnit: "أيام", fromDate: "2026-10-09", toDate: "2026-10-09", status: "مرفوضة", documentNumber: "طلب-2026-41", notes: "طلب مرفوض تجريبياً" }
      ],
      employeesStatus: {
        "1001": { status: "نشط", date: "2025-01-01" },
        "1002": { status: "نشط", date: "2025-01-01" },
        "1003": { status: "نشط", date: "2025-01-01" },
        "1004": { status: "نشط", date: "2025-01-01" },
        "1005": { status: "نشط", date: "2025-01-01" },
        "1006": { status: "نشط", date: "2025-01-01" },
        "1007": { status: "منقول", date: "2026-09-15" },
        "1008": { status: "نشط", date: "2025-01-01" },
        "1009": { status: "نشط", date: "2025-01-01" },
        "1010": { status: "نشط", date: "2025-01-01" },
        "1011": { status: "نشط", date: "2025-01-01" }
      }
    };
    demo.attendance = demo.attendance.map(rec => ({
      ...rec,
      directorate: rec.directorate || rec.department,
      department: rec.department && rec.directorate ? rec.department : '',
      day: getArabicDay(rec.date)
    }));
    demo.employeesStatus = Object.fromEntries(
      Object.entries(demo.employeesStatus).map(([empId, status]) => [
        empId,
        {
          ...status,
          directorate: demo.attendance.find(rec => rec.empId === empId)?.directorate || 'عامة',
          department: ''
        }
      ])
    );
    const saved = await saveToDatabase(demo);
    if (!saved) return;
    setAppData(demo);
    setSettingsForm(demo.settings);
    if (notify && saved) showToastMsg("تم تحميل البيانات التجريبية بنجاح");
  };

  const timeToMinutes = (timeStr) => {
    if (typeof timeStr !== 'string') return null;
    const match = timeStr.match(/^(\d{1,2}):(\d{2})$/);
    if (!match) return null;
    const h = Number(match[1]);
    const m = Number(match[2]);
    if (h > 23 || m > 59) return null;
    return h * 60 + m;
  };

  const getArabicDay = (dateStr) => {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return '';
    const date = new Date(`${dateStr}T00:00:00Z`);
    if (Number.isNaN(date.getTime()) || date.toISOString().slice(0, 10) !== dateStr) return '';
    return ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"][date.getUTCDay()];
  };

  const getTodayDate = () => {
    const now = new Date();
    return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
  };

  const formatExcelDate = (value) => {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, '0')}-${String(value.getDate()).padStart(2, '0')}`;
    }
    if (typeof value === 'number' && Number.isFinite(value)) {
      if (value < 1) return '';
      const date = XLSX.SSF.parse_date_code(value);
      if (!date || !date.y || !date.m || !date.d) return '';
      return `${date.y}-${String(date.m).padStart(2, '0')}-${String(date.d).padStart(2, '0')}`;
    }
    const text = String(value ?? '').trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(text) && getArabicDay(text)) return text;
    const localizedDate = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
    if (localizedDate) {
      const [, day, month, year] = localizedDate;
      const date = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
      if (getArabicDay(date)) return date;
    }
    return '';
  };

  const formatExcelTime = (value) => {
    if (value === undefined || value === null || value === '') return '';
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return `${String(value.getHours()).padStart(2, '0')}:${String(value.getMinutes()).padStart(2, '0')}`;
    }
    if (typeof value === 'number' && Number.isFinite(value)) {
      const fraction = ((value % 1) + 1) % 1;
      const minutes = Math.round(fraction * 1440) % 1440;
      return `${String(Math.floor(minutes / 60)).padStart(2, '0')}:${String(minutes % 60).padStart(2, '0')}`;
    }
    const text = String(value).trim();
    const match = text.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?$/i);
    if (!match) return null;
    let hours = Number(match[1]);
    const minutes = Number(match[2]);
    const meridiem = match[3]?.toUpperCase();
    if (minutes > 59 || hours > (meridiem ? 12 : 23) || (meridiem && hours === 0)) return null;
    if (meridiem) hours = (hours % 12) + (meridiem === 'PM' ? 12 : 0);
    return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
  };

  const isDateInHolidays = (dateStr) => {
    if (!appData.settings.holidays) return false;
    return appData.settings.holidays.some(h => dateStr >= h.fromDate && dateStr <= h.toDate);
  };

  const getEmployeeActiveStatus = (empId, dateStr) => {
    const empStat = appData.employeesStatus[empId];
    if (!empStat) return 'نشط';
    if (empStat.status !== 'نشط' && (!empStat.date || dateStr >= empStat.date)) {
      return empStat.status;
    }
    return 'نشط';
  };

  const getEmployeeDirectorate = (empId, fallback = 'عامة') => {
    const employee = appData.employeesStatus[empId];
    return employee?.directorate || employee?.department || fallback;
  };

  const getEmployeeSection = (empId, fallback = '') => {
    return appData.employeesStatus[empId]?.department || fallback;
  };

  const getUncoveredMinutes = (start, end, leaveIntervals) => {
    if (end <= start) return 0;

    let uncovered = 0;
    let cursor = start;
    const intervals = leaveIntervals
      .map(({ start: leaveStart, end: leaveEnd }) => ({
        start: Math.max(start, leaveStart),
        end: Math.min(end, leaveEnd)
      }))
      .filter(interval => interval.end > interval.start)
      .sort((a, b) => a.start - b.start);

    intervals.forEach(interval => {
      if (interval.start > cursor) uncovered += interval.start - cursor;
      cursor = Math.max(cursor, interval.end);
    });

    return uncovered + Math.max(0, end - cursor);
  };

  const calculateAttendanceStatus = (rec) => {
    const empId = rec.empId;
    const dateStr = rec.date;

    const currentStatus = getEmployeeActiveStatus(empId, dateStr);
    if (currentStatus !== 'نشط') {
      return { status: currentStatus, lateMinutes: 0, earlyMinutes: 0, totalDiff: 0 };
    }

    if (isDateInHolidays(dateStr)) {
      return { status: "عطلة رسمية", lateMinutes: 0, earlyMinutes: 0, totalDiff: 0 };
    }

    const day = getArabicDay(dateStr);
    if (day && Array.isArray(appData.settings.workDays) && !appData.settings.workDays.includes(day)) {
      if (!rec.firstIn || !rec.lastOut) {
        return { status: "عطلة أسبوعية", lateMinutes: 0, earlyMinutes: 0, totalDiff: 0 };
      }
      const firstInM = timeToMinutes(rec.firstIn);
      const lastOutM = timeToMinutes(rec.lastOut);
      if (firstInM === null || lastOutM === null || lastOutM <= firstInM) {
        return { status: "تعذر الاحتساب", lateMinutes: 0, earlyMinutes: 0, totalDiff: 0 };
      }
      const workdayOvertimeMinutes = lastOutM - firstInM;
      return {
        status: "دوام في عطلة أسبوعية",
        lateMinutes: 0,
        earlyMinutes: 0,
        earlyOvertimeMinutes: 0,
        lateOvertimeMinutes: 0,
        overtimeMinutes: workdayOvertimeMinutes,
        workdayOvertimeMinutes,
        totalDiff: 0
      };
    }

    const activeLeave = appData.leaves.find(l =>
      l.empId === empId &&
      dateStr >= l.fromDate &&
      dateStr <= l.toDate &&
      l.status === 'معتمدة' &&
      l.leaveUnit !== 'ساعية'
    );
    if (activeLeave) {
      return { status: "إجازة", lateMinutes: 0, earlyMinutes: 0, totalDiff: 0, leaveType: activeLeave.type };
    }

    const hourlyLeaveIntervals = appData.leaves
      .filter(l => l.empId === empId && l.fromDate === dateStr && l.status === 'معتمدة' && l.leaveUnit === 'ساعية')
      .map(l => ({ start: timeToMinutes(l.fromTime), end: timeToMinutes(l.toTime) }))
      .filter(interval => interval.start !== null && interval.end !== null && interval.end > interval.start);

    const startM = timeToMinutes(appData.settings.startTime);
    const endM = timeToMinutes(appData.settings.endTime);
    if (startM === null || endM === null || endM <= startM) {
      return { status: "تعذر الاحتساب", lateMinutes: 0, earlyMinutes: 0, totalDiff: 0 };
    }
    const graceLate = appData.settings.graceLate || 0;
    const graceEarly = appData.settings.graceEarly || 0;

    const firstInM = timeToMinutes(rec.firstIn);
    const lastOutM = timeToMinutes(rec.lastOut);

    if (firstInM === null || lastOutM === null || rec.firstIn === "" || rec.lastOut === "") {
      return { status: "غياب بدون تبرير", lateMinutes: 0, earlyMinutes: 0, earlyOvertimeMinutes: 0, lateOvertimeMinutes: 0, overtimeMinutes: 0, totalDiff: 0 };
    }

    const lateUncovered = getUncoveredMinutes(startM, Math.min(firstInM, endM), hourlyLeaveIntervals);
    const earlyUncovered = getUncoveredMinutes(Math.max(lastOutM, startM), endM, hourlyLeaveIntervals);
    const lateMinutes = lateUncovered > graceLate ? lateUncovered : 0;
    const earlyMinutes = earlyUncovered > graceEarly ? earlyUncovered : 0;
    const earlyOvertimeMinutes = Math.max(0, startM - firstInM);
    const lateOvertimeMinutes = Math.max(0, lastOutM - endM);
    const overtimeMinutes = earlyOvertimeMinutes + lateOvertimeMinutes;

    let status = "حاضر";
    if (lateMinutes > 0 && earlyMinutes > 0) {
      status = "متأخر وخروج مبكر";
    } else if (lateMinutes > 0) {
      status = "متأخر";
    } else if (earlyMinutes > 0) {
      status = "خروج مبكر";
    }

    return {
      status,
      lateMinutes,
      earlyMinutes,
      earlyOvertimeMinutes,
      lateOvertimeMinutes,
      overtimeMinutes,
      totalDiff: lateMinutes + earlyMinutes
    };
  };

  const getStatusBadgeClass = (status) => {
    switch(status) {
      case 'حاضر': return 'text-emerald-700';
      case 'متأخر': return 'text-amber-700';
      case 'خروج مبكر': return 'text-amber-700';
      case 'متأخر وخروج مبكر': return 'text-rose-700';
      case 'غياب بدون تبرير': return 'text-red-700';
      case 'إجازة': return 'text-teal-700';
      case 'عطلة رسمية': return 'text-amber-800';
      case 'عطلة أسبوعية': return 'text-slate-700';
      case 'دوام في عطلة أسبوعية': return 'text-indigo-700';
      case 'تعذر الاحتساب': return 'text-red-700';
      default: return 'text-gray-700';
    }
  };

  const openReportsWithFilter = (status, directorate = '') => {
    setReportFilters(prev => ({ ...prev, status, directorate }));
    setActiveTab('reports');
  };

  const dashboardStats = useMemo(() => {
    let activeCount = 0;
    let presentCount = 0;
    let absentCount = 0;
    let lateCount = 0;
    let overtimeMinutes = 0;

    const employeeIds = new Set([
      ...appData.attendance.map(rec => rec.empId),
      ...Object.keys(appData.employeesStatus)
    ]);
    activeCount = [...employeeIds].filter(empId => getEmployeeActiveStatus(empId, getTodayDate()) === 'نشط').length;

    appData.attendance.forEach(rec => {
      if (getEmployeeActiveStatus(rec.empId, rec.date) !== 'نشط') return;
      const calc = calculateAttendanceStatus(rec);
      overtimeMinutes += calc.overtimeMinutes || 0;
      if (calc.status === 'حاضر' || calc.status === 'دوام في عطلة أسبوعية') presentCount++;
      if (calc.status === 'متأخر' || calc.status === 'خروج مبكر' || calc.status === 'متأخر وخروج مبكر') {
        presentCount++;
        lateCount++;
      }
      if (calc.status === 'غياب بدون تبرير') absentCount++;
    });

    return { activeCount, presentCount, absentCount, lateCount, overtimeMinutes };
  }, [appData]);

  useEffect(() => {
    if (activeTab === 'dashboard') {
      let present = 0, late = 0, absent = 0, leave = 0;
      appData.attendance.forEach(rec => {
        if (getEmployeeActiveStatus(rec.empId, rec.date) !== 'نشط') return;
        const calc = calculateAttendanceStatus(rec);
        if (calc.status === 'حاضر' || calc.status === 'دوام في عطلة أسبوعية') present++;
        else if (calc.status.includes('متأخر') || calc.status.includes('مبكر')) late++;
        else if (calc.status === 'غياب بدون تبرير') absent++;
        else if (calc.status === 'إجازة' || calc.status === 'عطلة رسمية' || calc.status === 'عطلة أسبوعية') leave++;
      });

      if (pieInstance.current) pieInstance.current.destroy();
      if (pieChartRef.current) {
        pieInstance.current = new Chart(pieChartRef.current, {
          type: 'doughnut',
          data: {
            labels: ['حضور، بما فيه دوام العطلة الأسبوعية', 'متأخر / مبكر', 'غياب بدون تبرير', 'إجازة / عطلة / راحة أسبوعية'],
            datasets: [{
              data: [present, late, absent, leave],
              backgroundColor: ['#054239', '#c18642', '#4a151e', '#b9a779'],
              borderWidth: 2,
              borderColor: '#ffffff'
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            onClick: (_event, elements) => {
              if (!elements.length) return;
              const filters = [
                'حضور',
                'متأخر/مبكر',
                'غياب بدون تبرير',
                'إجازة/عطلة'
              ];
              openReportsWithFilter(filters[elements[0].index]);
            },
            plugins: { legend: { position: 'bottom', labels: { font: { family: 'Cairo', size: 11 } } } }
          }
        });
      }

      const deptStats = {};
      appData.attendance.forEach(rec => {
        if (getEmployeeActiveStatus(rec.empId, rec.date) !== 'نشط') return;
        const calc = calculateAttendanceStatus(rec);
        const directorate = getEmployeeDirectorate(rec.empId, rec.directorate || rec.department);
        if (!deptStats[directorate]) deptStats[directorate] = { late: 0, absent: 0 };
        if (calc.status.includes('متأخر') || calc.status.includes('مبكر')) deptStats[directorate].late++;
        if (calc.status === 'غياب بدون تبرير') deptStats[directorate].absent++;
      });

      const deptLabels = Object.keys(deptStats);
      const lateData = deptLabels.map(d => deptStats[d].late);
      const absentData = deptLabels.map(d => deptStats[d].absent);

      if (barInstance.current) barInstance.current.destroy();
      if (barChartRef.current) {
        barInstance.current = new Chart(barChartRef.current, {
          type: 'bar',
          data: {
            labels: deptLabels,
            datasets: [
              { label: 'حالات التأخير والخروج المبكر', data: lateData, backgroundColor: '#c18642', borderRadius: 6 },
              { label: 'حالات الغياب', data: absentData, backgroundColor: '#4a151e', borderRadius: 6 }
            ]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            onClick: (_event, elements) => {
              if (!elements.length) return;
              const element = elements[0];
              const status = element.datasetIndex === 0 ? 'متأخر/مبكر' : 'غياب بدون تبرير';
              openReportsWithFilter(status, deptLabels[element.index]);
            },
            scales: {
              x: { ticks: { font: { family: 'Cairo', size: 10 } } },
              y: { ticks: { font: { family: 'Cairo', size: 10 } }, beginAtZero: true }
            },
            plugins: { legend: { position: 'bottom', labels: { font: { family: 'Cairo', size: 11 } } } }
          }
        });
      }
    }
  }, [activeTab, appData]);

  const handleFileUpload = (e) => {
    const input = e.target;
    const file = input.files[0];
    if (!file) return;
    setTempUploadedAttendance([]);

    const reader = new FileReader();
    reader.onload = function(evt) {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const firstSheetName = workbook.SheetNames[0];
        if (!firstSheetName) {
          showToastMsg("الملف لا يحتوي على ورقة بيانات", "error");
          return;
        }
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

        if (json.length < 2) {
          showToastMsg("الملف فارغ أو لا يحتوي على بيانات كافية", "error");
          return;
        }

        const normalizeHeader = value => String(value ?? '').replace(/[\s_-]/g, '').toLowerCase();
        const header = json[0].map(normalizeHeader);
        const rows = json.slice(1);
        const parsed = [];
        let invalidRows = 0;

        rows.forEach(row => {
          if (!row.some(value => value !== '' && value !== null && value !== undefined)) return;
          const getValue = (aliases, fallbackIndex) => {
            const aliasIndex = header.findIndex(column => aliases.includes(column));
            return aliasIndex >= 0 ? row[aliasIndex] : row[fallbackIndex];
          };
          const textValue = value => String(value ?? '').trim();
          const empId = textValue(getValue(['رقمالموظف', 'empid', 'employeeid', 'id'], 0));
          const name = textValue(getValue(['اسمالموظف', 'الاسم', 'name', 'employeename'], 1));
          const directorate = textValue(getValue(['المديرية', 'مديريةالموظف', 'directorate', 'department', 'dept'], 2)) || 'عامة';
          const department = textValue(getValue(['القسم', 'قسمالموظف', 'departmentsection', 'employeedepartment', 'section', 'unit'], -1));
          const date = formatExcelDate(getValue(['التاريخ', 'date'], 3));
          const firstIn = formatExcelTime(getValue(['أولدخول', 'firstin', 'in', 'checkin'], 5));
          const lastOut = formatExcelTime(getValue(['آخرخروج', 'lastout', 'out', 'checkout'], 6));

          if (!empId || !name || !date || firstIn === null || lastOut === null) {
            invalidRows++;
            return;
          }
          parsed.push({
            empId,
            name,
            directorate,
            department,
            date,
            day: getArabicDay(date),
            firstIn,
            lastOut
          });
        });

        if (parsed.length === 0) {
          showToastMsg("لم تُستورد أي سجلات صالحة. تحقق من رقم الموظف والاسم والتاريخ ووقتَي الدخول والخروج.", "error");
          return;
        }
        setTempUploadedAttendance(parsed);
        showToastMsg(`تمت قراءة ${parsed.length} سجل${invalidRows ? `، وتجاهل ${invalidRows} صف غير صالح` : ''}`, invalidRows ? 'warning' : 'success');
      } catch (err) {
        console.error(err);
        showToastMsg("خطأ في قراءة ملف الإكسل", "error");
      }
    };
    reader.onerror = () => showToastMsg("تعذر فتح الملف المحدد", "error");
    reader.readAsArrayBuffer(file);
    input.value = '';
  };

  const processAndSaveUploadedData = async () => {
    if (tempUploadedAttendance.length === 0) return;
    const updatedAttendance = [...appData.attendance, ...tempUploadedAttendance];
    const updatedStatus = { ...appData.employeesStatus };

    tempUploadedAttendance.forEach(r => {
      if (!updatedStatus[r.empId]) {
        updatedStatus[r.empId] = { status: "نشط", date: "2025-01-01", directorate: r.directorate, department: r.department };
      } else {
        updatedStatus[r.empId] = {
          ...updatedStatus[r.empId],
          directorate: r.directorate || updatedStatus[r.empId].directorate,
          department: r.department || updatedStatus[r.empId].department || ''
        };
      }
    });

    const newAppData = { ...appData, attendance: updatedAttendance, employeesStatus: updatedStatus };
    if (!await saveToDatabase(newAppData)) return;
    setAppData(newAppData);
    setTempUploadedAttendance([]);
    showToastMsg("تم اعتماد وحفظ بيانات الحضور بنجاح");
    setActiveTab('dashboard');
  };

  const downloadSampleExcel = () => {
    const sampleData = [
      { رقم_الموظف: "1001", اسم_الموظف: "أحمد محمد العلي", المديرية: "تقنية المعلومات", القسم: "الدعم الفني", التاريخ: "2026-10-07", أول_دخول: "08:10", آخر_خروج: "16:00" },
      { رقم_الموظف: "1002", اسم_الموظف: "سارة خالد الأحمد", المديرية: "الموارد البشرية", القسم: "شؤون العاملين", التاريخ: "2026-10-08", أول_دخول: "08:35", آخر_خروج: "15:45" }
    ];
    const ws = XLSX.utils.json_to_sheet(sampleData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "AttendanceSample");
    XLSX.writeFile(wb, "Xface_Attendance_Sample.xlsx");
    showToastMsg("تم تحميل نموذج الإكسل بنجاح");
  };

  const employeesListForDatalist = useMemo(() => {
    const empMap = {};
    appData.attendance.forEach(rec => {
      empMap[rec.empId] = rec.name;
    });
    Object.keys(appData.employeesStatus).forEach(empId => {
      if (!empMap[empId]) empMap[empId] = `موظف رقم ${empId}`;
    });
    return Object.entries(empMap).map(([id, name]) => ({ id, name }));
  }, [appData]);

  const filteredLeaveEmployees = employeesListForDatalist.filter(employee => {
    const query = leaveEmployeeSearch.trim().toLocaleLowerCase('ar');
    return !query ||
      employee.name.toLocaleLowerCase('ar').includes(query) ||
      employee.id.toLocaleLowerCase('ar').includes(query);
  });

  const selectLeaveEmployee = (employee) => {
    setLeaveForm(prev => ({ ...prev, empId: employee.id, empName: employee.name }));
    setLeaveEmployeeSearch(employee.name);
    setShowLeaveEmployeeSuggestions(false);
    setActiveLeaveEmployeeSuggestion(-1);
  };

  const saveLeaveRecord = async (e) => {
    e.preventDefault();
    if (!leaveForm.empId) {
      showToastMsg("الرجاء اختيار موظف صحيح من القائمة", "error");
      return;
    }
    if (!leaveForm.fromDate || (leaveForm.leaveUnit !== 'ساعية' && !leaveForm.toDate)) {
      showToastMsg("الرجاء إدخال تاريخ الإجازة", "error");
      return;
    }
    if (leaveForm.leaveUnit === 'ساعية') {
      const fromTime = timeToMinutes(leaveForm.fromTime);
      const toTime = timeToMinutes(leaveForm.toTime);
      if (fromTime === null || toTime === null || toTime <= fromTime) {
        showToastMsg("وقت نهاية الإجازة الساعية يجب أن يكون بعد وقت بدايتها", "error");
        return;
      }
    } else if (leaveForm.toDate < leaveForm.fromDate) {
      showToastMsg("تاريخ نهاية الإجازة يجب ألا يسبق تاريخ بدايتها", "error");
      return;
    }
    const conflictingLeave = appData.leaves.some(existing => {
      if (existing.empId !== leaveForm.empId || existing.status !== 'معتمدة') return false;
      const currentHourly = leaveForm.leaveUnit === 'ساعية';
      const existingHourly = existing.leaveUnit === 'ساعية';
      const currentEndDate = currentHourly ? leaveForm.fromDate : leaveForm.toDate;
      const existingStartDate = existing.fromDate || '';
      const existingEndDate = existing.toDate || existingStartDate;
      if (!existingStartDate || !existingEndDate) return true;
      if (leaveForm.fromDate > existingEndDate || currentEndDate < existingStartDate) return false;
      if (!currentHourly || !existingHourly) return true;
      const existingStartTime = timeToMinutes(existing.fromTime);
      const existingEndTime = timeToMinutes(existing.toTime);
      if (existingStartTime === null || existingEndTime === null) return true;
      return timeToMinutes(leaveForm.fromTime) < existingEndTime &&
        timeToMinutes(leaveForm.toTime) > existingStartTime;
    });
    if (leaveForm.status === 'معتمدة' && conflictingLeave) {
      showToastMsg("تتداخل هذه الإجازة مع إجازة معتمدة مسجلة للموظف", "error");
      return;
    }
    const newLeave = {
      id: 'L_' + Date.now(),
      ...leaveForm,
      toDate: leaveForm.leaveUnit === 'ساعية' ? leaveForm.fromDate : leaveForm.toDate,
      documentNumber: leaveForm.documentNumber.trim()
    };
    const newAppData = { ...appData, leaves: [...appData.leaves, newLeave] };
    if (!await saveToDatabase(newAppData)) return;
    setAppData(newAppData);
    setLeaveForm({ empName: '', empId: '', type: 'سنوية', leaveUnit: 'أيام', fromDate: '', toDate: '', fromTime: '', toTime: '', documentNumber: '', status: 'معتمدة', notes: '' });
    setLeaveEmployeeSearch('');
    setShowLeaveEmployeeSuggestions(false);
    showToastMsg("تم حفظ الإجازة بنجاح");
  };

  const deleteLeaveRecord = async (id) => {
    const newAppData = { ...appData, leaves: appData.leaves.filter(l => l.id !== id) };
    if (!await saveToDatabase(newAppData)) return;
    setAppData(newAppData);
    showToastMsg("تم حذف الإجازة بنجاح");
  };

  const saveEmployeeStatusModal = async (e) => {
    e.preventDefault();
    const updatedStatus = {
      ...appData.employeesStatus,
      [empModal.empId]: {
        ...appData.employeesStatus[empModal.empId],
        status: empModal.status,
        date: empModal.date,
        directorate: empModal.directorate.trim(),
        department: empModal.department.trim()
      }
    };
    const newAppData = { ...appData, employeesStatus: updatedStatus };
    if (!await saveToDatabase(newAppData)) return;
    setAppData(newAppData);
    setEmpModal({ isOpen: false, empId: '', status: 'نشط', date: '', directorate: '', department: '' });
    showToastMsg("تم تحديث بيانات الموظف بنجاح");
  };

  const saveSettings = async (e) => {
    e.preventDefault();
    const graceLate = Number(settingsForm.graceLate);
    const graceEarly = Number(settingsForm.graceEarly);
    const startTime = timeToMinutes(settingsForm.startTime);
    const endTime = timeToMinutes(settingsForm.endTime);
    if (startTime === null || endTime === null || endTime <= startTime) {
      showToastMsg("تحقق من وقتَي الدوام؛ يجب أن تكون نهاية الدوام بعد بدايته في اليوم نفسه", "error");
      return;
    }
    if (!Number.isInteger(graceLate) || graceLate < 0 || !Number.isInteger(graceEarly) || graceEarly < 0) {
      showToastMsg("فترات السماح يجب أن تكون أعداداً صحيحة غير سالبة", "error");
      return;
    }
    if (!Array.isArray(settingsForm.workDays) || settingsForm.workDays.length === 0) {
      showToastMsg("حدد يوماً واحداً على الأقل من أيام الدوام", "error");
      return;
    }
    const newSettings = {
      ...appData.settings,
      ...settingsForm,
      graceLate,
      graceEarly,
      workDays: [...new Set(settingsForm.workDays)].filter(day => ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"].includes(day))
    };
    const newAppData = { ...appData, settings: newSettings };
    if (!await saveToDatabase(newAppData)) return;
    setAppData(newAppData);
    showToastMsg("تم حفظ إعدادات الدوام بنجاح");
  };

  const addHolidayRecord = async (e) => {
    e.preventDefault();
    if (!holidayForm.title.trim() || !holidayForm.fromDate || !holidayForm.toDate || holidayForm.toDate < holidayForm.fromDate) {
      showToastMsg("تحقق من اسم العطلة وتاريخَي البداية والنهاية", "error");
      return;
    }
    const newHoliday = {
      id: 'H_' + Date.now(),
      ...holidayForm,
      title: holidayForm.title.trim()
    };
    const holidays = [...(appData.settings.holidays || []), newHoliday];
    const newAppData = { ...appData, settings: { ...appData.settings, holidays } };
    if (!await saveToDatabase(newAppData)) return;
    setAppData(newAppData);
    setHolidayForm({ title: '', fromDate: '', toDate: '' });
    showToastMsg("تم إضافة العطلة الرسمية بنجاح");
  };

  const deleteHolidayRecord = async (id) => {
    const holidays = (appData.settings.holidays || []).filter(h => h.id !== id);
    const newAppData = { ...appData, settings: { ...appData.settings, holidays } };
    if (!await saveToDatabase(newAppData)) return;
    setAppData(newAppData);
    showToastMsg("تم حذف العطلة بنجاح");
  };

  const filteredAttendanceRecords = appData.attendance.filter(rec => {
    if (getEmployeeActiveStatus(rec.empId, rec.date) !== 'نشط') return false;
    const calc = calculateAttendanceStatus(rec);
    const statusMatches = !reportFilters.status ||
      (reportFilters.status === 'متأخر/مبكر'
        ? calc.status.includes('متأخر') || calc.status.includes('مبكر')
        : reportFilters.status === 'إجازة/عطلة'
          ? calc.status === 'إجازة' || calc.status === 'عطلة رسمية' || calc.status === 'عطلة أسبوعية'
          : reportFilters.status === 'حضور'
              ? calc.status === 'حاضر' || calc.status === 'دوام في عطلة أسبوعية' || calc.status.includes('متأخر') || calc.status.includes('مبكر')
            : reportFilters.status === 'إضافي'
              ? (calc.overtimeMinutes || 0) > 0
              : calc.status === reportFilters.status);
    if (!statusMatches) return false;
    if (reportFilters.search && !rec.name.toLowerCase().includes(reportFilters.search.toLowerCase()) && !rec.empId.includes(reportFilters.search)) return false;
    if (reportFilters.directorate && getEmployeeDirectorate(rec.empId, rec.directorate || rec.department) !== reportFilters.directorate) return false;
    if (reportFilters.department && getEmployeeSection(rec.empId, rec.department) !== reportFilters.department) return false;
    if (reportFilters.startDate && rec.date < reportFilters.startDate) return false;
    if (reportFilters.endDate && rec.date > reportFilters.endDate) return false;
    return true;
  });

  const employeeDirectory = useMemo(() => {
    const employees = new Map();
    const ensureEmployee = (empId) => {
      const normalizedEmpId = String(empId || '');
      if (!employees.has(normalizedEmpId)) {
        const status = appData.employeesStatus[normalizedEmpId] || {};
        employees.set(normalizedEmpId, {
          empId: normalizedEmpId,
          name: '',
          directorate: status.directorate || status.department || 'عامة',
          department: status.department || '',
          status: status.status || 'نشط'
        });
      }
      return employees.get(normalizedEmpId);
    };

    appData.attendance.forEach(record => {
      const employee = ensureEmployee(record.empId);
      if (!employee.name && record.name) employee.name = record.name;
      if (!employee.directorate || employee.directorate === 'عامة') {
        employee.directorate = record.directorate || record.department || employee.directorate;
      }
      if (!employee.department && record.directorate) employee.department = record.department || '';
    });

    appData.leaves.forEach(leave => {
      const employee = ensureEmployee(leave.empId);
      if (!employee.name && leave.empName) employee.name = leave.empName;
    });

    Object.entries(appData.employeesStatus).forEach(([empId, status]) => {
      const employee = ensureEmployee(empId);
      employee.status = status?.status || employee.status;
      employee.directorate = status?.directorate || employee.directorate;
      employee.department = status?.department || employee.department;
    });

    return [...employees.values()]
      .map(employee => ({
        ...employee,
        name: employee.name || `موظف رقم ${employee.empId}`,
        attendanceCount: appData.attendance.filter(record => String(record.empId) === employee.empId).length
      }))
      .sort((a, b) => a.name.localeCompare(b.name, 'ar', { numeric: true }) || a.empId.localeCompare(b.empId, 'ar', { numeric: true }));
  }, [appData.attendance, appData.leaves, appData.employeesStatus]);

  const filteredEmployeeDirectory = employeeDirectory.filter(employee =>
    !employeeDirectorySearch ||
    employee.name.toLocaleLowerCase('ar').includes(employeeDirectorySearch.toLocaleLowerCase('ar')) ||
    employee.empId.includes(employeeDirectorySearch.trim())
  );

  const selectedDirectoryEmployee = employeeDirectory.find(employee => employee.empId === selectedDirectoryEmployeeId);
  const selectedEmployeeAttendance = selectedDirectoryEmployeeId
    ? appData.attendance.filter(record => String(record.empId) === selectedDirectoryEmployeeId)
    : [];

  const exportEmployeeAttendanceExcel = (employee) => {
    const employeeRecords = appData.attendance
      .filter(record => String(record.empId) === String(employee.empId))
      .slice()
      .sort((a, b) => b.date.localeCompare(a.date));

    if (employeeRecords.length === 0) {
      showToastMsg("لا توجد سجلات حضور لهذا الموظف", "warning");
      return;
    }

    const exportData = employeeRecords.map(record => {
      const calculation = calculateAttendanceStatus(record);
      return {
        "رقم الموظف": record.empId,
        "اسم الموظف": record.name,
        "المديرية": getEmployeeDirectorate(record.empId, record.directorate || record.department),
        "القسم": getEmployeeSection(record.empId, record.department),
        "التاريخ": record.date,
        "اليوم": getArabicDay(record.date),
        "أول دخول": record.firstIn || '-',
        "آخر خروج": record.lastOut || '-',
        "الحالة": calculation.status,
        "التأخير (دقيقة)": calculation.lateMinutes || 0,
        "الخروج المبكر (دقيقة)": calculation.earlyMinutes || 0,
        "إضافي قبل الدوام (دقيقة)": calculation.earlyOvertimeMinutes || 0,
        "إضافي بعد الدوام (دقيقة)": calculation.lateOvertimeMinutes || 0,
        "إجمالي الوقت الإضافي (دقيقة)": calculation.overtimeMinutes || 0
      };
    });

    try {
      const ws = XLSX.utils.json_to_sheet(exportData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "EmployeeAttendance");
      const safeName = String(employee.name).replace(/[\\/:*?"<>|]/g, '-').trim() || employee.empId;
      XLSX.writeFile(wb, `سجل_حضور_${safeName}_${employee.empId}.xlsx`);
      showToastMsg(`تم تصدير سجل حضور ${employee.name} بنجاح`);
    } catch (err) {
      console.error(err);
      showToastMsg("تعذر تصدير سجل حضور الموظف", "error");
    }
  };

  const exportReportsExcel = () => {
    const exportData = [];
    const recordsToExport = activeTab === 'reports' ? filteredAttendanceRecords : appData.attendance;
    recordsToExport.forEach(rec => {
      const activeStat = getEmployeeActiveStatus(rec.empId, rec.date);
      if (activeStat !== 'نشط') return;
      const calc = calculateAttendanceStatus(rec);
      exportData.push({
        "رقم الموظف": rec.empId,
        "اسم الموظف": rec.name,
        "المديرية": getEmployeeDirectorate(rec.empId, rec.directorate || rec.department),
        "القسم": getEmployeeSection(rec.empId, rec.department),
        "التاريخ": rec.date,
        "أول دخول": rec.firstIn,
        "آخر خروج": rec.lastOut,
        "الحالة": calc.status,
        "التأخير (دقيقة)": calc.lateMinutes,
        "الخروج المبكر (دقيقة)": calc.earlyMinutes,
        "إضافي قبل الدوام (دقيقة)": calc.earlyOvertimeMinutes || 0,
        "إضافي بعد الدوام (دقيقة)": calc.lateOvertimeMinutes || 0,
        "إجمالي الوقت الإضافي (دقيقة)": calc.overtimeMinutes || 0
      });
    });

    if (exportData.length === 0) {
      showToastMsg("لا توجد سجلات مطابقة للتصدير", "warning");
      return;
    }
    try {
      const ws = XLSX.utils.json_to_sheet(exportData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "AttendanceReports");
      XLSX.writeFile(wb, "Xface_Detailed_Attendance_Report.xlsx");
      showToastMsg("تم تصدير التقرير إلى Excel بنجاح");
    } catch (err) {
      console.error(err);
      showToastMsg("تعذر إنشاء ملف التقرير", "error");
    }
  };

  if (!isDataReady) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-[#edebe0] text-[#002623]" dir="rtl" role="status" aria-live="polite">
        <div className="flex items-center gap-3 rounded-2xl bg-white px-6 py-5 shadow-sm border border-[#b9a779]/50">
          <i className="fa-solid fa-database text-[#054239]" aria-hidden="true"></i>
          <span className="text-sm font-semibold">جارٍ تحميل البيانات المحلية...</span>
        </div>
      </div>
    );
  }

  return (
    <div className="bg-[#edebe0]/50 text-[#3d3a3b] antialiased min-h-screen flex flex-col font-sans" dir="rtl">
      <div className="flex h-screen overflow-hidden">
        
        {/* Sidebar Navigation */}
        {sidebarOpen && <button type="button" aria-label="إغلاق القائمة" onClick={() => setSidebarOpen(false)} className="fixed inset-0 bg-[#002623]/50 z-10 md:hidden" />}
        <aside id="main-navigation" data-collapsed={sidebarCollapsed} className={`w-64 ${sidebarCollapsed ? 'md:w-20' : 'md:w-64'} bg-[#002623] text-[#edebe0] flex flex-col flex-shrink-0 shadow-2xl z-20 transition-all duration-300 ${sidebarOpen ? 'translate-x-0' : 'translate-x-full md:translate-x-0'} fixed md:static inset-y-0 right-0`}>
          <div className="p-5 flex items-center justify-between border-b border-[#054239]/60">
            <div className="flex items-center space-x-3 space-x-reverse">
              <button
                type="button"
                className="sidebar-brand-mark"
                title={sidebarCollapsed ? "توسيع القائمة الجانبية" : "طي القائمة الجانبية"}
                aria-label={sidebarCollapsed ? "توسيع القائمة الجانبية" : "طي القائمة الجانبية"}
                aria-expanded={!sidebarCollapsed}
                aria-controls="main-navigation"
                onClick={() => setSidebarCollapsed(collapsed => !collapsed)}
              >
                <img src={syrianEagleLogo} alt="شعار النسر السوري" />
              </button>
                
              <div className={sidebarCollapsed ? 'md:hidden' : ''}>
                <h1 className="font-bold text-white text-base">مديرية التنمية الإدارية</h1>
                <p className="text-xs text-[#edebe0]/80">محافظة حمص</p>
              </div>
            </div>
          </div>

          <nav className="flex-1 px-4 py-6 space-y-1.5 overflow-y-auto">
            <button type="button" title="لوحة القيادة" aria-label="لوحة القيادة" aria-current={activeTab === 'dashboard' ? 'page' : undefined} onClick={() => { setActiveTab('dashboard'); setSidebarOpen(false); }} className={`w-full flex items-center ${sidebarCollapsed ? 'md:justify-center md:px-2' : 'px-4'} py-3 text-sm font-medium rounded-xl transition-all ${activeTab === 'dashboard' ? 'bg-[#054239] text-[#edebe0] shadow-md border border-[#b9a779]/20' : 'hover:bg-[#054239]/50 hover:text-white text-[#edebe0]/90'}`}>
              <i className="fa-solid fa-chart-pie w-6 text-lg text-[#b9a779]"></i>
              <span className={sidebarCollapsed ? 'md:hidden' : ''}>لوحة القيادة</span>
            </button>
            <button type="button" title="رفع وتحليل ملف البصمة" aria-label="رفع وتحليل ملف البصمة" aria-current={activeTab === 'attendance' ? 'page' : undefined} onClick={() => { setActiveTab('attendance'); setSidebarOpen(false); }} className={`w-full flex items-center ${sidebarCollapsed ? 'md:justify-center md:px-2' : 'px-4'} py-3 text-sm font-medium rounded-xl transition-all ${activeTab === 'attendance' ? 'bg-[#054239] text-[#edebe0] shadow-md border border-[#b9a779]/20' : 'hover:bg-[#054239]/50 hover:text-white text-[#edebe0]/90'}`}>
              <i className="fa-solid fa-file-excel w-6 text-lg text-[#b9a779] font-bold"></i>
              <span className={sidebarCollapsed ? 'md:hidden' : ''}>رفع وتحليل ملف البصمة</span>
            </button>
            <button type="button" title="الإجازات والمغادرات" aria-label="الإجازات والمغادرات" aria-current={activeTab === 'leaves' ? 'page' : undefined} onClick={() => { setActiveTab('leaves'); setSidebarOpen(false); }} className={`w-full flex items-center ${sidebarCollapsed ? 'md:justify-center md:px-2' : 'px-4'} py-3 text-sm font-medium rounded-xl transition-all ${activeTab === 'leaves' ? 'bg-[#054239] text-[#edebe0] shadow-md border border-[#b9a779]/20' : 'hover:bg-[#054239]/50 hover:text-white text-[#edebe0]/90'}`}>
              <i className="fa-solid fa-calendar-days w-6 text-lg text-[#988561]"></i>
              <span className={sidebarCollapsed ? 'md:hidden' : ''}>الإجازات والمغادرات</span>
            </button>
            <button type="button" title="حالة الموظفين الوظيفية" aria-label="حالة الموظفين الوظيفية" aria-current={activeTab === 'employees' ? 'page' : undefined} onClick={() => { setActiveTab('employees'); setSidebarOpen(false); }} className={`w-full flex items-center ${sidebarCollapsed ? 'md:justify-center md:px-2' : 'px-4'} py-3 text-sm font-medium rounded-xl transition-all ${activeTab === 'employees' ? 'bg-[#054239] text-[#edebe0] shadow-md border border-[#b9a779]/20' : 'hover:bg-[#054239]/50 hover:text-white text-[#edebe0]/90'}`}>
              <i className="fa-solid fa-users-gear w-6 text-lg text-[#428177]"></i>
              <span className={sidebarCollapsed ? 'md:hidden' : ''}>حالة الموظفين الوظيفية</span>
            </button>
            <button type="button" title="جميع الموظفين وسجلاتهم" aria-label="جميع الموظفين وسجلاتهم" aria-current={activeTab === 'employeeDirectory' ? 'page' : undefined} onClick={() => { setActiveTab('employeeDirectory'); setSelectedDirectoryEmployeeId(''); setSidebarOpen(false); }} className={`w-full flex items-center ${sidebarCollapsed ? 'md:justify-center md:px-2' : 'px-4'} py-3 text-sm font-medium rounded-xl transition-all ${activeTab === 'employeeDirectory' ? 'bg-[#054239] text-[#edebe0] shadow-md border border-[#b9a779]/20' : 'hover:bg-[#054239]/50 hover:text-white text-[#edebe0]/90'}`}>
              <i className="fa-solid fa-address-book w-6 text-lg text-[#b9a779]"></i>
              <span className={sidebarCollapsed ? 'md:hidden' : ''}>جميع الموظفين وسجلاتهم</span>
            </button>
            <button type="button" title="التقارير التفصيلية والمخالفات" aria-label="التقارير التفصيلية والمخالفات" aria-current={activeTab === 'reports' ? 'page' : undefined} onClick={() => { setActiveTab('reports'); setSidebarOpen(false); }} className={`w-full flex items-center ${sidebarCollapsed ? 'md:justify-center md:px-2' : 'px-4'} py-3 text-sm font-medium rounded-xl transition-all ${activeTab === 'reports' ? 'bg-[#054239] text-[#edebe0] shadow-md border border-[#b9a779]/20' : 'hover:bg-[#054239]/50 hover:text-white text-[#edebe0]/90'}`}>
              <i className="fa-solid fa-file-lines w-6 text-lg text-[#c18642]"></i>
              <span className={sidebarCollapsed ? 'md:hidden' : ''}>التقارير التفصيلية والمخالفات</span>
            </button>
            <button type="button" title="إعدادات الدوام والعطل" aria-label="إعدادات الدوام والعطل" aria-current={activeTab === 'settings' ? 'page' : undefined} onClick={() => { setActiveTab('settings'); setSidebarOpen(false); }} className={`w-full flex items-center ${sidebarCollapsed ? 'md:justify-center md:px-2' : 'px-4'} py-3 text-sm font-medium rounded-xl transition-all ${activeTab === 'settings' ? 'bg-[#054239] text-[#edebe0] shadow-md border border-[#b9a779]/20' : 'hover:bg-[#054239]/50 hover:text-white text-[#edebe0]/90'}`}>
              <i className="fa-solid fa-gears w-6 text-lg text-[#6b1f2a]"></i>
              <span className={sidebarCollapsed ? 'md:hidden' : ''}>إعدادات الدوام والعطل</span>
            </button>
          </nav>

          <div className="p-3 border-t border-[#054239]/60 text-xs text-[#edebe0]/70 flex items-center justify-between gap-2">
            <span className={sidebarCollapsed ? 'md:hidden' : ''}>Developed by Eng.Muhammad Wazier</span>
          </div>
        </aside>

        {/* Main Content */}
        <main className="flex-1 flex flex-col overflow-hidden bg-[#edebe0]/40">
          <header className="bg-white shadow-sm min-h-16 flex flex-wrap items-center justify-between gap-2 px-3 sm:px-6 py-2 z-10 border-b border-[#b9a779]/40">
            <div className="flex items-center space-x-4 space-x-reverse">
              <button type="button" aria-label={sidebarOpen ? "إغلاق القائمة" : "فتح القائمة"} aria-expanded={sidebarOpen} aria-controls="main-navigation" onClick={() => setSidebarOpen(!sidebarOpen)} className="text-[#3d3a3b] hover:text-[#054239] focus:outline-none focus:ring-2 focus:ring-[#054239] rounded md:hidden">
                <i className="fa-solid fa-bars text-xl" aria-hidden="true"></i>
              </button>
              <h2 className="text-sm sm:text-lg font-bold text-[#002623]">
                {activeTab === 'dashboard' && 'لوحة القيادة والتحليلات العامة'}
                {activeTab === 'attendance' && 'رفع وتحليل ملف البصمة (Xface 2000)'}
                {activeTab === 'leaves' && 'إدارة الإجازات والمغادرات'}
                {activeTab === 'employees' && 'حالة الموظفين الوظيفية'}
                {activeTab === 'employeeDirectory' && 'دليل الموظفين وسجلات الحضور'}
                {activeTab === 'reports' && 'التقارير التفصيلية والمخالفات الزمنية'}
                {activeTab === 'settings' && 'إعدادات الدوام والعطل الرسمية'}
              </h2>
            </div>
            <div className="flex items-center space-x-3 space-x-reverse">
              <button type="button" onClick={() => {
                if (appData.attendance.length || appData.leaves.length) {
                  if (!window.confirm("تحميل البيانات التجريبية سيستبدل بيانات الحضور والإجازات الحالية. هل تريد المتابعة؟")) return;
                }
                loadDemoData();
              }} className="px-2 sm:px-3.5 py-2 bg-[#edebe0] hover:bg-[#b9a779]/30 text-[#002623] text-xs font-semibold rounded-lg transition-all flex items-center space-x-1.5 space-x-reverse border border-[#b9a779]">
                <i className="fa-solid fa-database text-[#054239]"></i>
                <span>تحميل بيانات تجريبية</span>
              </button>
              <button type="button" onClick={exportReportsExcel} className="px-2 sm:px-3.5 py-2 bg-[#054239] hover:bg-[#002623] text-[#edebe0] text-xs font-semibold rounded-lg transition-all flex items-center space-x-1.5 space-x-reverse shadow-sm border border-[#428177]">
                <i className="fa-solid fa-file-excel text-[#edebe0]"></i>
                <span>تصدير التقارير</span>
              </button>
            </div>
          </header>

          <div className="app-content flex-1 overflow-y-auto p-6 space-y-6">
            
            {/* DASHBOARD VIEW */}
            {activeTab === 'dashboard' && (
              <div className="space-y-6">
                <div className="dashboard-filter-grid grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-5">
                  <button type="button" onClick={() => setActiveTab('employees')} className="text-right bg-white p-5 rounded-2xl shadow-sm border border-[#b9a779]/40 flex items-center justify-between hover:-translate-y-1 hover:shadow-md transition-all focus:outline-none focus:ring-2 focus:ring-[#054239]">
                    <div>
                      <p className="text-xs font-semibold text-[#3d3a3b]/70">الموظفون النشطون حالياً</p>
                      <h3 className="text-2xl font-black text-[#002623] mt-1">{dashboardStats.activeCount}</h3>
                    </div>
                    <div className="w-12 h-12 bg-[#edebe0] text-[#054239] rounded-xl flex items-center justify-center text-xl shadow-inner border border-[#b9a779]">
                      <i className="fa-solid fa-users"></i>
                    </div>
                  </button>
                  <button type="button" onClick={() => openReportsWithFilter('حضور')} className="text-right bg-white p-5 rounded-2xl shadow-sm border border-[#b9a779]/40 flex items-center justify-between hover:-translate-y-1 hover:shadow-md transition-all focus:outline-none focus:ring-2 focus:ring-[#054239]">
                    <div>
                      <p className="text-xs font-semibold text-[#3d3a3b]/70">إجمالي سجلات الحضور</p>
                      <h3 className="text-2xl font-black text-[#428177] mt-1">{dashboardStats.presentCount}</h3>
                    </div>
                    <div className="w-12 h-12 bg-[#054239]/10 text-[#428177] rounded-xl flex items-center justify-center text-xl shadow-inner border border-[#428177]/30">
                      <i className="fa-solid fa-user-check"></i>
                    </div>
                  </button>
                  <button type="button" onClick={() => openReportsWithFilter('غياب بدون تبرير')} className="text-right bg-white p-5 rounded-2xl shadow-sm border border-[#b9a779]/40 flex items-center justify-between hover:-translate-y-1 hover:shadow-md transition-all focus:outline-none focus:ring-2 focus:ring-[#054239]">
                    <div>
                      <p className="text-xs font-semibold text-[#3d3a3b]/70">إجمالي حالات الغياب</p>
                      <h3 className="text-2xl font-black text-[#4a151e] mt-1">{dashboardStats.absentCount}</h3>
                    </div>
                    <div className="w-12 h-12 bg-[#6b1f2a]/10 text-[#6b1f2a] rounded-xl flex items-center justify-center text-xl shadow-inner border border-[#6b1f2a]/20">
                      <i className="fa-solid fa-user-xmark"></i>
                    </div>
                  </button>
                  <button type="button" onClick={() => openReportsWithFilter('متأخر/مبكر')} className="text-right bg-white p-5 rounded-2xl shadow-sm border border-[#b9a779]/40 flex items-center justify-between hover:-translate-y-1 hover:shadow-md transition-all focus:outline-none focus:ring-2 focus:ring-[#054239]">
                    <div>
                      <p className="text-xs font-semibold text-[#3d3a3b]/70">إجمالي حالات التأخير والمخالفة</p>
                      <h3 className="text-2xl font-black text-[#988561] mt-1">{dashboardStats.lateCount}</h3>
                    </div>
                    <div className="w-12 h-12 bg-[#b9a779]/20 text-[#988561] rounded-xl flex items-center justify-center text-xl shadow-inner border border-[#b9a779]">
                      <i className="fa-solid fa-clock-rotate-left"></i>
                    </div>
                  </button>
                  <button type="button" onClick={() => openReportsWithFilter('إضافي')} className="text-right bg-white p-5 rounded-2xl shadow-sm border border-[#b9a779]/40 flex items-center justify-between hover:-translate-y-1 hover:shadow-md transition-all focus:outline-none focus:ring-2 focus:ring-[#054239]">
                    <div>
                      <p className="text-xs font-semibold text-[#3d3a3b]/70">الوقت الإضافي (دقيقة)</p>
                      <h3 className="text-2xl font-black text-[#054239] mt-1">{dashboardStats.overtimeMinutes}</h3>
                    </div>
                    <div className="w-12 h-12 bg-[#054239]/10 text-[#054239] rounded-xl flex items-center justify-center text-xl shadow-inner border border-[#428177]/30">
                      <i className="fa-solid fa-stopwatch"></i>
                    </div>
                  </button>
                </div>

                <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                  <div className="bg-white p-6 rounded-2xl shadow-sm border border-[#b9a779]/40">
                    <h3 className="font-bold text-[#002623] mb-4 flex items-center">
                      <i className="fa-solid fa-chart-pie text-[#054239] ml-2"></i>
                      توزيع حالات الحضور والانصراف الكلية
                    </h3>
                    <div className="relative h-72 flex items-center justify-center">
                      <canvas ref={pieChartRef} className="cursor-pointer"></canvas>
                    </div>
                  </div>
                  <div className="bg-white p-6 rounded-2xl shadow-sm border border-[#b9a779]/40">
                    <h3 className="font-bold text-[#002623] mb-4 flex items-center">
                      <i className="fa-solid fa-chart-column text-[#054239] ml-2"></i>
                      عدد حالات التأخير والغياب حسب المديريات
                    </h3>
                    <div className="relative h-72 flex items-center justify-center">
                      <canvas ref={barChartRef} className="cursor-pointer"></canvas>
                    </div>
                  </div>
                </div>

                <div className="bg-white rounded-2xl shadow-sm border border-[#b9a779]/40 p-6">
                  <div className="flex items-center justify-between mb-4">
                    <h3 className="font-bold text-[#002623]">أحدث السجلات</h3>
                    <button onClick={() => openReportsWithFilter('')} className="text-xs text-[#054239] hover:text-[#002623] font-bold">عرض الكل &larr;</button>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="attendance-report-table w-full text-right border-collapse">
                      <thead>
                        <tr className="bg-[#edebe0] border-b border-[#b9a779]/40 text-xs font-bold text-[#002623]">
                          <th className="p-3 rounded-r-xl">رقم الموظف</th>
                          <th className="p-3">اسم الموظف</th>
                          <th className="p-3">المديرية</th>
                          <th className="p-3">القسم</th>
                          <th className="p-3">التاريخ</th>
                          <th className="p-3">أول دخول</th>
                          <th className="p-3">آخر خروج</th>
                          <th className="p-3">الحالة النهائية</th>
                          <th className="p-3 rounded-l-xl">الوقت الإضافي (د)</th>
                        </tr>
                      </thead>
                      <tbody className="text-xs divide-y divide-[#b9a779]/20">
                        {appData.attendance.slice(-5).reverse().map((rec, idx) => {
                          const calc = calculateAttendanceStatus(rec);
                          return (
                            <tr key={idx} className="hover:bg-[#edebe0]/30">
                              <td className="p-3 font-mono font-bold text-[#002623]">{rec.empId}</td>
                              <td className="p-3 font-semibold">{rec.name}</td>
                              <td className="p-3">{getEmployeeDirectorate(rec.empId, rec.directorate || rec.department)}</td>
                              <td className="p-3">{getEmployeeSection(rec.empId, rec.department) || '-'}</td>
                              <td className="p-3 font-mono">{rec.date}</td>
                              <td className="p-3 font-mono text-[#3a6b5a]">{rec.firstIn || '-'}</td>
                              <td className="p-3 font-mono text-[#6b1f2a]">{rec.lastOut || '-'}</td>
                              <td className="p-3">
                                <span className={`status-badge ${getStatusBadgeClass(calc.status)}`}>{calc.status}</span>
                              </td>
                              <td className="p-3 font-mono text-xs">{calc.overtimeMinutes > 0 ? calc.overtimeMinutes : '-'}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}

            {/* ATTENDANCE UPLOAD VIEW */}
            {activeTab === 'attendance' && (
              <div className="space-y-6">
                <div className="bg-white p-6 rounded-2xl shadow-sm border border-[#b9a779]/40">
                  <div className="max-w-xl mx-auto text-center py-8 px-4 border-2 border-dashed border-[#b9a779] rounded-2xl bg-[#edebe0]/50 hover:bg-[#edebe0] transition-all cursor-pointer relative">
                    <input type="file" accept=".xlsx, .xls, .csv" onChange={handleFileUpload} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" />
                    <div className="w-16 h-16 bg-[#054239]/10 text-[#054239] rounded-2xl flex items-center justify-center mx-auto mb-4 shadow-inner text-2xl border border-[#428177]/30">
                      <i className="fa-solid fa-file-arrow-up"></i>
                    </div>
                    <h3 className="font-bold text-[#002623] text-lg mb-1">اسحب ملف البصمة هنا أو انقر للاختيار</h3>
                    <p className="text-xs text-[#3d3a3b]/70 mb-4">يدعم ملفات Excel (XLSX, XLS) والصيغ القياسية لأجهزة Xface</p>
                    <span className="inline-block px-4 py-2 bg-[#054239] text-[#edebe0] text-xs font-bold rounded-xl shadow-sm">تصفح الملفات</span>
                  </div>
                  <div className="mt-4 flex items-center justify-between text-xs text-[#3d3a3b]/70 px-2">
                    <span>الأعمدة المطلوبة: رقم الموظف / اسم الموظف / المديرية / القسم (اختياري) / أول دخول / آخر خروج / التاريخ</span>
                    <button onClick={downloadSampleExcel} className="text-[#054239] hover:underline font-bold flex items-center gap-1">
                      <i className="fa-solid fa-download"></i> تحميل ملف Excel نموذجي
                    </button>
                  </div>
                </div>

                {tempUploadedAttendance.length > 0 && (
                  <div className="bg-white rounded-2xl shadow-sm border border-[#b9a779]/40 p-6">
                    <div className="flex items-center justify-between mb-4">
                      <div>
                        <h3 className="font-bold text-[#002623] text-base">معاينة وتدقيق بيانات البصمة المحملة</h3>
                        <p className="text-xs text-[#3d3a3b]/70 mt-0.5">تم تنظيف البيانات ومعالجة القيم وتطبيق القواعد بنجاح</p>
                      </div>
                      <div className="flex items-center space-x-2 space-x-reverse">
                        <button onClick={processAndSaveUploadedData} className="px-4 py-2 bg-[#054239] hover:bg-[#002623] text-[#edebe0] text-xs font-bold rounded-xl shadow-sm flex items-center gap-1.5 border border-[#428177]">
                          <i className="fa-solid fa-check-circle"></i> اعتماد وحفظ البيانات في النظام
                        </button>
                        <button onClick={() => setTempUploadedAttendance([])} className="px-3 py-2 bg-[#6b1f2a]/10 hover:bg-[#6b1f2a]/20 text-[#6b1f2a] text-xs font-bold rounded-xl">
                          إلغاء
                        </button>
                      </div>
                    </div>
                    <div className="overflow-x-auto max-h-96">
                      <table className="attendance-report-table w-full text-right border-collapse">
                        <thead>
                          <tr className="bg-[#edebe0] border-b border-[#b9a779]/40 text-xs font-bold text-[#002623] sticky top-0">
                            <th className="p-3 rounded-r-xl">رقم الموظف</th>
                            <th className="p-3">اسم الموظف</th>
                            <th className="p-3">المديرية</th>
                            <th className="p-3">القسم</th>
                            <th className="p-3">التاريخ</th>
                            <th className="p-3">أول دخول</th>
                            <th className="p-3">آخر خروج</th>
                            <th className="p-3">الحالة المحسوبة</th>
                            <th className="p-3 rounded-l-xl">الوقت الإضافي (د)</th>
                          </tr>
                        </thead>
                        <tbody className="text-xs divide-y divide-[#b9a779]/20">
                          {tempUploadedAttendance.map((rec, idx) => {
                            const calc = calculateAttendanceStatus(rec);
                            return (
                              <tr key={idx} className="hover:bg-[#edebe0]/30">
                                <td className="p-3 font-mono font-bold text-[#002623]">{rec.empId}</td>
                                <td className="p-3 font-semibold">{rec.name}</td>
                                <td className="p-3">{getEmployeeDirectorate(rec.empId, rec.directorate || rec.department)}</td>
                                <td className="p-3">{getEmployeeSection(rec.empId, rec.department) || '-'}</td>
                                <td className="p-3 font-mono">{rec.date}</td>
                                <td className="p-3 font-mono text-[#3a6b5a]">{rec.firstIn || '-'}</td>
                                <td className="p-3 font-mono text-[#6b1f2a]">{rec.lastOut || '-'}</td>
                                <td className="p-3">
                                  <span className={`status-badge ${getStatusBadgeClass(calc.status)}`}>{calc.status}</span>
                                </td>
                                <td className="p-3 font-mono text-xs">{calc.overtimeMinutes > 0 ? calc.overtimeMinutes : '-'}</td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* LEAVES VIEW */}
            {activeTab === 'leaves' && (
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                <div className="bg-white p-6 rounded-2xl shadow-sm border border-[#b9a779]/40 lg:col-span-3">
                  <h3 className="font-bold text-[#002623] text-base mb-4 flex items-center">
                    <i className="fa-solid fa-calendar-plus text-[#988561] ml-2"></i>
                    إدخال إجازة أو مغادرة جديدة
                  </h3>
                  <form onSubmit={saveLeaveRecord} className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 2xl:grid-cols-6 items-end gap-4">
                    <div className="relative">
                      <label className="block text-xs font-semibold text-[#3d3a3b]/80 mb-1">الموظف</label>
                      <input
                        type="text"
                        value={leaveEmployeeSearch}
                        onChange={event => {
                          const value = event.target.value;
                          setLeaveEmployeeSearch(value);
                          setLeaveForm(prev => ({ ...prev, empId: '', empName: '' }));
                          setActiveLeaveEmployeeSuggestion(-1);
                          setShowLeaveEmployeeSuggestions(true);
                        }}
                        onFocus={() => {
                          setActiveLeaveEmployeeSuggestion(-1);
                          setShowLeaveEmployeeSuggestions(true);
                        }}
                        onBlur={() => setShowLeaveEmployeeSuggestions(false)}
                        onKeyDown={event => {
                          if (event.key === 'ArrowDown' && filteredLeaveEmployees.length) {
                            event.preventDefault();
                            setShowLeaveEmployeeSuggestions(true);
                            setActiveLeaveEmployeeSuggestion(index => Math.min(index + 1, filteredLeaveEmployees.length - 1));
                          } else if (event.key === 'ArrowUp' && filteredLeaveEmployees.length) {
                            event.preventDefault();
                            setActiveLeaveEmployeeSuggestion(index => Math.max(index - 1, 0));
                          } else if (event.key === 'Enter' && showLeaveEmployeeSuggestions && filteredLeaveEmployees.length) {
                            event.preventDefault();
                            selectLeaveEmployee(filteredLeaveEmployees[activeLeaveEmployeeSuggestion] || filteredLeaveEmployees[0]);
                          } else if (event.key === 'Escape') {
                            setShowLeaveEmployeeSuggestions(false);
                          }
                        }}
                        role="combobox"
                        aria-autocomplete="list"
                        aria-expanded={showLeaveEmployeeSuggestions}
                        aria-controls="leave-employee-options"
                        aria-activedescendant={showLeaveEmployeeSuggestions && activeLeaveEmployeeSuggestion >= 0 && filteredLeaveEmployees.length ? `leave-employee-option-${activeLeaveEmployeeSuggestion}` : undefined}
                        required
                        className="w-full px-3 py-2 text-xs bg-[#edebe0]/50 border border-[#b9a779] rounded-xl focus:ring-2 focus:ring-[#054239] focus:outline-none"
                        placeholder="ابحث باسم الموظف أو رقمه"
                      />
                      {showLeaveEmployeeSuggestions && (
                        <ul id="leave-employee-options" role="listbox" className="absolute inset-x-0 top-full z-20 mt-1 max-h-56 overflow-y-auto rounded-xl border border-[#b9a779] bg-white py-1 shadow-lg">
                          {filteredLeaveEmployees.length ? filteredLeaveEmployees.map((employee, index) => (
                            <li
                              key={employee.id}
                              id={`leave-employee-option-${index}`}
                              role="option"
                              aria-selected={index === activeLeaveEmployeeSuggestion}
                              onMouseDown={event => event.preventDefault()}
                              onClick={() => selectLeaveEmployee(employee)}
                              className={`cursor-pointer px-3 py-2 text-xs ${index === activeLeaveEmployeeSuggestion ? 'bg-[#edebe0] text-[#002623]' : 'text-[#3d3a3b] hover:bg-[#edebe0]/70'}`}
                            >
                              <span className="font-semibold">{employee.name}</span>
                              <span className="mr-2 font-mono text-[#988561]">{employee.id}</span>
                            </li>
                          )) : (
                            <li className="px-3 py-2 text-xs text-[#3d3a3b]/70">لا توجد نتائج مطابقة</li>
                          )}
                        </ul>
                      )}
                    </div>
                    <div>
                      <label className="block text-xs font-semibold text-[#3d3a3b]/80 mb-1">رقم الموظف (يُجلب تلقائياً)</label>
                      <input type="text" value={leaveForm.empId} readOnly required className="w-full px-3 py-2 text-xs bg-[#edebe0] border border-[#b9a779] rounded-xl text-[#002623] font-mono font-bold" placeholder="يُحدد تلقائياً" />
                    </div>
                    <div>
                      <label className="block text-xs font-semibold text-[#3d3a3b]/80 mb-1">نوع الإجازة / المغادرة</label>
                      <select value={leaveForm.type} onChange={e => setLeaveForm({ ...leaveForm, type: e.target.value })} className="w-full px-3 py-2 text-xs bg-[#edebe0]/50 border border-[#b9a779] rounded-xl focus:ring-2 focus:ring-[#054239] focus:outline-none">
                        <option value="سنوية">سنوية</option>
                        <option value="مرضية">مرضية</option>
                        <option value="صحية">صحية</option>
                        <option value="أمومة">أمومة</option>
                        <option value="أبوة">أبوة</option>
                        <option value="زواج">زواج</option>
                        <option value="وفاة">وفاة</option>
                        <option value="حج">حج</option>
                        <option value="بدون راتب">بدون راتب</option>
                        <option value="مهمة / مغادرة رسمية">مهمة / مغادرة رسمية</option>
                        <option value="أخرى">أخرى</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-semibold text-[#3d3a3b]/80 mb-1">مدة الإجازة</label>
                      <select value={leaveForm.leaveUnit} onChange={e => setLeaveForm({ ...leaveForm, leaveUnit: e.target.value })} className="w-full px-3 py-2 text-xs bg-[#edebe0]/50 border border-[#b9a779] rounded-xl focus:ring-2 focus:ring-[#054239] focus:outline-none">
                        <option value="أيام">إجازة بالأيام</option>
                        <option value="ساعية">إجازة ساعية</option>
                      </select>
                    </div>
                    <div className={`sm:col-span-2 grid ${leaveForm.leaveUnit === 'ساعية' ? 'grid-cols-1' : 'grid-cols-2'} gap-3`}>
                      <div>
                        <label className="block text-xs font-semibold text-[#3d3a3b]/80 mb-1">{leaveForm.leaveUnit === 'ساعية' ? 'تاريخ الإجازة الساعية' : 'من تاريخ'}</label>
                        <input type="date" value={leaveForm.fromDate} onChange={e => setLeaveForm(prev => ({ ...prev, fromDate: e.target.value, toDate: prev.leaveUnit === 'ساعية' ? e.target.value : prev.toDate }))} required className="w-full px-3 py-2 text-xs bg-[#edebe0]/50 border border-[#b9a779] rounded-xl focus:ring-2 focus:ring-[#054239] focus:outline-none" />
                      </div>
                      {leaveForm.leaveUnit !== 'ساعية' && (
                        <div>
                          <label className="block text-xs font-semibold text-[#3d3a3b]/80 mb-1">إلى تاريخ</label>
                          <input type="date" value={leaveForm.toDate} onChange={e => setLeaveForm({ ...leaveForm, toDate: e.target.value })} required className="w-full px-3 py-2 text-xs bg-[#edebe0]/50 border border-[#b9a779] rounded-xl focus:ring-2 focus:ring-[#054239] focus:outline-none" />
                        </div>
                      )}
                    </div>
                    {leaveForm.leaveUnit === 'ساعية' && (
                      <div className="sm:col-span-2 grid grid-cols-2 gap-3">
                        <div>
                          <label className="block text-xs font-semibold text-[#3d3a3b]/80 mb-1">من الساعة</label>
                          <input type="time" value={leaveForm.fromTime} onChange={e => setLeaveForm({ ...leaveForm, fromTime: e.target.value })} required className="w-full px-3 py-2 text-xs bg-[#edebe0]/50 border border-[#b9a779] rounded-xl focus:ring-2 focus:ring-[#054239] focus:outline-none" />
                        </div>
                        <div>
                          <label className="block text-xs font-semibold text-[#3d3a3b]/80 mb-1">إلى الساعة</label>
                          <input type="time" value={leaveForm.toTime} onChange={e => setLeaveForm({ ...leaveForm, toTime: e.target.value })} required className="w-full px-3 py-2 text-xs bg-[#edebe0]/50 border border-[#b9a779] rounded-xl focus:ring-2 focus:ring-[#054239] focus:outline-none" />
                        </div>
                      </div>
                    )}
                    <div>
                      <label className="block text-xs font-semibold text-[#3d3a3b]/80 mb-1">رقم الوثيقة</label>
                      <input type="text" value={leaveForm.documentNumber} onChange={e => setLeaveForm({ ...leaveForm, documentNumber: e.target.value })} className="w-full px-3 py-2 text-xs bg-[#edebe0]/50 border border-[#b9a779] rounded-xl focus:ring-2 focus:ring-[#054239] focus:outline-none" placeholder="رقم القرار أو الوثيقة (اختياري)" />
                    </div>
                    <div>
                      <label className="block text-xs font-semibold text-[#3d3a3b]/80 mb-1">الحالة / الاعتماد</label>
                      <select value={leaveForm.status} onChange={e => setLeaveForm({ ...leaveForm, status: e.target.value })} className="w-full px-3 py-2 text-xs bg-[#edebe0]/50 border border-[#b9a779] rounded-xl focus:ring-2 focus:ring-[#054239] focus:outline-none">
                        <option value="معتمدة">معتمدة (تعتبر غياب مبرر)</option>
                        <option value="مرفوضة">مرفوضة</option>
                      </select>
                    </div>
                    <div className="sm:col-span-2">
                      <label className="block text-xs font-semibold text-[#3d3a3b]/80 mb-1">ملاحظات</label>
                      <textarea rows="2" value={leaveForm.notes} onChange={e => setLeaveForm({ ...leaveForm, notes: e.target.value })} className="w-full px-3 py-2 text-xs bg-[#edebe0]/50 border border-[#b9a779] rounded-xl focus:ring-2 focus:ring-[#054239] focus:outline-none" placeholder="سبب الإجازة أو تفاصيل إضافية..."></textarea>
                    </div>
                    <button type="submit" className="w-full min-h-10 py-2.5 bg-[#054239] hover:bg-[#002623] text-[#edebe0] text-xs font-bold rounded-xl shadow-sm transition-all border border-[#428177]">
                      حفظ وإضافة الإجازة
                    </button>
                  </form>
                </div>

                <div className="bg-white p-6 rounded-2xl shadow-sm border border-[#b9a779]/40 lg:col-span-3 flex min-w-0 flex-col">
                  <div className="flex items-center justify-between mb-4">
                    <h3 className="font-bold text-[#002623] text-base">سجل الإجازات والمغادرات المدخلة</h3>
                    <span className="px-2.5 py-1 bg-[#edebe0] text-[#988561] text-xs font-bold rounded-lg border border-[#b9a779]">{appData.leaves.length} إجازة</span>
                  </div>
                  <div className="flex-1 overflow-x-auto max-h-[500px]">
                    <table id="leave-records-table" className="w-full text-right border-collapse">
                      <thead>
                        <tr className="bg-[#edebe0] border-b border-[#b9a779]/40 text-xs font-bold text-[#002623] sticky top-0">
                          <th className="p-3 rounded-r-xl">رقم الموظف</th>
                          <th className="p-3">الاسم</th>
                          <th className="p-3">المديرية</th>
                          <th className="p-3">القسم</th>
                          <th className="p-3">النوع</th>
                          <th className="p-3">رقم الوثيقة</th>
                          <th className="p-3">من (تاريخ / وقت)</th>
                          <th className="p-3">إلى (تاريخ / وقت)</th>
                          <th className="p-3">الحالة</th>
                          <th className="p-3 rounded-l-xl">إجراء</th>
                        </tr>
                      </thead>
                      <tbody className="text-xs divide-y divide-[#b9a779]/20">
                        {appData.leaves.map((leave) => (
                          <tr key={leave.id} className="hover:bg-[#edebe0]/30">
                            <td className="p-3 font-mono font-bold text-[#002623]">{leave.empId}</td>
                            <td className="p-3 font-semibold">{leave.empName}</td>
                            <td className="p-3">{getEmployeeDirectorate(leave.empId, appData.attendance.find(rec => rec.empId === leave.empId)?.directorate || appData.attendance.find(rec => rec.empId === leave.empId)?.department || 'عامة')}</td>
                            <td className="p-3">{getEmployeeSection(leave.empId, appData.attendance.find(rec => rec.empId === leave.empId)?.department) || '-'}</td>
                            <td className="p-3">{leave.type}{leave.leaveUnit === 'ساعية' ? ' (ساعية)' : ''}</td>
                            <td className="p-3 font-mono">{leave.documentNumber || '-'}</td>
                            <td className="p-3 font-mono">{leave.leaveUnit === 'ساعية' ? `${leave.fromDate} ${leave.fromTime || ''}` : leave.fromDate}</td>
                            <td className="p-3 font-mono">{leave.leaveUnit === 'ساعية' ? `${leave.fromDate} ${leave.toTime || ''}` : leave.toDate}</td>
                            <td className="p-3">
                              <span className={`status-badge ${leave.status === 'معتمدة' ? 'text-emerald-700' : 'text-amber-700'}`}>{leave.status}</span>
                            </td>
                            <td className="p-3">
                              <button onClick={() => deleteLeaveRecord(leave.id)} className="text-[#6b1f2a] hover:text-[#4a151e] p-1"><i className="fa-solid fa-trash"></i></button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}

            {/* EMPLOYEES STATUS VIEW */}
            {activeTab === 'employees' && (
              <div className="bg-white p-6 rounded-2xl shadow-sm border border-[#b9a779]/40">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
                  <div>
                    <h3 className="font-bold text-[#002623] text-base">إدارة الحالة الوظيفية للموظفين</h3>
                    <p className="text-xs text-[#3d3a3b]/70 mt-0.5">الموظفون غير النشطين (مفصول، مستقيل، منقول) يُخفون تلقائياً من التقارير الحالية</p>
                  </div>
                  <div className="flex items-center space-x-3 space-x-reverse">
                    <label className="flex items-center space-x-2 space-x-reverse text-xs cursor-pointer font-semibold bg-[#edebe0] px-3 py-2 rounded-xl border border-[#b9a779]">
                      <input type="checkbox" checked={showInactiveToggle} onChange={e => setShowInactiveToggle(e.target.checked)} className="rounded text-[#054239] focus:ring-[#054239]" />
                      <span>عرض الموظفين غير النشطين</span>
                    </label>
                  </div>
                </div>
                <div className="overflow-x-auto max-h-[550px]">
                  <table className="attendance-report-table w-full text-right border-collapse">
                    <thead>
                      <tr className="bg-[#edebe0] border-b border-[#b9a779]/40 text-xs font-bold text-[#002623] sticky top-0">
                        <th className="p-3 rounded-r-xl">رقم الموظف</th>
                        <th className="p-3">اسم الموظف</th>
                        <th className="p-3">المديرية</th>
                        <th className="p-3">القسم</th>
                        <th className="p-3">الحالة الحالية</th>
                        <th className="p-3">تاريخ تغيير الحالة</th>
                        <th className="p-3 rounded-l-xl">تعديل بيانات الموظف</th>
                      </tr>
                    </thead>
                    <tbody className="text-xs divide-y divide-[#b9a779]/20">
                      {(() => {
                        const empMap = {};
                        appData.attendance.forEach(rec => {
                          if (!empMap[rec.empId]) empMap[rec.empId] = { empId: rec.empId, name: rec.name, directorate: rec.directorate, department: rec.department };
                        });
                        Object.keys(appData.employeesStatus).forEach(empId => {
                          if (!empMap[empId]) empMap[empId] = { empId, name: `موظف رقم ${empId}`, directorate: "عامة", department: "" };
                        });

                        return Object.values(empMap)
                          .filter(emp => {
                            const statusInfo = appData.employeesStatus[emp.empId] || { status: "نشط" };
                            return showInactiveToggle || statusInfo.status === 'نشط';
                          })
                          .map((emp) => {
                            const statusInfo = appData.employeesStatus[emp.empId] || { status: "نشط", date: "2025-01-01" };
                            return (
                              <tr key={emp.empId} className="hover:bg-[#edebe0]/30">
                                <td className="p-3 font-mono font-bold text-[#002623]">{emp.empId}</td>
                                <td className="p-3 font-semibold">{emp.name}</td>
                                <td className="p-3">{getEmployeeDirectorate(emp.empId, emp.directorate || emp.department)}</td>
                                <td className="p-3">{getEmployeeSection(emp.empId, emp.department) || '-'}</td>
                                <td className="p-3">
                                  <span className={`status-badge ${statusInfo.status === 'نشط' ? 'text-emerald-700' : 'text-red-700'}`}>{statusInfo.status}</span>
                                </td>
                                <td className="p-3 font-mono">{statusInfo.date || '-'}</td>
                                <td className="p-3">
                                  <button onClick={() => setEmpModal({ isOpen: true, empId: emp.empId, status: statusInfo.status, date: statusInfo.date || new Date().toISOString().split('T')[0], directorate: getEmployeeDirectorate(emp.empId, emp.directorate || emp.department), department: getEmployeeSection(emp.empId, emp.department) })} className="px-3 py-1.5 bg-[#054239]/10 hover:bg-[#054239]/20 text-[#054239] text-xs font-bold rounded-lg">تعديل بيانات الموظف</button>
                                </td>
                              </tr>
                            );
                          });
                      })()}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {activeTab === 'employeeDirectory' && (
              <div className="space-y-6">
                {selectedDirectoryEmployee ? (
                  <section className="bg-white p-6 rounded-2xl shadow-sm border border-[#b9a779]/40 space-y-5">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                      <div>
                        <h3 className="font-bold text-[#002623] text-lg">سجل حضور {selectedDirectoryEmployee.name}</h3>
                        <p className="text-xs text-[#3d3a3b]/70 mt-1">
                          الرقم الوظيفي: <span className="font-bold">{selectedDirectoryEmployee.empId}</span>
                          <span className="mx-2">•</span>
                          {selectedDirectoryEmployee.attendanceCount} سجل حضور
                        </p>
                      </div>
                      <div className="flex flex-wrap items-center gap-2">
                        <button type="button" onClick={() => exportEmployeeAttendanceExcel(selectedDirectoryEmployee)} className="min-h-10 px-4 py-2 bg-[#054239] hover:bg-[#002623] text-[#edebe0] text-xs font-bold rounded-xl border border-[#428177] shadow-sm transition-colors">
                          <i className="fa-solid fa-file-excel ml-2" aria-hidden="true"></i>
                          تصدير سجل الموظف
                        </button>
                        <button type="button" onClick={() => setSelectedDirectoryEmployeeId('')} className="min-h-10 px-4 py-2 bg-[#edebe0] hover:bg-[#b9a779]/30 text-[#002623] text-xs font-bold rounded-xl border border-[#b9a779] transition-colors">
                          <i className="fa-solid fa-arrow-right ml-2" aria-hidden="true"></i>
                          العودة إلى جميع الموظفين
                        </button>
                      </div>
                    </div>
                    {selectedEmployeeAttendance.length > 0 ? (
                      <div className="overflow-x-auto max-h-[650px]">
                        <table className="compact-attendance-table w-full text-right border-collapse">
                          <thead>
                            <tr>
                              <th>التاريخ</th>
                              <th>أول دخول</th>
                              <th>آخر خروج</th>
                              <th>المديرية</th>
                              <th>القسم</th>
                              <th>الحالة</th>
                              <th>الوقت الإضافي (د)</th>
                            </tr>
                          </thead>
                          <tbody className="text-xs divide-y divide-[#b9a779]/20">
                            {selectedEmployeeAttendance
                              .slice()
                              .sort((a, b) => b.date.localeCompare(a.date))
                              .map((record, index) => {
                                const calculation = calculateAttendanceStatus(record);
                                return (
                                  <tr key={`${record.date}-${index}`}>
                                    <td className="font-mono whitespace-nowrap">{record.date}</td>
                                    <td className="font-mono text-[#3a6b5a]">{record.firstIn || '-'}</td>
                                    <td className="font-mono text-[#6b1f2a]">{record.lastOut || '-'}</td>
                                    <td>{getEmployeeDirectorate(record.empId, record.directorate || record.department)}</td>
                                    <td>{getEmployeeSection(record.empId, record.department) || '-'}</td>
                                    <td><span className={`status-badge ${getStatusBadgeClass(calculation.status)}`}>{calculation.status}</span></td>
                                    <td className="font-mono font-bold text-emerald-700">{calculation.overtimeMinutes > 0 ? calculation.overtimeMinutes : '-'}</td>
                                  </tr>
                                );
                              })}
                          </tbody>
                        </table>
                      </div>
                    ) : (
                      <div className="rounded-xl border border-dashed border-[#b9a779] bg-[#edebe0]/40 px-4 py-10 text-center">
                        <i className="fa-regular fa-calendar-xmark text-2xl text-[#988561]" aria-hidden="true"></i>
                        <p className="mt-3 font-bold text-[#002623]">لا توجد سجلات حضور لهذا الموظف</p>
                        <p className="mt-1 text-xs text-[#3d3a3b]/70">سيظهر سجل الحضور هنا بعد استيراد بياناته.</p>
                      </div>
                    )}
                  </section>
                ) : (
                  <section className="bg-white p-6 rounded-2xl shadow-sm border border-[#b9a779]/40 space-y-5">
                    <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                      <div>
                        <h3 className="font-bold text-[#002623] text-lg">جميع الموظفين في النظام</h3>
                        <p className="mt-1 text-xs text-[#3d3a3b]/70">اختر اسم الموظف لعرض سجلات حضوره، بما في ذلك الموظفون الذين لا يملكون سجلات بعد.</p>
                      </div>
                      <span className="px-3 py-1.5 bg-[#edebe0] text-[#054239] text-xs font-bold rounded-xl border border-[#b9a779]/60">{employeeDirectory.length} موظف</span>
                    </div>
                    <div className="relative max-w-md">
                      <i className="fa-solid fa-magnifying-glass absolute right-3 top-1/2 -translate-y-1/2 text-[#988561]" aria-hidden="true"></i>
                      <input type="search" value={employeeDirectorySearch} onChange={event => setEmployeeDirectorySearch(event.target.value)} placeholder="ابحث باسم الموظف أو رقمه الوظيفي" className="w-full pr-10 pl-4 py-2.5 text-sm bg-[#edebe0]/50 border border-[#b9a779]/70 rounded-xl focus:outline-none focus:ring-2 focus:ring-[#054239]" />
                    </div>
                    <div className="overflow-x-auto max-h-[650px]">
                      <table className="w-full text-right border-collapse">
                        <thead>
                          <tr>
                            <th>رقم الموظف</th>
                            <th>اسم الموظف</th>
                            <th>المديرية</th>
                            <th>القسم</th>
                            <th>الحالة الوظيفية</th>
                            <th>سجلات الحضور</th>
                            <th>التصدير</th>
                          </tr>
                        </thead>
                        <tbody className="text-xs divide-y divide-[#b9a779]/20">
                          {filteredEmployeeDirectory.map(employee => (
                            <tr key={employee.empId}>
                              <td className="font-mono font-bold text-[#002623]">{employee.empId}</td>
                              <td>
                                <button type="button" onClick={() => setSelectedDirectoryEmployeeId(employee.empId)} className="font-bold text-[#054239] hover:text-[#002623] hover:underline underline-offset-4 focus-visible:rounded-sm">
                                  {employee.name}
                                </button>
                              </td>
                              <td>{employee.directorate}</td>
                              <td>{employee.department || '-'}</td>
                              <td><span className={`status-badge ${employee.status === 'نشط' ? 'text-emerald-700' : 'text-red-700'}`}>{employee.status}</span></td>
                              <td className="font-mono">{employee.attendanceCount}</td>
                              <td>
                                <button
                                  type="button"
                                  onClick={() => exportEmployeeAttendanceExcel(employee)}
                                  disabled={employee.attendanceCount === 0}
                                  title={employee.attendanceCount === 0 ? 'لا توجد سجلات حضور للتصدير' : `تصدير سجل حضور ${employee.name}`}
                                  aria-label={`تصدير سجل حضور ${employee.name}`}
                                  className="inline-flex min-h-9 items-center gap-1.5 rounded-lg border border-[#054239]/25 bg-[#054239]/10 px-3 py-2 text-xs font-bold text-[#054239] transition-colors hover:bg-[#054239] hover:text-white disabled:cursor-not-allowed disabled:opacity-45 disabled:hover:bg-[#054239]/10 disabled:hover:text-[#054239]"
                                >
                                  <i className="fa-solid fa-file-excel" aria-hidden="true"></i>
                                  Excel
                                </button>
                              </td>
                            </tr>
                          ))}
                          {filteredEmployeeDirectory.length === 0 && (
                            <tr>
                              <td colSpan={7} className="py-10 text-center text-sm text-[#3d3a3b]/70">لا توجد نتائج مطابقة للبحث.</td>
                            </tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                  </section>
                )}
              </div>
            )}

            {/* REPORTS VIEW */}
            {activeTab === 'reports' && (
              <div className="space-y-6">
                <div className="bg-white p-6 rounded-2xl shadow-sm border border-[#b9a779]/40 space-y-4">
                  <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
                    <h3 className="font-bold text-[#002623] text-base flex items-center">
                      <i className="fa-solid fa-filter text-[#054239] ml-2"></i>
                      فلاتر البحث والتقارير المتقدمة
                    </h3>
                    <div className="flex flex-wrap items-center gap-2">
                      <button onClick={exportReportsExcel} className="min-h-10 whitespace-nowrap px-4 py-2 bg-[#054239] hover:bg-[#002623] text-[#edebe0] text-xs font-bold rounded-xl shadow-sm flex items-center gap-1.5 border border-[#428177]">
                        <i className="fa-solid fa-file-excel text-[#edebe0]"></i> تصدير لـ Excel
                      </button>
                      <button onClick={() => window.print()} className="min-h-10 whitespace-nowrap px-4 py-2 bg-[#161616] hover:bg-black text-[#edebe0] text-xs font-bold rounded-xl shadow-sm flex items-center gap-1.5">
                        <i className="fa-solid fa-print"></i> طباعة / PDF
                      </button>
                    </div>
                  </div>
                  <div className="report-filter-grid grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 2xl:grid-cols-7 gap-3 pt-2">
                    <div>
                      <label className="report-filter-label text-xs font-semibold text-[#3d3a3b]/80">بحث بالاسم أو الرقم</label>
                      <input type="text" value={reportFilters.search} onChange={e => setReportFilters({ ...reportFilters, search: e.target.value })} placeholder="ابحث..." className="report-filter-control px-3 text-xs" />
                    </div>
                    <div>
                      <label className="report-filter-label text-xs font-semibold text-[#3d3a3b]/80">المديرية</label>
                      <select value={reportFilters.directorate} onChange={e => setReportFilters({ ...reportFilters, directorate: e.target.value })} className="report-filter-control px-3 text-xs">
                        <option value="">جميع المديريات</option>
                        {[...new Set([
                          ...appData.attendance.map(r => getEmployeeDirectorate(r.empId, r.directorate || r.department)),
                          ...Object.values(appData.employeesStatus).map(status => status.directorate).filter(Boolean)
                        ])].map((d, i) => (
                          <option key={i} value={d}>{d}</option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label className="report-filter-label text-xs font-semibold text-[#3d3a3b]/80">القسم</label>
                      <select value={reportFilters.department} onChange={e => setReportFilters({ ...reportFilters, department: e.target.value })} className="report-filter-control px-3 text-xs">
                        <option value="">جميع الأقسام</option>
                        {[...new Set([
                          ...appData.attendance.map(r => getEmployeeSection(r.empId, r.department)),
                          ...Object.values(appData.employeesStatus).map(status => status.department)
                        ].filter(Boolean))].map((d, i) => (
                          <option key={i} value={d}>{d}</option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label className="report-filter-label text-xs font-semibold text-[#3d3a3b]/80">الحالة</label>
                      <select value={reportFilters.status} onChange={e => setReportFilters({ ...reportFilters, status: e.target.value })} className="report-filter-control px-3 text-xs">
                        <option value="">جميع الحالات</option>
                        <option value="حضور">الحضور (بما في ذلك المخالفات)</option>
                        <option value="حاضر">حاضر</option>
                        <option value="متأخر/مبكر">متأخر أو خروج مبكر</option>
                        <option value="متأخر">متأخر</option>
                        <option value="خروج مبكر">خروج مبكر</option>
                        <option value="متأخر وخروج مبكر">متأخر وخروج مبكر</option>
                        <option value="غياب بدون تبرير">غياب بدون تبرير</option>
                        <option value="إجازة/عطلة">إجازة أو عطلة رسمية أو راحة أسبوعية</option>
                        <option value="إجازة">إجازة</option>
                        <option value="مغادرة رسمية">مغادرة رسمية</option>
                        <option value="عطلة رسمية">عطلة رسمية</option>
                        <option value="عطلة أسبوعية">عطلة أسبوعية</option>
                        <option value="دوام في عطلة أسبوعية">دوام في عطلة أسبوعية</option>
                        <option value="إضافي">لديهم وقت إضافي</option>
                      </select>
                    </div>
                    <div>
                      <label className="report-filter-label text-xs font-semibold text-[#3d3a3b]/80">من تاريخ</label>
                      <input type="date" value={reportFilters.startDate} onChange={e => setReportFilters({ ...reportFilters, startDate: e.target.value })} className="report-filter-control px-3 text-xs" />
                    </div>
                    <div>
                      <label className="report-filter-label text-xs font-semibold text-[#3d3a3b]/80">إلى تاريخ</label>
                      <input type="date" value={reportFilters.endDate} onChange={e => setReportFilters({ ...reportFilters, endDate: e.target.value })} className="report-filter-control px-3 text-xs" />
                    </div>
                    <div className="flex items-end">
                      <button onClick={() => setReportFilters({ search: '', directorate: '', department: '', status: '', startDate: '', endDate: '' })} className="report-filter-control report-filter-reset text-xs font-bold">
                        إعادة ضبط
                      </button>
                    </div>
                  </div>
                </div>

                <div className="bg-white rounded-2xl shadow-sm border border-[#b9a779]/40 p-6">
                  <div className="flex items-center justify-between mb-4">
                    <h3 className="font-bold text-[#002623]">سجلات الحضور التفصيلية والمخالفات</h3>
                    <span className="text-xs font-semibold text-[#3d3a3b]/70">
                      {filteredAttendanceRecords.length} سجل مطابق
                    </span>
                  </div>
                  <div className="overflow-x-auto max-h-[600px]">
                    <table className="attendance-report-table w-full text-right border-collapse">
                      <thead>
                        <tr className="bg-[#edebe0] border-b border-[#b9a779]/40 text-xs font-bold text-[#002623] sticky top-0">
                          <th className="p-3 rounded-r-xl">رقم الموظف</th>
                          <th className="p-3">اسم الموظف</th>
                          <th className="p-3">المديرية</th>
                          <th className="p-3">القسم</th>
                          <th className="p-3">التاريخ</th>
                          <th className="p-3">أول دخول</th>
                          <th className="p-3">آخر خروج</th>
                          <th className="p-3">الحالة النهائية</th>
                          <th className="p-3">التأخير (د)</th>
                          <th className="p-3">الخروج المبكر (د)</th>
                          <th className="p-3 rounded-l-xl">الوقت الإضافي (د)</th>
                        </tr>
                      </thead>
                      <tbody className="text-xs divide-y divide-[#b9a779]/20">
                        {filteredAttendanceRecords.map((rec, idx) => {
                            const calc = calculateAttendanceStatus(rec);
                            return (
                              <tr key={idx} className="hover:bg-[#edebe0]/30">
                                <td className="p-3 font-mono font-bold text-[#002623]">{rec.empId}</td>
                                <td className="p-3 font-semibold">{rec.name}</td>
                                <td className="p-3">{getEmployeeDirectorate(rec.empId, rec.directorate || rec.department)}</td>
                                <td className="p-3">{getEmployeeSection(rec.empId, rec.department) || '-'}</td>
                                <td className="p-3 font-mono">{rec.date}</td>
                                <td className="p-3 font-mono text-[#3a6b5a]">{rec.firstIn || '-'}</td>
                                <td className="p-3 font-mono text-[#6b1f2a]">{rec.lastOut || '-'}</td>
                                <td className="p-3">
                                  <span className={`status-badge ${getStatusBadgeClass(calc.status)}`}>{calc.status}</span>
                                </td>
                                <td className="p-3 font-mono">{calc.lateMinutes > 0 ? calc.lateMinutes : '-'}</td>
                                <td className="p-3 font-mono">{calc.earlyMinutes > 0 ? calc.earlyMinutes : '-'}</td>
                                <td className="p-3 font-mono font-bold text-emerald-700">{calc.overtimeMinutes > 0 ? calc.overtimeMinutes : '-'}</td>
                              </tr>
                            );
                          })}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}

            {/* SETTINGS VIEW */}
            {activeTab === 'settings' && (
              <div className="bg-white p-6 rounded-2xl shadow-sm border border-[#b9a779]/40 max-w-3xl mx-auto space-y-6">
                <div>
                  <h3 className="font-bold text-[#002623] text-lg mb-2 flex items-center">
                    <i className="fa-solid fa-gears text-[#6b1f2a] ml-2"></i>
                    إعدادات الدوام وقواعد احتساب المخالفات
                  </h3>
                  <p className="text-xs text-[#3d3a3b]/70">تُطبق هذه الإعدادات تلقائياً عند تحليل ملفات البصمة وحساب التأخير والغياب</p>
                </div>
                
                <form onSubmit={saveSettings} className="space-y-5 border-b border-[#b9a779]/40 pb-6">
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <label className="block text-xs font-semibold text-[#3d3a3b] mb-1">وقت بداية الدوام الرسمي</label>
                      <input type="time" value={settingsForm.startTime} onChange={e => setSettingsForm({ ...settingsForm, startTime: e.target.value })} required className="w-full px-3 py-2 text-xs bg-[#edebe0]/50 border border-[#b9a779] rounded-xl focus:ring-2 focus:ring-[#054239] focus:outline-none" />
                    </div>
                    <div>
                      <label className="block text-xs font-semibold text-[#3d3a3b] mb-1">وقت نهاية الدوام الرسمي</label>
                      <input type="time" value={settingsForm.endTime} onChange={e => setSettingsForm({ ...settingsForm, endTime: e.target.value })} required className="w-full px-3 py-2 text-xs bg-[#edebe0]/50 border border-[#b9a779] rounded-xl focus:ring-2 focus:ring-[#054239] focus:outline-none" />
                    </div>
                  </div>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <label className="block text-xs font-semibold text-[#3d3a3b] mb-1">فترة سماحية الدخول المتأخر (بالدقائق)</label>
                      <input type="number" min="0" value={settingsForm.graceLate} onChange={e => setSettingsForm({ ...settingsForm, graceLate: e.target.value })} required className="w-full px-3 py-2 text-xs bg-[#edebe0]/50 border border-[#b9a779] rounded-xl focus:ring-2 focus:ring-[#054239] focus:outline-none" />
                    </div>
                    <div>
                      <label className="block text-xs font-semibold text-[#3d3a3b] mb-1">فترة سماحية الخروج المبكر (بالدقائق)</label>
                      <input type="number" min="0" value={settingsForm.graceEarly} onChange={e => setSettingsForm({ ...settingsForm, graceEarly: e.target.value })} required className="w-full px-3 py-2 text-xs bg-[#edebe0]/50 border border-[#b9a779] rounded-xl focus:ring-2 focus:ring-[#054239] focus:outline-none" />
                    </div>
                  </div>
                  <div>
                    <p className="block text-xs font-semibold text-[#3d3a3b] mb-2">أيام الدوام الأسبوعية</p>
                    <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                      {["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"].map(day => (
                        <label key={day} className="flex items-center gap-2 text-xs bg-[#edebe0]/50 border border-[#b9a779]/60 rounded-lg px-3 py-2 cursor-pointer">
                          <input
                            type="checkbox"
                            checked={settingsForm.workDays.includes(day)}
                            onChange={e => setSettingsForm(prev => ({
                              ...prev,
                              workDays: e.target.checked
                                ? [...prev.workDays, day]
                                : prev.workDays.filter(workDay => workDay !== day)
                            }))}
                            className="rounded text-[#054239] focus:ring-[#054239]"
                          />
                          <span>{day}</span>
                        </label>
                      ))}
                    </div>
                    <p className="text-[11px] text-[#3d3a3b]/65 mt-2">لن تُحتسب السجلات الواقعة خارج هذه الأيام كغياب أو مخالفة.</p>
                  </div>
                  <div className="pt-2 flex justify-end">
                    <button type="submit" className="px-6 py-2.5 bg-[#054239] hover:bg-[#002623] text-[#edebe0] text-xs font-bold rounded-xl shadow-sm transition-all border border-[#428177]">
                      حفظ إعدادات الدوام العامة
                    </button>
                  </div>
                </form>

                <div className="space-y-4 pt-2">
                  <div>
                    <h4 className="font-bold text-[#002623] text-sm flex items-center">
                      <i className="fa-solid fa-calendar-check text-[#054239] ml-2"></i>
                      إدارة العطل الرسمية
                    </h4>
                    <p className="text-xs text-[#3d3a3b]/70">العطل المضافة لا تُحسب أبداً كغياب وتُعامل كعطلة رسمية مدفوعة</p>
                  </div>

                  <form onSubmit={addHolidayRecord} className="grid grid-cols-1 sm:grid-cols-4 gap-3 bg-[#edebe0]/50 p-4 rounded-xl border border-[#b9a779]">
                    <div className="sm:col-span-1">
                      <label className="block text-xs font-semibold text-[#3d3a3b]/80 mb-1">اسم العطلة</label>
                      <input type="text" value={holidayForm.title} onChange={e => setHolidayForm({ ...holidayForm, title: e.target.value })} required placeholder="مثال: عيد الجلاء" className="w-full px-3 py-2 text-xs bg-white border border-[#b9a779] rounded-lg focus:outline-none focus:ring-2 focus:ring-[#054239]" />
                    </div>
                    <div className="sm:col-span-1">
                      <label className="block text-xs font-semibold text-[#3d3a3b]/80 mb-1">من تاريخ</label>
                      <input type="date" value={holidayForm.fromDate} onChange={e => setHolidayForm({ ...holidayForm, fromDate: e.target.value })} required className="w-full px-3 py-2 text-xs bg-white border border-[#b9a779] rounded-lg focus:outline-none focus:ring-2 focus:ring-[#054239]" />
                    </div>
                    <div className="sm:col-span-1">
                      <label className="block text-xs font-semibold text-[#3d3a3b]/80 mb-1">إلى تاريخ</label>
                      <input type="date" value={holidayForm.toDate} onChange={e => setHolidayForm({ ...holidayForm, toDate: e.target.value })} required className="w-full px-3 py-2 text-xs bg-white border border-[#b9a779] rounded-lg focus:outline-none focus:ring-2 focus:ring-[#054239]" />
                    </div>
                    <div className="sm:col-span-1 flex items-end">
                      <button type="submit" className="w-full py-2 bg-[#054239] hover:bg-[#002623] text-[#edebe0] text-xs font-bold rounded-lg shadow-sm transition-all border border-[#428177]">
                        إضافة عطلة
                      </button>
                    </div>
                  </form>

                  <div className="overflow-x-auto max-h-48 border border-[#b9a779] rounded-xl">
                    <table className="w-full text-right border-collapse text-xs">
                      <thead>
                        <tr className="bg-[#edebe0] border-b border-[#b9a779] text-[#002623] font-bold sticky top-0">
                          <th className="p-2.5">اسم العطلة / الملاحظة</th>
                          <th className="p-2.5">من تاريخ</th>
                          <th className="p-2.5">إلى تاريخ</th>
                          <th className="p-2.5">إجراء</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-[#b9a779]/20">
                        {(appData.settings.holidays || []).map(h => (
                          <tr key={h.id} className="hover:bg-[#edebe0]/30">
                            <td className="p-2.5 font-semibold">{h.title}</td>
                            <td className="p-2.5 font-mono">{h.fromDate}</td>
                            <td className="p-2.5 font-mono">{h.toDate}</td>
                            <td className="p-2.5">
                              <button onClick={() => deleteHolidayRecord(h.id)} className="text-[#6b1f2a] hover:text-[#4a151e] p-1"><i className="fa-solid fa-trash"></i></button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>

                <section className="space-y-4 border-t border-[#b9a779]/50 pt-6" aria-labelledby="data-management-title">
                  <div>
                    <h4 id="data-management-title" className="font-bold text-[#002623] text-sm flex items-center">
                      <i className="fa-solid fa-database text-[#054239] ml-2" aria-hidden="true"></i>
                      إدارة البيانات والنسخ الاحتياطي
                    </h4>
                    <p className="text-xs text-[#3d3a3b]/70 mt-1">
                      تُحفظ البيانات في قاعدة IndexedDB داخل هذا المتصفح وعلى هذا الجهاز. أنشئ ملف JSON لحفظ نسخة احتياطية أو استعد بيانات من نسخة سابقة.
                    </p>
                  </div>
                  <input
                    ref={backupFileInputRef}
                    type="file"
                    accept=".json,application/json"
                    onChange={importBackup}
                    className="hidden"
                    aria-label="اختيار ملف نسخة احتياطية"
                  />
                  <div className="flex flex-col sm:flex-row flex-wrap gap-3">
                    <button
                      type="button"
                      onClick={createBackup}
                      className="min-h-10 px-4 py-2 bg-[#054239] hover:bg-[#002623] text-[#edebe0] text-xs font-bold rounded-xl border border-[#428177] shadow-sm transition-colors"
                    >
                      <i className="fa-solid fa-download ml-2" aria-hidden="true"></i>
                      إنشاء نسخة احتياطية
                    </button>
                    <button
                      type="button"
                      onClick={() => backupFileInputRef.current?.click()}
                      className="min-h-10 px-4 py-2 bg-white hover:bg-[#edebe0] text-[#054239] text-xs font-bold rounded-xl border border-[#054239] shadow-sm transition-colors"
                    >
                      <i className="fa-solid fa-upload ml-2" aria-hidden="true"></i>
                      استيراد نسخة احتياطية
                    </button>
                    <button
                      type="button"
                      onClick={deleteAllData}
                      className="min-h-10 px-4 py-2 bg-[#6b1f2a] hover:bg-[#4a151e] text-white text-xs font-bold rounded-xl border border-[#6b1f2a] shadow-sm transition-colors"
                    >
                      <i className="fa-solid fa-trash-can ml-2" aria-hidden="true"></i>
                      حذف جميع البيانات
                    </button>
                  </div>
                  <p className="text-[11px] text-[#6b1f2a]">
                    الحذف يمسح السجلات والإعدادات والعطل من هذا المتصفح بعد التأكيد، ولا يمكن التراجع عنه.
                  </p>
                </section>
              </div>
            )}

          </div>
        </main>
      </div>

      {/* Employee Status Edit Modal */}
      {empModal.isOpen && (
        <div className="fixed inset-0 bg-[#002623]/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-xl max-w-md w-full p-6 space-y-4 border border-[#b9a779]">
            <h3 className="font-bold text-[#002623] text-base">تعديل بيانات الموظف</h3>
            <form onSubmit={saveEmployeeStatusModal} className="space-y-4">
              <div>
                <label className="block text-xs font-semibold text-[#3d3a3b]/80 mb-1">المديرية</label>
                <input type="text" value={empModal.directorate} onChange={e => setEmpModal({ ...empModal, directorate: e.target.value })} required className="w-full px-3 py-2 text-xs bg-[#edebe0]/50 border border-[#b9a779] rounded-xl focus:ring-2 focus:ring-[#054239] focus:outline-none" placeholder="أدخل المديرية" />
              </div>
              <div>
                <label className="block text-xs font-semibold text-[#3d3a3b]/80 mb-1">القسم</label>
                <input type="text" value={empModal.department} onChange={e => setEmpModal({ ...empModal, department: e.target.value })} className="w-full px-3 py-2 text-xs bg-[#edebe0]/50 border border-[#b9a779] rounded-xl focus:ring-2 focus:ring-[#054239] focus:outline-none" placeholder="أدخل قسم الموظف (اختياري)" />
              </div>
              <div>
                <label className="block text-xs font-semibold text-[#3d3a3b]/80 mb-1">حالة الموظف الوظيفية</label>
                <select value={empModal.status} onChange={e => setEmpModal({ ...empModal, status: e.target.value })} className="w-full px-3 py-2 text-xs bg-[#edebe0]/50 border border-[#b9a779] rounded-xl focus:ring-2 focus:ring-[#054239] focus:outline-none">
                  <option value="نشط">نشط</option>
                  <option value="مفصول">مفصول</option>
                  <option value="مستقيل">مستقيل</option>
                  <option value="منقول">منقول</option>
                </select>
              </div>
              <div>
                <label className="block text-xs font-semibold text-[#3d3a3b]/80 mb-1">تاريخ تغيير الحالة</label>
                <input type="date" value={empModal.date} onChange={e => setEmpModal({ ...empModal, date: e.target.value })} required className="w-full px-3 py-2 text-xs bg-[#edebe0]/50 border border-[#b9a779] rounded-xl focus:ring-2 focus:ring-[#054239] focus:outline-none" />
              </div>
              <div className="flex justify-end space-x-2 space-x-reverse pt-2">
                <button type="button" onClick={() => setEmpModal({ isOpen: false, empId: '', status: 'نشط', date: '', directorate: '', department: '' })} className="px-4 py-2 bg-[#edebe0] hover:bg-[#b9a779] text-[#3d3a3b] text-xs font-bold rounded-xl">إلغاء</button>
                <button type="submit" className="px-4 py-2 bg-[#054239] hover:bg-[#002623] text-[#edebe0] text-xs font-bold rounded-xl shadow-sm border border-[#428177]">حفظ البيانات</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Notification Toast */}
      {toast.show && (
        <div role={toast.type === 'error' ? 'alert' : 'status'} aria-live={toast.type === 'error' ? 'assertive' : 'polite'} className="fixed bottom-5 left-5 bg-[#002623] text-[#edebe0] px-5 py-3 rounded-2xl shadow-2xl text-xs flex items-center space-x-3 space-x-reverse z-50 border border-[#428177]">
          <i className={`fa-solid ${toast.type === 'error' ? 'fa-triangle-exclamation text-rose-400' : toast.type === 'warning' ? 'fa-circle-exclamation text-amber-300' : 'fa-check-circle text-emerald-400'} text-base`} aria-hidden="true"></i>
          <span>{toast.message}</span>
        </div>
      )}
    </div>
  );
}