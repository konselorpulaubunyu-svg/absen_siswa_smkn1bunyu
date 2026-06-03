/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo } from 'react';
import { 
  Users, 
  LayoutDashboard, 
  ClipboardCheck, 
  LogOut, 
  FileUp, 
  Search, 
  Calendar as CalendarIcon, 
  ChevronRight,
  School,
  CheckCircle2,
  AlertCircle,
  FileSpreadsheet,
  Filter,
  Download,
  Eye,
  EyeOff,
  Trash2,
  Edit,
  Plus,
  X,
  GraduationCap,
  Clock,
  Camera,
  RotateCcw
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { 
  Document, 
  Packer, 
  Paragraph, 
  TextRun, 
  Table, 
  TableRow, 
  TableCell, 
  WidthType, 
  AlignmentType,
  HeadingLevel
} from 'docx';
import { saveAs } from 'file-saver';

// --- Types ---
type Student = {
  nomor: string;
  nama: string;
  kelas: string;
  jurusan: string;
};

type AttendanceStatus = 'H' | 'S' | 'I' | 'A' | 'T' | '';

type AttendanceItem = {
  nama: string;
  status: AttendanceStatus;
  evidence?: string; // Base64 image
  notes?: string;
};

type User = {
  name: string;
  username: string;
};

type Page = 'landing' | 'register' | 'login' | 'dashboard' | 'students' | 'attendance' | 'recap' | 'teachers';

type TeacherAttendance = {
  id: string;
  nama: string;
  jamMasuk: string;
  keterangan: string;
  hariTanggal: string; // YYYY-MM-DD
};

type TeacherMaster = {
  nomor: string;
  nama: string;
  foto?: string; // base64 string
};

const DEFAULT_TEACHERS: TeacherMaster[] = [
  { nomor: "1", nama: "Kokom Komariyah, S.Pd", foto: "" },
  { nomor: "2", nama: "Rismadamayanti, M.Pd", foto: "" },
  { nomor: "3", nama: "Nurlaelah, S.Pd", foto: "" },
  { nomor: "4", nama: "Kristiani Nainggolan, S.Pd", foto: "" },
  { nomor: "5", nama: "Donmarth Sianturi, M.Pd", foto: "" },
  { nomor: "6", nama: "Mursaling, M.Pd", foto: "" },
  { nomor: "7", nama: "Eka Patria Krisna, ST", foto: "" },
  { nomor: "8", nama: "Robertus Lolok, S.Pd.K", foto: "" },
  { nomor: "9", nama: "Ana Puspitasari, S.Pd", foto: "" },
  { nomor: "10", nama: "Hotmauli Gultom, S.Pd", foto: "" },
  { nomor: "11", nama: "Raihanatun Nisa, S.Pd.I", foto: "" },
  { nomor: "12", nama: "Lia Anggraeni, S.Pd", foto: "" },
  { nomor: "13", nama: "Amelia Sri Rezky K., ST", foto: "" },
  { nomor: "14", nama: "Reymond Edward S., S.Pd", foto: "" },
  { nomor: "15", nama: "Anita Wahyuni, S.Pd", foto: "" },
  { nomor: "16", nama: "M. Furqon Izzul Jalali, ST", foto: "" },
  { nomor: "17", nama: "Syamsiar, S.Pd", foto: "" },
  { nomor: "18", nama: "Abd. Rajab, ST", foto: "" },
  { nomor: "19", nama: "Novianti, ST", foto: "" },
  { nomor: "20", nama: "Yenny Debora Ponni, S.Pd", foto: "" },
  { nomor: "21", nama: "Fathiyatu Rizqilah, SE", foto: "" },
  { nomor: "22", nama: "M. Nur Rafli, A.Md", foto: "" },
  { nomor: "23", nama: "Andrew Novandi", foto: "" }
];

const DEFAULT_USERS = [
  { name: "Andrian", username: "Andrian", password: "Andrian2026" }
];

const DEFAULT_STUDENTS: Student[] = [
  { nomor: "26001", nama: "Aditya Pratama", kelas: "XII", jurusan: "RPL 1" },
  { nomor: "26002", nama: "Amanda Putri", kelas: "XII", jurusan: "RPL 1" },
  { nomor: "26003", nama: "Bagas Saputra", kelas: "XII", jurusan: "RPL 2" },
  { nomor: "26004", nama: "Citra Lestari", kelas: "XII", jurusan: "RPL 2" },
  { nomor: "26005", nama: "Dimas Wijaya", kelas: "XI", jurusan: "RPL 1" },
  { nomor: "26006", nama: "Eka Rahmawati", kelas: "XI", jurusan: "RPL 1" },
  { nomor: "26007", nama: "Fajar Nugraha", kelas: "XI", jurusan: "RPL 2" },
  { nomor: "26008", nama: "Gita Cahyani", kelas: "XI", jurusan: "RPL 2" },
  { nomor: "26009", nama: "Hendra Wijaya", kelas: "X", jurusan: "RPL 1" },
  { nomor: "26010", nama: "Indah Permatasari", kelas: "X", jurusan: "RPL 1" }
];

const generateDefaultAttendance = (): Record<string, AttendanceItem[]> => {
  const dates: string[] = [];
  const today = new Date();
  for (let i = 0; i < 5; i++) {
    const d = new Date();
    d.setDate(today.getDate() - i);
    if (d.getDay() !== 0 && d.getDay() !== 6) {
      dates.push(d.toISOString().split('T')[0]);
    }
  }
  
  const history: Record<string, AttendanceItem[]> = {};
  dates.forEach((dateString) => {
    history[dateString] = [
      { nama: "Aditya Pratama", status: "H" },
      { nama: "Amanda Putri", status: "H" },
      { nama: "Bagas Saputra", status: "H" },
      { nama: "Citra Lestari", status: "S", notes: "Demam tinggi" },
      { nama: "Dimas Wijaya", status: "H" },
      { nama: "Eka Rahmawati", status: "I", notes: "Acara keluarga" },
      { nama: "Fajar Nugraha", status: "H" },
      { nama: "Gita Cahyani", status: "T" },
      { nama: "Hendra Wijaya", status: "H" },
      { nama: "Indah Permatasari", status: "H" }
    ];
  });
  
  return history;
};

const generateDefaultTeachersAttendance = (): TeacherAttendance[] => {
  const list: TeacherAttendance[] = [];
  const todayStr = new Date().toISOString().split('T')[0];
  
  list.push({
    id: "T-1",
    nama: "Kokom Komariyah, S.Pd",
    jamMasuk: "07:15",
    keterangan: "Mengajar Kelas XII RPL 1",
    hariTanggal: todayStr
  });
  list.push({
    id: "T-2",
    nama: "Nurlaelah, S.Pd",
    jamMasuk: "07:30",
    keterangan: "Rapat Guru",
    hariTanggal: todayStr
  });
  list.push({
    id: "T-3",
    nama: "Kristiani Nainggolan, S.Pd",
    jamMasuk: "07:45",
    keterangan: "Piket Sekolah",
    hariTanggal: todayStr
  });
  
  return list;
};

// --- Components ---

const Sidebar = ({ activePage, onNavigate, onLogout, onDeleteAccount, logo, isOpen, onClose }: { 
  activePage: Page, 
  onNavigate: (p: Page) => void, 
  onLogout: () => void, 
  onDeleteAccount: () => void, 
  logo: string | null,
  isOpen: boolean,
  onClose: () => void
}) => {
  const [showConfirmDelete, setShowConfirmDelete] = useState(false);
  const menuItems = [
    { id: 'dashboard', icon: LayoutDashboard, label: 'Dashboard' },
    { id: 'students', icon: Users, label: 'Data Siswa' },
    { id: 'attendance', icon: ClipboardCheck, label: 'Input Absensi' },
    { id: 'recap', icon: FileSpreadsheet, label: 'Rekap Absensi' },
    { id: 'teachers', icon: GraduationCap, label: 'Absensi Guru' },
  ];

  return (
    <>
      {/* Mobile Overlay */}
      <AnimatePresence>
        {isOpen && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            onClick={onClose}
            className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-30 lg:hidden"
          />
        )}
      </AnimatePresence>

      <motion.div 
        className={`w-64 h-screen bg-slate-900 flex flex-col fixed left-0 top-0 z-40 transition-transform duration-300 transform lg:translate-x-0 ${isOpen ? 'translate-x-0' : '-translate-x-full'}`}
      >
        <div className="p-6 border-b border-slate-800 flex items-center justify-between">
          <div className="flex items-center gap-3">
            {logo ? (
              <div className="w-10 h-10 bg-white/10 rounded-lg flex items-center justify-center overflow-hidden p-1">
                <img src={logo} className="w-full h-full object-contain" alt="Logo" />
              </div>
            ) : (
              <School className="text-white w-8 h-8" />
            )}
            <div>
              <h1 className="text-white font-bold text-sm tracking-tight leading-none uppercase">SMKN 1 Bunyu</h1>
              <p className="text-slate-500 text-[9px] mt-1 font-bold uppercase tracking-tighter">Absensi Digital</p>
            </div>
          </div>
          <button onClick={onClose} className="lg:hidden text-slate-400 hover:text-white">
            <X className="w-6 h-6" />
          </button>
        </div>
        
        <nav className="flex-1 px-4 py-6 space-y-1.5 overflow-y-auto">
          {menuItems.map((item) => (
            <button
              key={item.id}
              onClick={() => { onNavigate(item.id as Page); onClose(); }}
              className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all duration-200 group relative ${
                activePage === item.id 
                  ? 'bg-blue-600 text-white shadow-lg shadow-blue-600/20' 
                  : 'text-slate-400 hover:bg-slate-800/50 hover:text-white'
              }`}
            >
              <item.icon className={`w-5 h-5 ${activePage === item.id ? 'text-white' : 'text-slate-500 group-hover:text-white'}`} />
              <span className="font-bold text-sm">{item.label}</span>
              {activePage === item.id && (
                <motion.div layoutId="activeNav" className="absolute right-2 w-1.5 h-1.5 rounded-full bg-white opacity-40" />
              )}
            </button>
          ))}

          <div className="pt-4 mt-4 border-t border-slate-800/50">
            {!showConfirmDelete ? (
              <button 
                onClick={() => setShowConfirmDelete(true)}
                className="w-full flex items-center gap-3 px-4 py-3 text-slate-500 hover:bg-rose-500/10 hover:text-rose-400 rounded-xl transition-all font-bold text-xs uppercase tracking-widest group"
              >
                <Trash2 className="w-4 h-4 opacity-50 group-hover:opacity-100" />
                <span>Hapus Akun</span>
              </button>
            ) : (
              <div className="bg-rose-600/10 border border-rose-500/20 rounded-xl p-3 animate-in fade-in zoom-in duration-300">
                <p className="text-[9px] text-rose-400 font-black text-center mb-3 uppercase tracking-widest leading-tight">Yakin Hapus Permanen?</p>
                <div className="flex gap-2">
                  <button 
                    onClick={onDeleteAccount}
                    className="flex-1 bg-rose-600 text-white py-2 rounded-lg text-[9px] font-black hover:bg-rose-500 transition-colors uppercase tracking-widest"
                  >Hapus</button>
                  <button 
                    onClick={() => setShowConfirmDelete(false)}
                    className="flex-1 bg-slate-800 text-slate-300 py-2 rounded-lg text-[9px] font-black hover:bg-slate-700 transition-colors uppercase tracking-widest"
                  >Batal</button>
                </div>
              </div>
            )}
          </div>
          <div className="mt-auto p-4 flex flex-col items-center gap-2">
            <button 
              onClick={() => {
                if(confirm('Aplikasi akan dimuat ulang untuk memperbarui sistem. Lanjutkan?')) {
                  if ('serviceWorker' in navigator) {
                    navigator.serviceWorker.getRegistrations().then(registrations => {
                      for(let registration of registrations) {
                        registration.unregister();
                      }
                    });
                  }
                  window.location.reload();
                }
              }}
              className="text-[9px] font-black text-blue-400 hover:text-blue-300 uppercase tracking-widest bg-blue-500/10 px-3 py-1.5 rounded-lg transition-colors"
            >
              Perbarui Aplikasi
            </button>
            <div className="text-[8px] font-black text-slate-700 uppercase tracking-[0.2em] opacity-40">Version 2026.05.05.0.01</div>
          </div>
        </nav>

        <div className="p-4 mb-4 border-t border-slate-800/50">
          <button 
            onClick={onLogout}
            className="w-full flex items-center gap-3 px-4 py-3 text-slate-400 hover:bg-rose-500/10 hover:text-rose-400 rounded-xl transition-all font-bold group"
          >
            <div className="w-8 h-8 rounded-lg bg-slate-800 flex items-center justify-center group-hover:bg-rose-500/20 transition-colors">
              <LogOut className="w-4 h-4 group-hover:-translate-x-0.5 transition-transform" />
            </div>
            <span className="text-sm">Keluar Sesi</span>
          </button>
        </div>
      </motion.div>
    </>
  );
};


const Header = ({ title, user, onLogout, logo, onToggleSidebar }: { title: string, user: User | null, onLogout: () => void, logo: string | null, onToggleSidebar: () => void }) => (
  <header className="h-20 bg-white/80 backdrop-blur-md border-b border-slate-200 flex items-center justify-between px-4 sm:px-8 sticky top-0 z-20 w-full">
    <div className="flex items-center gap-3 sm:gap-4">
      <button 
        onClick={onToggleSidebar}
        className="lg:hidden p-2 text-slate-600 hover:bg-slate-100 rounded-lg transition-colors"
      >
        <LayoutDashboard className="w-6 h-6" />
      </button>

      {logo && (
        <div className="hidden xs:flex w-10 h-10 sm:w-12 sm:h-12 bg-slate-50 border border-slate-100 rounded-xl items-center justify-center overflow-hidden p-1 sm:p-2">
          <img src={logo} className="w-full h-full object-contain" alt="Logo" />
        </div>
      )}
      <div>
        <h2 className="text-sm sm:text-xl font-bold text-slate-800 line-clamp-1">{title}</h2>
        <p className="text-[10px] sm:text-xs text-slate-500 italic font-medium whitespace-nowrap">
          {new Intl.DateTimeFormat('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }).format(new Date())}
        </p>
      </div>
    </div>
    
    <div className="flex items-center gap-2 sm:gap-6">
      <div className="flex items-center gap-2 sm:gap-4 lg:border-r lg:border-slate-100 lg:pr-6">
        <div className="text-right hidden md:block">
          <p className="text-sm font-bold text-slate-800">{user?.name}</p>
          <p className="text-[10px] uppercase font-bold text-slate-400 tracking-wider font-mono">{user?.username}</p>
        </div>
        <div className="w-8 h-8 sm:w-10 sm:h-10 rounded-full bg-blue-600 shadow-lg shadow-blue-600/20 flex items-center justify-center font-bold text-white text-xs sm:text-sm">
          {user?.name?.[0]?.toUpperCase()}
        </div>
      </div>
      <button 
        onClick={onLogout}
        className="hidden sm:flex items-center gap-2 text-rose-500 hover:text-rose-600 font-bold text-sm transition-colors group"
      >
        <LogOut className="w-4 h-4 group-hover:-translate-x-1 transition-transform" />
        <span>Keluar</span>
      </button>
    </div>
  </header>
);

const DeleteAccountModal = ({ 
  isOpen, 
  onClose, 
  user, 
  onConfirm 
}: { 
  isOpen: boolean, 
  onClose: () => void, 
  user: User | null, 
  onConfirm: (password: string, clearAllData: boolean) => boolean 
}) => {
  const [password, setPassword] = useState('');
  const [clearAllData, setClearAllData] = useState(false);
  const [showPassword, setShowPassword] = useState(false);
  const [error, setError] = useState<string | null>(null);

  if (!isOpen || !user) return null;

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!password.trim()) {
      setError('Masukkan password Anda untuk konfirmasi!');
      return;
    }
    setError(null);
    const success = onConfirm(password, clearAllData);
    if (!success) {
      setError('Password salah! Gagal menghapus akun.');
    } else {
      setPassword('');
      setClearAllData(false);
    }
  };

  return (
    <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/65 backdrop-blur-md">
      <motion.div 
        initial={{ opacity: 0, scale: 0.95, y: 10 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.95, y: 10 }}
        className="w-full max-w-md bg-white rounded-3xl overflow-hidden shadow-2xl border border-slate-100 flex flex-col"
      >
        {/* Decorative Top header */}
        <div className="bg-gradient-to-r from-red-500 to-rose-600 p-6 text-white text-center relative">
          <div className="absolute top-4 right-4">
            <button 
              type="button"
              onClick={onClose}
              className="text-white/60 hover:text-white p-1 rounded-full hover:bg-white/10 transition-colors"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
          <div className="w-12 h-12 bg-white/20 rounded-full flex items-center justify-center mx-auto mb-3 shadow-inner">
            <Trash2 className="w-6 h-6 text-white" />
          </div>
          <h3 className="font-extrabold text-lg uppercase tracking-wider">Hapus Akun Permanen</h3>
          <p className="text-white/80 text-[11px] font-bold uppercase tracking-widest mt-1">Konfirmasi Keamanan</p>
        </div>

        <form onSubmit={handleSubmit} className="p-6 space-y-5">
          <div className="bg-red-50 border border-red-100 rounded-2xl p-4 flex gap-3">
            <AlertCircle className="w-5 h-5 text-red-600 shrink-0 mt-0.5" />
            <div className="space-y-1">
              <h4 className="text-xs font-black text-red-800 uppercase tracking-wide">Peringatan Kritis!</h4>
              <p className="text-xs text-red-700 leading-relaxed font-medium">
                Tindakan ini akan menghapus akun operator <span className="font-black">@{user.username} ({user.name})</span> secara permanen dari basis data lokal aplikasi ini.
              </p>
            </div>
          </div>

          <div className="space-y-4">
            {/* Password input */}
            <div className="space-y-1.5">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider block">Masukkan Password Anda</label>
              <div className="relative">
                <input 
                  type={showPassword ? 'text' : 'password'}
                  placeholder="Ketik password login Anda"
                  className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-3 outline-none focus:ring-2 focus:ring-red-500/20 text-slate-800 font-medium text-sm transition-all"
                  value={password}
                  onChange={(e) => {
                    setPassword(e.target.value);
                    setError(null);
                  }}
                  required
                />
                <button
                  type="button"
                  onClick={() => setShowPassword(!showPassword)}
                  className="absolute right-3.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 p-1"
                >
                  {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                </button>
              </div>
            </div>

            {/* Checkbox Wipe Data */}
            <label className="flex items-start gap-3 bg-slate-50 hover:bg-slate-100 p-3.5 rounded-2xl cursor-pointer transition-colors border border-slate-100 group select-none">
              <input 
                type="checkbox"
                className="mt-0.5 rounded text-red-600 focus:ring-red-500/20 w-4 h-4 cursor-pointer"
                checked={clearAllData}
                onChange={(e) => setClearAllData(e.target.checked)}
              />
              <div className="space-y-0.5">
                <span className="text-xs font-black text-slate-700 uppercase tracking-wide group-hover:text-amber-700 transition-colors">Ikut Hapus Semua Data Sekolah</span>
                <p className="text-[10px] text-slate-500 font-medium leading-relaxed">
                  Centang untuk menghapus seluruh data siswa, riwayat absensi harian, daftar guru, dan config logo sekolah dari perangkat ini (Wipe/Reset).
                </p>
              </div>
            </label>
          </div>

          {error && (
            <div className="text-xs text-red-600 bg-red-50 border border-red-100 p-2.5 rounded-xl font-bold flex items-center gap-2">
              <AlertCircle className="w-3.5 h-3.5 shrink-0" />
              <span>{error}</span>
            </div>
          )}

          {/* Action buttons */}
          <div className="flex gap-3 pt-2">
            <button
              type="button"
              onClick={() => {
                onClose();
                setPassword('');
                setClearAllData(false);
                setError(null);
              }}
              className="flex-1 bg-slate-100 hover:bg-slate-200 text-slate-600 font-bold py-3 rounded-xl text-xs transition-style text-center uppercase tracking-wider"
            >
              Batal
            </button>
            <button
              type="submit"
              className="flex-1 bg-rose-600 hover:bg-rose-700 text-white font-bold py-3 rounded-xl text-xs transition-all uppercase tracking-wider text-center shadow-lg shadow-rose-600/15"
            >
              Hapus Akun
            </button>
          </div>
        </form>
      </motion.div>
    </div>
  );
};

const LandingPage = ({ onStart, logo, onLogoChange }: { onStart: () => void, logo: string | null, onLogoChange: (file: File) => void }) => (
  <div className="min-h-screen bg-slate-900 flex flex-col items-center justify-center p-4 relative overflow-hidden text-white">
    <div className="absolute inset-0 z-0 opacity-20 pointer-events-none">
        <div className="absolute top-[-10%] left-[-10%] w-[40%] h-[40%] bg-blue-600 blur-[120px] rounded-full" />
        <div className="absolute bottom-[-10%] right-[-10%] w-[50%] h-[50%] bg-indigo-600 blur-[120px] rounded-full" />
    </div>
    
    <motion.div 
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      className="text-center z-10 max-w-2xl px-6 flex flex-col items-center"
    >
      <div className="relative group mb-8">
        <div className="w-40 h-40 bg-white rounded-3xl flex items-center justify-center shadow-2xl shadow-blue-600/20 overflow-hidden border-2 border-slate-700/50 p-5">
          {logo ? (
            <img src={logo} className="w-full h-full object-contain" alt="School Logo" />
          ) : (
            <div className="w-full h-full bg-slate-800 flex items-center justify-center">
              <School className="text-white w-14 h-14" />
            </div>
          )}
        </div>
        
        <label className="absolute -bottom-2 -right-2 w-10 h-10 bg-blue-600 hover:bg-blue-500 rounded-full flex items-center justify-center cursor-pointer shadow-lg transition-transform hover:scale-110 active:scale-95 border-4 border-slate-900">
          <Plus className="w-5 h-5 text-white" />
          <input 
            type="file" 
            className="hidden" 
            accept="image/*" 
            onChange={(e) => {
              const file = e.target.files?.[0];
              if (file) onLogoChange(file);
            }} 
          />
        </label>
        
        <div className="absolute -top-12 left-1/2 -translate-x-1/2 bg-slate-800 text-[10px] font-bold px-3 py-1.5 rounded-full border border-slate-700 opacity-0 group-hover:opacity-100 transition-opacity whitespace-nowrap pointer-events-none tracking-widest text-blue-400">
          UBAH LOGO SEKOLAH
        </div>
      </div>

      <h1 className="text-5xl sm:text-7xl font-bold tracking-tight mb-10">
        SMKN 1 BUNYU
      </h1>
      <button 
        type="button"
        onClick={onStart}
        className="px-10 py-4 bg-blue-600 text-white rounded-lg font-bold text-lg hover:bg-blue-700 hover:scale-105 active:scale-95 transition-all shadow-xl shadow-blue-900/20 flex items-center gap-2 mx-auto group"
      >
        <span>Masuk ke Sistem</span>
        <ChevronRight className="w-5 h-5 group-hover:translate-x-1 transition-transform" />
      </button>
    </motion.div>
  </div>
);

const AuthForm = ({ type, onToggle, onAuthSuccess, onBack, logo }: { type: 'login' | 'register', onToggle: () => void, onAuthSuccess: (user: User) => void, onBack: () => void, logo: string | null }) => {
  const [formData, setFormData] = useState({ name: '', username: '', password: '' });
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState<{ type: 'success' | 'error', text: string } | null>(null);
  const [showPassword, setShowPassword] = useState(false);
  const [registeredUsers, setRegisteredUsers] = useState<any[]>([]);

  // States for local password recovery
  const [isRecovering, setIsRecovering] = useState(false);
  const [recoveryUsername, setRecoveryUsername] = useState('');
  const [recoveryFullName, setRecoveryFullName] = useState('');
  const [recoveredPassword, setRecoveredPassword] = useState<string | null>(null);

  useEffect(() => {
    const saved = localStorage.getItem('app_users');
    if (saved) {
      setRegisteredUsers(JSON.parse(saved));
    } else {
      setRegisteredUsers(DEFAULT_USERS);
      localStorage.setItem('app_users', JSON.stringify(DEFAULT_USERS));
    }
  }, [type, isRecovering]);

  const selectUser = (user: any) => {
    setFormData(prev => ({ ...prev, username: user.username }));
    setMessage(null);
  };

  const handleRecoverPassword = (e: React.FormEvent) => {
    e.preventDefault();
    const users = JSON.parse(localStorage.getItem('app_users') || '[]');
    const targetUser = users.find((u: any) => 
      u.username.toLowerCase().trim() === recoveryUsername.toLowerCase().trim()
    );

    if (!targetUser) {
      setMessage({ type: 'error', text: 'Username tidak ditemukan!' });
      return;
    }

    if (targetUser.name.toLowerCase().trim() !== recoveryFullName.toLowerCase().trim()) {
      setMessage({ type: 'error', text: 'Nama Lengkap tidak cocok dengan username tersebut!' });
      return;
    }

    setRecoveredPassword(targetUser.password);
    setMessage({ type: 'success', text: 'Identitas Anda telah berhasil diverifikasi!' });
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    setMessage(null);
    
    try {
      const users = JSON.parse(localStorage.getItem('app_users') || '[]');
      
      if (type === 'register') {
        const exists = users.find((u: any) => u.username === formData.username);
        if (exists) {
          throw new Error('Username sudah terdaftar.');
        }
        
        const newUser = {
          name: formData.name,
          username: formData.username,
          password: formData.password // Simpan password simpel untuk demo lokal
        };
        
        users.push(newUser);
        localStorage.setItem('app_users', JSON.stringify(users));

        setMessage({ type: 'success', text: 'Registrasi berhasil!' });
        setTimeout(() => onToggle(), 1500);
      } else {
        // Login
        const found = users.find((u: any) => u.username === formData.username && u.password === formData.password);
        
        if (found) {
          const userData = { name: found.name, username: found.username };
          setMessage({ type: 'success', text: 'Login berhasil!' });
          setTimeout(() => onAuthSuccess(userData), 1000);
        } else {
          throw new Error('Username atau password salah.');
        }
      }
    } catch (err: any) {
      setMessage({ type: 'error', text: err.message || 'Terjadi kesalahan' });
    } finally {
      setLoading(false);
    }
  };

  return (
    <motion.div 
      initial={{ opacity: 0, scale: 0.95 }}
      animate={{ opacity: 1, scale: 1 }}
      className="max-w-md w-full p-6 sm:p-10 bg-white rounded-3xl shadow-2xl border border-slate-200 z-10 relative mx-4"
    >
      <div className="mb-10 text-center">
        <div className="w-24 h-24 bg-slate-50 border border-slate-100 rounded-2xl flex items-center justify-center mx-auto mb-4 overflow-hidden shadow-lg p-3">
          {logo ? (
            <img src={logo} className="w-full h-full object-contain" alt="Logo" />
          ) : (
            <div className="w-full h-full bg-blue-600 rounded-xl flex items-center justify-center">
              <School className="text-white w-8 h-8" />
            </div>
          )}
        </div>
        <h2 className="text-2xl font-bold text-slate-800 tracking-tight">
          {isRecovering 
            ? 'Pemulihan Sandi' 
            : type === 'login' 
              ? 'Selamat Datang' 
              : 'Registrasi Guru'}
        </h2>
        <p className="text-slate-500 text-sm font-medium mt-1">
          Sistem Absensi SMKN 1 BUNYU
        </p>

        {type === 'login' && !isRecovering && registeredUsers.length > 0 && (
          <div className="mt-8">
            <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-4">Pilih Akun Terdaftar</p>
            <div className="flex flex-wrap justify-center gap-4 max-h-40 overflow-y-auto py-2 px-1">
              {registeredUsers.map((u, i) => (
                <button
                  key={i}
                  type="button"
                  onClick={() => selectUser(u)}
                  className={`group flex flex-col items-center gap-2 p-2 rounded-xl transition-all duration-300 ${
                    formData.username === u.username 
                      ? 'bg-blue-600 shadow-lg shadow-blue-500/20 scale-105' 
                      : 'hover:bg-slate-50'
                  }`}
                >
                  <div className={`w-12 h-12 rounded-full flex items-center justify-center font-bold text-sm transition-all shadow-sm ${
                    formData.username === u.username
                      ? 'bg-white text-blue-600'
                      : 'bg-slate-100 text-slate-500 group-hover:bg-blue-50 group-hover:text-blue-600'
                  }`}>
                    {u.name?.[0]?.toUpperCase()}
                  </div>
                  <span className={`text-[10px] font-bold truncate max-w-[80px] ${
                    formData.username === u.username ? 'text-white' : 'text-slate-500'
                  }`}>
                    {u.name.split(' ')[0]}
                  </span>
                </button>
              ))}
            </div>
          </div>
        )}
      </div>

      {message && (
        <div className={`mb-6 p-4 rounded-lg flex items-center gap-3 text-sm font-semibold animate-in fade-in slide-in-from-top-4 duration-300 ${
          message.type === 'success' ? 'bg-emerald-50 text-emerald-700 border border-emerald-100' : 'bg-rose-50 text-rose-700 border border-rose-100'
        }`}>
          {message.type === 'success' ? <CheckCircle2 className="w-5 h-5 text-emerald-500" /> : <AlertCircle className="w-5 h-5 text-rose-500" />}
          {message.text}
        </div>
      )}

      {isRecovering ? (
        <form onSubmit={handleRecoverPassword} className="space-y-5">
          {recoveredPassword ? (
            <div className="space-y-4">
              <div className="bg-emerald-50 border border-emerald-100 rounded-2xl p-4 text-center space-y-2">
                <CheckCircle2 className="w-8 h-8 text-emerald-500 mx-auto" />
                <h4 className="text-xs font-bold text-emerald-800 uppercase tracking-wide">Kata Sandi Anda Ditemukan!</h4>
                <div className="bg-white rounded-xl border border-emerald-200 py-3 px-4 shadow-sm inline-block min-w-[200px]">
                  <span className="text-lg font-mono font-bold tracking-wider text-slate-800 select-all">
                    {recoveredPassword}
                  </span>
                </div>
                <p className="text-[10px] text-emerald-700 leading-relaxed max-w-xs mx-auto">
                  Silakan gunakan kata sandi di atas untuk masuk kembali ke sistem.
                </p>
              </div>

              <button
                type="button"
                onClick={() => {
                  setFormData(prev => ({ ...prev, username: recoveryUsername, password: recoveredPassword }));
                  setIsRecovering(false);
                  setRecoveredPassword(null);
                  setRecoveryUsername('');
                  setRecoveryFullName('');
                  setMessage(null);
                }}
                className="w-full py-4 bg-blue-600 text-white rounded-lg font-bold text-sm hover:bg-blue-700 shadow-lg shadow-blue-900/20 active:scale-[0.98] transition-all"
              >
                Kembali & Masuk
              </button>
            </div>
          ) : (
            <>
              <p className="text-xs text-slate-500 leading-relaxed text-center">
                Masukkan username dan nama lengkap terdaftar Anda untuk memulihkan kata sandi secara instan.
              </p>
              
              <div className="space-y-4 pt-2">
                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2 ml-1">Username Operator</label>
                  {registeredUsers.length > 0 ? (
                    <select
                      className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-lg focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 outline-none font-medium transition-all text-slate-700 cursor-pointer"
                      value={recoveryUsername}
                      onChange={e => {
                        setRecoveryUsername(e.target.value);
                        setMessage(null);
                      }}
                      required
                    >
                      <option value="">Pilih Username</option>
                      {registeredUsers.map(u => (
                        <option key={u.username} value={u.username}>@{u.username}</option>
                      ))}
                    </select>
                  ) : (
                    <input 
                      required
                      type="text" 
                      placeholder="Ketik username Anda"
                      className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-lg focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 outline-none font-medium transition-all text-slate-700"
                      value={recoveryUsername}
                      onChange={e => {
                        setRecoveryUsername(e.target.value);
                        setMessage(null);
                      }}
                    />
                  )}
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2 ml-1">Nama Lengkap Terdaftar</label>
                  <input 
                    required
                    type="text" 
                    placeholder="Contoh: Ahmad Syarif, S.Pd"
                    className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-lg focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 outline-none font-medium transition-all text-slate-700"
                    value={recoveryFullName}
                    onChange={e => {
                      setRecoveryFullName(e.target.value);
                      setMessage(null);
                    }}
                  />
                  <div className="text-[10px] text-slate-400 mt-1.5 leading-relaxed italic">
                    *Harus sama persis dengan nama lengkap saat pendaftaran akun.
                  </div>
                </div>
              </div>

              <div className="flex gap-3 pt-4">
                <button
                  type="button"
                  onClick={() => {
                    setIsRecovering(false);
                    setRecoveryUsername('');
                    setRecoveryFullName('');
                    setMessage(null);
                  }}
                  className="flex-1 py-3.5 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-lg font-bold text-xs uppercase tracking-wider transition-all"
                >
                  Batal
                </button>
                <button
                  type="submit"
                  className="flex-1 py-3.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-bold text-xs uppercase tracking-wider transition-all shadow-md shadow-blue-500/10"
                >
                  Tampilkan Sandi
                </button>
              </div>
            </>
          )}
        </form>
      ) : (
        <form onSubmit={handleSubmit} className="space-y-5">
          {type === 'register' && (
            <div>
              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2 ml-1">Nama Lengkap</label>
              <input 
                required
                type="text" 
                placeholder="Ahmad Syarif, S.Pd"
                className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-lg focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 outline-none font-medium transition-all text-slate-700"
                value={formData.name}
                onChange={e => setFormData({ ...formData, name: e.target.value })}
              />
            </div>
          )}
          <div>
            <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2 ml-1">Username</label>
            <input 
              required
              type="text" 
              placeholder="username_anda"
              className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-lg focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 outline-none font-medium transition-all text-slate-700"
              value={formData.username}
              onChange={e => setFormData({ ...formData, username: e.target.value })}
            />
          </div>
          <div>
            <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2 ml-1">Kata Sandi</label>
            <div className="relative">
              <input 
                required
                type={showPassword ? "text" : "password"} 
                placeholder="••••••••"
                className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-lg focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 outline-none font-medium transition-all text-slate-700 pr-12"
                value={formData.password}
                onChange={e => setFormData({ ...formData, password: e.target.value })}
              />
              <button 
                type="button"
                onClick={() => setShowPassword(!showPassword)}
                className="absolute right-3 top-1/2 -translate-y-1/2 p-1 text-slate-400 hover:text-slate-600 transition-colors"
              >
                {showPassword ? <EyeOff className="w-5 h-5" /> : <Eye className="w-5 h-5" />}
              </button>
            </div>
          </div>
          <button 
            type="submit"
            disabled={loading}
            className="w-full py-4 bg-blue-600 text-white rounded-lg font-bold text-sm hover:bg-blue-700 shadow-lg shadow-blue-900/20 active:scale-[0.98] transition-all disabled:opacity-50 mt-4"
          >
            {loading ? 'Memproses...' : (type === 'login' ? 'Masuk Sekarang' : 'Daftar Akun')}
          </button>
        </form>
      )}

      <div className="mt-8 text-center pt-8 border-t border-slate-100 flex flex-col gap-4">
        {!isRecovering && (
          <>
            <button 
              type="button"
              onClick={onToggle}
              className="text-blue-600 font-bold hover:text-blue-700 transition-colors text-sm"
            >
              {type === 'login' ? 'Belum punya akun? Silahkan daftar disini' : 'Sudah punya akun? Kembali ke Login'}
            </button>
            {type === 'login' && (
              <button 
                type="button"
                onClick={() => {
                  setIsRecovering(true);
                  setMessage(null);
                }}
                className="text-slate-400 font-medium hover:text-slate-600 transition-colors text-[10px] uppercase tracking-wider focus:outline-none"
              >
                Lupa Sandi?
              </button>
            )}
          </>
        )}
        <button 
          type="button"
          onClick={() => {
            if (isRecovering) {
              setIsRecovering(false);
              setRecoveredPassword(null);
              setRecoveryUsername('');
              setRecoveryFullName('');
              setMessage(null);
            } else {
              onBack();
            }
          }}
          className="text-slate-400 font-bold hover:text-slate-600 transition-colors text-xs uppercase tracking-widest"
        >
          {isRecovering ? 'Kembali ke Login' : 'Kembali ke Beranda'}
        </button>
      </div>
    </motion.div>
  );
};

const generateWordReport = async (title: string, date: string, data: { student: Student, att?: AttendanceItem }[]) => {
  const rows = [
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "No. Induk", bold: true })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Nama", bold: true })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Kelas/Jurusan", bold: true })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Status", bold: true })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Catatan", bold: true })] })] }),
      ],
    }),
  ];

  for (const item of data) {
    const { student, att } = item;
    const statusText = att?.status === 'H' ? 'Hadir' : 
                      att?.status === 'S' ? 'Sakit' : 
                      att?.status === 'I' ? 'Izin' : 
                      att?.status === 'A' ? 'Alfa' : 
                      att?.status === 'T' ? 'Terlambat' : 'Belum Absen';

    const notesText = att?.notes || att?.evidence || "-";

    rows.push(
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph(student.nomor)] }),
          new TableCell({ children: [new Paragraph(student.nama)] }),
          new TableCell({ children: [new Paragraph(`${student.kelas} - ${student.jurusan}`)] }),
          new TableCell({ children: [new Paragraph(statusText)] }),
          new TableCell({ children: [new Paragraph(notesText)] }),
        ],
      })
    );
  }

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({
            text: title,
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
          }),
          new Paragraph({
            text: `Tanggal: ${new Intl.DateTimeFormat('id-ID', { dateStyle: 'full' }).format(new Date(date))}`,
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 },
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: rows,
          }),
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, `Laporan_Absensi_${date}.docx`);
};

// --- Page Components ---

const Dashboard = ({ stats, deferredPrompt, onInstall }: { stats: any, deferredPrompt: any, onInstall: () => void }) => (
  <div className="p-4 sm:p-8 space-y-6 sm:space-y-8">
    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
      {[
        { label: 'Total Siswa', value: stats.totalStudents || 0, icon: Users, bgColor: 'bg-indigo-50', borderColor: 'border-indigo-100', textColor: 'text-indigo-700', valColor: 'text-indigo-800' },
        { label: 'Total Guru', value: stats.totalTeachers || 0, icon: GraduationCap, bgColor: 'bg-blue-50', borderColor: 'border-blue-100', textColor: 'text-blue-700', valColor: 'text-blue-800' },
      ].map((card, i) => (
        <motion.div 
          key={i}
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ delay: i * 0.1 }}
          className={`${card.bgColor} ${card.borderColor} border p-8 rounded-xl flex items-center justify-between group hover:shadow-md transition-all duration-300`}
        >
          <div>
            <p className={`text-xs font-bold uppercase ${card.textColor} tracking-wider mb-1`}>{card.label}</p>
            <h3 className={`text-4xl font-black ${card.valColor}`}>{card.value}</h3>
          </div>
          <div className="w-12 h-12 flex items-center justify-center opacity-40 group-hover:scale-110 transition-transform">
            <card.icon className={`w-8 h-8 ${card.textColor}`} />
          </div>
        </motion.div>
      ))}
    </div>

    {deferredPrompt && (
      <motion.div 
        initial={{ opacity: 0, scale: 0.95 }}
        animate={{ opacity: 1, scale: 1 }}
        className="bg-blue-600 rounded-2xl p-6 sm:p-8 text-white shadow-xl shadow-blue-500/20 flex flex-col sm:flex-row items-center justify-between gap-6"
      >
        <div className="flex items-center gap-4">
          <div className="w-16 h-16 bg-white/20 rounded-2xl flex items-center justify-center backdrop-blur-md">
            <Download className="w-8 h-8 text-white" />
          </div>
          <div>
            <h4 className="text-xl font-bold">Instal Aplikasi Presensi</h4>
            <p className="text-blue-100 text-sm font-medium mt-1">Akses lebih cepat & lancar langsung dari layar utama perangkat Anda.</p>
          </div>
        </div>
        <button 
          onClick={onInstall}
          className="w-full sm:w-auto px-8 py-3 bg-white text-blue-600 rounded-xl font-bold hover:bg-blue-50 transition-all shadow-lg active:scale-95 whitespace-nowrap"
        >
          Instal Sekarang
        </button>
      </motion.div>
    )}

    <div className="bg-white p-8 rounded-xl border border-slate-200 shadow-sm">
      <h3 className="text-xl font-bold text-slate-800 mb-4">Selamat Datang di Portal Guru</h3>
      <p className="text-slate-600 font-medium leading-relaxed max-w-2xl">
        Gunakan menu di samping untuk mengelola data siswa atau mencatat absensi harian. Pastikan data siswa sudah diimpor melalui file Excel sebelum memulai absensi.
      </p>
    </div>
  </div>
);

const StudentsPage = ({ students, onImport, onDeleteStudent, onDeleteAll, onUpdateStudent }: { 
  students: Student[], 
  onImport: (students: Student[]) => void,
  onDeleteStudent: (nomor: string) => void,
  onDeleteAll: () => void,
  onUpdateStudent: (oldNomor: string, updated: Student) => void
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [filter, setFilter] = useState('');
  const [editingStudent, setEditingStudent] = useState<Student | null>(null);
  const [isAdding, setIsAdding] = useState(false);
  const [formData, setFormData] = useState<Student>({ nomor: '', nama: '', kelas: '', jurusan: '' });

  const filteredStudents = useMemo(() => {
    return students.filter(s => {
      const matchFilter = filter ? s.kelas === filter : true;
      const matchSearch = searchTerm ? s.nama.toLowerCase().includes(searchTerm.toLowerCase()) : true;
      return matchFilter && matchSearch;
    });
  }, [students, filter, searchTerm]);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        let allStudents: Student[] = [];

        wb.SheetNames.forEach((sheetName) => {
          const ws = wb.Sheets[sheetName];
          const data = XLSX.utils.sheet_to_json(ws) as any[];
          
          if (data.length > 0) {
            const parsed: Student[] = data.map((item: any) => {
              const findKey = (search: string) => {
                const keys = Object.keys(item);
                return keys.find(k => k.toLowerCase() === search.toLowerCase());
              };

              const keyNomor = findKey('NOMOR') || findKey('NO') || findKey('NUMBER');
              const keyNama = findKey('NAMA') || findKey('NAME');
              const keyKelas = findKey('KELAS') || findKey('CLASS');
              const keyJurusan = findKey('JURUSAN') || findKey('MAJOR') || findKey('DEPARTMENT');

              return {
                nomor: String(item[keyNomor || ''] || ''),
                nama: String(item[keyNama || ''] || ''),
                kelas: String(item[keyKelas || ''] || ''),
                jurusan: String(item[keyJurusan || ''] || ''),
              };
            });
            
            allStudents = [...allStudents, ...parsed.filter(s => s.nama)];
          }
        });
        
        if (allStudents.length > 0) {
          onImport(allStudents);
          alert(`Berhasil mengimpor ${allStudents.length} siswa dari ${wb.SheetNames.length} sheet.`);
        } else {
          alert('Tidak ada data yang valid ditemukan di semua sheet. Pastikan ada kolom NAMA, NOMOR, KELAS, JURUSAN.');
        }
      } catch (err) {
        alert('Gagal membaca file Excel. Pastikan format file benar.');
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = '';
  };

  const handleEditClick = (s: Student) => {
    setEditingStudent(s);
    setFormData({ ...s });
  };

  const handleSave = () => {
    if (!formData.nomor || !formData.nama) {
      alert('Nama dan Nomor harus diisi');
      return;
    }
    if (editingStudent) {
      onUpdateStudent(editingStudent.nomor, formData);
    } else {
      onImport([...students, formData]);
    }
    setEditingStudent(null);
    setIsAdding(false);
    setFormData({ nomor: '', nama: '', kelas: '', jurusan: '' });
  };

  const handleExportStudents = () => {
    if (students.length === 0) {
      alert('Tidak ada data siswa untuk diunduh.');
      return;
    }
    try {
      const exportData = students.map(s => ({
        'NOMOR': s.nomor,
        'NAMA': s.nama,
        'KELAS': s.kelas,
        'JURUSAN': s.jurusan
      }));
      const ws = XLSX.utils.json_to_sheet(exportData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Siswa');
      XLSX.writeFile(wb, 'Database_Siswa.xlsx');
    } catch (err) {
      alert('Gagal mengunduh database siswa.');
    }
  };

  return (
    <div className="p-4 sm:p-8 space-y-6 sm:space-y-8">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex items-center gap-3 bg-slate-50 border border-slate-200 rounded-xl px-4 py-3 w-full md:max-w-xs focus-within:ring-2 focus-within:ring-blue-500/20 transition-all">
          <Search className="w-5 h-5 text-slate-400" />
          <input 
            type="text" 
            placeholder="Cari nama siswa..." 
            className="flex-1 bg-transparent outline-none font-medium text-slate-700 text-sm"
            value={searchTerm}
            onChange={e => setSearchTerm(e.target.value)}
          />
        </div>

        <div className="flex flex-wrap items-center gap-3">
          <div className="relative">
            <select 
              className="appearance-none bg-slate-50 border border-slate-200 text-slate-700 text-sm rounded-lg block w-44 p-2.5 pr-8 focus:ring-blue-500/20 focus:border-blue-500 outline-none cursor-pointer font-medium"
              value={filter}
              onChange={e => setFilter(e.target.value)}
            >
              <option value="">Semua Kelas</option>
              {Array.from(new Set(students.map(s => s.kelas))).sort().map(k => (
                <option key={k} value={k}>{k}</option>
              ))}
            </select>
            <div className="pointer-events-none absolute inset-y-0 right-0 flex items-center px-2 text-slate-400">
              <Filter className="w-4 h-4" />
            </div>
          </div>
          
          <button 
            onClick={() => { setIsAdding(true); setFormData({ nomor: '', nama: '', kelas: '', jurusan: '' }); }}
            className="flex items-center gap-2 bg-green-600 hover:bg-green-700 text-white px-4 py-2.5 rounded-lg text-sm font-semibold shadow-sm transition-all"
          >
            <Plus className="w-4 h-4" />
            Tambah
          </button>

          <label className="flex items-center gap-2 bg-blue-600 hover:bg-blue-700 text-white px-4 py-2.5 rounded-lg text-sm font-semibold shadow-sm transition-all cursor-pointer">
            <FileUp className="w-4 h-4" />
            <span>Import Excel</span>
            <input type="file" className="hidden" accept=".xlsx, .xls" onChange={handleFileUpload} />
          </label>

          <button 
            onClick={handleExportStudents}
            disabled={students.length === 0}
            className="flex items-center gap-2 bg-amber-600 hover:bg-amber-700 text-white px-4 py-2.5 rounded-lg text-sm font-semibold shadow-sm transition-all disabled:opacity-50"
          >
            <Download className="w-4 h-4" />
            <span>Download Database</span>
          </button>

          <button 
            onClick={() => {
              if(confirm('Yakin ingin menghapus semua data siswa?')) onDeleteAll();
            }}
            className="flex items-center gap-2 bg-red-50 hover:bg-red-100 text-red-600 border border-red-200 px-4 py-2.5 rounded-lg text-sm font-semibold transition-all"
          >
            <Trash2 className="w-4 h-4" />
            <span>Kosongkan</span>
          </button>
        </div>
      </div>

      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm">
        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead className="bg-slate-50 border-b border-slate-200">
              <tr>
                <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider">Nomor</th>
                <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider">Nama Lengkap</th>
                <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider">Kelas</th>
                <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider">Jurusan</th>
                <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider text-right">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {filteredStudents.length > 0 ? filteredStudents.map((s, i) => (
                <tr key={`${s.nomor}-${i}`} className="hover:bg-slate-50 transition-colors">
                  <td className="px-6 py-4 text-sm font-mono text-slate-500">{s.nomor}</td>
                  <td className="px-6 py-4 font-semibold text-slate-800">{s.nama}</td>
                  <td className="px-6 py-4">
                    <span className="px-2.5 py-1 bg-slate-100 rounded-md text-[10px] font-bold text-slate-600 uppercase">{s.kelas}</span>
                  </td>
                  <td className="px-6 py-4 text-sm text-slate-600">{s.jurusan}</td>
                  <td className="px-6 py-4 text-sm text-right">
                    <div className="flex justify-end gap-1">
                      <button 
                        onClick={() => handleEditClick(s)}
                        className="p-2 text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
                      >
                        <Edit className="w-4 h-4" />
                      </button>
                      <button 
                        onClick={() => {
                          if(confirm(`Hapus siswa ${s.nama}?`)) onDeleteStudent(s.nomor);
                        }}
                        className="p-2 text-red-600 hover:bg-red-50 rounded-lg transition-colors"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </div>
                  </td>
                </tr>
              )) : (
                <tr>
                  <td colSpan={5} className="px-6 py-16 text-center text-slate-400 font-medium italic">
                    {searchTerm || filter ? 'Data tidak ditemukan.' : 'Belum ada data siswa. Silakan klik "Import Excel" atau "Tambah".'}
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      <AnimatePresence>
        {(editingStudent || isAdding) && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-sm">
            <motion.div 
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-2xl shadow-xl w-full max-w-md overflow-hidden border border-slate-200"
            >
              <div className="px-6 py-4 bg-slate-50 border-b border-slate-200 flex justify-between items-center">
                <h3 className="font-bold text-slate-900">{isAdding ? 'Tambah Siswa' : 'Edit Siswa'}</h3>
                <button onClick={() => { setEditingStudent(null); setIsAdding(false); }} className="text-slate-400 hover:text-slate-600"><X className="w-5 h-5"/></button>
              </div>
              <div className="p-6 space-y-4">
                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase mb-1">No. Induk</label>
                  <input 
                    className="w-full px-4 py-2 border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 text-sm"
                    value={formData.nomor}
                    onChange={e => setFormData({...formData, nomor: e.target.value})}
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase mb-1">Nama Lengkap</label>
                  <input 
                    className="w-full px-4 py-2 border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 text-sm"
                    value={formData.nama}
                    onChange={e => setFormData({...formData, nama: e.target.value})}
                  />
                </div>
                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="block text-xs font-bold text-slate-500 uppercase mb-1">Kelas</label>
                    <input 
                      className="w-full px-4 py-2 border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 text-sm"
                      value={formData.kelas}
                      onChange={e => setFormData({...formData, kelas: e.target.value})}
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-slate-500 uppercase mb-1">Jurusan</label>
                    <input 
                      className="w-full px-4 py-2 border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 text-sm"
                      value={formData.jurusan}
                      onChange={e => setFormData({...formData, jurusan: e.target.value})}
                    />
                  </div>
                </div>
              </div>
              <div className="px-6 py-4 bg-slate-50 border-t border-slate-200 flex gap-2">
                <button 
                  onClick={() => { setEditingStudent(null); setIsAdding(false); }}
                  className="flex-1 py-2.5 text-sm font-bold text-slate-600 hover:bg-slate-100 rounded-lg transition-all"
                >Batal</button>
                <button 
                  onClick={handleSave}
                  className="flex-1 py-2.5 text-sm font-bold bg-blue-600 text-white hover:bg-blue-700 rounded-lg shadow-lg shadow-blue-900/10 transition-all"
                >
                  {isAdding ? 'Tambah Siswa' : 'Simpan Perubahan'}
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
};

const AttendancePage = ({ students, attendance, date, onDateChange, onStatusChange, onSave, onEvidenceChange, onDeleteStudent, onUpdateStudent }: { 
  students: Student[], 
  attendance: AttendanceItem[], 
  date: string, 
  onDateChange: (d: string) => void,
  onStatusChange: (nama: string, status: AttendanceStatus) => void,
  onSave: () => void,
  onEvidenceChange: (nama: string, evidence: string) => void,
  onDeleteStudent: (id: string) => void,
  onUpdateStudent: (id: string, data: Student) => void
}) => {
  const [filter, setFilter] = useState('');
  const [previewImage, setPreviewImage] = useState<string | null>(null);
  const [editingStudent, setEditingStudent] = useState<Student | null>(null);
  const [formData, setFormData] = useState<Student>({ nomor: '', nama: '', kelas: '', jurusan: '' });

  const filteredStudents = useMemo(() => {
    return students.filter(s => filter ? s.kelas === filter : true);
  }, [students, filter]);

  const handleImageUpload = (nama: string, e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const base64 = evt.target?.result as string;
      onEvidenceChange(nama, base64);
    };
    reader.readAsDataURL(file);
  };

  const handleEditClick = (s: Student) => {
    setEditingStudent(s);
    setFormData(s);
  };

  const handleSaveEdit = () => {
    if (editingStudent) {
      onUpdateStudent(editingStudent.nomor, formData);
      setEditingStudent(null);
    }
  };

  const handleDownload = () => {
    if (!students.length) return;
    const dataForExcel = students.filter(s => filter ? s.kelas === filter : true).map(s => {
      const statusItem = attendance.find(a => a.nama === s.nama);
      const status = statusItem ? statusItem.status : '';
      return {
        'No. Induk': s.nomor,
        'Nama Siswa': s.nama,
        'Kelas': s.kelas,
        'Jurusan': s.jurusan,
        'Status': status === 'H' ? 'Hadir' : 
                  status === 'S' ? 'Sakit' : 
                  status === 'I' ? 'Izin' : 
                  status === 'A' ? 'Alfa' : 
                  status === 'T' ? 'Terlambat' : 'Belum Absen',
        'Catatan': status === 'T' ? (statusItem?.notes || statusItem?.evidence || '') : '-'
      };
    });
    const ws = XLSX.utils.json_to_sheet(dataForExcel);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rekap Absensi");
    XLSX.writeFile(wb, `Rekap_Absensi_${date}.xlsx`);
  };

  const handleDownloadPDF = () => {
    const doc = new jsPDF();
    const formattedDate = new Intl.DateTimeFormat('id-ID', { dateStyle: 'full' }).format(new Date(date));
    
    doc.setFontSize(16);
    doc.text('LAPORAN ABSENSI HARIAN', 105, 15, { align: 'center' });
    doc.setFontSize(12);
    doc.text(`SMKN 1 BUNYU - ${formattedDate}`, 105, 25, { align: 'center' });
    
    const tableData = students
      .filter(s => filter ? s.kelas === filter : true)
      .map((s, i) => {
        const att = attendance.find(a => a.nama === s.nama);
        const statusText = att?.status === 'H' ? 'Hadir' : 
                           att?.status === 'S' ? 'Sakit' : 
                           att?.status === 'I' ? 'Izin' : 
                           att?.status === 'A' ? 'Alfa' : 
                           att?.status === 'T' ? `Terlambat (${att?.notes || att?.evidence || ''})` : '-';
        return [
          i + 1,
          s.nama,
          s.kelas,
          statusText
        ];
      });

    autoTable(doc, {
      startY: 35,
      head: [['No', 'Nama Siswa', 'Kelas', 'Status']],
      body: tableData,
      theme: 'grid',
      headStyles: { fillColor: [37, 99, 235] }
    });

    doc.save(`Absensi_${date}.pdf`);
  };

  const handleDownloadWord = async () => {
    if (!students.length) return;
    const data = students
      .filter(s => filter ? s.kelas === filter : true)
      .map(s => {
        const att = attendance.find(a => a.nama === s.nama);
        return { student: s, att };
      });
    await generateWordReport("Laporan Absensi Harian", date, data);
  };

  return (
    <div className="p-4 sm:p-8 space-y-6 sm:space-y-8">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
        <div className="flex items-center gap-3">
          <div className="relative">
            <CalendarIcon className="w-4 h-4 absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" />
            <input 
              type="date" 
              className="pl-10 pr-4 py-2.5 bg-slate-50 border border-slate-200 text-slate-700 text-sm rounded-lg outline-none font-medium focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500"
              value={date}
              onChange={e => onDateChange(e.target.value)}
            />
          </div>
          <div className="relative">
            <select 
              className="appearance-none bg-slate-50 border border-slate-200 text-slate-700 text-sm rounded-lg block w-40 p-2.5 pr-8 focus:ring-blue-500/20 focus:border-blue-500 outline-none cursor-pointer font-medium"
              value={filter}
              onChange={e => setFilter(e.target.value)}
            >
              <option value="">Semua Kelas</option>
              <option value="X">Kelas X</option>
              <option value="XI">Kelas XI</option>
              <option value="XII">Kelas XII</option>
              <option value="XIII">Kelas XIII</option>
            </select>
            <div className="pointer-events-none absolute inset-y-0 right-0 flex items-center px-2 text-slate-400">
              <Filter className="w-4 h-4" />
            </div>
          </div>
        </div>

        <div className="flex items-center gap-2">
          <button 
            onClick={handleDownload}
            disabled={students.length === 0}
            className="bg-emerald-600 hover:bg-emerald-700 text-white px-5 py-2.5 rounded-lg text-sm font-semibold shadow-sm transition-all disabled:opacity-50 flex items-center gap-2"
          >
            <Download className="w-4 h-4" />
            <span>Excel</span>
          </button>

          <button 
            onClick={handleDownloadWord}
            disabled={students.length === 0}
            className="bg-blue-600 hover:bg-blue-700 text-white px-5 py-2.5 rounded-lg text-sm font-semibold shadow-sm transition-all disabled:opacity-50 flex items-center gap-2"
          >
            <Download className="w-4 h-4" />
            <span>Word</span>
          </button>
          
          <button 
            onClick={handleDownloadPDF}
            disabled={students.length === 0}
            className="bg-rose-600 hover:bg-rose-700 text-white px-5 py-2.5 rounded-lg text-sm font-semibold shadow-sm transition-all disabled:opacity-50 flex items-center gap-2"
          >
            <Download className="w-4 h-4" />
            <span>PDF</span>
          </button>
        </div>
      </div>

      <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead className="bg-slate-50 border-b border-slate-200">
              <tr>
                <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider">Nama Siswa</th>
                <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider text-center">Kelas</th>
                <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider text-center">Status Absensi</th>
                <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider text-center">Catatan</th>
                <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider text-right">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {filteredStudents.length > 0 ? filteredStudents.map((s, i) => {
                const att = attendance.find(a => a.nama === s.nama);
                const status = att?.status || '';
                const evidence = att?.evidence;
                
                const getStatusStyle = () => {
                  switch(status) {
                    case 'H': return 'bg-emerald-50 text-emerald-800 border-emerald-200';
                    case 'S': return 'bg-indigo-50 text-indigo-800 border-indigo-200';
                    case 'I': return 'bg-amber-50 text-amber-800 border-amber-200';
                    case 'A': return 'bg-rose-50 text-rose-800 border-rose-200';
                    case 'T': return 'bg-orange-50 text-orange-800 border-orange-200';
                    default: return 'bg-slate-50 text-slate-400 border-slate-200';
                  }
                };
                return (
                  <tr key={`${s.nomor}-${i}`} className="hover:bg-slate-50 transition-colors">
                    <td className="px-6 py-4 font-semibold text-slate-800">{s.nama}</td>
                    <td className="px-6 py-4 text-center">
                      <span className="px-2.5 py-1 bg-slate-100 rounded-md text-[10px] font-bold text-slate-600 uppercase">{s.kelas}</span>
                    </td>
                    <td className="px-6 py-4 flex justify-center">
                      <select 
                        className={`appearance-none text-xs font-bold rounded-full px-5 py-1.5 border cursor-pointer outline-none transition-all ${getStatusStyle()}`}
                        value={status}
                        onChange={e => onStatusChange(s.nama, e.target.value as AttendanceStatus)}
                      >
                        <option value="">-- Pilih --</option>
                        <option value="H">H (Hadir)</option>
                        <option value="S">S (Sakit)</option>
                        <option value="I">I (Izin)</option>
                        <option value="A">A (Alpa)</option>
                        <option value="T">T (Terlambat)</option>
                      </select>
                    </td>
                    <td className="px-6 py-4">
                      <div className="flex flex-col items-center gap-2">
                        {status === 'T' ? (
                          <input 
                            type="text" 
                            placeholder="Alasan terlambat..." 
                            className="w-full max-w-[200px] px-3.5 py-2 border border-slate-250 bg-slate-50 text-slate-800 rounded-xl outline-none focus:ring-2 focus:ring-orange-500/20 focus:border-orange-500 text-xs font-semibold focus:bg-white transition-all shadow-inner"
                            value={att?.notes || att?.evidence || ''}
                            onChange={(e) => onEvidenceChange(s.nama, e.target.value)}
                          />
                        ) : status ? (
                          <span className="text-slate-350 text-xs">—</span>
                        ) : (
                          <span className="text-[10px] text-slate-400 italic">Pilih status dulu</span>
                        )}
                      </div>
                    </td>
                    <td className="px-6 py-4 text-right">
                      <div className="flex justify-end gap-1">
                        <button 
                          onClick={() => handleEditClick(s)}
                          className="p-1.5 text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
                          title="Edit Siswa"
                        >
                          <Edit className="w-3.5 h-3.5" />
                        </button>
                        <button 
                          onClick={() => {
                            if(confirm(`Hapus siswa ${s.nama}?`)) onDeleteStudent(s.nomor);
                          }}
                          className="p-1.5 text-red-600 hover:bg-red-50 rounded-lg transition-colors"
                          title="Hapus Siswa"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              }) : (
                <tr>
                  <td colSpan={5} className="px-6 py-16 text-center text-slate-400 font-medium italic">
                    {students.length === 0 ? 'Belum ada data siswa. Impor terlebih dahulu di menu Data Siswa.' : 'Tidak ada siswa di kelas ini.'}
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      <div className="flex justify-end pt-2">
        <button 
          onClick={onSave}
          disabled={students.length === 0}
          className="bg-blue-600 hover:bg-blue-700 text-white px-8 py-3 rounded-xl text-sm font-bold shadow-md hover:shadow-lg transition-all disabled:opacity-50 flex items-center gap-2"
        >
          <CheckCircle2 className="w-4 h-4" />
          <span>Simpan Absensi</span>
        </button>
      </div>

      <AnimatePresence>
        {previewImage && (
          <div className="fixed inset-0 z-[60] flex items-center justify-center p-8 bg-slate-900/90 backdrop-blur-md">
            <motion.div 
              initial={{ opacity: 0, scale: 0.9 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.9 }}
              className="relative max-w-4xl w-full flex flex-col items-center bg-white rounded-3xl p-4 overflow-hidden"
            >
              <button 
                onClick={() => setPreviewImage(null)}
                className="absolute top-4 right-4 p-3 bg-white/20 hover:bg-white/40 text-slate-800 rounded-full transition-all shadow-xl z-20"
              >
                <X className="w-6 h-6" />
              </button>
              <div className="aspect-[3/4] max-h-[80vh] w-full bg-slate-100 rounded-2xl overflow-hidden flex items-center justify-center">
                <img src={previewImage} alt="Surat Keterangan Sakit" className="max-w-full max-h-full object-contain" />
              </div>
              <p className="mt-4 font-bold text-slate-900 text-lg uppercase tracking-wider">Surat Keterangan Sakit</p>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {editingStudent && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-sm">
            <motion.div 
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-2xl shadow-xl w-full max-w-md overflow-hidden border border-slate-200"
            >
              <div className="px-6 py-4 bg-slate-50 border-b border-slate-200 flex justify-between items-center">
                <h3 className="font-bold text-slate-900">Edit Siswa</h3>
                <button onClick={() => setEditingStudent(null)} className="text-slate-400 hover:text-slate-600"><X className="w-5 h-5"/></button>
              </div>
              <div className="p-6 space-y-4">
                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase mb-1">No. Induk</label>
                  <input 
                    className="w-full px-4 py-2 border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 text-sm"
                    value={formData.nomor}
                    onChange={e => setFormData({...formData, nomor: e.target.value})}
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase mb-1">Nama Lengkap</label>
                  <input 
                    className="w-full px-4 py-2 border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 text-sm"
                    value={formData.nama}
                    onChange={e => setFormData({...formData, nama: e.target.value})}
                  />
                </div>
                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="block text-xs font-bold text-slate-500 uppercase mb-1">Kelas</label>
                    <input 
                      className="w-full px-4 py-2 border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 text-sm"
                      value={formData.kelas}
                      onChange={e => setFormData({...formData, kelas: e.target.value})}
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-slate-500 uppercase mb-1">Jurusan</label>
                    <input 
                      className="w-full px-4 py-2 border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 text-sm"
                      value={formData.jurusan}
                      onChange={e => setFormData({...formData, jurusan: e.target.value})}
                    />
                  </div>
                </div>
              </div>
              <div className="px-6 py-4 bg-slate-50 border-t border-slate-200 flex gap-2">
                <button 
                  onClick={() => setEditingStudent(null)}
                  className="flex-1 py-2.5 text-sm font-bold text-slate-600 hover:bg-slate-100 rounded-lg transition-all"
                >Batal</button>
                <button 
                  onClick={handleSaveEdit}
                  className="flex-1 py-2.5 text-sm font-bold bg-blue-600 text-white hover:bg-blue-700 rounded-lg shadow-lg shadow-blue-900/10 transition-all"
                >Simpan Perubahan</button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
};

const RecapPage = ({ students, attendanceHistory, onSelectDate, onPreviewImage, onUpdateAttendanceRecord, onDeleteAttendanceRecord, onDeleteMonthAttendance }: { 
  students: Student[], 
  attendanceHistory: Record<string, AttendanceItem[]>,
  onSelectDate: (d: string) => void,
  onPreviewImage: (img: string) => void,
  onUpdateAttendanceRecord: (date: string, nama: string, status: AttendanceStatus) => void,
  onDeleteAttendanceRecord: (date: string, nama: string) => void,
  onDeleteMonthAttendance: (month: string, nama: string) => void
}) => {
  const [filterKelas, setFilterKelas] = useState('');
  const [filterJurusan, setFilterJurusan] = useState('');
  const [selectedMonth, setSelectedMonth] = useState(new Date().toISOString().slice(0, 7)); // YYYY-MM
  const [editingRecord, setEditingRecord] = useState<{ date: string, nama: string, status: AttendanceStatus } | null>(null);

  const dates = useMemo(() => {
    return Object.keys(attendanceHistory)
      .filter(d => d.startsWith(selectedMonth))
      .sort((a, b) => b.localeCompare(a));
  }, [attendanceHistory, selectedMonth]);

  const classes = useMemo(() => Array.from(new Set(students.map(s => s.kelas))).sort(), [students]);
  const majors = useMemo(() => Array.from(new Set(students.map(s => s.jurusan))).sort(), [students]);

  const handleDownloadPDF = () => {
    const doc = new jsPDF();
    const monthName = new Intl.DateTimeFormat('id-ID', { month: 'long', year: 'numeric' }).format(new Date(selectedMonth));
    
    doc.setFontSize(18);
    doc.text('LAPORAN REKAPITULASI ABSENSI SISWA', 105, 15, { align: 'center' });
    doc.setFontSize(14);
    doc.text(`SMKN 1 BUNYU - Periode: ${monthName}`, 105, 25, { align: 'center' });
    
    const summaryData = studentSummary.map(({ student, summary }) => [
      student.nomor,
      student.nama,
      student.kelas,
      summary.H,
      summary.S,
      summary.I,
      summary.A,
      summary.T
    ]);

    autoTable(doc, {
      startY: 35,
      head: [['No', 'Nama Siswa', 'Kelas', 'H', 'S', 'I', 'A', 'T']],
      body: summaryData,
      theme: 'grid',
      headStyles: { fillColor: [37, 99, 235], halign: 'center' },
      columnStyles: {
        3: { halign: 'center' },
        4: { halign: 'center' },
        5: { halign: 'center' },
        6: { halign: 'center' },
        7: { halign: 'center' },
      }
    });

    doc.save(`Rekap_Absensi_${selectedMonth}.pdf`);
  };

  const filteredStudents = useMemo(() => {
    return students.filter(s => {
      const matchKelas = filterKelas ? s.kelas === filterKelas : true;
      const matchJurusan = filterJurusan ? s.jurusan === filterJurusan : true;
      return matchKelas && matchJurusan;
    });
  }, [students, filterKelas, filterJurusan]);

  const studentSummary = useMemo(() => {
    return filteredStudents.map(s => {
      const summary = { H: 0, S: 0, I: 0, A: 0, T: 0 };
      dates.forEach(date => {
        const att = attendanceHistory[date]?.find(a => a.nama === s.nama);
        if (att?.status && (att.status in summary)) {
          summary[att.status as keyof typeof summary]++;
        }
      });
      return { student: s, summary };
    });
  }, [filteredStudents, dates, attendanceHistory]);

  return (
    <div className="p-4 sm:p-8 space-y-6 sm:space-y-8">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-6 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
        <div className="flex flex-wrap items-center gap-4">
          <div className="space-y-1">
            <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider">Bulan</label>
            <input 
              type="month" 
              className="px-4 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm font-medium outline-none"
              value={selectedMonth}
              onChange={e => setSelectedMonth(e.target.value)}
            />
          </div>
          <div className="space-y-1">
            <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider">Kelas</label>
            <select 
              className="px-4 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm font-medium outline-none min-w-[120px]"
              value={filterKelas}
              onChange={e => setFilterKelas(e.target.value)}
            >
              <option value="">Semua Kelas</option>
              {classes.map(k => <option key={k} value={k}>{k}</option>)}
            </select>
          </div>
          <div className="space-y-1">
            <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider">Jurusan</label>
            <select 
              className="px-4 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm font-medium outline-none min-w-[150px]"
              value={filterJurusan}
              onChange={e => setFilterJurusan(e.target.value)}
            >
              <option value="">Semua Jurusan</option>
              {majors.map(j => <option key={j} value={j}>{j}</option>)}
            </select>
          </div>
        </div>
        <div className="flex flex-wrap gap-2">
          <button 
            onClick={handleDownloadPDF}
            className="flex items-center gap-2 bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-bold shadow-sm transition-all"
          >
            <Download className="w-4 h-4" />
            <span>Unduh Laporan PDF</span>
          </button>
          <button 
            onClick={() => {
              const data = studentSummary.map(({ student, summary }) => ({
                'No. Induk': student.nomor,
                'Nama': student.nama,
                'Hadir': summary.H,
                'Sakit': summary.S,
                'Izin': summary.I,
                'Alfa': summary.A,
                'Terlambat': summary.T,
                'Total Hari': dates.length
              }));
              const ws = XLSX.utils.json_to_sheet(data);
              const wb = XLSX.utils.book_new();
              XLSX.utils.book_append_sheet(wb, ws, "Rekap Bulanan");
              XLSX.writeFile(wb, `Rekap_Bulanan_${selectedMonth}.xlsx`);
            }}
            className="flex items-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg text-sm font-bold shadow-sm transition-all"
          >
            <Download className="w-4 h-4" />
            <span>Unduh Excel Bulanan</span>
          </button>
        </div>
      </div>

      <div className="grid grid-cols-1 gap-6">
        <motion.div 
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          className="bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-sm"
        >
          <div className="p-4 bg-slate-50 border-b border-slate-200">
            <h4 className="font-bold text-slate-800 flex items-center gap-2">
              <FileSpreadsheet className="w-5 h-5 text-blue-600" />
              Rekap Kehadiran Bulanan
            </h4>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-left text-sm">
              <thead className="bg-white text-[10px] font-bold text-slate-400 uppercase tracking-widest border-b border-slate-100">
                <tr>
                  <th className="px-6 py-4">Nama Siswa</th>
                  <th className="px-6 py-4 text-center bg-emerald-50/30 text-emerald-600">Hadir (H)</th>
                  <th className="px-6 py-4 text-center bg-blue-50/30 text-blue-600">Sakit (S)</th>
                  <th className="px-6 py-4 text-center bg-amber-50/30 text-amber-600">Izin (I)</th>
                  <th className="px-6 py-4 text-center bg-rose-50/30 text-rose-600">Alfa (A)</th>
                  <th className="px-6 py-4 text-center bg-orange-50/30 text-orange-600">Terlambat (T)</th>
                  <th className="px-6 py-4 text-right">Aksi</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-50">
                {studentSummary.length > 0 ? studentSummary.map(({ student, summary }, i) => (
                  <tr key={`${student.nomor}-${i}`} className="hover:bg-slate-50/50">
                    <td className="px-6 py-4">
                      <div className="font-semibold text-slate-700">{student.nama}</div>
                      <div className="text-[10px] text-slate-400 font-medium uppercase">{student.kelas} • {student.jurusan}</div>
                    </td>
                    <td className="px-6 py-4 text-center font-bold text-emerald-600">{summary.H}</td>
                    <td className="px-6 py-4 text-center font-bold text-blue-600">{summary.S}</td>
                    <td className="px-6 py-4 text-center font-bold text-amber-600">{summary.I}</td>
                    <td className="px-6 py-4 text-center font-bold text-rose-600">{summary.A}</td>
                    <td className="px-6 py-4 text-center font-bold text-orange-600">{summary.T}</td>
                    <td className="px-6 py-4 text-right">
                      <button 
                        onClick={() => { if(confirm(`Hapus SEMUA data absensi ${student.nama} untuk bulan ${selectedMonth}?`)) onDeleteMonthAttendance(selectedMonth, student.nama); }}
                        className="p-1.5 text-rose-600 hover:bg-rose-50 rounded-lg transition-colors border border-rose-100"
                        title="Hapus Semua Absensi Bulan Ini"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </td>
                  </tr>
                )) : (
                  <tr>
                    <td colSpan={7} className="px-6 py-10 text-center text-slate-400 italic">Belum ada data siswa.</td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </motion.div>

        {/* Daily History with Evidence */}
        <div className="space-y-6">
          <h3 className="text-xl font-bold text-slate-800 flex items-center gap-2 px-1">
            <CalendarIcon className="w-6 h-6 text-blue-600" />
            Riwayat Absensi Harian
          </h3>
          
          {dates.length > 0 ? dates.map(date => (
            <motion.div 
              key={date}
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              className="bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-sm"
            >
              <div className="p-4 bg-slate-50 border-b border-slate-200 flex justify-between items-center">
                <div className="flex items-center gap-3">
                  <span className="bg-blue-600 text-white px-3 py-1 rounded-full text-[10px] font-black uppercase tracking-widest">
                    {new Intl.DateTimeFormat('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }).format(new Date(date))}
                  </span>
                  <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">
                    {attendanceHistory[date]?.filter(a => a.status).length || 0} Siswa Tercatat
                  </span>
                </div>
                <button 
                  onClick={() => onSelectDate(date)}
                  className="text-[10px] font-bold text-blue-600 hover:text-blue-700 uppercase tracking-widest flex items-center gap-1"
                >
                  Edit Hari Ini <ChevronRight className="w-3 h-3" />
                </button>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-left text-sm">
                  <thead className="bg-slate-50/50 text-[9px] font-bold text-slate-400 uppercase tracking-widest border-b border-slate-100">
                    <tr>
                      <th className="px-6 py-3">Nama Siswa</th>
                      <th className="px-6 py-3 text-center">Status</th>
                      <th className="px-6 py-3 text-center">Catatan</th>
                      <th className="px-6 py-3 text-right">Aksi</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-50">
                    {filteredStudents.map((s, i) => {
                      const att = attendanceHistory[date]?.find(a => a.nama === s.nama);
                      if (!att || !att.status) return null;

                      return (
                        <tr key={`${date}-${s.nomor}-${i}`} className="hover:bg-slate-50/30">
                          <td className="px-6 py-3 font-medium text-slate-700">{s.nama}</td>
                          <td className="px-6 py-3 text-center">
                            <span className={`px-2 py-0.5 rounded-full text-[9px] font-bold uppercase border ${
                              att.status === 'H' ? 'bg-emerald-50 text-emerald-700 border-emerald-100' :
                              att.status === 'S' ? 'bg-indigo-50 text-indigo-700 border-indigo-100' :
                              att.status === 'I' ? 'bg-amber-50 text-amber-700 border-amber-100' :
                              att.status === 'T' ? 'bg-orange-50 text-orange-700 border-orange-100' :
                              'bg-rose-50 text-rose-700 border-rose-100'
                            }`}>
                              {att.status === 'H' ? 'Hadir' : 
                               att.status === 'S' ? 'Sakit' : 
                               att.status === 'I' ? 'Izin' : 
                               att.status === 'T' ? 'Terlambat' : 'Alfa'}
                            </span>
                          </td>
                          <td className="px-6 py-3 text-center font-semibold text-slate-600 text-xs text-center">
                            {(att.notes || att.evidence) ? (
                              <span className="bg-slate-50 px-2.5 py-1 rounded-lg border border-slate-200 inline-block max-w-[200px] truncate" title={att.notes || att.evidence}>
                                {att.notes || att.evidence}
                              </span>
                            ) : (
                              <span className="text-slate-300 text-[10px]">—</span>
                            )}
                          </td>
                          <td className="px-6 py-3 text-right">
                            <button 
                              onClick={() => { if(confirm('Hapus record ini?')) onDeleteAttendanceRecord(date, s.nama); }}
                              className="p-1.5 text-slate-300 hover:text-rose-500 hover:bg-rose-50 rounded-lg transition-all"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </motion.div>
          )) : (
            <div className="bg-white border-2 border-dashed border-slate-200 rounded-2xl p-12 text-center text-slate-400">
              <CalendarIcon className="w-12 h-12 mx-auto mb-4 opacity-20" />
              <p className="font-bold uppercase tracking-widest text-xs">Belum Ada Riwayat untuk Bulan Ini</p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

const generateTeacherWordReport = async (title: string, data: TeacherAttendance[]) => {
  const rows = [
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "No", bold: true })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Hari, Tanggal", bold: true })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Nama Guru", bold: true })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Jam Masuk", bold: true })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Keterangan", bold: true })] })] }),
      ],
    }),
  ];

  let index = 1;
  for (const item of data) {
    const formattedDate = new Intl.DateTimeFormat('id-ID', { dateStyle: 'full' }).format(new Date(item.hariTanggal));
    rows.push(
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph(String(index++))] }),
          new TableCell({ children: [new Paragraph(formattedDate)] }),
          new TableCell({ children: [new Paragraph(item.nama)] }),
          new TableCell({ children: [new Paragraph(item.jamMasuk || '-')] }),
          new TableCell({ children: [new Paragraph(item.keterangan || '-')] }),
        ],
      })
    );
  }

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({
            text: title,
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: rows,
          }),
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, `Rekap_Absensi_Guru_${new Date().toISOString().split('T')[0]}.docx`);
};

const handleExportExcel = (filteredData: TeacherAttendance[]) => {
  try {
    const list = filteredData.map((item, idx) => ({
      'NO': idx + 1,
      'HARI & TANGGAL': new Intl.DateTimeFormat('id-ID', { dateStyle: 'full' }).format(new Date(item.hariTanggal)),
      'NAMA GURU': item.nama,
      'JAM MASUK': item.jamMasuk,
      'KETERANGAN': item.keterangan
    }));
    const ws = XLSX.utils.json_to_sheet(list);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Absensi Guru');
    XLSX.writeFile(wb, `Rekap_Absensi_Guru_${new Date().toISOString().split('T')[0]}.xlsx`);
  } catch (err) {
    alert('Gagal mengunduh rekapan Excel.');
  }
};

const handleExportPDF = (filteredData: TeacherAttendance[]) => {
  try {
    const doc = new jsPDF();
    doc.text('REKAP ABSENSI GURU', 14, 15);
    doc.setFontSize(10);
    doc.text(`Dicetak pada: ${new Intl.DateTimeFormat('id-ID', { dateStyle: 'full' }).format(new Date())}`, 14, 22);
    
    const tableData = filteredData.map((item, idx) => [
      idx + 1,
      new Intl.DateTimeFormat('id-ID', { dateStyle: 'full' }).format(new Date(item.hariTanggal)),
      item.nama,
      item.jamMasuk,
      item.keterangan
    ]);

    autoTable(doc, {
      startY: 28,
      head: [['No', 'Hari, Tanggal', 'Nama Guru', 'Jam Masuk', 'Keterangan']],
      body: tableData,
      theme: 'striped',
      headStyles: { fillColor: [59, 130, 246] }
    });

    doc.save(`Rekap_Absensi_Guru_${new Date().toISOString().split('T')[0]}.pdf`);
  } catch (err) {
    alert('Gagal mengunduh rekapan PDF.');
  }
};

const TeachersPage = ({ 
  teachersAttendance, 
  onAddAttendance, 
  onDeleteAttendance, 
  onUpdateAttendance,
  teachersList,
  onImportTeachersList,
  onAddTeacher,
  onDeleteTeacher,
  onClearTeachersList,
  onUpdateTeacherPhoto
}: {
  teachersAttendance: TeacherAttendance[];
  onAddAttendance: (item: Omit<TeacherAttendance, 'id'>) => void;
  onDeleteAttendance: (id: string) => void;
  onUpdateAttendance: (id: string, updated: Omit<TeacherAttendance, 'id'>) => void;
  teachersList: TeacherMaster[];
  onImportTeachersList: (teachers: TeacherMaster[]) => void;
  onAddTeacher: (teacher: TeacherMaster) => void;
  onDeleteTeacher: (nomor: string) => void;
  onClearTeachersList: () => void;
  onUpdateTeacherPhoto: (nomor: string, base64: string) => void;
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [filterMonth, setFilterMonth] = useState(new Date().toISOString().slice(0, 7)); // YYYY-MM
  const [filterDate, setFilterDate] = useState(''); // YYYY-MM-DD
  const [editingItem, setEditingItem] = useState<TeacherAttendance | null>(null);

  // Form State
  const [nama, setNama] = useState('');
  const [hariTanggal, setHariTanggal] = useState(() => new Date().toISOString().split('T')[0]);
  const [jamMasuk, setJamMasuk] = useState('');
  const [keterangan, setKeterangan] = useState('Hadir');
  const [keteranganCustom, setKeteranganCustom] = useState('');

  // Manual Master Add state
  const [manualNomor, setManualNomor] = useState('');
  const [manualNama, setManualNama] = useState('');
  const [isMasterOpen, setIsMasterOpen] = useState(false);

  const handlePhotoUpload = (nomor: string, e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const base64 = evt.target?.result as string;
      onUpdateTeacherPhoto(nomor, base64);
    };
    reader.readAsDataURL(file);
    e.target.value = '';
  };

  const handleManualAddTeacher = () => {
    if (!manualNama.trim()) {
      alert('Nama guru tidak boleh kosong');
      return;
    }
    const finalNomor = manualNomor.trim() || String(teachersList.length + 1);

    if (teachersList.some(t => t.nomor === finalNomor)) {
      alert(`Nomor "${finalNomor}" sudah terdaftar untuk guru lain.`);
      return;
    }

    onAddTeacher({
      nomor: finalNomor,
      nama: manualNama.trim(),
      foto: ''
    });

    setManualNama('');
    setManualNomor('');
  };

  // Suggestions (combined from registered master list + history logs)
  const uniqueTeacherNames = useMemo(() => {
    const fromMaster = teachersList.map(t => t.nama);
    const fromLogs = teachersAttendance.map(t => t.nama);
    return Array.from(new Set([...fromMaster, ...fromLogs])).sort();
  }, [teachersList, teachersAttendance]);

  const handleTeacherExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const arrayBuffer = evt.target?.result as ArrayBuffer;
        const dataBytes = new Uint8Array(arrayBuffer);
        const wb = XLSX.read(dataBytes, { type: 'array' });
        let importedTeachers: TeacherMaster[] = [];

        wb.SheetNames.forEach((sheetName) => {
          const ws = wb.Sheets[sheetName];
          const data = XLSX.utils.sheet_to_json(ws) as any[];
          
          if (data.length > 0) {
            const parsed = data.map((item: any, idx: number) => {
              const keys = Object.keys(item);
              
              const findKey = (searchTerms: string[]) => {
                return keys.find(k => {
                  const cleaned = k.toLowerCase().replace(/[^a-z0-9]/gi, '').trim();
                  return searchTerms.some(term => cleaned === term || cleaned.replace(/[\s\r\n\t_]/g, '') === term);
                });
              };

              const keyNama = findKey(['nama', 'name', 'namaguru']);
              const keyNomor = findKey(['no', 'nomor', 'number', 'id', 'nomer']);
              
              const namaVal = keyNama ? String(item[keyNama]).trim() : '';
              const nomorVal = keyNomor ? String(item[keyNomor]).trim() : String(idx + 1);

              return {
                nomor: nomorVal,
                nama: namaVal,
                foto: ''
              };
            }).filter(t => t.nama.length > 0);

            importedTeachers = [...importedTeachers, ...parsed];
          }
        });

        if (importedTeachers.length > 0) {
          // Normalize and prevent duplicates
          const seen = new Set<string>();
          const finalImport: TeacherMaster[] = [];
          importedTeachers.forEach((t, i) => {
            const numKey = t.nomor || String(i + 1);
            if (!seen.has(numKey)) {
              seen.add(numKey);
              finalImport.push({
                nomor: numKey,
                nama: t.nama,
                foto: ''
              });
            } else {
              // fallback to generate sequential if collision
              let nextNum = i + 1;
              while (seen.has(String(nextNum))) {
                nextNum++;
              }
              seen.add(String(nextNum));
              finalImport.push({
                nomor: String(nextNum),
                nama: t.nama,
                foto: ''
              });
            }
          });

          onImportTeachersList(finalImport);
          alert(`Berhasil mengimpor ${finalImport.length} guru dari file Excel dangan format kolom Nomor dan Nama.`);
        } else {
          alert('Tidak ada data guru yang valid. Pastikan file Excel Anda setidaknya memiliki kolom bernama "Nama".');
        }
      } catch (err) {
        alert('Gagal membaca file Excel. Pastikan format file benar, berisi kolom "No" dan "Nama".');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!nama.trim()) {
      alert('Nama guru tidak boleh kosong.');
      return;
    }
    const finalKeterangan = keterangan === 'Lainnya' ? (keteranganCustom || 'Lainnya') : keterangan;
    
    if (editingItem) {
      onUpdateAttendance(editingItem.id, {
        nama: nama.trim(),
        hariTanggal,
        jamMasuk,
        keterangan: finalKeterangan
      });
      setEditingItem(null);
      alert('Berhasil mengupdate data absensi guru.');
    } else {
      onAddAttendance({
        nama: nama.trim(),
        hariTanggal,
        jamMasuk,
        keterangan: finalKeterangan
      });
      alert('Berhasil mencatat absensi guru.');
    }

    // Reset Form (except date to make repetitive logging faster)
    setNama('');
    setJamMasuk('');
    setKeterangan('Hadir');
    setKeteranganCustom('');
  };

  const handleEditClick = (item: TeacherAttendance) => {
    setEditingItem(item);
    setNama(item.nama);
    setHariTanggal(item.hariTanggal);
    setJamMasuk(item.jamMasuk);
    if (['Hadir', 'Sakit', 'Izin', 'Tugas Luar', 'Terlambat'].includes(item.keterangan)) {
      setKeterangan(item.keterangan);
      setKeteranganCustom('');
    } else {
      setKeterangan('Lainnya');
      setKeteranganCustom(item.keterangan);
    }
  };

  const handleCancelEdit = () => {
    setEditingItem(null);
    setNama('');
    setHariTanggal(new Date().toISOString().split('T')[0]);
    setJamMasuk('');
    setKeterangan('Hadir');
    setKeteranganCustom('');
  };

  // Filtered list
  const filteredData = useMemo(() => {
    return teachersAttendance.filter(item => {
      const matchSearch = searchTerm ? item.nama.toLowerCase().includes(searchTerm.toLowerCase()) : true;
      const matchMonth = filterMonth ? item.hariTanggal.startsWith(filterMonth) : true;
      const matchDate = filterDate ? item.hariTanggal === filterDate : true;
      return matchSearch && matchMonth && matchDate;
    }).sort((a, b) => b.hariTanggal.localeCompare(a.hariTanggal) || b.jamMasuk.localeCompare(a.jamMasuk));
  }, [teachersAttendance, searchTerm, filterMonth, filterDate]);

  // Statistics
  const stats = useMemo(() => {
    const today = new Date().toISOString().split('T')[0];
    const todayLogs = teachersAttendance.filter(item => item.hariTanggal === today);
    const totalToday = todayLogs.length;
    const hadirToday = todayLogs.filter(item => item.keterangan === 'Hadir' || item.keterangan === 'Terlambat' || item.keterangan === 'Tugas Luar').length;
    return {
      totalLogs: teachersAttendance.length,
      todayTotal: totalToday,
      todayHadir: hadirToday,
    };
  }, [teachersAttendance]);

  return (
    <div className="p-4 sm:p-8 space-y-6 sm:space-y-8">
      {/* Overview cards */}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm flex items-center gap-4">
          <div className="p-3 bg-blue-50 text-blue-600 rounded-xl">
            <GraduationCap className="w-6 h-6" />
          </div>
          <div>
            <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">Total Absensi Guru</p>
            <h3 className="text-2xl font-bold text-slate-800 mt-1">{stats.totalLogs} <span className="text-sm font-medium text-slate-500 font-bold uppercase tracking-wider">Log</span></h3>
          </div>
        </div>

        <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm flex items-center gap-4">
          <div className="p-3 bg-green-50 text-green-600 rounded-xl">
            <CheckCircle2 className="w-6 h-6" />
          </div>
          <div>
            <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">Guru Aktif Hari Ini</p>
            <h3 className="text-2xl font-bold text-slate-800 mt-1">{stats.todayHadir} / {stats.todayTotal} <span className="text-sm font-medium text-slate-500 font-bold uppercase tracking-wider">Guru</span></h3>
          </div>
        </div>

        <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm flex items-center gap-4">
          <div className="p-3 bg-amber-50 text-amber-600 rounded-xl">
            <Clock className="w-6 h-6" />
          </div>
          <div>
            <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">Log Hari Ini</p>
            <h3 className="text-2xl font-bold text-slate-800 mt-1">{stats.todayTotal} <span className="text-sm font-medium text-slate-500 font-bold uppercase tracking-wider">Record</span></h3>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 sm:gap-8">
        {/* Form panel & Master Database */}
        <div className="lg:col-span-4 flex flex-col gap-6 self-start">
          {/* Add/Edit Form Card */}
          <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm space-y-6">
            <div>
              <h2 className="text-sm font-bold text-slate-800 uppercase tracking-wider flex items-center gap-2">
                <Plus className="w-5 h-5 text-blue-600" />
                <span>{editingItem ? 'Edit Kehadiran Guru' : 'Tambah Kehadiran Guru'}</span>
              </h2>
              <p className="text-xs text-slate-500 mt-1">Hari & Tanggal disiapkan otomatis, silakan lengkapi data lainnya.</p>
            </div>

            <form onSubmit={handleSubmit} className="space-y-4">
              {/* Nama Input */}
              <div className="space-y-1.5">
                <div className="flex items-center justify-between">
                  <label className="text-xs font-bold text-slate-500 uppercase">Nama Guru</label>
                  <button
                    type="button"
                    onClick={() => setIsMasterOpen(prev => !prev)}
                    className="text-[11px] text-blue-600 hover:text-blue-800 font-bold focus:outline-none flex items-center gap-1 transition-all"
                  >
                    <span>{isMasterOpen ? 'Sembunyikan Master' : 'Kelola Database'}</span>
                  </button>
                </div>
                <select 
                  className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 outline-none focus:ring-2 focus:ring-blue-500/20 text-slate-800 font-medium text-sm transition-all cursor-pointer"
                  value={nama}
                  onChange={e => setNama(e.target.value)}
                  required
                >
                  <option value="">Pilih Guru</option>
                  {uniqueTeacherNames.map(name => (
                    <option key={name} value={name}>{name}</option>
                  ))}
                </select>
              </div>

              {/* Jam Masuk */}
              <div className="space-y-1.5">
                <label className="text-xs font-bold text-slate-500 uppercase">Jam Masuk</label>
                <div className="flex gap-2">
                  <input 
                    type="text"
                    placeholder="Contoh: 07:15"
                    className="flex-grow bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 outline-none focus:ring-2 focus:ring-blue-500/20 text-slate-800 font-medium text-sm transition-all"
                    value={jamMasuk}
                    onChange={e => setJamMasuk(e.target.value)}
                    required
                  />
                  <button
                    type="button"
                    onClick={() => {
                      const now = new Date();
                      setJamMasuk(`${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`);
                    }}
                    className="bg-slate-100 hover:bg-slate-200 text-slate-600 px-3 py-2.5 rounded-xl text-xs font-bold transition-all shrink-0 flex items-center justify-center gap-1 border border-slate-200"
                    title="Gunakan jam sekarang"
                  >
                    <Clock className="w-4 h-4 text-slate-500" />
                    <span>Sekarang</span>
                  </button>
                </div>
              </div>

              {/* Tanggal */}
              <div className="space-y-1.5">
                <label className="text-xs font-bold text-slate-500 uppercase">Hari & Tanggal</label>
                <input 
                  type="date"
                  className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 outline-none focus:ring-2 focus:ring-blue-500/20 text-slate-800 font-medium text-sm transition-all cursor-pointer"
                  value={hariTanggal}
                  onChange={e => setHariTanggal(e.target.value)}
                  required
                />
              </div>

              {/* Keterangan */}
              <div className="space-y-1.5">
                <label className="text-xs font-bold text-slate-500 uppercase">Keterangan</label>
                <div className="grid grid-cols-2 gap-2">
                  {['Hadir', 'Sakit', 'Izin', 'Tugas Luar', 'Terlambat', 'Lainnya'].map(status => (
                    <button
                      key={status}
                      type="button"
                      onClick={() => { setKeterangan(status); if(status !== 'Lainnya') setKeteranganCustom(''); }}
                      className={`px-3 py-2 text-xs font-semibold rounded-lg border text-center transition-all ${
                        keterangan === status 
                          ? 'bg-blue-600 border-blue-600 text-white shadow-md shadow-blue-600/10' 
                          : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                      }`}
                    >
                      {status}
                    </button>
                  ))}
                </div>

                {keterangan === 'Lainnya' && (
                  <input 
                    type="text"
                    placeholder="Isi keterangan / alasan lainnya"
                    className="w-full mt-2 bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 outline-none focus:ring-2 focus:ring-blue-500/20 text-slate-800 font-medium text-sm transition-all"
                    value={keteranganCustom}
                    onChange={e => setKeteranganCustom(e.target.value)}
                    required
                  />
                )}
              </div>

              {/* Action buttons */}
              <div className="flex gap-2 pt-2">
                {editingItem && (
                  <button
                    type="button"
                    onClick={handleCancelEdit}
                    className="flex-1 border border-slate-200 hover:bg-slate-50 text-slate-600 py-2.5 rounded-xl text-sm font-semibold transition-all"
                  >
                    Batal
                  </button>
                )}
                <button
                  type="submit"
                  className="flex-1 bg-blue-600 hover:bg-blue-700 text-white py-2.5 rounded-xl text-sm font-semibold shadow-md shadow-blue-600/15 transition-all"
                >
                  {editingItem ? 'Simpan' : 'Simpan Kehadiran'}
                </button>
              </div>
            </form>
          </div>

          {/* Database Master Guru Card */}
          <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm space-y-4">
            <button
              type="button"
              onClick={() => setIsMasterOpen(prev => !prev)}
              className="w-full flex items-center justify-between text-left focus:outline-none"
            >
              <h2 className="text-sm font-bold text-slate-800 uppercase tracking-wider flex items-center gap-2">
                <Users className="w-4 h-4 text-blue-600" />
                <span>Database Guru (Master)</span>
              </h2>
              <div className="flex items-center gap-2">
                <span className="text-[10px] bg-blue-50 text-blue-600 px-2 py-0.5 rounded-full font-bold">
                  {teachersList.length} Guru
                </span>
                <span className="text-xs text-slate-400 font-bold">
                  {isMasterOpen ? '▲ Sembunyikan' : '▼ Kelola'}
                </span>
              </div>
            </button>
            
            {isMasterOpen && (
              <div className="space-y-4 pt-2 border-t border-slate-100">
                <p className="text-xs text-slate-500 leading-relaxed">
                  Upload file Excel dengan kolom <span className="font-bold underline text-slate-700">nomor (no)</span> dan <span className="font-bold underline text-slate-700">nama</span> saja. Guru-guru yang diimpor dapat ditambahkan fotonya langsung melalui tombol kamera di tabel.
                </p>

                {/* Upload Button */}
                <div className="flex flex-wrap gap-2">
                  <label className="flex-grow flex items-center justify-center gap-2 bg-slate-50 hover:bg-slate-100 border border-slate-200 text-slate-600 px-4 py-2.5 rounded-xl text-xs font-semibold cursor-pointer transition-all">
                    <FileUp className="w-4 h-4 text-blue-600" />
                    <span>Upload Excel Guru</span>
                    <input 
                      type="file" 
                      className="hidden" 
                      accept=".xlsx, .xls" 
                      onChange={handleTeacherExcelUpload} 
                    />
                  </label>

                  <button
                    type="button"
                    onClick={() => {
                      if (confirm('Yakin ingin memuat ulang daftar 23 Guru Default sekolah? Daftar guru saat ini akan digantikan.')) {
                        onImportTeachersList(DEFAULT_TEACHERS);
                        alert('Berhasil memuat ulang 23 guru default sekolah!');
                      }
                    }}
                    className="flex items-center justify-center gap-2 bg-amber-50 hover:bg-amber-100 border border-amber-200 text-amber-700 px-4 py-2.5 rounded-xl text-xs font-semibold transition-all"
                    title="Muat Ulang 23 Guru Default"
                  >
                    <RotateCcw className="w-4 h-4 text-amber-600" />
                    <span>Daftar Asli (23 Guru)</span>
                  </button>

                  {teachersList.length > 0 && (
                    <button
                      type="button"
                      onClick={() => { if(confirm('Yakin ingin menghapus semua database guru master?')) onClearTeachersList(); }}
                      className="bg-rose-50 hover:bg-rose-100 text-rose-600 border border-rose-100 px-3.5 rounded-xl transition-all"
                      title="Hapus Semua Guru"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  )}
                </div>

                {/* Mini manual add with No and Nama */}
                <div className="space-y-2 border-t border-slate-50 pt-3">
                  <span className="text-[10px] text-slate-400 block font-bold uppercase">Tambah Guru Manual</span>
                  <div className="flex gap-2">
                    <input 
                      type="text"
                      placeholder="No"
                      className="w-16 bg-slate-50 border border-slate-200 rounded-lg px-2.5 py-1.5 outline-none text-xs focus:ring-2 focus:ring-blue-500/20 text-slate-800 font-mono"
                      value={manualNomor}
                      onChange={(e) => setManualNomor(e.target.value)}
                    />
                    <input 
                      type="text"
                      placeholder="Nama Lengkap Guru"
                      className="flex-1 bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 outline-none text-xs focus:ring-2 focus:ring-blue-500/20 text-slate-800"
                      value={manualNama}
                      onChange={(e) => setManualNama(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') {
                          handleManualAddTeacher();
                        }
                      }}
                    />
                    <button
                      type="button"
                      onClick={handleManualAddTeacher}
                      className="bg-blue-600 hover:bg-blue-700 text-white px-3.5 py-1.5 rounded-lg text-xs font-semibold transition-all shrink-0"
                    >
                      Tambah
                    </button>
                  </div>
                </div>

                {/* Elegant Table: Nomor, Nama, Foto */}
                {teachersList.length > 0 ? (
                  <div className="border border-slate-100 rounded-xl overflow-hidden">
                    <div className="max-h-72 overflow-y-auto">
                      <table className="w-full text-left text-slate-600 border-collapse">
                        <thead className="bg-slate-50 text-[10px] font-bold text-slate-400 uppercase tracking-wider sticky top-0 bg-slate-50 border-b border-slate-100 z-10">
                          <tr>
                            <th className="py-2 px-3 text-center w-12">No</th>
                            <th className="py-2 px-3">Nama</th>
                            <th className="py-2 px-3 text-center w-16">Foto</th>
                            <th className="py-2 px-2 text-center w-8"></th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100 text-xs">
                          {teachersList.map((item) => (
                            <tr key={item.nomor} className="hover:bg-slate-50/50 transition-colors">
                              <td className="py-2 px-3 text-center font-mono font-bold text-slate-500 bg-slate-50/20 w-12">
                                {item.nomor}
                              </td>
                              <td className="py-2 px-3 text-slate-700 font-bold truncate max-w-[130px]" title={item.nama}>
                                {item.nama}
                              </td>
                              <td className="py-2 px-3 text-center w-16">
                                <div className="relative inline-block group">
                                  <input 
                                    type="file" 
                                    id={`master-photo-${item.nomor}`} 
                                    className="hidden" 
                                    accept="image/*" 
                                    onChange={(e) => handlePhotoUpload(item.nomor, e)} 
                                  />
                                  <button
                                    type="button"
                                    onClick={() => document.getElementById(`master-photo-${item.nomor}`)?.click()}
                                    className="w-8 h-8 rounded-full border border-slate-200 bg-slate-50 hover:bg-slate-100 flex items-center justify-center overflow-hidden transition-all relative group"
                                    title="Klik untuk upload/ganti foto"
                                  >
                                    {item.foto ? (
                                      <img 
                                        src={item.foto} 
                                        alt={item.nama} 
                                        className="w-full h-full object-cover" 
                                        referrerPolicy="no-referrer" 
                                      />
                                    ) : (
                                      <Camera className="w-3.5 h-3.5 text-slate-400 group-hover:text-blue-500 group-hover:scale-110 transition-transform" />
                                    )}
                                    <div className="absolute inset-0 bg-black/40 text-white flex items-center justify-center opacity-0 group-hover:opacity-100 transition-opacity">
                                      <Camera className="w-3 h-3" />
                                    </div>
                                  </button>
                                </div>
                              </td>
                              <td className="py-2 px-2 text-center w-8">
                                <button
                                  type="button"
                                  onClick={() => onDeleteTeacher(item.nomor)}
                                  className="text-slate-400 hover:text-rose-600 p-1 rounded transition-colors"
                                  title="Hapus dari master"
                                >
                                  <X className="w-3.5 h-3.5" />
                                </button>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                ) : (
                  <div className="border border-dashed border-slate-200 rounded-xl p-6 text-center">
                    <Users className="w-8 h-8 text-slate-300 mx-auto mb-1.5 animate-pulse" />
                    <p className="text-[11px] text-slate-400 font-bold uppercase tracking-wider">Belum Ada Database Guru</p>
                    <p className="text-[10px] text-slate-400 mt-1">Impor Excel (nama) atau tambah secara manual di atas.</p>
                  </div>
                )}

                {/* Save & Finish Button right below manual add menu & teacher table */}
                <div className="border-t border-slate-100 pt-3 mt-1">
                  <button
                    type="button"
                    onClick={() => {
                      alert('Data guru berhasil disimpan!');
                      setIsMasterOpen(false);
                    }}
                    className="w-full bg-blue-600 hover:bg-blue-700 text-white py-2.5 rounded-xl text-xs font-bold transition-all flex items-center justify-center gap-1.5 shadow-md shadow-blue-600/10 focus:ring-2 focus:ring-blue-500/20"
                  >
                    <CheckCircle2 className="w-4 h-4" />
                    <span>Simpan & Selesai</span>
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>

        {/* List Table panel */}
        <div className="lg:col-span-8 bg-white p-6 rounded-2xl border border-slate-100 shadow-sm space-y-6">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <div>
              <h2 className="text-lg font-bold text-slate-800">Daftar Kehadiran Guru</h2>
              <p className="text-xs text-slate-500 mt-1">Daftar rekapitulasi kehadiran guru SMKN 1 Bunyu.</p>
            </div>

            {/* Download Buttons */}
            <div className="flex flex-wrap items-center gap-2">
              <button
                onClick={() => handleExportExcel(filteredData)}
                disabled={filteredData.length === 0}
                className="bg-[#107C41] hover:bg-[#0E6C38] text-white px-3.5 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-all disabled:opacity-50"
                title="Download Excel"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Excel</span>
              </button>

              <button
                onClick={() => generateTeacherWordReport("REKAP ABSENSI GURU SMKN 1 BUNYU", filteredData)}
                disabled={filteredData.length === 0}
                className="bg-[#2B579A] hover:bg-[#244A83] text-white px-3.5 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-all disabled:opacity-50"
                title="Download Word (DOCX)"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Word</span>
              </button>

              <button
                onClick={() => handleExportPDF(filteredData)}
                disabled={filteredData.length === 0}
                className="bg-[#CC0000] hover:bg-[#B30000] text-white px-3.5 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-all disabled:opacity-50"
                title="Download PDF"
              >
                <Download className="w-3.5 h-3.5" />
                <span>PDF</span>
              </button>
            </div>
          </div>

          {/* Filters & Search */}
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
            <div className="flex items-center gap-2 bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 w-full focus-within:ring-2 focus-within:ring-blue-500/20 transition-all">
              <Search className="w-4 h-4 text-slate-400" />
              <input 
                type="text" 
                placeholder="Cari nama guru..." 
                className="flex-1 bg-transparent outline-none font-medium text-slate-700 text-xs"
                value={searchTerm}
                onChange={e => setSearchTerm(e.target.value)}
              />
            </div>

            <div className="relative">
              <input 
                type="month"
                className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 outline-none text-slate-600 font-medium text-xs cursor-pointer"
                value={filterMonth}
                onChange={e => { setFilterMonth(e.target.value); setFilterDate(''); }}
              />
            </div>

            <div className="relative">
              <input 
                type="date"
                className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 outline-none text-slate-600 font-medium text-xs cursor-pointer"
                value={filterDate}
                onChange={e => { setFilterDate(e.target.value); setFilterMonth(''); }}
              />
              { (filterMonth || filterDate) && (
                <button 
                  onClick={() => { setFilterMonth(''); setFilterDate(''); }}
                  className="absolute right-2.5 top-1/2 -translate-y-1/2 text-[10px] bg-slate-200 hover:bg-slate-300 px-1.5 py-0.5 rounded text-slate-600 font-bold"
                >
                  Clear
                </button>
              )}
            </div>
          </div>

          {/* Table list */}
          <div className="overflow-x-auto rounded-xl border border-slate-100">
            <table className="w-full table-auto text-left text-sm text-slate-600">
              <thead className="bg-slate-50 text-slate-500 font-bold text-xs uppercase tracking-wider">
                <tr>
                  <th className="px-5 py-3">No</th>
                  <th className="px-5 py-3">Hari & Tanggal</th>
                  <th className="px-5 py-3">Nama Guru</th>
                  <th className="px-5 py-3">Jam Masuk</th>
                  <th className="px-5 py-3">Keterangan</th>
                  <th className="px-5 py-3 text-right">Aksi</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filteredData.length > 0 ? (
                  filteredData.map((item, idx) => {
                    const formattedDate = new Intl.DateTimeFormat('id-ID', { dateStyle: 'full' }).format(new Date(item.hariTanggal));
                    
                    const getKeteranganBadge = (ket: string) => {
                      if (ket === 'Hadir') return 'bg-green-50 text-green-700 border-green-200';
                      if (ket === 'Sakit') return 'bg-blue-50 text-blue-700 border-blue-200';
                      if (ket === 'Izin') return 'bg-yellow-50 text-yellow-700 border-yellow-200';
                      if (ket === 'Tugas Luar') return 'bg-indigo-50 text-indigo-700 border-indigo-200';
                      if (ket === 'Terlambat') return 'bg-orange-50 text-orange-700 border-orange-200';
                      return 'bg-slate-50 text-slate-700 border-slate-200';
                    };

                    return (
                      <tr key={item.id} className="hover:bg-slate-50/50 transition-all group">
                        <td className="px-5 py-4 font-bold text-slate-400 text-xs">{idx + 1}</td>
                        <td className="px-5 py-4 font-semibold text-slate-700 text-xs">{formattedDate}</td>
                        <td className="px-5 py-4 font-bold text-slate-800 text-sm whitespace-nowrap">{item.nama}</td>
                        <td className="px-5 py-4 font-mono text-slate-600 text-xs whitespace-nowrap">
                          <span className="flex items-center gap-1.5">
                            <Clock className="w-3.5 h-3.5 text-slate-400" />
                            {item.jamMasuk}
                          </span>
                        </td>
                        <td className="px-5 py-4">
                          <span className={`px-2.5 py-1 text-[11px] font-bold rounded-lg border uppercase tracking-wider inline-block ${getKeteranganBadge(item.keterangan)}`}>
                            {item.keterangan}
                          </span>
                        </td>
                        <td className="px-5 py-4 text-right">
                          <div className="flex items-center justify-end gap-1.5 opacity-0 group-hover:opacity-100 transition-opacity">
                            <button
                              onClick={() => handleEditClick(item)}
                              className="p-1.5 text-blue-500 hover:bg-blue-50 rounded-lg transition-colors"
                              title="Edit Rekor"
                            >
                              <Edit className="w-3.5 h-3.5" />
                            </button>
                            <button
                              onClick={() => { if(confirm('Hapus record absensi ini?')) onDeleteAttendance(item.id); }}
                              className="p-1.5 text-rose-500 hover:bg-rose-50 rounded-lg transition-colors"
                              title="Hapus Rekor"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        </td>
                      </tr>
                    );
                  })
                ) : (
                  <tr>
                    <td colSpan={6} className="px-5 py-12 text-center text-slate-400">
                      <GraduationCap className="w-12 h-12 mx-auto mb-3 opacity-20" />
                      <p className="font-bold text-xs uppercase tracking-widest text-slate-400">Belum Ada Data Absensi Guru</p>
                      <p className="text-[11px] text-slate-400 mt-1">Silakan tambahkan data absensi menggunakan form di sebelah kiri.</p>
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
};

// --- Main App ---

export default function App() {
  const [currentPage, setCurrentPage] = useState<Page>('landing');
  const [authType, setAuthType] = useState<'login' | 'register'>('login');
  const [user, setUser] = useState<User | null>(() => {
    const saved = localStorage.getItem('user');
    return saved ? JSON.parse(saved) : null;
  });
  
  const [students, setStudents] = useState<Student[]>([]);
  const [attendance, setAttendance] = useState<AttendanceItem[]>([]);
  const [attendanceHistory, setAttendanceHistory] = useState<Record<string, AttendanceItem[]>>({});
  const [teachersAttendance, setTeachersAttendance] = useState<TeacherAttendance[]>([]);
  const [teachersList, setTeachersList] = useState<TeacherMaster[]>([]);
  const [currentDate, setCurrentDate] = useState(new Date().toISOString().split('T')[0]);
  const [stats, setStats] = useState({ totalStudents: 0, totalTeachers: 0, attendanceToday: 0, historyCount: 0 });
  const [globalPreviewImage, setGlobalPreviewImage] = useState<string | null>(null);
  const [logo, setLogo] = useState<string | null>(() => localStorage.getItem('app_logo'));
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [isDeleteModalOpen, setIsDeleteModalOpen] = useState(false);
  const [deferredPrompt, setDeferredPrompt] = useState<any>(null);

  useEffect(() => {
    const handleBeforeInstallPrompt = (e: any) => {
      e.preventDefault();
      setDeferredPrompt(e);
    };

    window.addEventListener('beforeinstallprompt', handleBeforeInstallPrompt);

    return () => {
      window.removeEventListener('beforeinstallprompt', handleBeforeInstallPrompt);
    };
  }, []);

  const handleInstallApp = async () => {
    if (!deferredPrompt) return;
    deferredPrompt.prompt();
    const { outcome } = await deferredPrompt.userChoice;
    if (outcome === 'accepted') {
      setDeferredPrompt(null);
    }
  };

  useEffect(() => {
    if (user) {
      if (currentPage === 'landing' || currentPage === 'login' || currentPage === 'register') {
        setCurrentPage('dashboard');
      }
      fetchStudents();
      fetchAttendance(currentDate);
      fetchAttendanceHistory();
      fetchTeachersAttendance();
      fetchTeachersList();
    }
  }, [user]);

  useEffect(() => {
    if (currentPage === 'attendance' && user) {
      fetchAttendance(currentDate);
    }
    if (currentPage === 'recap' && user) {
      fetchAttendanceHistory();
    }
    if (currentPage === 'teachers' && user) {
      fetchTeachersAttendance();
      fetchTeachersList();
    }
    if (currentPage === 'dashboard' && user) {
      updateStats(students, attendanceHistory);
    }
  }, [currentPage, currentDate]);

  const updateStats = (currentStudents: Student[], history: Record<string, AttendanceItem[]>) => {
    const today = new Date().toISOString().split("T")[0];
    const todayAttendance = history[today] || [];
    
    const validHistoryCount = Object.values(history).filter((monthData: any) => 
      Array.isArray(monthData) && monthData.some((item: any) => item.status)
    ).length;

    const savedTeachersRaw = localStorage.getItem('app_teachers_list');
    const teachersCount = savedTeachersRaw ? JSON.parse(savedTeachersRaw).length : DEFAULT_TEACHERS.length;

    setStats({
      totalStudents: currentStudents.length,
      totalTeachers: teachersCount,
      attendanceToday: todayAttendance.filter((a: any) => a.status).length,
      historyCount: validHistoryCount
    });
  };

  const fetchStudents = () => {
    const saved = localStorage.getItem('app_students');
    if (saved) {
      const data = JSON.parse(saved);
      setStudents(data);
      updateStats(data, attendanceHistory);
    } else {
      setStudents(DEFAULT_STUDENTS);
      localStorage.setItem('app_students', JSON.stringify(DEFAULT_STUDENTS));
      updateStats(DEFAULT_STUDENTS, attendanceHistory);
    }
  };

  const fetchStats = () => {
    updateStats(students, attendanceHistory);
  };

  const fetchAttendance = (date: string) => {
    const saved = localStorage.getItem('app_attendance');
    if (saved) {
      const history = JSON.parse(saved);
      setAttendance(history[date] || []);
    } else {
      const defaultHistory = generateDefaultAttendance();
      setAttendance(defaultHistory[date] || []);
      localStorage.setItem('app_attendance', JSON.stringify(defaultHistory));
    }
  };

  const fetchAttendanceHistory = () => {
    const saved = localStorage.getItem('app_attendance');
    if (saved) {
      const history = JSON.parse(saved);
      setAttendanceHistory(history);
      updateStats(students, history);
    } else {
      const defaultHistory = generateDefaultAttendance();
      setAttendanceHistory(defaultHistory);
      localStorage.setItem('app_attendance', JSON.stringify(defaultHistory));
      updateStats(students, defaultHistory);
    }
  };

  const fetchTeachersAttendance = () => {
    const saved = localStorage.getItem('app_teachers_attendance');
    if (saved) {
      setTeachersAttendance(JSON.parse(saved));
    } else {
      const defaultList = generateDefaultTeachersAttendance();
      setTeachersAttendance(defaultList);
      localStorage.setItem('app_teachers_attendance', JSON.stringify(defaultList));
    }
  };

  const handleAddTeacherAttendance = (item: Omit<TeacherAttendance, 'id'>) => {
    const newVal: TeacherAttendance = {
      ...item,
      id: crypto.randomUUID ? crypto.randomUUID() : String(Date.now() + Math.random())
    };
    setTeachersAttendance(prev => {
      const next = [newVal, ...prev];
      localStorage.setItem('app_teachers_attendance', JSON.stringify(next));
      return next;
    });
  };

  const handleDeleteTeacherAttendance = (id: string) => {
    setTeachersAttendance(prev => {
      const next = prev.filter(t => t.id !== id);
      localStorage.setItem('app_teachers_attendance', JSON.stringify(next));
      return next;
    });
  };

  const handleUpdateTeacherAttendance = (id: string, updated: Omit<TeacherAttendance, 'id'>) => {
    setTeachersAttendance(prev => {
      const next = prev.map(t => t.id === id ? { ...updated, id } : t);
      localStorage.setItem('app_teachers_attendance', JSON.stringify(next));
      return next;
    });
  };

  const fetchTeachersList = () => {
    const saved = localStorage.getItem('app_teachers_list');
    if (saved) {
      setTeachersList(JSON.parse(saved));
    } else {
      setTeachersList(DEFAULT_TEACHERS);
      localStorage.setItem('app_teachers_list', JSON.stringify(DEFAULT_TEACHERS));
    }
  };

  const handleImportTeachersList = (teachers: TeacherMaster[]) => {
    setTeachersList(teachers);
    localStorage.setItem('app_teachers_list', JSON.stringify(teachers));
  };

  const handleAddTeacher = (teacher: TeacherMaster) => {
    setTeachersList(prev => {
      const filtered = prev.filter(t => t.nomor !== teacher.nomor);
      const next = [...filtered, teacher].sort((a, b) => a.nama.localeCompare(b.nama));
      localStorage.setItem('app_teachers_list', JSON.stringify(next));
      return next;
    });
  };

  const handleDeleteTeacher = (nomor: string) => {
    setTeachersList(prev => {
      const next = prev.filter(t => t.nomor !== nomor);
      localStorage.setItem('app_teachers_list', JSON.stringify(next));
      return next;
    });
  };

  const handleClearTeachersList = () => {
    setTeachersList([]);
    localStorage.removeItem('app_teachers_list');
  };

  const handleUpdateTeacherPhoto = (nomor: string, base64: string) => {
    setTeachersList(prev => {
      const next = prev.map(t => t.nomor === nomor ? { ...t, foto: base64 } : t);
      localStorage.setItem('app_teachers_list', JSON.stringify(next));
      return next;
    });
  };

  const handleAuthSuccess = (userData: User) => {
    setUser(userData);
    localStorage.setItem('user', JSON.stringify(userData));
  };

  const handleLogout = () => {
    setUser(null);
    localStorage.removeItem('user');
    setCurrentPage('landing');
  };

  const handleImport = (newStudents: Student[]) => {
    localStorage.setItem('app_students', JSON.stringify(newStudents));
    setStudents(newStudents);
    fetchStats();
  };

  const handleDeleteStudent = (nomor: string) => {
    const next = students.filter(s => s.nomor !== nomor);
    localStorage.setItem('app_students', JSON.stringify(next));
    setStudents(next);
    fetchStats();
  };

  const handleDeleteAll = () => {
    localStorage.setItem('app_students', JSON.stringify([]));
    setStudents([]);
    fetchStats();
  };

  const handleUpdateStudent = (oldNomor: string, updated: Student) => {
    const next = students.map(s => s.nomor === oldNomor ? updated : s);
    localStorage.setItem('app_students', JSON.stringify(next));
    setStudents(next);
  };

  const handleStatusChange = (nama: string, status: AttendanceStatus) => {
    setAttendance(prev => {
      const existing = prev.find(a => a.nama === nama);
      if (existing) {
        return prev.map(a => a.nama === nama ? { ...a, status } : a);
      } else {
        return [...prev, { nama, status }];
      }
    });
  };

  const handleEvidenceChange = (nama: string, notesVal: string) => {
    setAttendance(prev => {
      const existing = prev.find(a => a.nama === nama);
      if (existing) {
        return prev.map(a => a.nama === nama ? { ...a, evidence: notesVal, notes: notesVal } : a);
      } else {
        return [...prev, { nama, status: 'T', evidence: notesVal, notes: notesVal }];
      }
    });
  };

  const handleSaveAttendance = () => {
    const history = JSON.parse(localStorage.getItem('app_attendance') || '{}');
    history[currentDate] = attendance;
    localStorage.setItem('app_attendance', JSON.stringify(history));
    
    alert('Data absensi hari ini berhasil disimpan!');
    
    setAttendanceHistory(history);
    updateStats(students, history);
  };

  const handleUpdateRecord = (date: string, nama: string, status: AttendanceStatus) => {
    setAttendanceHistory(prev => {
      const records = prev[date] || [];
      const updatedRecords = records.map(r => r.nama === nama ? { ...r, status } : r);
      if (!records.find(r => r.nama === nama)) {
        updatedRecords.push({ nama, status });
      }
      
      const next = { ...prev, [date]: updatedRecords };
      localStorage.setItem('app_attendance', JSON.stringify(next));
      updateStats(students, next);
      return next;
    });
  };

  const handleDeleteRecord = (date: string, nama: string) => {
    setAttendanceHistory(prev => {
      const records = prev[date] || [];
      const updatedRecords = records.filter(r => r.nama !== nama);
      const next = { ...prev, [date]: updatedRecords };
      localStorage.setItem('app_attendance', JSON.stringify(next));
      updateStats(students, next);
      return next;
    });
  };

  const handleDeleteMonthAttendance = (month: string, nama: string) => {
    setAttendanceHistory(prev => {
      const next = { ...prev };
      Object.keys(next).forEach(date => {
        if (date.startsWith(month)) {
          next[date] = next[date].filter(r => r.nama !== nama);
        }
      });
      localStorage.setItem('app_attendance', JSON.stringify(next));
      updateStats(students, next);
      return next;
    });
  };

  const handleDeleteAccount = () => {
    setIsDeleteModalOpen(true);
  };

  const handleConfirmDeleteAccount = (password: string, clearAllData: boolean) => {
    if (!user) return false;
    
    // Verifikasi password
    const users = JSON.parse(localStorage.getItem('app_users') || '[]');
    const foundUser = users.find((u: any) => u.username === user.username);
    
    if (!foundUser || foundUser.password !== password) {
      alert('Password konfirmasi tidak benar!');
      return false;
    }
    
    // Hapus user dari database lokal
    const nextUsers = users.filter((u: any) => u.username !== user.username);
    localStorage.setItem('app_users', JSON.stringify(nextUsers));
    
    // Jika user memilih untuk menghapus seluruh data sekolah juga
    if (clearAllData) {
      localStorage.removeItem('app_students');
      localStorage.removeItem('app_attendance');
      localStorage.removeItem('app_teachers_attendance');
      localStorage.removeItem('app_teachers_list');
      localStorage.removeItem('app_logo');
      
      // Reset state instan
      setStudents([]);
      setAttendance([]);
      setAttendanceHistory({});
      setTeachersAttendance([]);
      setTeachersList([]);
      setLogo(null);
      setStats({ totalStudents: 0, attendanceToday: 0, historyCount: 0 });
    }
    
    setIsDeleteModalOpen(false);
    handleLogout();
    alert('Akun Anda berhasil dihapus dari sistem secara permanen.');
    return true;
  };

  const handleLogoUpdate = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const base64 = e.target?.result as string;
      setLogo(base64);
      localStorage.setItem('app_logo', base64);
    };
    reader.readAsDataURL(file);
  };

  const renderContent = () => {
    switch (currentPage) {
      case 'landing': return <LandingPage onStart={() => setCurrentPage('login')} logo={logo} onLogoChange={handleLogoUpdate} />;
      case 'login': 
      case 'register': 
        return (
          <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4">
             <div className="fixed inset-0 z-0 overflow-hidden pointer-events-none">
                <div className="absolute top-[-10%] right-[-10%] w-[40%] h-[40%] bg-blue-100/40 blur-[120px] rounded-full" />
                <div className="absolute bottom-[-10%] left-[-10%] w-[50%] h-[50%] bg-indigo-100/40 blur-[120px] rounded-full" />
            </div>
            <AuthForm 
              type={currentPage as 'login' | 'register'} 
              onToggle={() => setCurrentPage(currentPage === 'login' ? 'register' : 'login')}
              onAuthSuccess={handleAuthSuccess}
              onBack={() => setCurrentPage('landing')}
              logo={logo}
            />
          </div>
        );
      default:
        return (
          <div className="min-h-screen bg-[#F5F5F0] flex">
            <Sidebar 
              activePage={currentPage as any} 
              onNavigate={setCurrentPage} 
              onLogout={handleLogout} 
              onDeleteAccount={handleDeleteAccount} 
              logo={logo} 
              isOpen={isSidebarOpen}
              onClose={() => setIsSidebarOpen(false)}
            />
            <main className="flex-1 lg:ml-64 min-h-screen w-full overflow-x-hidden">
              <Header 
                title={
                  currentPage === 'dashboard' ? 'Overview' : 
                  currentPage === 'students' ? 'Database Siswa' : 
                  currentPage === 'attendance' ? 'Input Absensi' : 
                  currentPage === 'teachers' ? 'Absensi Khusus Guru' : 'Rekap Kehadiran'
                } 
                user={user} 
                onLogout={handleLogout}
                logo={logo}
                onToggleSidebar={() => setIsSidebarOpen(!isSidebarOpen)}
              />
              <AnimatePresence mode="wait">
                <motion.div
                  key={currentPage}
                  initial={{ opacity: 0, x: 20 }}
                  animate={{ opacity: 1, x: 0 }}
                  exit={{ opacity: 0, x: -20 }}
                  transition={{ duration: 0.2 }}
                >
                  {currentPage === 'dashboard' && <Dashboard stats={stats} deferredPrompt={deferredPrompt} onInstall={handleInstallApp} />}
                  {currentPage === 'students' && (
                    <StudentsPage 
                      students={students} 
                      onImport={handleImport} 
                      onDeleteStudent={handleDeleteStudent}
                      onDeleteAll={handleDeleteAll}
                      onUpdateStudent={handleUpdateStudent}
                    />
                  )}
                  {currentPage === 'attendance' && (
                    <AttendancePage 
                      students={students} 
                      attendance={attendance} 
                      date={currentDate} 
                      onDateChange={setCurrentDate}
                      onStatusChange={handleStatusChange}
                      onSave={handleSaveAttendance}
                      onEvidenceChange={handleEvidenceChange}
                      onDeleteStudent={handleDeleteStudent}
                      onUpdateStudent={handleUpdateStudent}
                    />
                  )}
                  {currentPage === 'recap' && (
                    <RecapPage 
                      students={students}
                      attendanceHistory={attendanceHistory}
                      onSelectDate={(d) => { setCurrentDate(d); setCurrentPage('attendance'); }}
                      onPreviewImage={setGlobalPreviewImage}
                      onUpdateAttendanceRecord={handleUpdateRecord}
                      onDeleteAttendanceRecord={handleDeleteRecord}
                      onDeleteMonthAttendance={handleDeleteMonthAttendance}
                    />
                  )}
                  {currentPage === 'teachers' && (
                    <TeachersPage 
                      teachersAttendance={teachersAttendance}
                      onAddAttendance={handleAddTeacherAttendance}
                      onDeleteAttendance={handleDeleteTeacherAttendance}
                      onUpdateAttendance={handleUpdateTeacherAttendance}
                      teachersList={teachersList}
                      onImportTeachersList={handleImportTeachersList}
                      onAddTeacher={handleAddTeacher}
                      onDeleteTeacher={handleDeleteTeacher}
                      onClearTeachersList={handleClearTeachersList}
                      onUpdateTeacherPhoto={handleUpdateTeacherPhoto}
                    />
                  )}
                </motion.div>
              </AnimatePresence>
            </main>

            <AnimatePresence>
              {isDeleteModalOpen && (
                <DeleteAccountModal 
                  isOpen={isDeleteModalOpen}
                  onClose={() => setIsDeleteModalOpen(false)}
                  user={user}
                  onConfirm={handleConfirmDeleteAccount}
                />
              )}
              {globalPreviewImage && (
                <div className="fixed inset-0 z-[100] flex items-center justify-center p-8 bg-slate-900/90 backdrop-blur-md">
                  <motion.div 
                    initial={{ opacity: 0, scale: 0.9 }}
                    animate={{ opacity: 1, scale: 1 }}
                    exit={{ opacity: 0, scale: 0.9 }}
                    className="relative max-w-4xl w-full flex flex-col items-center bg-white rounded-3xl p-4 overflow-hidden shadow-2xl"
                  >
                    <button 
                      onClick={() => setGlobalPreviewImage(null)}
                      className="absolute top-4 right-4 p-3 bg-slate-100 hover:bg-slate-200 text-slate-800 rounded-full transition-all shadow-xl z-20"
                    >
                      <X className="w-6 h-6" />
                    </button>
                    <div className="aspect-[3/4] max-h-[70vh] w-full bg-slate-50 rounded-2xl overflow-hidden flex items-center justify-center border border-slate-100">
                      <img src={globalPreviewImage} alt="Lampiran" className="max-w-full max-h-full object-contain" />
                    </div>
                    <div className="mt-6 flex flex-col items-center gap-1">
                      <p className="font-black text-slate-900 text-2xl uppercase tracking-tighter">Lampiran Absensi</p>
                      <p className="text-slate-400 text-xs font-bold uppercase tracking-widest">Detail Dokumen Terlampir</p>
                    </div>
                  </motion.div>
                </div>
              )}
            </AnimatePresence>
          </div>
        );
    }
  };

  return (
    <div className="font-sans antialiased text-[#141414] selection:bg-[#141414] selection:text-white">
      {renderContent()}
    </div>
  );
}
