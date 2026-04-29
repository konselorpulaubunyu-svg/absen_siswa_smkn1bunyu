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
  X
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';

// --- Types ---
type Student = {
  nomor: string;
  nama: string;
  kelas: string;
  jurusan: string;
};

type AttendanceStatus = 'H' | 'S' | 'I' | 'A' | '';

type AttendanceItem = {
  nama: string;
  status: AttendanceStatus;
};

type User = {
  name: string;
  username: string;
};

type Page = 'landing' | 'register' | 'login' | 'dashboard' | 'students' | 'attendance';

// --- Components ---

const Sidebar = ({ activePage, onNavigate, onLogout }: { activePage: Page, onNavigate: (p: Page) => void, onLogout: () => void }) => {
  const menuItems = [
    { id: 'dashboard', icon: LayoutDashboard, label: 'Dashboard' },
    { id: 'students', icon: Users, label: 'Data Siswa' },
    { id: 'attendance', icon: ClipboardCheck, label: 'Absensi & Rekap' },
  ];

  return (
    <div className="w-64 h-screen bg-slate-900 flex flex-col fixed left-0 top-0 z-20">
      <div className="p-6 border-b border-slate-800">
        <h1 className="text-white font-bold text-lg tracking-tight">SMKN 1 BUNYU</h1>
        <p className="text-slate-400 text-xs mt-1 font-medium">Sistem Absensi Digital</p>
      </div>
      
      <nav className="flex-1 px-4 py-6 space-y-2">
        {menuItems.map((item) => (
          <button
            key={item.id}
            onClick={() => onNavigate(item.id as Page)}
            className={`w-full flex items-center gap-3 px-4 py-3 rounded-lg transition-all duration-200 group ${
              activePage === item.id 
                ? 'bg-blue-600 text-white shadow-lg shadow-blue-900/20' 
                : 'text-slate-400 hover:bg-slate-800 hover:text-white'
            }`}
          >
            <item.icon className={`w-5 h-5 ${activePage === item.id ? 'text-white' : 'text-slate-400 group-hover:text-white'}`} />
            <span className="font-medium">{item.label}</span>
          </button>
        ))}
      </nav>

      <div className="p-4 border-t border-slate-800">
        <button 
          onClick={onLogout}
          className="w-full flex items-center gap-3 px-4 py-3 text-rose-400 hover:bg-rose-500/10 rounded-lg transition-colors font-medium"
        >
          <LogOut className="w-5 h-5" />
          <span>Logout</span>
        </button>
      </div>
    </div>
  );
};

const Header = ({ title, user }: { title: string, user: User | null }) => (
  <header className="h-20 bg-white border-b border-slate-200 flex items-center justify-between px-8 sticky top-0 z-10 w-full">
    <div>
      <h2 className="text-xl font-bold text-slate-800">{title}</h2>
      <p className="text-slate-500 text-xs italic font-medium">
        {new Intl.DateTimeFormat('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }).format(new Date())}
      </p>
    </div>
    <div className="flex items-center gap-4">
      <div className="text-right hidden sm:block">
        <p className="text-sm font-bold text-slate-800">{user?.name}</p>
        <p className="text-[10px] uppercase font-bold text-slate-400 tracking-wider font-mono">{user?.username}</p>
      </div>
      <div className="w-10 h-10 rounded-full bg-slate-100 border border-slate-200 flex items-center justify-center font-bold text-slate-600 text-sm">
        {user?.name?.[0]?.toUpperCase()}
      </div>
    </div>
  </header>
);

const LandingPage = ({ onStart }: { onStart: () => void }) => (
  <div className="min-h-screen bg-slate-900 flex flex-col items-center justify-center p-4 relative overflow-hidden text-white">
    <div className="absolute inset-0 z-0 opacity-20 pointer-events-none">
        <div className="absolute top-[-10%] left-[-10%] w-[40%] h-[40%] bg-blue-600 blur-[120px] rounded-full" />
        <div className="absolute bottom-[-10%] right-[-10%] w-[50%] h-[50%] bg-indigo-600 blur-[120px] rounded-full" />
    </div>
    
    <motion.div 
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      className="text-center z-10 max-w-2xl px-6"
    >
      <div className="w-20 h-20 bg-blue-600 rounded-2xl flex items-center justify-center mx-auto mb-8 shadow-2xl shadow-blue-600/20">
        <School className="text-white w-10 h-10" />
      </div>
      <h1 className="text-5xl sm:text-7xl font-bold tracking-tight mb-4">
        SMKN 1 BUNYU
      </h1>
      <p className="text-slate-400 font-medium text-lg mb-12 max-w-md mx-auto">
        Sistem Absensi Digital Terpadu untuk Manajemen Sekolah yang Lebih Efisien dan Presisi.
      </p>
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

const AuthForm = ({ type, onToggle, onAuthSuccess }: { type: 'login' | 'register', onToggle: () => void, onAuthSuccess: (user: User) => void }) => {
  const [formData, setFormData] = useState({ name: '', username: '', password: '' });
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState<{ type: 'success' | 'error', text: string } | null>(null);
  const [showPassword, setShowPassword] = useState(false);
  const [recentUsers, setRecentUsers] = useState<any[]>([]);

  useEffect(() => {
    const fetchUsers = async () => {
      try {
        const res = await fetch('/api/users');
        if (res.ok) {
          const data = await res.json();
          setRecentUsers(data);
        }
      } catch (err) {}
    };
    if (type === 'login') fetchUsers();
  }, [type]);

  const handleQuickLogin = (user: any) => {
    setFormData({ ...formData, username: user.username });
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    setMessage(null);

    const endpoint = type === 'login' ? '/api/login' : '/api/register';
    
    try {
      const res = await fetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(formData),
      });
      const data = await res.json();
      
      if (res.ok) {
        setMessage({ type: 'success', text: data.message });
        if (type === 'login') {
          setTimeout(() => onAuthSuccess(data.user), 1000);
        } else {
          setTimeout(() => onToggle(), 1500);
        }
      } else {
        setMessage({ type: 'error', text: data.error });
      }
    } catch (err) {
      setMessage({ type: 'error', text: 'Connection failed' });
    } finally {
      setLoading(false);
    }
  };

  return (
    <motion.div 
      initial={{ opacity: 0, scale: 0.95 }}
      animate={{ opacity: 1, scale: 1 }}
      className="max-w-md w-full p-10 bg-white rounded-2xl shadow-xl border border-slate-200 z-10 relative"
    >
      <div className="mb-10 text-center">
        <div className="w-12 h-12 bg-blue-600 rounded-xl flex items-center justify-center mx-auto mb-4">
          <School className="text-white w-6 h-6" />
        </div>
        <h2 className="text-2xl font-bold text-slate-800 tracking-tight">
          {type === 'login' ? 'Selamat Datang' : 'Registrasi Guru'}
        </h2>
        <p className="text-slate-500 text-sm font-medium mt-1">
          Sistem Absensi SMKN 1 BUNYU
        </p>
        
        {type === 'login' && recentUsers.length > 0 && (
          <div className="mt-6 space-y-3">
            <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest text-center">Klik nama untuk masuk cepat:</p>
            <div className="flex flex-wrap justify-center gap-2">
              {recentUsers.map((u, i) => (
                <button
                  key={i}
                  type="button"
                  onClick={() => handleQuickLogin(u)}
                  className="px-3 py-1.5 bg-slate-100 hover:bg-blue-50 text-slate-700 hover:text-blue-700 rounded-full text-xs font-bold border border-slate-200 hover:border-blue-200 transition-all shrink-0 uppercase"
                >
                  {u.name}
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

      <div className="mt-8 text-center pt-8 border-t border-slate-100">
        <button 
          type="button"
          onClick={onToggle}
          className="text-blue-600 font-bold hover:text-blue-700 transition-colors text-sm"
        >
          {type === 'login' ? 'Belum punya akun? Silahkan daftar disini' : 'Sudah punya akun? Kembali ke Login'}
        </button>
      </div>
    </motion.div>
  );
};

// --- Page Components ---

const Dashboard = ({ stats }: { stats: any }) => (
  <div className="p-8 space-y-8">
    <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
      {[
        { label: 'Total Siswa', value: stats.totalStudents, icon: Users, bgColor: 'bg-indigo-50', borderColor: 'border-indigo-100', textColor: 'text-indigo-700', valColor: 'text-indigo-800' },
        { label: 'Hadir Hari Ini', value: stats.attendanceToday, icon: ClipboardCheck, bgColor: 'bg-emerald-50', borderColor: 'border-emerald-100', textColor: 'text-emerald-700', valColor: 'text-emerald-800' },
        { label: 'Rekap Absensi', value: stats.historyCount, icon: FileSpreadsheet, bgColor: 'bg-amber-50', borderColor: 'border-amber-100', textColor: 'text-amber-700', valColor: 'text-amber-800' },
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
        console.error('Excel parse error:', err);
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

  return (
    <div className="p-8 space-y-8">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
        <div className="flex items-center gap-3 bg-slate-50 border border-slate-200 rounded-lg px-4 py-2.5 w-full max-w-md focus-within:ring-2 focus-within:ring-blue-500/20 transition-all">
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
                <tr key={i} className="hover:bg-slate-50 transition-colors">
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
                >Simpan</button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
};

const AttendancePage = ({ students, attendance, date, onDateChange, onStatusChange, onSave }: { 
  students: Student[], 
  attendance: AttendanceItem[], 
  date: string, 
  onDateChange: (d: string) => void,
  onStatusChange: (nama: string, status: AttendanceStatus) => void,
  onSave: () => void 
}) => {
  const [filter, setFilter] = useState('');

  const filteredStudents = useMemo(() => {
    return students.filter(s => filter ? s.kelas === filter : true);
  }, [students, filter]);

  const handleDownload = () => {
    if (!students.length) return;
    const dataForExcel = students.filter(s => filter ? s.kelas === filter : true).map(s => {
      const statusItem = attendance.find(a => a.nama === s.nama);
      const status = statusItem ? statusItem.status : '';
      return {
        'Nama Siswa': s.nama,
        'Kelas': s.kelas,
        'Status': status === 'H' ? 'Hadir' : 
                  status === 'S' ? 'Sakit' : 
                  status === 'I' ? 'Izin' : 
                  status === 'A' ? 'Alfa' : 'Belum Absen'
      };
    });
    const ws = XLSX.utils.json_to_sheet(dataForExcel);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rekap Absensi");
    XLSX.writeFile(wb, `Rekap_Absensi_${date}.xlsx`);
  };

  return (
    <div className="p-8 space-y-8">
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
            className="bg-green-600 hover:bg-green-700 text-white px-5 py-2.5 rounded-lg text-sm font-semibold shadow-sm transition-all disabled:opacity-50 flex items-center gap-2"
          >
            <Download className="w-4 h-4" />
            <span>Unduh Excel</span>
          </button>
          
          <button 
            onClick={onSave}
            disabled={students.length === 0}
            className="bg-blue-600 hover:bg-blue-700 text-white px-6 py-2.5 rounded-lg text-sm font-semibold shadow-sm transition-all disabled:opacity-50 flex items-center gap-2"
          >
            <CheckCircle2 className="w-4 h-4" />
            <span>Simpan Absensi</span>
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
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {filteredStudents.length > 0 ? filteredStudents.map((s, i) => {
                const status = attendance.find(a => a.nama === s.nama)?.status || '';
                const getStatusStyle = () => {
                  switch(status) {
                    case 'H': return 'bg-emerald-50 text-emerald-800 border-emerald-200';
                    case 'S': return 'bg-indigo-50 text-indigo-800 border-indigo-200';
                    case 'I': return 'bg-amber-50 text-amber-800 border-amber-200';
                    case 'A': return 'bg-rose-50 text-rose-800 border-rose-200';
                    default: return 'bg-slate-50 text-slate-400 border-slate-200';
                  }
                };
                return (
                  <tr key={i} className="hover:bg-slate-50 transition-colors">
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
                      </select>
                    </td>
                  </tr>
                );
              }) : (
                <tr>
                  <td colSpan={3} className="px-6 py-16 text-center text-slate-400 font-medium italic">
                    {students.length === 0 ? 'Belum ada data siswa. Impor terlebih dahulu di menu Data Siswa.' : 'Tidak ada siswa di kelas ini.'}
                  </td>
                </tr>
              )}
            </tbody>
          </table>
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
  const [currentDate, setCurrentDate] = useState(new Date().toISOString().split('T')[0]);
  const [stats, setStats] = useState({ totalStudents: 0, attendanceToday: 0, historyCount: 0 });

  useEffect(() => {
    if (user) {
      if (currentPage === 'landing' || currentPage === 'login' || currentPage === 'register') {
        setCurrentPage('dashboard');
      }
      fetchData();
    }
  }, [user]);

  useEffect(() => {
    if (currentPage === 'attendance' && user) {
      fetchAttendance(currentDate);
    }
    if (currentPage === 'dashboard' && user) {
      fetchStats();
    }
  }, [currentPage, currentDate]);

  const fetchData = async () => {
    try {
      const res = await fetch('/api/students');
      if (res.ok) {
        const data = await res.json();
        setStudents(data);
      }
      fetchStats();
    } catch (err) {
      console.error('Fetch failed', err);
    }
  };

  const fetchStats = async () => {
    try {
      const res = await fetch('/api/stats');
      if (res.ok) setStats(await res.json());
    } catch (err) {}
  };

  const fetchAttendance = async (date: string) => {
    try {
      const res = await fetch(`/api/attendance/${date}`);
      if (res.ok) {
        const data = await res.json();
        setAttendance(data);
      }
    } catch (err) {}
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

  const handleImport = async (newStudents: Student[]) => {
    try {
      const res = await fetch('/api/students/import', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ students: newStudents }),
      });
      if (res.ok) {
        setStudents(newStudents);
        fetchStats();
      }
    } catch (err) {}
  };

  const handleDeleteStudent = async (nomor: string) => {
    try {
      const res = await fetch(`/api/students/${nomor}`, { method: 'DELETE' });
      if (res.ok) {
        setStudents(prev => prev.filter(s => s.nomor !== nomor));
        fetchStats();
      }
    } catch (err) {}
  };

  const handleDeleteAll = async () => {
    try {
      const res = await fetch('/api/students', { method: 'DELETE' });
      if (res.ok) {
        setStudents([]);
        fetchStats();
      }
    } catch (err) {}
  };

  const handleUpdateStudent = async (oldNomor: string, updated: Student) => {
    try {
      const res = await fetch(`/api/students/${oldNomor}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(updated),
      });
      if (res.ok) {
        setStudents(prev => prev.map(s => s.nomor === oldNomor ? updated : s));
      }
    } catch (err) {}
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

  const handleSaveAttendance = async () => {
    try {
      const res = await fetch('/api/attendance', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ date: currentDate, data: attendance }),
      });
      if (res.ok) {
        alert('Data absensi hari ini berhasil disimpan!');
        fetchStats();
      }
    } catch (err) {}
  };

  const renderContent = () => {
    switch (currentPage) {
      case 'landing': return <LandingPage onStart={() => setCurrentPage('login')} />;
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
            />
          </div>
        );
      default:
        return (
          <div className="min-h-screen bg-[#F5F5F0] flex">
            <Sidebar activePage={currentPage} onNavigate={setCurrentPage} onLogout={handleLogout} />
            <main className="flex-1 ml-64 min-h-screen">
              <Header 
                title={
                  currentPage === 'dashboard' ? 'Overview' : 
                  currentPage === 'students' ? 'Database Siswa' : 'Rekap Kehadiran'
                } 
                user={user} 
              />
              <AnimatePresence mode="wait">
                <motion.div
                  key={currentPage}
                  initial={{ opacity: 0, x: 20 }}
                  animate={{ opacity: 1, x: 0 }}
                  exit={{ opacity: 0, x: -20 }}
                  transition={{ duration: 0.2 }}
                >
                  {currentPage === 'dashboard' && <Dashboard stats={stats} />}
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
                    />
                  )}
                </motion.div>
              </AnimatePresence>
            </main>
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
