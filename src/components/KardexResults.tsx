import React, { useState } from 'react';
import { KardexRecord } from '../types';
import { 
  FileText, Calendar, Hash, BookOpen, UserCircle2, Download, Printer, 
  Settings2, Check, Users, TrendingUp, PieChart, Cake, Search, Filter,
  School, UserCheck
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import 'jspdf-autotable';

interface Props {
  results: KardexRecord[];
  loading: boolean;
  searched: boolean;
}

const COLUMN_OPTIONS = [
  { key: 'folio', label: 'Folio' },
  { key: 'courseName', label: 'Curso' },
  { key: 'date', label: 'Mes/Año' },
  { key: 'grade', label: 'Calif.' },
  { key: 'section', label: 'Sección' },
  { key: 'curp', label: 'CURP' },
  { key: 'sexo', label: 'Sexo' },
  { key: 'edad', label: 'Edad' },
  { key: 'fechaNacimiento', label: 'F. Nacimiento' },
  { key: 'semestre', label: 'Semestre' },
  { key: 'aprobo', label: 'Aprobó' },
  { key: 'local', label: 'Local' },
  { key: 'numInterno', label: 'Núm. Interno' },
  { key: 'folioConstancia', label: 'Folio Const.' },
  { key: 'tipoCurso', label: 'Tipo Curso' },
  { key: 'instructor', label: 'Instructor' },
  { key: 'periodoImparticion', label: 'Periodo' },
];

export default function KardexResults({ results, loading, searched }: Props) {
  const [visibleColumns, setVisibleColumns] = useState<string[]>([
    'folio', 'courseName', 'grade', 'section', 'date'
  ]);
  const [showConfig, setShowConfig] = useState(false);
  const [filters, setFilters] = useState({
    age: 'all',
    gender: 'all',
    type: 'all'
  });

  if (loading) return null;

  const toggleColumn = (key: string) => {
    setVisibleColumns(prev => 
      prev.includes(key) ? prev.filter(c => c !== key) : [...prev, key]
    );
  };

  const currentMonth = new Date().getMonth() + 1;

  const filteredResults = results.filter(r => {
    const age = parseInt(r.edad || '0');
    if (filters.age === 'minors' && (age <= 0 || age >= 17)) return false;
    if (filters.age === 'adults' && age < 17) return false;
    if (filters.age === 'birthday') {
      let bMonth = 0;
      if (r.fechaNacimiento?.includes('/')) {
        bMonth = parseInt(r.fechaNacimiento.split('/')[1]);
      } else if (r.curp && r.curp.length >= 10) {
        bMonth = parseInt(r.curp.substring(6, 8));
      }
      if (bMonth !== currentMonth) return false;
    }
    if (filters.gender !== 'all') {
      const g = r.sexo?.toLowerCase() || '';
      if (filters.gender === 'male' && !g.startsWith('h')) return false;
      if (filters.gender === 'female' && !g.startsWith('m')) return false;
    }
    if (filters.type !== 'all') {
      const name = (r.courseName || '').toLowerCase();
      if (filters.type === 'taller' && !name.includes('taller')) return false;
      if (filters.type === 'curso' && name.includes('taller')) return false;
    }
    return true;
  });

  const recordsByUser = filteredResults.reduce((acc, current) => {
    if (!acc[current.userName]) acc[current.userName] = [];
    acc[current.userName].push(current);
    return acc;
  }, {} as Record<string, KardexRecord[]>);

  const exportToExcel = (userName: string, userRecords: KardexRecord[]) => {
    const data = userRecords.map(r => {
      const row: any = {};
      COLUMN_OPTIONS.forEach(opt => {
        if (visibleColumns.includes(opt.key)) row[opt.label] = (r as any)[opt.key];
      });
      return row;
    });
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Kardex');
    XLSX.writeFile(wb, `Kardex_${userName}.xlsx`);
  };

  const exportToPDF = (userName: string, userRecords: KardexRecord[]) => {
    const doc = new jsPDF();
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();

    // ── Encabezado — estilo app ──
    // Fondo blanco ya es default. Franja roja izquierda (ícono simulado)
    // Cuadro rojo redondeado como ícono de la app
    doc.setFillColor(226, 31, 38);
    doc.roundedRect(14, 8, 14, 14, 3, 3, 'F');

    // Cruz / heartbeat: línea horizontal y vertical simples centradas en el cuadro
    doc.setDrawColor(255, 255, 255);
    doc.setLineWidth(1.8);
    doc.line(17, 15, 25, 15);   // horizontal
    doc.line(21, 11, 21, 19);   // vertical

    // Título principal negro, bold, grande — igual que la app
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(14);
    doc.setTextColor(15, 15, 15);
    doc.text('CRUZ ROJA MEXICANA', 33, 14);

    // Badge "DURANGO" — pastilla roja
    doc.setFillColor(255, 237, 237);
    doc.setDrawColor(255, 204, 204);
    doc.setLineWidth(0.4);
    doc.roundedRect(33, 16.5, 22, 5.5, 2, 2, 'FD');
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(6.5);
    doc.setTextColor(226, 31, 38);
    doc.text('DURANGO', 44, 20.3, { align: 'center' });

    // "SISTEMA DE KARDEX" en gris pequeño al lado del badge
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(6.5);
    doc.setTextColor(160, 160, 160);
    doc.text('SISTEMA DE KARDEX', 58, 20.3);

    // Línea separadora
    doc.setDrawColor(230, 230, 230);
    doc.setLineWidth(0.4);
    doc.line(14, 26, pageW - 14, 26);

    // Nombre del alumno
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(10);
    doc.setTextColor(30, 30, 30);
    doc.text(`Kardex del alumno: ${userName}`, 14, 33);

    // ── Tabla ──
    const head = [visibleColumns.map(c => COLUMN_OPTIONS.find(o => o.key === c)?.label || c)];
    const body = userRecords.map(r => visibleColumns.map(c => (r as any)[c] ?? ''));

    (doc as any).autoTable({
      head,
      body,
      startY: 37,
      margin: { left: 14, right: 14 },
      headStyles: {
        fillColor: [226, 31, 38],
        textColor: 255,
        fontStyle: 'bold',
        fontSize: 8,
      },
      bodyStyles: {
        fontSize: 7.5,
        textColor: [40, 40, 40],
      },
      alternateRowStyles: {
        fillColor: [252, 245, 245],
      },
      // ── Pie de página en cada hoja ──
      didDrawPage: () => {
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(7);
        doc.setTextColor(170, 170, 170);
        doc.text(
          'Desarrollado por Alexa Calderón Vázquez  ©  2026 Cruz Roja Mexicana Delegación Durango - Área de Capacitación',
          pageW / 2,
          pageH - 6,
          { align: 'center' }
        );
      },
    });

    doc.save(`Kardex_${userName}.pdf`);
  };

  return (
    <div className="space-y-8 pb-20">
      <div className="flex items-center justify-between">
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 flex-1">
          <motion.button
            whileHover={{ scale: 1.02 }}
            whileTap={{ scale: 0.98 }}
            onClick={() => setFilters(prev => ({ ...prev, age: 'all', gender: 'all' }))}
            className={`p-4 rounded-3xl border transition-all text-left ${filters.age === 'all' && filters.gender === 'all' ? 'bg-indigo-600 border-indigo-600 text-white' : 'bg-white border-slate-200 text-slate-600'}`}
          >
            <TrendingUp className={`w-5 h-5 mb-2 ${filters.age === 'all' && filters.gender === 'all' ? 'text-indigo-200' : 'text-indigo-600'}`} />
            <p className="text-[10px] font-black uppercase tracking-wider opacity-80">Vista Actual</p>
            <h4 className="text-xl font-black">{filteredResults.length} de {results.length}</h4>
          </motion.button>
          <motion.button
            onClick={() => setFilters(prev => ({ ...prev, age: prev.age === 'minors' ? 'all' : 'minors' }))}
            className={`p-4 rounded-3xl border transition-all text-left ${filters.age === 'minors' ? 'bg-indigo-600 border-indigo-600 text-white' : 'bg-white border-slate-200 text-slate-600'}`}
          >
            <Users className="w-5 h-5 mb-2" />
            <p className="text-[10px] font-black uppercase tracking-wider">Menores de 17</p>
            <h4 className="text-xl font-black">{results.filter(r => parseInt(r.edad || '0') < 17 && parseInt(r.edad || '0') > 0).length}</h4>
          </motion.button>
          <motion.button
            onClick={() => setFilters(prev => ({ ...prev, age: prev.age === 'adults' ? 'all' : 'adults' }))}
            className={`p-4 rounded-3xl border transition-all text-left ${filters.age === 'adults' ? 'bg-indigo-600 border-indigo-600 text-white' : 'bg-white border-slate-200 text-slate-600'}`}
          >
            <School className="w-5 h-5 mb-2" />
            <p className="text-[10px] font-black uppercase tracking-wider">Mayores de 17</p>
            <h4 className="text-xl font-black">{results.filter(r => parseInt(r.edad || '0') >= 17).length}</h4>
          </motion.button>
          <motion.button
            onClick={() => setFilters(prev => ({ ...prev, age: prev.age === 'birthday' ? 'all' : 'birthday' }))}
            className={`p-4 rounded-3xl border transition-all text-left ${filters.age === 'birthday' ? 'bg-indigo-600 border-indigo-600 text-white' : 'bg-white border-slate-200 text-slate-600'}`}
          >
            <Cake className="w-5 h-5 mb-2" />
            <p className="text-[10px] font-black uppercase tracking-wider">Cumples Mes</p>
            <h4 className="text-xl font-black">{results.filter(r => {
               const m = r.fechaNacimiento?.includes('/') ? parseInt(r.fechaNacimiento.split('/')[1]) : ((r.curp && r.curp.length >= 10) ? parseInt(r.curp.substring(6, 8)) : 0);
               return m === currentMonth;
            }).length}</h4>
          </motion.button>
        </div>
        <button 
          onClick={() => setShowConfig(!showConfig)}
          className={`ml-4 p-4 rounded-3xl border transition-all flex flex-col items-center justify-center ${showConfig ? 'bg-[#E21F26] border-[#E21F26] text-white shadow-lg shadow-red-100' : 'bg-white border-slate-200 text-slate-600 hover:border-red-300 hover:bg-red-50/30'}`}
        >
          <Settings2 className="w-6 h-6 mb-1" />
          <span className="text-[10px] font-black uppercase">Columnas</span>
        </button>
      </div>

      <AnimatePresence>
        {showConfig && (
          <motion.div
            initial={{ height: 0, opacity: 0 }}
            animate={{ height: 'auto', opacity: 1 }}
            exit={{ height: 0, opacity: 0 }}
            className="overflow-hidden"
          >
            <div className="bg-white p-6 rounded-[2rem] border border-red-50 shadow-sm">
              <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-4">Seleccionar Datos a Mostrar</p>
              <div className="flex flex-wrap gap-2">
                {COLUMN_OPTIONS.map(opt => (
                  <button
                    key={opt.key}
                    onClick={() => toggleColumn(opt.key)}
                    className={`px-4 py-2 rounded-xl text-xs font-bold transition-all border ${visibleColumns.includes(opt.key) ? 'bg-red-50 border-red-100 text-[#E21F26] shadow-sm' : 'bg-white border-slate-100 text-slate-400 hover:border-slate-200'}`}
                  >
                    {visibleColumns.includes(opt.key) && <Check className="w-3 h-3 inline-block mr-1" />}
                    {opt.label}
                  </button>
                ))}
              </div>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
        <div className="bg-white p-6 rounded-3xl border border-slate-200">
          <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-4">Filtro por Sexo</p>
          <div className="flex gap-2">
            <button onClick={() => setFilters(p => ({...p, gender: 'female'}))} className={`flex-1 py-3 rounded-2xl border font-bold text-xs transition-all ${filters.gender === 'female' ? 'bg-indigo-50 border-indigo-200 text-indigo-700' : 'bg-white border-slate-100'}`}>MUJERES</button>
            <button onClick={() => setFilters(p => ({...p, gender: 'male'}))} className={`flex-1 py-3 rounded-2xl border font-bold text-xs transition-all ${filters.gender === 'male' ? 'bg-indigo-50 border-indigo-200 text-indigo-700' : 'bg-white border-slate-100'}`}>HOMBRES</button>
          </div>
        </div>
        <div className="bg-white p-6 rounded-3xl border border-slate-200">
          <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-4">Tipo de Capacitación</p>
          <div className="flex gap-2">
            <button onClick={() => setFilters(p => ({...p, type: 'curso'}))} className={`flex-1 py-3 rounded-2xl border font-bold text-xs transition-all ${filters.type === 'curso' ? 'bg-emerald-50 border-emerald-200 text-emerald-700' : 'bg-white border-slate-100'}`}>CURSOS</button>
            <button onClick={() => setFilters(p => ({...p, type: 'taller'}))} className={`flex-1 py-3 rounded-2xl border font-bold text-xs transition-all ${filters.type === 'taller' ? 'bg-emerald-50 border-emerald-200 text-emerald-700' : 'bg-white border-slate-100'}`}>TALLERES</button>
          </div>
        </div>
      </div>

      {Object.entries(recordsByUser).map(([user, records]) => (
        <motion.div key={user} initial={{ opacity: 0 }} animate={{ opacity: 1 }} className="bg-white rounded-3xl shadow-sm border border-slate-200 overflow-hidden">
          <div className="bg-slate-50 p-6 border-b border-slate-100 flex items-center justify-between">
            <div className="flex items-center gap-3">
              <UserCircle2 className="w-6 h-6 text-indigo-600" />
              <h3 className="text-lg font-black text-slate-800 uppercase">{user}</h3>
            </div>
            <div className="flex gap-2">
              <button onClick={() => exportToExcel(user, records)} className="p-2 hover:bg-slate-200 rounded-xl transition-all"><Download className="w-5 h-5 text-slate-600" /></button>
              <button onClick={() => exportToPDF(user, records)} className="p-2 hover:bg-slate-200 rounded-xl transition-all"><Printer className="w-5 h-5 text-slate-600" /></button>
            </div>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="bg-slate-50/50">
                  {COLUMN_OPTIONS.filter(o => visibleColumns.includes(o.key)).map(opt => (
                    <th key={opt.key} className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">{opt.label}</th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {records.map((r, i) => (
                  <tr key={i} className="hover:bg-slate-50/30 transition-colors">
                    {COLUMN_OPTIONS.filter(o => visibleColumns.includes(o.key)).map(opt => (
                      <td key={opt.key} className="px-6 py-4 text-sm font-medium text-slate-600">{(r as any)[opt.key]}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </motion.div>
      ))}
    </div>
  );
}
