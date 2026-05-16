import React, { useState } from 'react';
import { Search, Loader2, Calendar, Filter } from 'lucide-react';
import { collection, query, where, getDocs, limit } from 'firebase/firestore';
import { db, handleFirestoreError, normalizeText } from '../lib/firebase';
import { KardexRecord, OperationType } from '../types';

interface Props {
  onResults: (results: KardexRecord[]) => void;
  setLoading: (loading: boolean) => void;
  setSearched: (searched: boolean) => void;
}

const YEARS = ['2026', '2025', '2024', '2023', '2022', '2021', '2020'];
const MONTHS = [
  { value: '1', label: 'Enero' },
  { value: '2', label: 'Febrero' },
  { value: '3', label: 'Marzo' },
  { value: '4', label: 'Abril' },
  { value: '5', label: 'Mayo' },
  { value: '6', label: 'Junio' },
  { value: '7', label: 'Julio' },
  { value: '8', label: 'Agosto' },
  { value: '9', label: 'Septiembre' },
  { value: '10', label: 'Octubre' },
  { value: '11', label: 'Noviembre' },
  { value: '12', label: 'Diciembre' }
];

export default function KardexSearch({ onResults, setLoading, setSearched }: Props) {
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedYear, setSelectedYear] = useState('');
  const [selectedMonth, setSelectedMonth] = useState('');

  const handleSearch = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!searchTerm.trim() && !selectedYear) return;

    setLoading(true);
    setSearched(true);
    
    try {
      const qText = normalizeText(searchTerm);
      const keywords = qText.split(/\s+/).filter(k => k.length > 0);

      let q;
      if (keywords.length > 0) {
        q = query(
          collection(db, 'kardex'),
          where('searchKeywords', 'array-contains', keywords[0]),
          limit(300)
        );
      } else if (selectedYear && selectedMonth) {
        q = query(
          collection(db, 'kardex'),
          where('searchKeywords', 'array-contains', `${selectedMonth}-${selectedYear}`),
          limit(500)
        );
      } else if (selectedYear) {
        q = query(
          collection(db, 'kardex'),
          where('searchKeywords', 'array-contains', selectedYear),
          limit(500)
        );
      } else {
        q = query(collection(db, 'kardex'), limit(500));
      }

      const querySnapshot = await getDocs(q);
      let results: KardexRecord[] = [];
      querySnapshot.forEach((doc) => {
        results.push({ id: doc.id, ...doc.data() } as KardexRecord);
      });

      // Local refinement
      if (keywords.length > 1) {
        results = results.filter(r => {
          const combined = normalizeText(`${r.userName} ${r.folio}`);
          return keywords.every(k => combined.includes(k));
        });
      }

      if (selectedYear) {
        results = results.filter(r => {
          if (!r.date || r.date === 'N/A') return false;
          // date is usually MM-yyyy or M-yyyy
          const parts = r.date.split('-');
          if (parts.length !== 2) return false;
          const [m, y] = parts;
          
          if (y !== selectedYear) return false;
          if (selectedMonth && m !== selectedMonth) return false;
          
          return true;
        });
      }

      onResults(results);
    } catch (error) {
      handleFirestoreError(error, OperationType.LIST, 'kardex');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="bg-white p-8 rounded-[2.5rem] shadow-xl shadow-slate-200/50 border border-slate-100">
      <div className="flex items-center gap-3 mb-6">
        <div className="w-8 h-8 bg-red-50 rounded-lg flex items-center justify-center">
          <Search className="text-[#E21F26] w-5 h-5" />
        </div>
        <h2 className="text-xl font-black text-slate-800 tracking-tight">Panel de Búsqueda</h2>
      </div>
      <form onSubmit={handleSearch} className="flex flex-col gap-4">
        
        <div className="flex flex-col md:flex-row gap-4">
          <div className="relative flex-1">
            <Search className="absolute left-5 top-1/2 -translate-y-1/2 text-slate-400 w-5 h-5" />
            <input
              type="text"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              placeholder="Ingrese Nombre del Alumno o Folio..."
              className="w-full pl-14 pr-6 py-4 bg-slate-50 border-2 border-slate-50 rounded-2xl focus:outline-none focus:border-red-100 focus:bg-white focus:ring-4 focus:ring-red-50/50 font-bold text-slate-700 placeholder:text-slate-400 transition-all text-base"
            />
          </div>
          <div className="flex gap-4">
            <div className="relative w-32">
              <Calendar className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400 w-4 h-4" />
              <select
                value={selectedYear}
                onChange={(e) => setSelectedYear(e.target.value)}
                className="w-full pl-10 pr-8 py-4 bg-slate-50 border-2 border-slate-50 rounded-2xl focus:outline-none focus:border-red-100 focus:bg-white focus:ring-4 focus:ring-red-50/50 font-bold text-slate-700 appearance-none text-base cursor-pointer"
              >
                <option value="">AÑO</option>
                {YEARS.map(y => <option key={y} value={y}>{y}</option>)}
              </select>
            </div>
            <div className="relative w-40">
              <Filter className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400 w-4 h-4" />
              <select
                value={selectedMonth}
                onChange={(e) => setSelectedMonth(e.target.value)}
                disabled={!selectedYear}
                className="w-full pl-10 pr-8 py-4 bg-slate-50 border-2 border-slate-50 rounded-2xl focus:outline-none focus:border-red-100 focus:bg-white focus:ring-4 focus:ring-red-50/50 font-bold text-slate-700 appearance-none text-base cursor-pointer disabled:opacity-50 disabled:cursor-not-allowed"
              >
                <option value="">MES (Opcional)</option>
                {MONTHS.map(m => <option key={m.value} value={m.value}>{m.label}</option>)}
              </select>
            </div>
          </div>
        </div>

        <button
          type="submit"
          className="w-full md:w-auto px-10 py-5 bg-[#E21F26] hover:bg-[#c41a21] text-white font-black rounded-2xl shadow-xl shadow-red-200 transition-all active:scale-95 flex items-center justify-center gap-2 uppercase tracking-wider mt-2"
        >
          <Search className="w-5 h-5" />
          Realizar Búsqueda
        </button>
      </form>
    </div>
  );
}
