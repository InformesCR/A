import React, { useState } from 'react';
import { Search, Loader2 } from 'lucide-react';
import { collection, query, where, getDocs, limit } from 'firebase/firestore';
import { db, handleFirestoreError, normalizeText } from '../lib/firebase';
import { KardexRecord, OperationType } from '../types';

interface Props {
  onResults: (results: KardexRecord[]) => void;
  setLoading: (loading: boolean) => void;
  setSearched: (searched: boolean) => void;
}

export default function KardexSearch({ onResults, setLoading, setSearched }: Props) {
  const [searchTerm, setSearchTerm] = useState('');

  const handleSearch = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!searchTerm.trim()) return;

    setLoading(true);
    setSearched(true);
    
    try {
      const qText = normalizeText(searchTerm);
      const keywords = qText.split(/\s+/).filter(k => k.length > 0);

      // Simple implementation: Search by first keyword in searchKeywords array
      const q = query(
        collection(db, 'kardex'),
        where('searchKeywords', 'array-contains', keywords[0]),
        limit(200)
      );

      const querySnapshot = await getDocs(q);
      const results: KardexRecord[] = [];
      querySnapshot.forEach((doc) => {
        results.push({ id: doc.id, ...doc.data() } as KardexRecord);
      });

      // Local refinement if multiple keywords
      const filtered = keywords.length > 1 
        ? results.filter(r => {
            const combined = normalizeText(`${r.userName} ${r.folio}`);
            return keywords.every(k => combined.includes(k));
          })
        : results;

      onResults(filtered);
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
      <form onSubmit={handleSearch} className="flex flex-col md:flex-row gap-4">
        <div className="relative flex-1">
          <Search className="absolute left-5 top-1/2 -translate-y-1/2 text-slate-400 w-5 h-5" />
          <input
            type="text"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            placeholder="Ingrese Nombre del Alumno o Folio de Consulta..."
            className="w-full pl-14 pr-6 py-5 bg-slate-50 border-2 border-slate-50 rounded-2xl focus:outline-none focus:border-red-100 focus:bg-white focus:ring-4 focus:ring-red-50/50 font-bold text-slate-700 placeholder:text-slate-400 transition-all text-lg"
          />
        </div>
        <button
          type="submit"
          className="px-10 py-5 bg-[#E21F26] hover:bg-[#c41a21] text-white font-black rounded-2xl shadow-xl shadow-red-200 transition-all active:scale-95 flex items-center justify-center gap-2 uppercase tracking-wider"
        >
          <Search className="w-5 h-5" />
          Realizar Búsqueda
        </button>
      </form>
    </div>
  );
}
