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
    <div className="bg-white p-8 rounded-3xl shadow-xl border border-slate-200">
      <h2 className="text-2xl font-black text-slate-800 mb-6">Consulta de Kardex</h2>
      <form onSubmit={handleSearch} className="flex gap-4">
        <div className="relative flex-1">
          <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400 w-5 h-5" />
          <input
            type="text"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            placeholder="Buscar por Nombre o Folio..."
            className="w-full pl-12 pr-4 py-4 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:ring-2 focus:ring-indigo-500 font-medium transition-all"
          />
        </div>
        <button
          type="submit"
          className="px-8 py-4 bg-indigo-600 hover:bg-indigo-700 text-white font-black rounded-2xl shadow-lg shadow-indigo-200 transition-all active:scale-95"
        >
          BUSCAR
        </button>
      </form>
    </div>
  );
}
