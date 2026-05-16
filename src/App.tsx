import React, { useState } from 'react';
import ExcelUploader from './components/ExcelUploader';
import KardexSearch from './components/KardexSearch';
import KardexResults from './components/KardexResults';
import { KardexRecord } from './types';
import { Database, Search, LayoutDashboard } from 'lucide-react';

export default function App() {
  const [results, setResults] = useState<KardexRecord[]>([]);
  const [loading, setLoading] = useState(false);
  const [searched, setSearched] = useState(false);
  const [view, setView] = useState<'search' | 'upload'>('search');

  return (
    <div className="min-h-screen pb-20">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-indigo-600 rounded-xl flex items-center justify-center">
              <Database className="text-white w-6 h-6" />
            </div>
            <div>
              <h1 className="text-xl font-black text-slate-800 tracking-tight">KARDEX SYSTEM</h1>
              <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Gestión de Capacitación</p>
            </div>
          </div>

          <nav className="flex bg-slate-100 p-1 rounded-2xl">
            <button
              onClick={() => setView('search')}
              className={`flex items-center gap-2 px-6 py-2 rounded-xl text-sm font-black transition-all ${view === 'search' ? 'bg-white text-indigo-600 shadow-sm' : 'text-slate-500 hover:text-slate-700'}`}
            >
              <Search className="w-4 h-4" />
              CONSULTA
            </button>
            <button
              onClick={() => setView('upload')}
              className={`flex items-center gap-2 px-6 py-2 rounded-xl text-sm font-black transition-all ${view === 'upload' ? 'bg-white text-indigo-600 shadow-sm' : 'text-slate-500 hover:text-slate-700'}`}
            >
              <LayoutDashboard className="w-4 h-4" />
              CARGA
            </button>
          </nav>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-6 pt-12">
        {view === 'upload' ? (
          <div className="max-w-2xl mx-auto">
            <ExcelUploader />
          </div>
        ) : (
          <div className="space-y-12">
            <KardexSearch 
              onResults={setResults} 
              setLoading={setLoading} 
              setSearched={setSearched} 
            />
            <KardexResults 
              results={results} 
              loading={loading} 
              searched={searched} 
            />
          </div>
        )}
      </main>
    </div>
  );
}
