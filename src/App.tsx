import React, { useState } from 'react';
import ExcelUploader from './components/ExcelUploader';
import KardexSearch from './components/KardexSearch';
import KardexResults from './components/KardexResults';
import { KardexRecord } from './types';
import { Search, LayoutDashboard, HeartPulse, ShieldCheck, FileText, Database } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { db } from './lib/firebase';
import { collection, writeBatch } from 'firebase/firestore';

export default function App() {
  const [results, setResults] = useState<KardexRecord[]>([]);
  const [loading, setLoading] = useState(false);
  const [searched, setSearched] = useState(false);
  const [view, setView] = useState<'search' | 'upload'>('search');

  const handleSecretClear = async () => {
    const pwd = window.prompt("¿De verdad desea eliminar las bases de datos? Ingrese la clave para continuar:");
    if (pwd !== "*******") return;

    if (!window.confirm("¿Está seguro? Esta acción es irreversible.")) return;

    try {
      setLoading(true);
      const { getDocs, query } = await import('firebase/firestore');
      const q = query(collection(db, 'kardex'));
      const snapshot = await getDocs(q);
      
      let deleted = 0;
      const BATCH_SIZE = 450;
      let batch = writeBatch(db);
      
      for (const docSnapshot of snapshot.docs) {
        batch.delete(docSnapshot.ref);
        deleted++;
        
        if (deleted % BATCH_SIZE === 0) {
          await batch.commit();
          batch = writeBatch(db);
        }
      }
      if (deleted % BATCH_SIZE !== 0) {
        await batch.commit();
      }
      alert(`Se borraron ${deleted} registros exitosamente.`);
    } catch (error) {
      console.error(error);
      alert('Error al borrar la base de datos.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-[#F8F9FA] pb-20 font-sans">
      {/* Dynamic Header / Red Cross Band */}
      <div className="bg-[#E21F26] h-2 w-full fixed top-0 z-[60]" />
      
      <header className="bg-white border-b border-slate-200 sticky top-2 z-50 shadow-sm">
        <div className="max-w-7xl mx-auto px-6 py-4 flex flex-col md:flex-row items-center justify-between gap-4">
          <div className="flex items-center gap-4">
            <div className="w-12 h-12 bg-[#E21F26] rounded-2xl flex items-center justify-center shadow-lg shadow-red-200">
              <HeartPulse className="text-white w-7 h-7" />
            </div>
            <div>
              <h1 className="text-xl font-black text-slate-900 tracking-tight flex items-center gap-2">
                CRUZ ROJA MEXICANA
              </h1>
              <div className="flex items-center gap-2">
                <span 
                  onClick={handleSecretClear}
                  className="text-[10px] font-black bg-red-50 text-[#E21F26] px-2 py-0.5 rounded-full border border-red-100 cursor-pointer"
                >
                  DURANGO
                </span>
                <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest hidden md:block">SISTEMA DE KARDEX</p>
              </div>
            </div>
          </div>

          <div className="flex items-center gap-4">
            <nav className="flex bg-slate-100 p-1 rounded-2xl border border-slate-200">
              <button
                onClick={() => setView('search')}
                className={`flex items-center gap-2 px-4 md:px-6 py-2.5 rounded-xl text-[10px] md:text-xs font-black transition-all ${view === 'search' ? 'bg-white text-[#E21F26] shadow-md' : 'text-slate-500 hover:text-slate-700'}`}
              >
                <Search className="w-4 h-4" />
                <span className="hidden sm:inline">CONSULTAR REGISTROS</span>
                <span className="sm:hidden">CONSULTA</span>
              </button>
              <button
                onClick={() => setView('upload')}
                className={`flex items-center gap-2 px-4 md:px-6 py-2.5 rounded-xl text-[10px] md:text-xs font-black transition-all ${view === 'upload' ? 'bg-white text-[#E21F26] shadow-md' : 'text-slate-500 hover:text-slate-700'}`}
              >
                <LayoutDashboard className="w-4 h-4" />
                <span className="hidden sm:inline">CARGA MASIVA</span>
                <span className="sm:hidden">CARGA</span>
              </button>
            </nav>
          </div>
        </div>
      </header>


      <main className="max-w-7xl mx-auto px-6 pt-10">
        <AnimatePresence mode="wait">
          {view === 'upload' ? (
            <motion.div 
              key="upload"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="max-w-3xl mx-auto"
            >
              <div className="mb-8 text-center">
                <h2 className="text-3xl font-black text-slate-900">Módulo de Carga Masiva</h2>
                <p className="text-slate-400 font-medium mt-2">Sincroniza múltiples archivos Excel de instructores y cursos</p>
              </div>
              <ExcelUploader onBack={() => setView('search')} />
              
              <div className="mt-12 grid grid-cols-1 md:grid-cols-3 gap-6">
                <div className="p-6 bg-white rounded-3xl border border-slate-200">
                  <ShieldCheck className="w-8 h-8 text-green-500 mb-4" />
                  <h4 className="font-bold text-slate-800">Seguridad Total</h4>
                  <p className="text-xs text-slate-500 mt-2">Toda la información se cifra y almacena de forma segura en la nube.</p>
                </div>
                <div className="p-6 bg-white rounded-3xl border border-slate-200">
                  <FileText className="w-8 h-8 text-blue-500 mb-4" />
                  <h4 className="font-bold text-slate-800">Formatos Excel</h4>
                  <p className="text-xs text-slate-500 mt-2">Compatible con formatos .xlsx y .xls de registros internos.</p>
                </div>
                <div className="p-6 bg-white rounded-3xl border border-slate-200">
                  <Database className="w-8 h-8 text-purple-500 mb-4" />
                  <h4 className="font-bold text-slate-800">Carga Múltiple</h4>
                  <p className="text-xs text-slate-500 mt-2">Ahora puedes seleccionar varios archivos a la vez para agilizar el proceso.</p>
                </div>
              </div>
            </motion.div>
          ) : (
            <motion.div 
              key="search"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="space-y-12"
            >
              {!searched && results.length === 0 && (
                <div className="bg-white p-12 rounded-[3rem] border border-slate-100 text-center shadow-xl shadow-slate-200/50 max-w-4xl mx-auto mb-12">
                   <div className="w-20 h-20 bg-red-50 rounded-full flex items-center justify-center mx-auto mb-6">
                     <Search className="w-10 h-10 text-[#E21F26]" />
                   </div>
                   <h2 className="text-4xl font-black text-slate-900 mb-4 tracking-tight">Panel de Consulta de Kardex</h2>
                   <p className="text-slate-500 max-w-xl mx-auto text-lg leading-relaxed">
                     Bienvenido al sistema de control académico de la <span className="font-bold text-[#E21F26]">Cruz Roja Mexicana Delegación Durango</span>. Busca por nombre de alumno o folio.
                   </p>
                </div>
              )}
              
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
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      <footer className="fixed bottom-0 w-full bg-white/80 backdrop-blur-md border-t border-slate-200 py-3 z-40">
        <div className="max-w-7xl mx-auto px-6 flex justify-between items-center">
          <p className="text-[10px] font-bold text-slate-400 tracking-widest uppercase">
            © {new Date().getFullYear()} Cruz Roja Mexicana Delegación Durango - Área de Capacitación
          </p>
        </div>
      </footer>
    </div>
  );
}
