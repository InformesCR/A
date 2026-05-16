import React, { useCallback, useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { db, normalizeText, auth } from '../lib/firebase';
import { collection, writeBatch, doc } from 'firebase/firestore';
import { signInWithPopup, GoogleAuthProvider, signOut, User } from 'firebase/auth';
import { KardexRecord } from '../types';
import { Upload, FileSpreadsheet, Loader2, CheckCircle2, AlertCircle, LogOut, ShieldAlert } from 'lucide-react';
import { motion } from 'motion/react';

const ADMIN_EMAIL = 'alexa.calderon@itdurango.edu.mx';

export default function ExcelUploader() {
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState({ current: 0, total: 0 });
  const [status, setStatus] = useState<{ type: 'success' | 'error'; message: string } | null>(null);
  
  const [user, setUser] = useState<User | null>(null);
  const [authLoading, setAuthLoading] = useState(true);

  useEffect(() => {
    const unsubscribe = auth.onAuthStateChanged((u) => {
      setUser(u);
      setAuthLoading(false);
    });
    return () => unsubscribe();
  }, []);

  const handleLogin = async () => {
    try {
      const provider = new GoogleAuthProvider();
      await signInWithPopup(auth, provider);
    } catch (error) {
      console.error("Login failed", error);
      setStatus({ type: 'error', message: 'Error al iniciar sesión.' });
    }
  };

  const handleLogout = async () => {
    await signOut(auth);
  };

  const onFileUpload = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;

    setLoading(true);
    setStatus(null);
    setProgress({ current: 0, total: 0 });

    try {
      const workbooks: XLSX.WorkBook[] = [];
      const genDataMap = new Map<string, any>();

      // 1. Read all files into memory once
      for (const file of files) {
        const buffer = await file.arrayBuffer();
        workbooks.push(XLSX.read(buffer, { type: 'array', cellDates: true }));
      }

      const findKey = (obj: any, candidates: string[]) => {
        if (!obj) return undefined;
        return Object.keys(obj).find(k => candidates.includes(normalizeText(k)));
      };

      // 2. First pass: Gather metadata (Instructors/Tipo) from ALL sheets in ALL files
      for (const workbook of workbooks) {
        for (const sName of workbook.SheetNames) {
          const sheet = workbook.Sheets[sName];
          const data = XLSX.utils.sheet_to_json<any>(sheet);
          if (data.length === 0) continue;

          const keys = {
            folio: findKey(data[0], ['folio', 'informe', 'no', 'num', 'id', 'a.']),
            tipo: findKey(data[0], ['tipo de curso', 'c.']),
            instructorKeys: Object.keys(data[0]).filter(k => {
              const nk = normalizeText(k);
              return nk === 'ec' || nk === 'ed' || nk === 'ee' || nk.includes('instructor');
            })
          };

          if (!keys.folio) continue;

          data.forEach((row) => {
            const folioRaw = row[keys.folio!];
            if (!folioRaw) return;

            const fStr = String(folioRaw).trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
            const existing = genDataMap.get(fStr) || { date: 'N/A', section: 'N/A', tipoCurso: 'N/A', instructor: 'N/A' };
            
            const tipoCurso = keys.tipo ? String(row[keys.tipo] || existing.tipoCurso) : existing.tipoCurso;
            let instructor = existing.instructor;
            if (keys.instructorKeys.length > 0) {
              const names = keys.instructorKeys.map(k => String(row[k] || '').trim()).filter(n => n.length > 2);
              if (names.length > 0) instructor = names.join(', ');
            }

            genDataMap.set(fStr, {
              date: existing.date,
              section: existing.section,
              tipoCurso: String(tipoCurso),
              instructor: String(instructor)
            });
          });
        }
      }

      // 3. Second pass: Gather student records from ALL sheets in ALL files
      let totalRecordsToUpload: KardexRecord[] = [];
      for (const workbook of workbooks) {
        for (const sName of workbook.SheetNames) {
          const sheet = workbook.Sheets[sName];
          const rows = XLSX.utils.sheet_to_json<any>(sheet);
          if (rows.length === 0) continue;

          const keys = {
            folio: findKey(rows[0], ['folio', 'informe', 'id', 'no']),
            nombres: findKey(rows[0], ['nombre', 'nombres']),
            pApellido: findKey(rows[0], ['primer apellido', 'apellido paterno', 'paterno']),
            sApellido: findKey(rows[0], ['segundo apellido', 'apellido materno', 'materno']),
            curso: findKey(rows[0], ['curso', 'nombre curso', 'capacitacion']),
            grade: findKey(rows[0], ['calificacion', 'promedio', 'nota']),
            curp: findKey(rows[0], ['curp']),
            sexo: findKey(rows[0], ['sexo', 'genero']),
            edad: findKey(rows[0], ['edad']),
            nacimiento: findKey(rows[0], ['fecha de nacimiento', 'nacimiento']),
            semestre: findKey(rows[0], ['semestre', 'nivel']),
            aprobo: findKey(rows[0], ['aprobo', 'resultado', 'estatus']),
            folioConst: findKey(rows[0], ['folio constancia', 'num constancia']),
            interno: findKey(rows[0], ['num interno', 'interno']),
            pref: findKey(rows[0], ['nombre preferencia', 'alias'])
          };

          // Heuristic: If it has folio and nombres, it's a student record sheet
          if (!keys.folio || !keys.nombres) continue;

          rows.forEach((row) => {
            const folio = String(row[keys.folio!] || '');
            if (!folio) return;

            const fStr = folio.trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
            const nombres = String(row[keys.nombres!] || '');
            const pApe = keys.pApellido ? String(row[keys.pApellido] || '') : '';
            const sApe = keys.sApellido ? String(row[keys.sApellido] || '') : '';
            const fullName = `${nombres} ${pApe} ${sApe}`.trim().replace(/\s+/g, ' ');

            const genData = genDataMap.get(fStr) || { date: 'N/A', section: 'N/A', tipoCurso: 'N/A', instructor: 'N/A' };

            let finalDate = genData.date;
            const rMonth = row['Mes'] || row['MES'] || row['mes'];
            const rYear = row['Año'] || row['AÑO'] || row['Ano'] || row['ANO'] || row['ano'];
            
            if (rMonth || rYear) {
              let m = '??';
              let y = '????';
              if (rMonth instanceof Date) m = String(rMonth.getMonth() + 1);
              else if (rMonth) m = String(rMonth).trim().padStart(2, '0').replace(/^0+/, '');
              
              if (rYear instanceof Date) y = String(rYear.getFullYear());
              else if (rYear) y = String(rYear).trim().split('-')[0];
              
              if (m !== '??' || y !== '????') finalDate = `${m}-${y}`;
            }

            totalRecordsToUpload.push({
              userName: fullName,
              folio: fStr,
              courseName: keys.curso ? String(row[keys.curso] || 'N/A') : 'N/A',
              grade: keys.grade ? String(row[keys.grade] || 'N/A') : 'N/A',
              section: genData.section,
              date: finalDate,
              curp: keys.curp ? String(row[keys.curp] || '').toUpperCase() : '',
              sexo: keys.sexo ? String(row[keys.sexo] || '') : '',
              edad: keys.edad ? String(row[keys.edad] || '') : '',
              fechaNacimiento: keys.nacimiento ? (row[keys.nacimiento] instanceof Date ? row[keys.nacimiento].toLocaleDateString() : String(row[keys.nacimiento])) : '',
              semestre: keys.semestre ? String(row[keys.semestre] || '') : '',
              aprobo: keys.aprobo ? String(row[keys.aprobo] || '') : '',
              folioConstancia: keys.folioConst ? String(row[keys.folioConst] || '') : '',
              numInterno: keys.interno ? String(row[keys.interno] || '') : '',
              nombrePreferencia: keys.pref ? String(row[keys.pref] || '') : '',
              tipoCurso: genData.tipoCurso,
              instructor: genData.instructor,
              searchKeywords: [
                ...normalizeText(fullName).split(/\s+/),
                fStr.toLowerCase()
              ].filter(w => w.length > 0),
              uploadedAt: new Date().toISOString()
            });
          });
        }
      }

      const total = totalRecordsToUpload.length;
      if (total === 0) {
        setStatus({ type: 'error', message: 'No se encontraron registros válidos para cargar.' });
        setLoading(false);
        return;
      }

      setProgress({ current: 0, total });

      // 3. Batched upload to Firestore
      const BATCH_SIZE = 450;
      for (let i = 0; i < total; i += BATCH_SIZE) {
        const batch = writeBatch(db);
        const chunk = totalRecordsToUpload.slice(i, i + BATCH_SIZE);
        
        chunk.forEach((record) => {
          const docRef = doc(collection(db, 'kardex'));
          batch.set(docRef, record);
        });

        await batch.commit();
        setProgress(prev => ({ ...prev, current: i + chunk.length }));
      }

      setStatus({ type: 'success', message: `Se cargaron ${total} registros de ${files.length} archivos exitosamente.` });
    } catch (err) {
      console.error(err);
      setStatus({ type: 'error', message: 'Error procesando los archivos.' });
    } finally {
      setLoading(false);
      e.target.value = '';
    }
  }, []);

  if (authLoading) {
    return (
      <div className="bg-white p-8 rounded-[2.5rem] shadow-xl border border-slate-100 flex flex-col items-center justify-center h-64">
        <Loader2 className="w-8 h-8 text-[#E21F26] animate-spin" />
      </div>
    );
  }

  if (!user) {
    return (
      <div className="bg-white p-10 rounded-[2.5rem] shadow-xl border border-slate-100 text-center flex flex-col items-center">
        <div className="w-20 h-20 bg-red-50 rounded-full flex items-center justify-center mb-6">
          <ShieldAlert className="w-10 h-10 text-[#E21F26]" />
        </div>
        <h2 className="text-3xl font-black text-slate-900 mb-3 tracking-tight">Acceso Restringido</h2>
        <p className="text-slate-500 mb-8 max-w-md text-lg">Para acceder al módulo de carga masiva, por favor inicie sesión con su cuenta institucional.</p>
        <button
          onClick={handleLogin}
          className="px-8 py-4 bg-[#E21F26] hover:bg-[#c41a21] text-white font-black rounded-2xl shadow-lg shadow-red-200 transition-all active:scale-95 flex items-center gap-3 uppercase tracking-wide text-sm"
        >
          <LogOut className="w-5 h-5 -scale-x-100" />
          Iniciar Sesión
        </button>
      </div>
    );
  }

  if (user.email !== ADMIN_EMAIL) {
    return (
      <div className="bg-white p-10 rounded-[2.5rem] shadow-xl border border-slate-100 text-center flex flex-col items-center">
        <div className="w-20 h-20 bg-amber-50 rounded-full flex items-center justify-center mb-6">
          <AlertCircle className="w-10 h-10 text-amber-500" />
        </div>
        <h2 className="text-3xl font-black text-slate-900 mb-3 tracking-tight">Usuario No Autorizado</h2>
        <p className="text-slate-500 mb-8 max-w-md text-lg">La cuenta <span className="font-bold text-slate-700">{user.email}</span> no cuenta con los permisos de administrador necesarios.</p>
        <button
          onClick={handleLogout}
          className="px-6 py-4 bg-slate-100 hover:bg-slate-200 text-slate-600 font-bold rounded-2xl transition-all active:scale-95 flex items-center gap-2 text-sm uppercase tracking-wide"
        >
          <LogOut className="w-4 h-4" />
          Cerrar Sesión
        </button>
      </div>
    );
  }

  return (
    <div className="bg-white p-8 rounded-[2.5rem] shadow-xl border border-slate-100">
      <div className="flex items-center gap-4 mb-8">
        <div className="p-4 bg-red-50 rounded-2xl">
          <FileSpreadsheet className="w-8 h-8 text-[#E21F26]" />
        </div>
        <div className="flex-1">
          <h2 className="text-2xl font-black text-slate-900">Módulo de Carga</h2>
          <p className="text-slate-500 text-sm font-medium mt-1">Sube archivos Excel (.xls, .xlsx) para sincronizar registros</p>
        </div>
        <button 
          onClick={handleLogout}
          className="p-4 bg-slate-50 hover:bg-rose-50 text-slate-400 hover:text-[#E21F26] rounded-2xl transition-all flex items-center gap-2"
          title="Cerrar sesión"
        >
          <LogOut className="w-5 h-5" />
          <span className="text-xs font-bold uppercase hidden md:inline">Salir</span>
        </button>
      </div>

      <div className="relative group">
        <input
          type="file"
          accept=".xlsx, .xls"
          multiple
          onChange={onFileUpload}
          disabled={loading}
          className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
        />
        <div className={`border-2 border-dashed rounded-[2rem] p-12 flex flex-col items-center justify-center transition-all ${loading ? 'bg-slate-50 border-slate-200' : 'border-slate-300 group-hover:border-[#E21F26] group-hover:bg-red-50/50'}`}>
          {loading ? (
            <div className="flex flex-col items-center w-full">
              <Loader2 className="w-12 h-12 text-[#E21F26] animate-spin mb-6" />
              {progress.total > 0 && (
                <div className="w-full max-w-sm bg-slate-100 rounded-full h-3 mb-3 overflow-hidden shadow-inner">
                  <div 
                    className="bg-[#E21F26] h-full transition-all duration-300 rounded-full" 
                    style={{ width: `${(progress.current / progress.total) * 100}%` }}
                  />
                </div>
              )}
            </div>
          ) : (
            <Upload className="w-14 h-14 text-slate-300 group-hover:text-[#E21F26] mb-4 transition-colors group-hover:scale-110 duration-300" />
          )}
          <p className="font-bold text-slate-700 text-lg text-center">
            {loading 
              ? progress.total > 0 
                ? `Procesando: ${progress.current} / ${progress.total} registros...`
                : 'Analizando archivos...'
              : 'Selecciona o arrastra archivos Excel aquí'}
          </p>
          <p className="text-sm text-slate-400 mt-2 font-medium">Archivos compatibles: .xlsx y .xls</p>
        </div>
      </div>

      {status && (
        <motion.div
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          className={`mt-4 p-4 rounded-2xl flex items-center gap-3 ${status.type === 'success' ? 'bg-emerald-50 text-emerald-700 border border-emerald-100' : 'bg-rose-50 text-rose-700 border border-rose-100'}`}
        >
          {status.type === 'success' ? <CheckCircle2 className="w-5 h-5" /> : <AlertCircle className="w-5 h-5" />}
          <span className="text-sm font-semibold">{status.message}</span>
        </motion.div>
      )}
    </div>
  );
}
