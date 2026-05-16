import React, { useCallback, useState } from 'react';
import * as XLSX from 'xlsx';
import { db, normalizeText } from '../lib/firebase';
import { collection, writeBatch, doc } from 'firebase/firestore';
import { KardexRecord } from '../types';
import { Upload, FileSpreadsheet, Loader2, CheckCircle2, AlertCircle, ArrowLeft } from 'lucide-react';
import { motion } from 'motion/react';

interface Props {
  onBack?: () => void;
}

export default function ExcelUploader({ onBack }: Props) {
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState({ current: 0, total: 0 });
  const [status, setStatus] = useState<{ type: 'success' | 'error'; message: string } | null>(null);

  const clearDatabase = async () => {
    if (!window.confirm("¿Está seguro que desea borrar TODOS los registros de la base de datos? Esta acción no se puede deshacer.")) return;
    
    setLoading(true);
    setStatus(null);
    try {
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
      
      setStatus({ type: 'success', message: `Se borraron ${deleted} registros de la base de datos exitosamente.` });
    } catch (error) {
      console.error(error);
      setStatus({ type: 'error', message: 'Error al borrar la base de datos.' });
    } finally {
      setLoading(false);
    }
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
            mes: findKey(data[0], ['mes']),
            year: findKey(data[0], ['año', 'ano']),
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
            
            let m = '??';
            let y = '????';
            
            if (keys.mes && row[keys.mes]) {
                const rMonth = row[keys.mes];
                if (rMonth instanceof Date) m = String(rMonth.getMonth() + 1);
                else m = String(rMonth).trim().padStart(2, '0').replace(/^0+/, '');
            }
            if (keys.year && row[keys.year]) {
                const rYear = row[keys.year];
                if (rYear instanceof Date) y = String(rYear.getFullYear());
                else y = String(rYear).trim().split('-')[0];
            }
            
            let date = existing.date;
            if (m !== '??' || y !== '????') {
                date = `${m}-${y}`;
            }

            genDataMap.set(fStr, {
              date: date !== 'N/A' ? date : existing.date,
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
            
            let m = '??';
            let y = '????';
            
            // First try to use genData date if available
            if (finalDate !== 'N/A') {
               const parts = finalDate.split('-');
               if (parts.length === 2) {
                 m = parts[0];
                 y = parts[1];
               }
            }

            // Override with current row's date if available
            if (rMonth || rYear) {
              if (rMonth instanceof Date) m = String(rMonth.getMonth() + 1);
              else if (rMonth) m = String(rMonth).trim().padStart(2, '0').replace(/^0+/, '');
              
              if (rYear instanceof Date) y = String(rYear.getFullYear());
              else if (rYear) y = String(rYear).trim().split('-')[0];
            } else if (y === '????') {
              // Try to extract year from folio (e.g. DGO-DGO-24-047 -> 24 -> 2024)
              const parts = folio.trim().toUpperCase().split('-');
              if (parts.length >= 3) {
                 const yearPart = parts[2];
                 if (yearPart.length === 2 && !isNaN(Number(yearPart))) {
                   y = `20${yearPart}`;
                 }
              }
            }

            if (m !== '??' || y !== '????') {
              finalDate = `${m}-${y}`;
            }

            const searchKeywords = [
              ...normalizeText(fullName).split(/\s+/),
              fStr.toLowerCase(),
              folio.trim().toLowerCase()
            ];
            if (y !== '????') searchKeywords.push(y);
            if (m !== '??' && y !== '????') searchKeywords.push(`${m}-${y}`);

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
              searchKeywords: searchKeywords.filter(w => w.length > 0),
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
          // Create a deterministic ID to avoid duplicates
          const rawId = `${record.folio}-${record.userName}-${record.courseName}-${record.date}`;
          const safeId = normalizeText(rawId).replace(/[^a-zA-Z0-9]/g, '');
          const finalId = safeId.length > 50 ? safeId.substring(0, 50) : safeId;
          
          const docRef = doc(db, 'kardex', finalId);
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

  return (
    <div className="bg-white p-8 rounded-[2.5rem] shadow-xl border border-slate-100">
      <div className="flex flex-col md:flex-row items-center gap-4 mb-8">
        <div className="p-4 bg-red-50 rounded-2xl">
          <FileSpreadsheet className="w-8 h-8 text-[#E21F26]" />
        </div>
        <div className="flex-1 text-center md:text-left">
          <h2 className="text-2xl font-black text-slate-900">Módulo de Carga</h2>
          <p className="text-slate-500 text-sm font-medium mt-1">Sube archivos Excel (.xls, .xlsx) para sincronizar registros</p>
        </div>
        <div className="flex items-center gap-2 w-full md:w-auto">
          {onBack && (
            <button 
              onClick={onBack}
              className="flex-1 md:flex-none justify-center p-4 bg-slate-50 hover:bg-slate-100 text-slate-600 rounded-2xl transition-all flex items-center gap-2"
              title="Volver"
            >
              <ArrowLeft className="w-5 h-5" />
              <span className="text-xs font-bold uppercase hidden md:inline">Volver</span>
            </button>
          )}
          <button 
            onClick={clearDatabase}
            disabled={loading}
            className="flex-1 md:flex-none justify-center p-4 bg-orange-50 hover:bg-orange-100 text-orange-600 rounded-2xl transition-all flex items-center gap-2 disabled:opacity-50"
            title="Limpiar Base de Datos"
          >
            <AlertCircle className="w-5 h-5" />
            <span className="text-xs font-bold uppercase hidden md:inline">Limpiar DB</span>
          </button>
        </div>
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
