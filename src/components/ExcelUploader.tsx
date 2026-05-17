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

// ─── Utilidad: encuentra una clave en un objeto por lista de candidatos normalizados ───
const findKey = (obj: any, candidates: string[]): string | undefined => {
  if (!obj) return undefined;
  return Object.keys(obj).find(k => candidates.some(c => normalizeText(k).includes(c)));
};

// ─── Genera un ID seguro y único para Firestore ───
const makeDocId = (folio: string, userName: string, courseName: string, date: string): string => {
  const raw = `${folio}-${normalizeText(userName)}-${normalizeText(courseName)}-${date}`;
  const safe = raw.replace(/[^a-zA-Z0-9]/g, '').substring(0, 100);
  return safe || `unknown-${Date.now()}`;
};

// ─── Extrae año como string "YYYY" desde distintos formatos ───
const extractYear = (val: any): string => {
  if (!val) return '????';
  if (val instanceof Date) return String(val.getFullYear());
  const s = String(val).trim();
  if (/^\d{4}$/.test(s)) return s;                    // "2026"
  if (/^\d{2}$/.test(s)) return `20${s}`;             // "26"
  const m = s.match(/\b(20\d{2})\b/);
  if (m) return m[1];
  return '????';
};

// ─── Extrae mes como string sin padding (ej. "3", "11") ───
const MONTH_NAMES: Record<string, string> = {
  enero:'1', febrero:'2', marzo:'3', abril:'4', mayo:'5', junio:'6',
  julio:'7', agosto:'8', septiembre:'9', octubre:'10', noviembre:'11', diciembre:'12'
};

const extractMonth = (val: any): string => {
  if (!val) return '??';
  if (val instanceof Date) return String(val.getMonth() + 1);
  const s = normalizeText(String(val).trim());
  if (MONTH_NAMES[s]) return MONTH_NAMES[s];           // "MARZO" → "3"
  const n = parseInt(s);
  if (!isNaN(n) && n >= 1 && n <= 12) return String(n);
  return '??';
};

// ─── Extrae el año desde el folio (ej. "DGO-DGO-26-006" → "2026") ───
const yearFromFolio = (folio: string): string => {
  const parts = folio.trim().toUpperCase().split('-');
  if (parts.length >= 3) {
    const p = parts[2];
    if (/^\d{2}$/.test(p)) return `20${p}`;
    if (/^\d{4}$/.test(p)) return p;
  }
  return '????';
};

export default function ExcelUploader({ onBack }: Props) {
  const [loading, setLoading] = useState(false);
  const [loadingPhase, setLoadingPhase] = useState('');
  const [progress, setProgress] = useState({ current: 0, total: 0 });
  const [status, setStatus] = useState<{ type: 'success' | 'error'; message: string } | null>(null);
  const [diagnostics, setDiagnostics] = useState<string[]>([]);

  const onFileUpload = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;

    setLoading(true);
    setStatus(null);
    setDiagnostics([]);
    setProgress({ current: 0, total: files.length });
    setLoadingPhase('Leyendo archivos...');

    const logs: string[] = [];

    try {
      // ── PASO 1: Leer todos los archivos en memoria ──
      const workbooks: { wb: XLSX.WorkBook; name: string }[] = [];
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const buffer = await file.arrayBuffer();
        workbooks.push({
          wb: XLSX.read(buffer, { type: 'array', cellDates: true }),
          name: file.name
        });
        setProgress({ current: i + 1, total: files.length });
        await new Promise(r => setTimeout(r, 20));
      }

      // ── PASO 2: Primer pase — metadatos generales (hoja GENERAL) ──
      // Mapa: folio normalizado → { date, tipoCurso, instructor, section, periodoImparticion }
      const genDataMap = new Map<string, {
        date: string; tipoCurso: string; instructor: string; section: string; periodoImparticion: string;
      }>();

      setLoadingPhase('Analizando datos maestros...');
      setProgress({ current: 0, total: workbooks.length });

      for (let wi = 0; wi < workbooks.length; wi++) {
        const { wb, name } = workbooks[wi];
        for (const sName of wb.SheetNames) {
          const sheet = wb.Sheets[sName];
          const rawRows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
          if (rawRows.length < 2) continue;

          // ── Buscar la fila principal de encabezado (contiene "FOLIO DE INFORME") ──
          let headerRowIdx = -1;
          for (let ri = 0; ri < Math.min(rawRows.length, 10); ri++) {
            const rowNorm = rawRows[ri].map((c: any) => normalizeText(String(c)));
            if (rowNorm.some(c => c.includes('folio') && c.includes('informe'))) {
              headerRowIdx = ri;
              break;
            }
          }
          if (headerRowIdx === -1) continue;

          // ── Construir mapa combinando TODAS las filas de encabezado (1 hasta headerRowIdx) ──
          // El Excel tiene encabezados en varias filas fusionadas; combinamos todas.
          const colIdx: Record<string, number> = {};
          for (let ri = 0; ri <= headerRowIdx; ri++) {
            rawRows[ri].forEach((cell: any, idx: number) => {
              const norm = normalizeText(String(cell));
              if (norm.length > 1 && !colIdx[norm]) colIdx[norm] = idx;
            });
          }

          // ── Detectar columnas por candidatos ──
          const findCol = (candidates: string[]): number => {
            const key = Object.keys(colIdx).find(k => candidates.some(c => k.includes(c)));
            return key !== undefined ? colIdx[key] : -1;
          };

          const folioCol   = findCol(['folio']);
          if (folioCol < 0) continue;

          const anoCol     = findCol(['ano']);
          const mesCol     = findCol(['mes']);
          const tipoCol    = findCol(['linea', 'gestion', 'inclusion']);
          const periodoCol = findCol(['periodo', 'imparticion', 'dia o']);
          // Instructor: puede estar como "nombre(s)" dentro de sección de instructores
          // Buscar columna que tenga "nombre" Y esté después de la columna 100 (zona instructores)
          const instructorCol = (() => {
            const candidates = Object.keys(colIdx).filter(k =>
              (k.includes('nombre') || k.includes('primer apellido')) && colIdx[k] > 50
            );
            return candidates.length > 0 ? colIdx[candidates[0]] : -1;
          })();

          logs.push(`[${name}/${sName}] GENERAL: Folio@${folioCol}, Año@${anoCol}, Mes@${mesCol}, Tipo@${tipoCol}, Periodo@${periodoCol}, Instructor@${instructorCol}`);

          // ── Iterar filas de datos ──
          for (let ri = headerRowIdx + 1; ri < rawRows.length; ri++) {
            const row = rawRows[ri];
            const folioRaw = String(row[folioCol] || '').trim();
            if (!folioRaw || normalizeText(folioRaw).includes('folio')) continue;

            const fStr = folioRaw.toUpperCase().replace(/[^A-Z0-9]/g, '');
            const m = mesCol >= 0 ? extractMonth(row[mesCol]) : '??';
            const y = anoCol >= 0 ? extractYear(row[anoCol]) : yearFromFolio(folioRaw);
            const date = (m !== '??' || y !== '????') ? `${m}-${y}` : 'N/A';
            const tipoCurso = tipoCol >= 0 ? String(row[tipoCol] || '').trim() || 'N/A' : 'N/A';
            const periodoImparticion = periodoCol >= 0 ? String(row[periodoCol] || '').trim() : '';

            // Instructor: buscar apellidos y nombre en columnas de instructores
            let instructor = genDataMap.get(fStr)?.instructor || '';
            if (instructorCol >= 0) {
              // Tomar apellido1, apellido2, nombre desde columnas consecutivas al instructorCol
              const ap1  = String(row[instructorCol]     || '').trim();
              const ap2  = String(row[instructorCol + 1] || '').trim();
              const nom  = String(row[instructorCol + 2] || '').trim();
              const full = [nom, ap1, ap2].filter(Boolean).join(' ');
              if (full.length > 2) instructor = full;
            }

            genDataMap.set(fStr, {
              date,
              tipoCurso,
              periodoImparticion: periodoImparticion || genDataMap.get(fStr)?.periodoImparticion || '',
              instructor: instructor || 'N/A',
              section: genDataMap.get(fStr)?.section || 'N/A',
            });
          }
        }
        setProgress(prev => ({ ...prev, current: wi + 1 }));
        await new Promise(r => setTimeout(r, 20));
      }

      const sampleKeys = [...genDataMap.keys()].slice(0, 5);
      logs.push(`Folios en mapa general: ${genDataMap.size}. Ejemplos: ${sampleKeys.join(', ')}`);

      // ── PASO 3: Segundo pase — registros de alumnos (hoja SOLO PREI u otras) ──
      setLoadingPhase('Extrayendo registros de alumnos...');
      setProgress({ current: 0, total: workbooks.length });

      const totalRecords: KardexRecord[] = [];

      for (let wi = 0; wi < workbooks.length; wi++) {
        const { wb, name } = workbooks[wi];
        for (const sName of wb.SheetNames) {
          const sheet = wb.Sheets[sName];
          const rawRows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
          if (rawRows.length < 2) continue;

          // Buscar la fila con los encabezados de alumnos
          let headerRowIdx = -1;
          let headerRow: any[] = [];
          for (let ri = 0; ri < Math.min(rawRows.length, 10); ri++) {
            const row = rawRows[ri];
            const norm = row.map((c: any) => normalizeText(String(c)));
            const hasFolio   = norm.some(c => c.includes('folio') && c.includes('informe'));
            const hasNombre  = norm.some(c => c.includes('nombre') && !c.includes('curso') && !c.includes('taller'));
            if (hasFolio && hasNombre) {
              headerRowIdx = ri;
              headerRow = row;
              break;
            }
          }
          if (headerRowIdx === -1) continue;

          // Mapa columna normalizada → índice
          const colIdx: Record<string, number> = {};
          headerRow.forEach((cell: any, idx: number) => {
            const norm = normalizeText(String(cell));
            if (norm && !colIdx[norm]) colIdx[norm] = idx;
          });

          // ── Detectar columnas clave ──
          const col = (candidates: string[]): number => {
            const key = Object.keys(colIdx).find(k => candidates.some(c => k.includes(c)));
            return key !== undefined ? colIdx[key] : -1;
          };

          const cols = {
            folio:        col(['folio']),
            nombres:      col(['nombre']),       // "NOMBRE(S)" o "Nombre de Preferencia" — tomamos el primero
            pApellido:    col(['primer apellido', 'apellido paterno', 'paterno']),
            sApellido:    col(['segundo apellido', 'apellido materno', 'materno']),
            curso:        col(['nombre completo del curso', 'nombre del curso', 'curso', 'nombre curso', 'capacitacion']),
            grade:        col(['calificacion', 'promedio', 'nota']),
            curp:         col(['curp']),
            sexo:         col(['sexo', 'genero']),
            edad:         col(['edad']),
            nacimiento:   col(['fecha de nacimiento', 'nacimiento']),
            semestre:     col(['semestre', 'nivel']),
            aprobo:       col(['aprobo', 'resultado', 'estatus']),
            folioConst:   col(['folio constancia', 'num constancia']),
            interno:      col(['numero interno', 'num interno', 'interno']),
            pref:         col(['nombre de preferencia', 'preferencia', 'alias']),
            mes:          col(['mes']),
            ano:          col(['ano', 'año']),
            seccion:      col(['seccion', 'sección']),
          };

          // Si no tiene folio ni nombres, esta hoja no es de alumnos
          if (cols.folio < 0 || cols.nombres < 0) {
            logs.push(`[${name}/${sName}] Hoja sin columnas de alumnos, omitida.`);
            continue;
          }

          logs.push(`[${name}/${sName}] Hoja de alumnos detectada. Folio@${cols.folio}, Nombres@${cols.nombres}, Curso@${cols.curso}`);

          let sheetCount = 0;
          for (let ri = headerRowIdx + 1; ri < rawRows.length; ri++) {
            const row = rawRows[ri];
            const folioRaw = String(row[cols.folio] || '').trim();
            if (!folioRaw) continue;

            const fStr = folioRaw.toUpperCase().replace(/[^A-Z0-9]/g, '');
            // Folio visible: conserva guiones y formato original
            const folioVisible = folioRaw.trim().toUpperCase();
            const nombres  = cols.nombres   >= 0 ? String(row[cols.nombres]   || '').trim() : '';
            const pApe     = cols.pApellido >= 0 ? String(row[cols.pApellido] || '').trim() : '';
            const sApe     = cols.sApellido >= 0 ? String(row[cols.sApellido] || '').trim() : '';
            const fullName = `${nombres} ${pApe} ${sApe}`.replace(/\s+/g, ' ').trim();
            if (!fullName) continue;

            // Fecha: prioridad hoja > mapa general > folio
            // Lookup en mapa general: intentar con y sin guiones
            const genData = genDataMap.get(fStr) ?? genDataMap.get(folioVisible.replace(/[^A-Z0-9]/g, ''));
            let m = cols.mes >= 0 ? extractMonth(row[cols.mes]) : '??';
            let y = cols.ano >= 0 ? extractYear(row[cols.ano]) : '????';
            if (m === '??' && genData?.date && genData.date !== 'N/A') m = genData.date.split('-')[0];
            if (y === '????' && genData?.date && genData.date !== 'N/A') y = genData.date.split('-')[1];
            if (y === '????') y = yearFromFolio(folioRaw);
            const finalDate = (m !== '??' || y !== '????') ? `${m}-${y}` : 'N/A';

            const courseName = cols.curso >= 0
              ? String(row[cols.curso] || '').trim() || genData?.tipoCurso || 'N/A'
              : genData?.tipoCurso || 'N/A';

            const fechaNac = cols.nacimiento >= 0
              ? (row[cols.nacimiento] instanceof Date
                  ? row[cols.nacimiento].toLocaleDateString('es-MX')
                  : String(row[cols.nacimiento] || '').trim())
              : '';

            const searchKeywords = [
              ...normalizeText(fullName).split(/\s+/).filter(w => w.length > 1),
              fStr.toLowerCase(),                          // sin guiones: dgodgo25004
              folioVisible.toLowerCase(),                  // con guiones: dgo-dgo-25-004
              normalizeText(folioVisible),                 // normalizado
            ];
            if (y !== '????') searchKeywords.push(y);
            if (m !== '??' && y !== '????') searchKeywords.push(`${m}-${y}`);

            totalRecords.push({
              userName:         fullName,
              folio:            folioVisible,
              courseName,
              grade:            cols.grade      >= 0 ? String(row[cols.grade]      || 'N/A').trim() : 'N/A',
              section:          cols.seccion    >= 0 ? String(row[cols.seccion]    || '').trim()    : genData?.section || 'N/A',
              date:             finalDate,
              curp:             cols.curp       >= 0 ? String(row[cols.curp]       || '').toUpperCase().trim() : '',
              sexo:             cols.sexo       >= 0 ? String(row[cols.sexo]       || '').trim()    : '',
              edad:             cols.edad       >= 0 ? String(row[cols.edad]       || '').trim()    : '',
              fechaNacimiento:  fechaNac,
              semestre:         cols.semestre   >= 0 ? String(row[cols.semestre]   || '').trim()    : '',
              aprobo:           cols.aprobo     >= 0 ? String(row[cols.aprobo]     || '').trim()    : '',
              folioConstancia:  cols.folioConst >= 0 ? String(row[cols.folioConst] || '').trim()    : '',
              numInterno:       cols.interno    >= 0 ? String(row[cols.interno]    || '').trim()    : '',
              nombrePreferencia:cols.pref       >= 0 ? String(row[cols.pref]       || '').trim()    : '',
              tipoCurso:        genData?.tipoCurso || 'N/A',
              instructor:       genData?.instructor || 'N/A',
              periodoImparticion: genData?.periodoImparticion || '',
              searchKeywords:   [...new Set(searchKeywords.filter(w => w.length > 0))],
              uploadedAt:       new Date().toISOString(),
            });
            sheetCount++;

            if (ri % 500 === 0) await new Promise(r => setTimeout(r, 0));
          }

          const matched = totalRecords.slice(-sheetCount).filter(r => r.tipoCurso !== 'N/A' || r.instructor !== 'N/A' || r.periodoImparticion).length;
          logs.push(`[${name}/${sName}] ${sheetCount} registros extraídos. ${matched} con datos de hoja GENERAL (tipo/instructor/periodo).`);
          if (sheetCount > 0 && matched === 0) {
            logs.push(`⚠️ [${name}/${sName}] Ningún folio de alumnos cruzó con la hoja GENERAL. Verifica que los folios coincidan.`);
          }
        }
        setProgress(prev => ({ ...prev, current: wi + 1 }));
        await new Promise(r => setTimeout(r, 20));
      }

      const total = totalRecords.length;
      logs.push(`Total registros a subir: ${total}`);
      setDiagnostics(logs);

      if (total === 0) {
        setStatus({
          type: 'error',
          message: 'No se encontraron registros válidos. Revisa el diagnóstico abajo para ver qué columnas detectó el sistema.'
        });
        setLoading(false);
        return;
      }

      // ── PASO 4: Subida por lotes a Firestore ──
      setLoadingPhase('Subiendo registros a Firebase...');
      setProgress({ current: 0, total });

      const BATCH_SIZE = 450;
      const TIMEOUT_MS = 30_000; // 30 s por batch

      for (let i = 0; i < total; i += BATCH_SIZE) {
        const batch = writeBatch(db);
        const chunk = totalRecords.slice(i, i + BATCH_SIZE);

        chunk.forEach(record => {
          const docId = makeDocId(record.folio, record.userName, record.courseName, record.date);
          const cleanRecord = Object.fromEntries(
            Object.entries(record).map(([k, v]) => [k, v === undefined ? null : v])
          );
          batch.set(doc(db, 'kardex', docId), cleanRecord);
        });

        await Promise.race([
          batch.commit(),
          new Promise<never>((_, reject) =>
            setTimeout(() => reject(new Error(`Timeout en batch ${i / BATCH_SIZE + 1}`)), TIMEOUT_MS)
          ),
        ]);

        setProgress(prev => ({ ...prev, current: i + chunk.length }));
      }

      setStatus({
        type: 'success',
        message: `✅ Se cargaron ${total} registros de ${files.length} archivo(s) exitosamente.`,
      });
    } catch (err: any) {
      console.error(err);
      const msg = err?.message || 'Error desconocido';
      setStatus({
        type: 'error',
        message: `Error: ${msg}. Revisa el diagnóstico abajo o la consola del navegador (F12).`,
      });
      setDiagnostics(prev => [...prev, `ERROR: ${msg}`]);
    } finally {
      setLoading(false);
      e.target.value = '';
    }
  }, []);

  return (
    <div className="bg-white p-8 rounded-[2.5rem] shadow-xl border border-slate-100 space-y-6">
      {/* Header */}
      <div className="flex flex-col md:flex-row items-center gap-4">
        <div className="p-4 bg-red-50 rounded-2xl">
          <FileSpreadsheet className="w-8 h-8 text-[#E21F26]" />
        </div>
        <div className="flex-1 text-center md:text-left">
          <h2 className="text-2xl font-black text-slate-900">Módulo de Carga</h2>
          <p className="text-slate-500 text-sm font-medium mt-1">Sube archivos Excel (.xls, .xlsx) para sincronizar registros</p>
        </div>
        {onBack && (
          <button
            onClick={onBack}
            className="p-4 bg-slate-50 hover:bg-slate-100 text-slate-600 rounded-2xl transition-all flex items-center gap-2"
          >
            <ArrowLeft className="w-5 h-5" />
            <span className="text-xs font-bold uppercase hidden md:inline">Volver</span>
          </button>
        )}
      </div>

      {/* Drop zone */}
      <div className="relative group">
        <input
          type="file"
          accept=".xlsx,.xls"
          multiple
          onChange={onFileUpload}
          disabled={loading}
          className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10 disabled:cursor-not-allowed"
        />
        <div className={`border-2 border-dashed rounded-[2rem] p-12 flex flex-col items-center justify-center transition-all
          ${loading ? 'bg-slate-50 border-slate-200' : 'border-slate-300 group-hover:border-[#E21F26] group-hover:bg-red-50/50'}`}>
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
              <p className="font-bold text-slate-700 text-lg text-center">
                {loadingPhase} {progress.total > 0 ? `${progress.current} / ${progress.total}` : ''}
              </p>
            </div>
          ) : (
            <>
              <Upload className="w-14 h-14 text-slate-300 group-hover:text-[#E21F26] mb-4 transition-all group-hover:scale-110 duration-300" />
              <p className="font-bold text-slate-700 text-lg text-center">Selecciona o arrastra archivos Excel aquí</p>
              <p className="text-sm text-slate-400 mt-2 font-medium">Archivos compatibles: .xlsx y .xls</p>
            </>
          )}
        </div>
      </div>

      {/* Status */}
      {status && (
        <motion.div
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          className={`p-4 rounded-2xl flex items-start gap-3
            ${status.type === 'success'
              ? 'bg-emerald-50 text-emerald-700 border border-emerald-100'
              : 'bg-rose-50 text-rose-700 border border-rose-100'}`}
        >
          {status.type === 'success'
            ? <CheckCircle2 className="w-5 h-5 mt-0.5 shrink-0" />
            : <AlertCircle  className="w-5 h-5 mt-0.5 shrink-0" />}
          <span className="text-sm font-semibold">{status.message}</span>
        </motion.div>
      )}

      {/* Diagnóstico (visible siempre que haya logs) */}
      {diagnostics.length > 0 && (
        <details className="bg-slate-50 border border-slate-200 rounded-2xl p-4">
          <summary className="text-xs font-black text-slate-500 uppercase tracking-widest cursor-pointer select-none">
            Diagnóstico de carga ({diagnostics.length} eventos)
          </summary>
          <div className="mt-3 space-y-1 max-h-48 overflow-y-auto">
            {diagnostics.map((log, i) => (
              <p key={i} className={`text-[11px] font-mono ${log.startsWith('ERROR') ? 'text-red-600' : 'text-slate-500'}`}>
                {log}
              </p>
            ))}
          </div>
        </details>
      )}
    </div>
  );
}
