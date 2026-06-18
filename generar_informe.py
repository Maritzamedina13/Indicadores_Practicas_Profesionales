"""
Generador de Informe de Indicadores de Prácticas Profesionales ITM
Versión 2 – Diseño web moderno, filtros por facultad y programa
"""
import pandas as pd
import json
import base64
import re
from collections import Counter

# ─── Helpers ───────────────────────────────────────────────────────────────────

MESES_ES = {1:'Enero',2:'Febrero',3:'Marzo',4:'Abril',5:'Mayo',6:'Junio',
            7:'Julio',8:'Agosto',9:'Septiembre',10:'Octubre',11:'Noviembre',12:'Diciembre'}

def sem_label(d):
    if pd.isna(d): return 'Sin fecha'
    return f'{d.year}-S{"1" if d.month<=6 else "2"}'

def clean_text(s):
    if not isinstance(s, str): return '' if pd.isna(s) else str(s)
    return re.sub(r'\s+', ' ', s.strip())

# ─── Carga y Limpieza ──────────────────────────────────────────────────────────

def load_practicantes():
    df = pd.read_excel('BD PRACTICANTES.xlsx', engine='openpyxl')
    df.columns = [c.strip() for c in df.columns]
    rename = {df.columns[0]:'PROGRAMA', df.columns[1]:'FACULTAD',
              df.columns[2]:'EMPRESA_RAW', df.columns[3]:'NIT_RAW',
              df.columns[4]:'EMPRESA_NUEVA', df.columns[5]:'TIPO_CONTRATO',
              df.columns[6]:'FECHA_INICIO', df.columns[7]:'FECHA_FIN',
              df.columns[8]:'ESTADO', df.columns[9]:'ASESOR', df.columns[10]:'MODALIDAD'}
    df.rename(columns=rename, inplace=True)
    df.drop(columns=['EMPRESA_RAW','NIT_RAW'], inplace=True)
    for c in ['PROGRAMA','FACULTAD','EMPRESA_NUEVA','TIPO_CONTRATO','ESTADO','ASESOR','MODALIDAD']:
        df[c] = df[c].apply(clean_text)
    df['FECHA_INICIO'] = pd.to_datetime(df['FECHA_INICIO'], errors='coerce')
    df['FECHA_FIN']    = pd.to_datetime(df['FECHA_FIN'],    errors='coerce')
    df['FECHA_FIN_STR']= df['FECHA_FIN'].dt.strftime('%Y-%m-%d').fillna('')
    df['ANIO']      = df['FECHA_INICIO'].dt.year.fillna(0).astype(int)
    df['MES']       = df['FECHA_INICIO'].dt.month.fillna(0).astype(int)
    df['MES_LABEL'] = df['MES'].map(lambda x: MESES_ES.get(x, 'Sin mes'))
    df['SEMESTRE']  = df['FECHA_INICIO'].apply(sem_label)
    return df

def load_disponibles():
    df = pd.read_excel('BD DISPONIBLES.xlsx', engine='openpyxl')
    df.columns = [c.strip() for c in df.columns]
    rename = {df.columns[0]:'MODALIDAD', df.columns[1]:'FECHA_SOLICITUD',
              df.columns[2]:'PROGRAMA', df.columns[3]:'FACULTAD',
              df.columns[4]:'ESTADO', df.columns[5]:'DISCAPACIDAD',
              df.columns[6]:'TIPO_DISCAPACIDAD', df.columns[7]:'UEN_REGION',
              df.columns[8]:'CREADO'}
    df.rename(columns=rename, inplace=True)
    for c in ['MODALIDAD','PROGRAMA','FACULTAD','ESTADO','DISCAPACIDAD','TIPO_DISCAPACIDAD']:
        df[c] = df[c].apply(clean_text)
    df['FECHA_SOLICITUD'] = pd.to_datetime(df['FECHA_SOLICITUD'], errors='coerce')
    df['ANIO']      = df['FECHA_SOLICITUD'].dt.year.fillna(0).astype(int)
    df['MES']       = df['FECHA_SOLICITUD'].dt.month.fillna(0).astype(int)
    df['MES_LABEL'] = df['MES'].map(lambda x: MESES_ES.get(x, 'Sin mes'))
    df['SEMESTRE']  = df['FECHA_SOLICITUD'].apply(sem_label)
    return df

def load_f082():
    df = pd.read_excel('BD F082 TRABAJOS ENTREGADOS.xlsx', sheet_name=0, engine='openpyxl')
    df.columns = [c.strip() for c in df.columns]
    df = df.iloc[1:].reset_index(drop=True)
    cols = list(df.columns)
    rename = {cols[0]:'MODALIDAD', cols[1]:'FECHA_INICIO', cols[2]:'FECHA_TERMINA',
              cols[3]:'EMPRESA_RAW', cols[4]:'CARGO', cols[5]:'RAZON_RAW',
              cols[6]:'ASESOR', cols[7]:'ACTIVIDADES', cols[8]:'DESCRIPCION',
              cols[9]:'PROMEDIO', cols[10]:'VINCULADO', cols[11]:'PROGRAMA',
              cols[12]:'FECHA_ENTREGA', cols[13]:'FACULTAD', cols[14]:'SISTEMATIZACION_RAW',
              cols[15]:'ES_OPTIMA', cols[16]:'POR_QUE_NO_OPTIMA', cols[17]:'ENTREGADO'}
    df.rename(columns=rename, inplace=True)
    df.drop(columns=['RAZON_RAW','SISTEMATIZACION_RAW'], inplace=True)
    df['EMPRESA'] = df['EMPRESA_RAW'].apply(clean_text)
    df.drop(columns=['EMPRESA_RAW'], inplace=True)
    for c in ['MODALIDAD','ASESOR','ACTIVIDADES','DESCRIPCION','VINCULADO','PROGRAMA','FACULTAD','ES_OPTIMA','ENTREGADO']:
        df[c] = df[c].apply(clean_text)
    df['FECHA_ENTREGA'] = pd.to_datetime(df['FECHA_ENTREGA'], errors='coerce')
    df['FECHA_INICIO']  = pd.to_datetime(df['FECHA_INICIO'], errors='coerce')
    df['PROMEDIO'] = pd.to_numeric(df['PROMEDIO'], errors='coerce')
    df['ANIO']      = df['FECHA_ENTREGA'].dt.year.fillna(0).astype(int)
    df['MES']       = df['FECHA_ENTREGA'].dt.month.fillna(0).astype(int)
    df['MES_LABEL'] = df['MES'].map(lambda x: MESES_ES.get(x, 'Sin mes'))
    df['SEMESTRE']  = df['FECHA_ENTREGA'].apply(sem_label)
    return df

def load_solicitud():
    df = pd.read_excel('Solicitud Empresas.xlsx', sheet_name=0, engine='openpyxl')
    df.columns = [c.strip() for c in df.columns]
    cols = list(df.columns)
    rename = {cols[0]:'EMPRESA', cols[1]:'HORA_INICIO', cols[2]:'NIT_RAW',
              cols[3]:'EMPRESA_NUEVA', cols[4]:'PROGRAMA_NUEVO',
              cols[5]:'PROGRAMA', cols[6]:'MODALIDAD', cols[7]:'FACULTAD'}
    df.rename(columns=rename, inplace=True)
    df.drop(columns=['NIT_RAW'], inplace=True)
    for c in ['EMPRESA','EMPRESA_NUEVA','PROGRAMA_NUEVO','PROGRAMA','MODALIDAD','FACULTAD']:
        df[c] = df[c].apply(clean_text)
    # Normalizar nombres de facultad al estándar del dashboard
    fac_norm = {
        'Facultad Ciencias Economicas Y Administrativas Itm': 'CIENCIAS ECONOMICAS Y ADMINISTRATIVAS',
        'Facultad De Ingenierias Itm': 'FACULTAD DE INGENIERIAS',
        'Facultad Ciencias Exactas Y Aplicadas Itm': 'CIENCIAS EXACTAS Y APLICADAS',
        'Facultad Artes Y Humanidades Itm': 'ARTES Y HUMANIDADES',
    }
    df['FACULTAD'] = df['FACULTAD'].replace(fac_norm)
    df['FACULTAD'] = df['FACULTAD'].apply(
        lambda x: x.upper().strip() if isinstance(x, str) and x.strip() else x)
    # Rellenar FACULTAD vacía usando el PROGRAMA como clave
    prog_fac = (df[df['FACULTAD'].notna() & (df['FACULTAD'].str.strip()!='')]
                .groupby('PROGRAMA')['FACULTAD']
                .agg(lambda x: x.mode().iloc[0] if len(x)>0 else None)
                .to_dict())
    df['FACULTAD'] = df.apply(
        lambda r: prog_fac.get(r['PROGRAMA'], r['FACULTAD'])
                  if (pd.isna(r['FACULTAD']) or str(r['FACULTAD']).strip()=='')
                  else r['FACULTAD'], axis=1)
    df['HORA_INICIO'] = pd.to_datetime(df['HORA_INICIO'], errors='coerce')
    df['ANIO']      = df['HORA_INICIO'].dt.year.fillna(0).astype(int)
    df['MES']       = df['HORA_INICIO'].dt.month.fillna(0).astype(int)
    df['MES_LABEL'] = df['MES'].map(lambda x: MESES_ES.get(x, 'Sin mes'))
    df['SEMESTRE']  = df['HORA_INICIO'].apply(sem_label)
    return df

def load_encuesta_est():
    fname = 'Copia de Encuesta satisfacción del estudiante con su práctica profesional (1-1934).xlsx'
    df = pd.read_excel(fname, engine='openpyxl')
    df.columns = [c.strip() for c in df.columns]
    cols = list(df.columns)
    rename = {
        cols[0]:'ID',     cols[1]:'FECHA',      cols[2]:'EMAIL',
        cols[3]:'PROGRAMA', cols[4]:'EMPRESA',   cols[5]:'NIT_RAW',
        cols[6]:'MODALIDAD', cols[7]:'PROYECTO_FUTURO',
        cols[8]:'CALIF_FORMACION',    cols[9]:'CALIF_ORIENTACION',
        cols[10]:'CALIF_PERTINENCIA', cols[11]:'CALIF_TUTOR',
        cols[12]:'CALIF_DESEMPENO',   cols[13]:'CALIF_PUESTO',
        cols[14]:'CALIF_ASESOR_ITM',  cols[15]:'ASESOR_NOMBRE',
        cols[16]:'CALIF_SEGUIMIENTO', cols[17]:'CALIF_INFORMACION',
        cols[18]:'RECOMIENDA',        cols[19]:'SATISFACCION_GENERAL',
    }
    df.rename(columns=rename, inplace=True)
    df.drop(columns=['ID','EMAIL','NIT_RAW'], inplace=True)
    calif_cols = ['CALIF_FORMACION','CALIF_ORIENTACION','CALIF_PERTINENCIA',
                  'CALIF_TUTOR','CALIF_DESEMPENO','CALIF_PUESTO',
                  'CALIF_ASESOR_ITM','CALIF_SEGUIMIENTO','CALIF_INFORMACION',
                  'SATISFACCION_GENERAL']
    for c in calif_cols:
        df[c] = df[c].apply(clean_text).str.strip().str.capitalize()
    for c in ['PROGRAMA','EMPRESA','MODALIDAD','PROYECTO_FUTURO','ASESOR_NOMBRE']:
        df[c] = df[c].apply(clean_text)
    df['PROGRAMA'] = df['PROGRAMA'].str.upper().apply(clean_text)
    df['ASESOR_NOMBRE'] = df['ASESOR_NOMBRE'].str.strip()
    df['RECOMIENDA'] = df['RECOMIENDA'].apply(clean_text).str.upper().str.strip()
    df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
    df['ANIO']     = df['FECHA'].dt.year.fillna(0).astype(int)
    df['MES']      = df['FECHA'].dt.month.fillna(0).astype(int)
    df['MES_LABEL']= df['MES'].map(lambda x: MESES_ES.get(x,'Sin mes'))
    df['SEMESTRE'] = df['FECHA'].apply(sem_label)
    # Score numérico para promedios
    score_map = {'Excelente':3,'Bueno':2,'Deficiente':1}
    df['SCORE_GLOBAL'] = df[calif_cols].map(lambda v: score_map.get(v,0)).replace(0,pd.NA).mean(axis=1)
    return df

def load_encuesta():
    df = pd.read_csv('EncuestaSatisfaccionEmpresarios.csv', encoding='latin-1', sep=';', on_bad_lines='skip')
    df.columns = [c.strip() for c in df.columns]
    cols = list(df.columns)
    # Renombrar columnas con tildes/ñ por posición para evitar problemas de encoding
    rename = {
        cols[10]: 'IMPACTO_SOCIAL',
        cols[11]: 'CALIF_FORMACION_EG',
        cols[12]: 'CALIF_DESEMPENO_EG',
        cols[13]: 'DEBILIDADES_EG',
        cols[14]: 'FORTALEZAS_EG',
        cols[15]: 'CALIF_FORMACION_PRAC',
        cols[16]: 'CALIF_DESEMPENO_PRAC',
        cols[17]: 'FORTALEZAS_PRAC',
        cols[18]: 'DEBILIDADES_PRAC',
        cols[19]: 'COMPETENCIAS',
        cols[20]: 'TENDENCIAS',
    }
    df.rename(columns=rename, inplace=True)
    for c in ['empresa','sector','mercado','tipo','programa','personal',
              'IMPACTO_SOCIAL','CALIF_FORMACION_EG','CALIF_DESEMPENO_EG',
              'DEBILIDADES_EG','FORTALEZAS_EG','CALIF_FORMACION_PRAC',
              'CALIF_DESEMPENO_PRAC','FORTALEZAS_PRAC','DEBILIDADES_PRAC',
              'COMPETENCIAS','vinculacionpracticantes']:
        if c in df.columns:
            df[c] = df[c].apply(clean_text)
    df['tipo'] = df['tipo'].str.strip().str.upper()
    df['programa'] = df['programa'].str.upper().apply(clean_text)
    df['vinculacionpracticantes'] = df['vinculacionpracticantes'].str.upper().str.strip()
    df['fechadiligenciamiento2'] = pd.to_datetime(df['fechadiligenciamiento2'], dayfirst=True, errors='coerce')
    df['ANIO']      = df['fechadiligenciamiento2'].dt.year.fillna(0).astype(int)
    df['MES']       = df['fechadiligenciamiento2'].dt.month.fillna(0).astype(int)
    df['MES_LABEL'] = df['MES'].map(lambda x: MESES_ES.get(x, 'Sin mes'))
    df['SEMESTRE']  = df['fechadiligenciamiento2'].apply(sem_label)
    return df

def load_aprobacion():
    df = pd.read_excel('aprobación de funciones.xlsx', sheet_name=0, engine='openpyxl')
    df.columns = [c.strip() for c in df.columns]
    df = df.iloc[1:].reset_index(drop=True)
    cols = list(df.columns)
    rename = {cols[0]:'PROGRAMA', cols[1]:'FUNCIONES', cols[2]:'EMPRESA',
              cols[3]:'NIT_RAW', cols[4]:'ESTADO_APROBACION',
              cols[5]:'APROBADOR', cols[6]:'COMENTARIOS', cols[7]:'CREADO'}
    df.rename(columns=rename, inplace=True)
    df.drop(columns=['NIT_RAW'], inplace=True)
    for c in ['PROGRAMA','FUNCIONES','EMPRESA','ESTADO_APROBACION','APROBADOR']:
        df[c] = df[c].apply(clean_text)
    df['CREADO'] = pd.to_datetime(df['CREADO'], errors='coerce')
    df['ANIO']      = df['CREADO'].dt.year.fillna(0).astype(int)
    df['MES']       = df['CREADO'].dt.month.fillna(0).astype(int)
    df['MES_LABEL'] = df['MES'].map(lambda x: MESES_ES.get(x, 'Sin mes'))
    df['SEMESTRE']  = df['CREADO'].apply(sem_label)
    return df

# ─── Utilidades de análisis ────────────────────────────────────────────────────

SW = {
    'de','la','el','en','y','a','los','las','del','se','por','con','para',
    'que','un','una','al','es','su','o','e','no','le','lo','sus','entre',
    'como','más','este','esta','son','han','ha','ser','fue','si','esto',
    'pero','también','sobre','ya','todo','bien','cada','así','nan','none',
    'the','of','and','to','in','is','it','for','on','with','its','are',
    'this','that','which','were','been','have','has','from','they','their',
    'una','unos','unas','estos','estas','ello','aquí','allí','cuando',
    'donde','cómo','qué','quién','cuál','tal','sino','cuales','debe',
    'dicho','dicha','mismo','misma','bajo','alto','gran','buen','cual',
    'esta','este','estos','estas','siendo','hacer','tener','poder',
    'área','nivel','través','parte','tipo','forma','manera','medio',
    'cargo','empresa','práctica','practicante','practicantes','periodo',
    'tiempo','trabajo','trabajos','proceso','procesos','actividad',
    'actividades','función','funciones','entre','para','hacia',
}

def tokenize(text):
    """Tokeniza texto en palabras limpias."""
    if not isinstance(text, str) or text.lower().strip() in ('nan','none',''): return []
    return [w for w in re.findall(r'\b[a-záéíóúñü]{4,}\b', text.lower()) if w not in SW]

def kw_extract(texts, stopwords_extra=None, n=15):
    sw = SW.copy()
    if stopwords_extra: sw |= set(stopwords_extra)
    words = []
    for t in texts:
        words.extend(tokenize(t))
    return Counter(words).most_common(n)

def areas_extract(texts, top=10):
    """
    Extrae top-10 áreas de desempeño combinando:
    - Bigramas significativos (frases de 2 palabras = área real)
    - Unigramas de respaldo
    Devuelve lista de {'area': str, 'count': int}.
    """
    unigrams  = Counter()
    bigrams   = Counter()

    for t in texts:
        tokens = tokenize(t)
        for w in tokens:
            unigrams[w] += 1
        for a, b in zip(tokens, tokens[1:]):
            bigrams[f'{a} {b}'] += 1

    # Puntuar bigramas: deben aparecer ≥2 veces
    scored = {}
    for bg, c in bigrams.items():
        if c >= 2:
            scored[bg] = c * 1.6   # peso mayor a bigramas
    # Añadir unigramas que no estén ya cubiertos por un bigrama
    covered = set(w for bg in scored for w in bg.split())
    for ug, c in unigrams.items():
        if ug not in covered:
            scored[ug] = c

    top_areas = sorted(scored.items(), key=lambda x: -x[1])[:top]
    return [{'area': a.title(), 'count': round(c)} for a, c in top_areas]

def activities_by_program(df, prog_col='PROGRAMA', act_col='ACTIVIDADES', top_progs=8):
    top = df[prog_col].value_counts().head(top_progs).index.tolist()
    result = {}
    for p in top:
        texts = df[df[prog_col]==p][act_col].dropna().tolist()
        kws = kw_extract(texts)
        result[p] = [{'word':w,'count':c} for w,c in kws]
    return result

def raw_records(df, cols):
    sub = df[[c for c in cols if c in df.columns]].copy()
    recs = []
    for row in sub.to_dict(orient='records'):
        r2 = {}
        for k, v in row.items():
            if hasattr(v, 'strftime'):
                try:    r2[k] = v.strftime('%Y-%m-%d')
                except: r2[k] = None
            elif isinstance(v, float) and v != v:
                r2[k] = None
            elif isinstance(v, (int, float)):
                r2[k] = v
            else:
                r2[k] = str(v).strip() if v is not None else None
        recs.append(r2)
    return recs

# ─── Build master data dict ────────────────────────────────────────────────────

def build_data(df1, df2, df3, df4, df5, df6, df7):
    data = {}

    # Facultad→Programas mapping (union de todas las BDs)
    fac_prog = {}
    for df, prog_col, fac_col in [
        (df1,'PROGRAMA','FACULTAD'), (df2,'PROGRAMA','FACULTAD'),
        (df3,'PROGRAMA','FACULTAD'), (df5,'PROGRAMA',None)
    ]:
        for _, row in df[[prog_col, fac_col] if fac_col else [prog_col]].drop_duplicates().iterrows():
            p = clean_text(str(row[prog_col])) if pd.notna(row[prog_col]) else ''
            f = clean_text(str(row[fac_col])) if fac_col and pd.notna(row[fac_col]) else 'SIN FACULTAD'
            if p and len(p) > 2:
                fac_prog.setdefault(f, set()).add(p)

    # también de solicitud (perfil/programa)
    for _, row in df4[['PROGRAMA','FACULTAD']].drop_duplicates().iterrows():
        p = clean_text(str(row['PROGRAMA'])) if pd.notna(row['PROGRAMA']) else ''
        f = clean_text(str(row['FACULTAD'])) if pd.notna(row['FACULTAD']) else ''
        if p and len(p) > 2:
            fac_prog.setdefault(f if f else 'SIN FACULTAD', set()).add(p)

    data['fac_prog'] = {k: sorted(v) for k, v in fac_prog.items() if k != 'SIN FACULTAD'}
    data['all_programas'] = sorted(set(p for v in fac_prog.values() for p in v))
    data['all_facultades'] = sorted([k for k in fac_prog.keys() if k != 'SIN FACULTAD'])



    # Años y meses disponibles
    all_anios = sorted(set(
        df1[df1['ANIO']>0]['ANIO'].tolist() + df2[df2['ANIO']>0]['ANIO'].tolist() +
        df3[df3['ANIO']>0]['ANIO'].tolist() + df4[df4['ANIO']>0]['ANIO'].tolist() +
        df5[df5['ANIO']>0]['ANIO'].tolist()
    ))
    data['filtros_anios']  = [int(x) for x in all_anios]
    data['filtros_meses']  = list(MESES_ES.items())

    # Semestres disponibles
    sems = set()
    for df in [df1, df2, df3, df4, df5]:
        sems |= set(df[df['SEMESTRE']!='Sin fecha']['SEMESTRE'].tolist())
    data['filtros_semestres'] = sorted(sems)

    # ── Raw rows para filtrado client-side ──────────────────────────────────────
    data['raw_practicantes'] = raw_records(df1,
        ['PROGRAMA','FACULTAD','EMPRESA_NUEVA','TIPO_CONTRATO','ESTADO','ASESOR','MODALIDAD','FECHA_FIN_STR','ANIO','MES','MES_LABEL','SEMESTRE'])
    data['raw_disponibles']  = raw_records(df2,
        ['MODALIDAD','PROGRAMA','FACULTAD','ESTADO','DISCAPACIDAD','TIPO_DISCAPACIDAD','ANIO','MES','MES_LABEL','SEMESTRE'])
    data['raw_f082']         = raw_records(df3,
        ['MODALIDAD','ASESOR','ACTIVIDADES','PROMEDIO','VINCULADO','EMPRESA','PROGRAMA','FACULTAD','ES_OPTIMA','ENTREGADO','ANIO','MES','MES_LABEL','SEMESTRE'])
    data['raw_solicitud']    = raw_records(df4,
        ['EMPRESA','EMPRESA_NUEVA','PROGRAMA_NUEVO','PROGRAMA','FACULTAD','MODALIDAD','ANIO','MES','MES_LABEL','SEMESTRE'])
    data['raw_aprobacion']   = raw_records(df5,
        ['PROGRAMA','FUNCIONES','EMPRESA','ESTADO_APROBACION','APROBADOR','ANIO','MES','MES_LABEL','SEMESTRE'])

    # ── Análisis combinado Solicitud vs Aprobación por año ────────────────────
    sol_anio = df4[df4['ANIO']>0].groupby('ANIO').size().to_dict()
    apr_anio = df5[df5['ANIO']>0].groupby('ANIO').size().to_dict()
    all_years = sorted(set(list(sol_anio.keys()) + list(apr_anio.keys())))
    combined = []
    for y in all_years:
        sol = int(sol_anio.get(y, 0))
        apr = int(apr_anio.get(y, 0))
        pct = round(apr / sol * 100, 1) if sol > 0 else 0
        combined.append({'anio': y, 'solicitudes': sol, 'aprobaciones': apr, 'pct': pct})
    data['sol_vs_aprob'] = combined

    # ── F082: áreas de desempeño por PROGRAMA y por FACULTAD ─────────────────
    all_progs_f082 = sorted(df3['PROGRAMA'].dropna().apply(clean_text).unique().tolist())
    all_facs_f082  = sorted(df3['FACULTAD'].dropna().apply(clean_text).unique().tolist())

    # Mapeo facultad → programas (para el dropdown encadenado)
    fac_to_prog_f082 = {}
    for fac in all_facs_f082:
        progs = sorted(df3[df3['FACULTAD']==fac]['PROGRAMA'].dropna().apply(clean_text).unique().tolist())
        fac_to_prog_f082[fac] = progs

    # Top-10 áreas por PROGRAMA (bigrams + unigrams)
    areas_por_prog = {}
    for p in all_progs_f082:
        mask  = df3['PROGRAMA'].apply(clean_text) == p
        texts = (df3[mask]['ACTIVIDADES'].dropna().tolist() +
                 df3[mask]['DESCRIPCION'].dropna().tolist())
        areas_por_prog[p] = areas_extract(texts, top=20)

    # Top-20 áreas por FACULTAD
    areas_por_fac = {}
    for f in all_facs_f082:
        mask  = df3['FACULTAD'].apply(clean_text) == f
        texts = (df3[mask]['ACTIVIDADES'].dropna().tolist() +
                 df3[mask]['DESCRIPCION'].dropna().tolist())
        areas_por_fac[f] = areas_extract(texts, top=20)

    data['f_areas_prog']      = areas_por_prog
    data['f_areas_fac']       = areas_por_fac
    data['f_programas_list']  = all_progs_f082
    data['f_facultades_list'] = all_facs_f082
    data['f_fac_to_prog']     = fac_to_prog_f082

    # Nube de palabras (unigramas) por programa — para visualización complementaria
    f_wc_all = {}
    for p in all_progs_f082:
        mask  = df3['PROGRAMA'].apply(clean_text) == p
        texts = (df3[mask]['ACTIVIDADES'].dropna().tolist() +
                 df3[mask]['DESCRIPCION'].dropna().tolist())
        kws = kw_extract(texts, n=20)
        f_wc_all[p] = [{'word': w, 'count': c} for w, c in kws]
    data['f_actividades_programa'] = f_wc_all

    # Todos los programas con funciones (sin límite de top 8)
    all_progs_aprob = sorted(df5['PROGRAMA'].dropna().apply(clean_text).unique().tolist())
    a_func_prog = {}
    for p in all_progs_aprob:
        if not p: continue
        texts = df5[df5['PROGRAMA'].apply(clean_text)==p]['FUNCIONES'].dropna().tolist()
        kws = kw_extract(texts, n=20)
        a_func_prog[p] = [{'word':w,'count':c} for w,c in kws]
    data['a_funciones_programa'] = a_func_prog

    # Palabras clave por facultad (usando fac_prog del mapa global)
    a_func_fac = {}
    for fac, progs in fac_prog.items():
        if fac == 'SIN FACULTAD': continue
        progs_clean = [clean_text(p) for p in progs]
        mask = df5['PROGRAMA'].apply(clean_text).isin(progs_clean)
        texts = df5[mask]['FUNCIONES'].dropna().tolist()
        if texts:
            kws = kw_extract(texts, n=20)
            a_func_fac[fac] = [{'word':w,'count':c} for w,c in kws]
    data['a_funciones_fac']      = a_func_fac

    data['a_funciones_kw']       = [{'word':w,'count':c} for w,c in kw_extract(df5['FUNCIONES'].tolist(), n=20)]

    # ── Encuesta Satisfacción Empresarios ─────────────────────────────────────
    def _clean_label(p):
        import re
        p = re.sub(r'^[A-Za-z]\.\s+', '', p)
        return p.capitalize() if p else p

    def _count_multi(series, top=12):
        c = Counter()
        for v in series.dropna():
            for p in str(v).split(';'):
                p = _clean_label(p.strip())
                if p and p.lower() not in ('nan', 'none', ''):
                    c[p] += 1
        return [{'label': k, 'count': v} for k, v in c.most_common(top)]

    def _count_calif(series):
        order = ['EXCELENTE', 'BUENO', 'REGULAR', 'DEFICIENTE']
        s = series.dropna().str.upper().str.strip()
        return {o: int((s == o).sum()) for o in order if int((s == o).sum()) > 0}

    # Normalizar programa en enc (puede venir en minúsculas o mixto) para que coincida con all_programas
    if 'programa' in df6.columns:
        df6['PROGRAMA'] = df6['programa'].str.upper().str.strip()
    data['enc_raw'] = raw_records(df6,
        ['empresa', 'sector', 'mercado', 'tipo', 'PROGRAMA',
         'IMPACTO_SOCIAL', 'CALIF_FORMACION_EG', 'CALIF_DESEMPENO_EG',
         'CALIF_FORMACION_PRAC', 'CALIF_DESEMPENO_PRAC',
         'FORTALEZAS_PRAC', 'DEBILIDADES_PRAC',
         'FORTALEZAS_EG', 'DEBILIDADES_EG',
         'COMPETENCIAS', 'vinculacionpracticantes',
         'ANIO', 'MES', 'MES_LABEL', 'SEMESTRE'])

    data['enc_calif_impacto']       = _count_calif(df6['IMPACTO_SOCIAL'])
    data['enc_calif_formacion_eg']  = _count_calif(df6['CALIF_FORMACION_EG'])
    data['enc_calif_desempeno_eg']  = _count_calif(df6['CALIF_DESEMPENO_EG'])
    data['enc_calif_formacion_prac']= _count_calif(df6['CALIF_FORMACION_PRAC'])
    data['enc_calif_desempeno_prac']= _count_calif(df6['CALIF_DESEMPENO_PRAC'])

    data['enc_fortalezas_prac']  = _count_multi(df6['FORTALEZAS_PRAC'])
    data['enc_debilidades_prac'] = _count_multi(df6['DEBILIDADES_PRAC'])
    data['enc_fortalezas_eg']    = _count_multi(df6['FORTALEZAS_EG'])
    data['enc_debilidades_eg']   = _count_multi(df6['DEBILIDADES_EG'])
    data['enc_competencias']     = _count_multi(df6['COMPETENCIAS'])

    # ── Encuesta Satisfacción Estudiantes ─────────────────────────────────────
    EST_CALIF = ['CALIF_FORMACION','CALIF_ORIENTACION','CALIF_PERTINENCIA',
                 'CALIF_TUTOR','CALIF_DESEMPENO','CALIF_PUESTO',
                 'CALIF_ASESOR_ITM','CALIF_SEGUIMIENTO','CALIF_INFORMACION',
                 'SATISFACCION_GENERAL']
    EST_SCORE_MAP = {'Excelente':3,'Bueno':2,'Deficiente':1}

    if 'PROGRAMA' in df7.columns:
        df7['PROGRAMA'] = df7['PROGRAMA'].str.upper().str.strip()
    data['est_raw'] = raw_records(df7,
        ['PROGRAMA','EMPRESA','MODALIDAD','PROYECTO_FUTURO','ASESOR_NOMBRE',
         'RECOMIENDA','SATISFACCION_GENERAL'] + EST_CALIF +
        ['ANIO','MES','MES_LABEL','SEMESTRE'])

    # Score promedio por programa
    prog_scores = {}
    for p in df7['PROGRAMA'].dropna().unique():
        p = clean_text(str(p))
        if not p: continue
        mask = df7['PROGRAMA'].apply(clean_text) == p
        vals = df7[mask][EST_CALIF].map(lambda v: EST_SCORE_MAP.get(str(v).strip().capitalize(),0))
        vals = vals.replace(0, pd.NA)
        avg = float(vals.stack().mean()) if not vals.stack().empty else 0
        n = int(mask.sum())
        prog_scores[p] = {'score': round(avg,2), 'n': n}
    data['est_prog_scores'] = dict(sorted(prog_scores.items(), key=lambda x:-x[1]['score']))

    # Score promedio por asesor
    asesor_scores = {}
    for a in df7['ASESOR_NOMBRE'].dropna().unique():
        a = clean_text(str(a))
        if not a or len(a) < 4: continue
        mask = df7['ASESOR_NOMBRE'].apply(clean_text) == a
        vals = df7[mask][EST_CALIF].map(lambda v: EST_SCORE_MAP.get(str(v).strip().capitalize(),0))
        vals = vals.replace(0, pd.NA)
        avg = float(vals.stack().mean()) if not vals.stack().empty else 0
        n = int(mask.sum())
        asesor_scores[a] = {'score': round(avg,2), 'n': n}
    data['est_asesor_scores'] = dict(sorted(asesor_scores.items(), key=lambda x:-x[1]['score']))

    return data

# ─── Logo ──────────────────────────────────────────────────────────────────────

def build_cuee_carousel():
    """Genera el HTML del carrusel con las fotos del CUEE embebidas en base64."""
    fotos = [
        ('DOCUMENTACION/fotos_cuee/word/media/image1.jpeg',  'Convocatoria Pasantías CUEE 2026',    'Campaña oficial · ¡La experiencia laboral ya comenzó!'),
        ('DOCUMENTACION/fotos_cuee/word/media/image2.jpeg',  'Inducción de Estudiantes',             'Sesión de inducción previa a la asignación de empresas'),
        ('DOCUMENTACION/fotos_cuee/word/media/image4.jpeg',  'Foto Grupal · Empresa Aliada',         'Estudiantes ITM en instalaciones de empresa del sector productivo'),
        ('DOCUMENTACION/fotos_cuee/word/media/image6.jpeg',  'Metro de Medellín',                    'Visita a instalaciones del Metro · #MiMetroMeMueve'),
        ('DOCUMENTACION/fotos_cuee/word/media/image8.jpeg',  'Visita a Planta de Producción',        'Estudiantes equipados para recorrido por planta de producción'),
        ('DOCUMENTACION/fotos_cuee/word/media/image10.jpeg', 'Sesión Empresarial · Día 1',           'Charla y presentación en empresa aliada · Línea Directa'),
        ('DOCUMENTACION/fotos_cuee/word/media/image12.jpeg', 'Colcafé · Espacio de Bienestar',       'Estudiantes ITM en instalaciones de Colcafé'),
        ('DOCUMENTACION/fotos_cuee/word/media/image14.jpeg', 'Protección S.A.',                      'Estudiantes en instalaciones de Protección · Grupo SURA'),
        ('DOCUMENTACION/fotos_cuee/word/media/image16.jpeg', 'Noel · Planta de Producción',          'Estudiantes con trajes de bioseguridad en Compañía de Galletas Noel'),
    ]
    slides_html = ''
    thumbs_html = ''
    loaded = 0
    for i, (path, titulo, desc) in enumerate(fotos):
        try:
            with open(path, 'rb') as f:
                b64 = base64.b64encode(f.read()).decode()
            src = f'data:image/jpeg;base64,{b64}'
            active = 'active' if i == 0 else ''
            slides_html += f'''
      <div class="cuee-slide {active}" data-index="{i}">
        <img src="{src}" alt="{titulo}" loading="lazy">
        <div class="cuee-caption">
          <div class="cuee-caption-title">{titulo}</div>
          <div class="cuee-caption-desc">{desc}</div>
        </div>
      </div>'''
            thumbs_html += f'''
        <div class="cuee-thumb {active}" data-index="{i}" onclick="cueeGoTo({i})">
          <img src="{src}" alt="{titulo}">
        </div>'''
            loaded += 1
        except Exception:
            pass
    if loaded == 0:
        return ''
    return f'''
<div class="card full" id="cuee-carousel-card">
  <div class="card-head">
    <h3>Galería Fotográfica · Pasantías CUEE 2026-1</h3>
    <span class="card-badge">{loaded} fotos · Mayo 2026</span>
  </div>
  <div class="card-body" style="padding-bottom:20px">
    <div class="sub-grid-1-1" style="display:grid;grid-template-columns:3fr 2fr;gap:20px;align-items:start">
      <!-- Carrusel -->
      <div>
        <div class="cuee-carousel">
          <div class="cuee-track">{slides_html}
          </div>
          <button class="cuee-btn cuee-prev" onclick="cueeMove(-1)" aria-label="Anterior">&#8249;</button>
          <button class="cuee-btn cuee-next" onclick="cueeMove(1)" aria-label="Siguiente">&#8250;</button>
          <div class="cuee-dots" id="cuee-dots"></div>
        </div>
        <div class="cuee-thumbs" id="cuee-thumbs">{thumbs_html}
        </div>
      </div>
      <!-- Empresas Aliadas -->
      <div>
        <div style="font-size:.72rem;font-weight:700;color:var(--text2,#475569);text-transform:uppercase;letter-spacing:.06em;margin-bottom:12px;padding-bottom:8px;border-bottom:2px solid var(--border,#e2e8f0)">
          🏢 Empresas Aliadas · Sector productivo Antioquia
        </div>
        <div id="cuee-empresas-grid" style="display:grid;grid-template-columns:1fr;gap:9px"></div>
      </div>
    </div>
  </div>
</div>'''

def get_logo():
    for path in ['DOCUMENTACION/LOGO ITM 2020-02.png', 'DOCUMENTACION/LOGO ITM 2020-03.png']:
        try:
            with open(path,'rb') as f:
                return 'data:image/png;base64,' + base64.b64encode(f.read()).decode()
        except: pass
    return ''

# ─── HTML ──────────────────────────────────────────────────────────────────────

HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Prácticas Profesionales ITM – Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
<style>
/* ── Variables ITM ─────────────────────────────── */
:root {
  --itm-blue:    #00539B;
  --itm-blue2:   #003d73;
  --itm-blue3:   #1a6eb5;
  --itm-gold:    #E8A000;
  --itm-gold2:   #b87e00;
  --itm-gold3:   #ffd060;
  --bg:          #EEF2F8;
  --surface:     #ffffff;
  --surface2:    #f8fafd;
  --border:      #dde4ee;
  --text:        #1a2540;
  --text2:       #4b5e7e;
  --text3:       #8fa3bf;
  --green:       #10b981;
  --red:         #ef4444;
  --purple:      #8b5cf6;
  --radius:      14px;
  --radius-sm:   8px;
  --shadow:      0 2px 16px rgba(0,83,155,.10);
  --shadow-md:   0 4px 28px rgba(0,83,155,.14);
  --transition:  all .22s cubic-bezier(.4,0,.2,1);
}
* { box-sizing:border-box; margin:0; padding:0 }
html { scroll-behavior:smooth }
body {
  font-family: 'Segoe UI', system-ui, -apple-system, Arial, sans-serif;
  background: var(--bg);
  color: var(--text);
  min-height: 100vh;
  font-size: 14px;
}

/* ── HEADER ───────────────────────────────────── */
header {
  background: linear-gradient(118deg, var(--itm-blue2) 0%, var(--itm-blue) 55%, #1a6eb5 100%);
  position: sticky; top:0; z-index:100;
  box-shadow: 0 2px 20px rgba(0,0,0,.25);
}
.header-top {
  display:flex; align-items:center; justify-content:space-between;
  padding: 10px 28px; gap:16px;
}
.header-brand { display:flex; align-items:center; gap:14px }
.header-brand img {
  height: 46px;
  filter: brightness(0) invert(1) drop-shadow(0 2px 6px rgba(0,0,0,.30));
}
.header-titles h1 {
  font-size: 1.05rem; font-weight:800; color:#fff;
  letter-spacing:.3px; line-height:1.2;
}
.header-titles p {
  font-size: .72rem; color: rgba(255,255,255,.65);
  margin-top:1px; letter-spacing:.2px;
}
.header-note {
  font-size:.63rem; color:rgba(255,255,255,.45);
  margin-top:2px; letter-spacing:.1px; font-style:italic;
  max-width:560px; line-height:1.4;
}
.header-badge {
  background: rgba(255,255,255,.12);
  border: 1px solid rgba(255,255,255,.3);
  color: #fff;
  font-size:.7rem; font-weight:700;
  padding:4px 10px; border-radius:20px;
  letter-spacing:.4px; text-transform:uppercase;
  white-space:nowrap;
}
/* Green stripe */
.header-stripe {
  height: 4px;
  background: linear-gradient(90deg, #065f46, #059669, #34d399, #059669);
}

/* ── NAVIGATION TABS ─────────────────────────── */
nav {
  background: rgba(0,28,60,.55);
  backdrop-filter: blur(6px);
  display:flex; gap:0; padding:0 22px;
  overflow-x:auto; scrollbar-width:none;
}
nav::-webkit-scrollbar { display:none }
.nav-btn {
  background:transparent; border:none;
  color: rgba(255,255,255,.6);
  padding: 11px 18px;
  cursor:pointer; font-size:.82rem; font-weight:600;
  border-bottom:3px solid transparent;
  transition: var(--transition);
  white-space:nowrap; display:flex; align-items:center; gap:6px;
  letter-spacing:.2px;
}
.nav-btn:hover { color:#fff; background:rgba(255,255,255,.07) }
.nav-btn.active {
  color: var(--itm-gold3);
  border-bottom-color: var(--itm-gold);
  background: rgba(255,255,255,.06);
}
.nav-btn .dot {
  width:7px; height:7px; border-radius:50%;
  background:var(--itm-gold); opacity:0;
  transition:opacity .2s;
}
.nav-btn.active .dot { opacity:1 }

/* ── FILTER BAR ──────────────────────────────── */
.filter-bar {
  background: var(--surface);
  border-bottom: 2px solid var(--border);
  padding: 8px 24px;
  display:flex; flex-wrap:wrap; align-items:center; gap:6px 0;
  position:sticky; top:0; z-index:100;
  box-shadow:0 2px 8px rgba(0,0,0,.06);
}
.filter-cluster {
  display:flex; align-items:center; gap:6px;
  padding: 2px 14px 2px 0; margin-right:14px;
  border-right: 2px solid var(--border);
}
.filter-cluster:last-of-type { border-right:none; margin-right:6px; }
.filter-cluster-label {
  font-size:.62rem; font-weight:800; color:var(--itm-blue);
  text-transform:uppercase; letter-spacing:.8px; white-space:nowrap;
  margin-right:4px; padding:2px 6px;
  background:rgba(0,83,155,.07); border-radius:4px;
}
.filter-group { display:flex; align-items:center; gap:4px }
.filter-label {
  font-size:.67rem; font-weight:700; color:var(--text2);
  text-transform:uppercase; letter-spacing:.4px; white-space:nowrap;
}
.filter-select {
  border: 1.5px solid var(--border); border-radius:var(--radius-sm);
  padding: 4px 24px 4px 8px; font-size:.78rem; color:var(--text);
  background: #f4f7fd url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%2300539B'/%3E%3C/svg%3E") no-repeat right 7px center;
  appearance:none; cursor:pointer;
  transition:border-color .2s, box-shadow .2s;
}
.filter-select:focus { outline:none; border-color:var(--itm-blue); box-shadow:0 0 0 3px rgba(0,83,155,.12) }
.filter-select.sel-sm  { max-width:90px }
.filter-select.sel-md  { max-width:160px }
.filter-select.sel-lg  { max-width:210px }
.btn-reset {
  background: linear-gradient(135deg, var(--itm-gold), var(--itm-gold2));
  color:#fff; border:none; border-radius:var(--radius-sm);
  padding:4px 12px; cursor:pointer; font-size:.75rem; font-weight:700;
  transition:var(--transition); box-shadow:0 2px 8px rgba(232,160,0,.3);
  letter-spacing:.2px; white-space:nowrap;
}
.btn-reset:hover { transform:translateY(-1px); box-shadow:0 4px 14px rgba(232,160,0,.4) }
.filter-info {
  font-size:.7rem; color:var(--itm-blue);
  background:rgba(0,83,155,.08); border-radius:20px;
  padding:2px 10px; font-weight:600; letter-spacing:.2px; white-space:nowrap;
}
.filter-divider { width:1px; height:20px; background:var(--border) }

/* ── MAIN CONTENT ────────────────────────────── */
main { padding: 24px 28px; max-width:1600px; margin:0 auto }

.section { display:none; animation:fadeIn .3s ease }
.section.active { display:block }
@keyframes fadeIn { from{opacity:0;transform:translateY(8px)} to{opacity:1;transform:none} }

/* ── SECTION HERO ────────────────────────────── */
.sec-hero {
  background: linear-gradient(118deg, var(--itm-blue2) 0%, var(--itm-blue) 70%, var(--itm-blue3) 100%);
  border-radius: var(--radius);
  padding: 22px 28px;
  margin-bottom: 22px;
  display:flex; align-items:center; justify-content:space-between; gap:20px;
  box-shadow: var(--shadow-md); position:relative; overflow:hidden;
}
.sec-hero::before {
  content:''; position:absolute; right:-40px; top:-40px;
  width:160px; height:160px; border-radius:50%;
  background:rgba(255,255,255,.05);
}
.sec-hero::after {
  content:''; position:absolute; right:60px; bottom:-60px;
  width:220px; height:220px; border-radius:50%;
  background:rgba(232,160,0,.08);
}
.sec-hero-left { position:relative; z-index:1 }
.sec-hero-icon {
  font-size:2rem; margin-bottom:6px;
}
.sec-hero-left h2 {
  font-size:1.25rem; font-weight:800; color:#fff; letter-spacing:.3px;
}
.sec-hero-left p { font-size:.8rem; color:rgba(255,255,255,.65); margin-top:3px }
.sec-hero-right {
  display:flex; gap:12px; flex-wrap:wrap; position:relative; z-index:1;
}
.hero-stat {
  background:rgba(255,255,255,.12);
  border: 1px solid rgba(255,255,255,.18);
  border-radius:var(--radius-sm);
  padding:10px 16px; text-align:center; min-width:90px;
  backdrop-filter:blur(4px);
}
.hero-stat-val {
  font-size:1.6rem; font-weight:900; color:#fff; line-height:1;
  font-variant-numeric:tabular-nums;
}
.hero-stat-val.gold { color:var(--itm-gold3) }
.hero-stat-val.green { color:#6ee7b7 }
.hero-stat-label { font-size:.65rem; color:rgba(255,255,255,.6); text-transform:uppercase; margin-top:3px; letter-spacing:.3px }

/* ── KPI CARDS ───────────────────────────────── */
.kpi-row {
  display:grid; grid-template-columns:repeat(auto-fill, minmax(160px,1fr));
  gap:14px; margin-bottom:22px;
}
.kpi {
  background: var(--surface);
  border-radius: var(--radius);
  padding: 16px 18px;
  box-shadow: var(--shadow);
  border-top: 4px solid var(--itm-blue);
  position:relative; overflow:hidden;
  transition:var(--transition);
}
.kpi:hover { transform:translateY(-2px); box-shadow:var(--shadow-md) }
.kpi::after {
  content:''; position:absolute; right:-12px; bottom:-12px;
  width:60px; height:60px; border-radius:50%;
  background:rgba(0,83,155,.05);
}
.kpi.gold  { border-top-color:var(--itm-gold) }
.kpi.green { border-top-color:var(--green) }
.kpi.red   { border-top-color:var(--red) }
.kpi.purple{ border-top-color:var(--purple) }
.kpi-num {
  font-size:2rem; font-weight:900; color:var(--itm-blue); line-height:1;
  letter-spacing:-1px; font-variant-numeric:tabular-nums;
}
.kpi.gold   .kpi-num { color:var(--itm-gold2) }
.kpi.green  .kpi-num { color:#059669 }
.kpi.red    .kpi-num { color:var(--red) }
.kpi.purple .kpi-num { color:var(--purple) }
.kpi-lbl {
  font-size:.7rem; font-weight:700; color:var(--text3);
  text-transform:uppercase; letter-spacing:.5px; margin-top:5px;
}
.kpi-sub { font-size:.72rem; color:var(--text2); margin-top:2px }

/* ── STAT BOX ────────────────────────────────── */
.stat-panel {
  background: var(--surface2);
  border:1px solid var(--border);
  border-left:5px solid var(--itm-blue);
  border-radius:var(--radius);
  padding:16px 20px;
  margin-bottom:22px;
  font-size:.84rem; line-height:1.8; color:var(--text2);
}
.stat-panel strong { color:var(--itm-blue); font-weight:700 }

/* ── SECTION LOCAL FILTER ────────────────────── */
.sec-filter-bar {
  display:flex; align-items:center; gap:8px; flex-wrap:wrap;
  background:rgba(0,83,155,.04); border:1.5px solid var(--border);
  border-radius:var(--radius); padding:8px 16px; margin-bottom:18px;
}
.sec-filter-bar .filter-label { font-size:.7rem; font-weight:700; color:var(--text2);
  text-transform:uppercase; letter-spacing:.4px; white-space:nowrap; }
.sec-filter-title { font-size:.72rem; font-weight:800; color:var(--itm-blue);
  text-transform:uppercase; letter-spacing:.6px; margin-right:4px; }
.sec-filter-chip {
  display:inline-flex; align-items:center; gap:4px;
  padding:4px 12px; border-radius:20px; font-size:.75rem; font-weight:600;
  border: 1.5px solid var(--border); background:var(--surface);
  cursor:pointer; transition:all .18s; white-space:nowrap; color:var(--text2);
}
.sec-filter-chip:hover { border-color:var(--itm-blue); color:var(--itm-blue); }
.sec-filter-chip.active { background:var(--itm-blue); color:#fff; border-color:var(--itm-blue); }

/* ── CHARTS GRID ─────────────────────────────── */
.charts-grid {
  display:grid;
  grid-template-columns:repeat(auto-fill, minmax(420px,1fr));
  gap:18px; margin-bottom:22px;
}
.card {
  background:var(--surface); border-radius:var(--radius);
  box-shadow:var(--shadow); overflow:hidden;
  transition:var(--transition);
}
.card:hover { box-shadow:var(--shadow-md) }
.card.full  { grid-column:1/-1 }
.card.half  { grid-column:span 1 }
.card-head {
  padding:14px 18px 0;
  display:flex; align-items:center; justify-content:space-between;
}
.card-head h3 {
  font-size:.85rem; font-weight:700; color:var(--itm-blue2);
  letter-spacing:.2px;
}
.card-badge {
  background:rgba(0,83,155,.1); color:var(--itm-blue);
  font-size:.68rem; font-weight:700; border-radius:20px;
  padding:2px 8px; letter-spacing:.3px;
}
.card-body { padding:14px 18px 18px }
.ch { position:relative }
.ch.h180 { height:180px }
.ch.h240 { height:240px }
.ch.h300 { height:300px }
.ch.h360 { height:360px }
.ch.h420 { height:420px }

/* ── TABLES ──────────────────────────────────── */
.tbl-scroll { overflow-x:auto; max-height:320px; overflow-y:auto }
table { width:100%; border-collapse:collapse; font-size:.8rem }
thead { position:sticky; top:0; z-index:2 }
th {
  background:var(--itm-blue); color:#fff;
  padding:8px 12px; text-align:left; font-weight:700;
  font-size:.72rem; text-transform:uppercase; letter-spacing:.3px;
  white-space:nowrap;
}
th:first-child { border-radius:0 }
td { padding:7px 12px; border-bottom:1px solid #eef1f7; vertical-align:middle }
tr:nth-child(even) td { background:#f8fafd }
tr:hover td { background:#e8f0fb }
.tbl-rank {
  background:var(--itm-blue); color:#fff;
  font-weight:800; font-size:.72rem;
  border-radius:50%; width:20px; height:20px;
  display:inline-flex; align-items:center; justify-content:center;
}
.tbl-rank.gold { background:var(--itm-gold) }
.tbl-rank.silver { background:#9ca3af }
.badge-val {
  background:rgba(0,83,155,.1); color:var(--itm-blue);
  font-weight:700; border-radius:6px;
  padding:2px 8px;
}
.badge-green { background:rgba(16,185,129,.12); color:#059669 }
.badge-orange { background:rgba(232,160,0,.12); color:var(--itm-gold2) }

/* ── WORD CLOUD ──────────────────────────────── */
.prog-pills { display:flex; flex-wrap:wrap; gap:6px; margin-bottom:12px }
.prog-pill {
  background:var(--bg); border:1.5px solid var(--border);
  border-radius:20px; padding:4px 12px;
  cursor:pointer; font-size:.76rem; color:var(--text2); font-weight:600;
  transition:var(--transition);
}
.prog-pill:hover { border-color:var(--itm-blue); color:var(--itm-blue) }
.prog-pill.active {
  background:var(--itm-blue); border-color:var(--itm-blue);
  color:#fff; box-shadow:0 2px 8px rgba(0,83,155,.3);
}
.word-cloud { display:flex; flex-wrap:wrap; gap:8px; padding:8px 0 }
.chip {
  border-radius:20px; padding:5px 14px;
  font-weight:700; cursor:default; transition:transform .15s;
  line-height:1.3;
}
.chip:hover { transform:scale(1.06) }
.chip.xl { font-size:1rem; background:var(--itm-blue2); color:#fff }
.chip.lg { font-size:.88rem; background:var(--itm-blue); color:#fff }
.chip.md { font-size:.78rem; background:var(--itm-blue3); color:#fff }
.chip.sm { font-size:.71rem; background:#4a7fb0; color:#fff }
.chip.xs { font-size:.65rem; background:var(--bg); color:var(--itm-blue); border:1.5px solid var(--border) }

/* ── PROGRESS BARS ───────────────────────────── */
.progress-list { display:flex; flex-direction:column; gap:8px }
.prog-item { display:flex; flex-direction:column; gap:3px }
.prog-item-head { display:flex; justify-content:space-between; align-items:center }
.prog-item-name { font-size:.76rem; font-weight:600; color:var(--text); max-width:240px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap }
.prog-item-val { font-size:.76rem; font-weight:800; color:var(--itm-blue) }
.prog-bar-bg { background:var(--bg); border-radius:6px; height:7px; overflow:hidden }
.prog-bar-fill { height:100%; border-radius:6px; background:linear-gradient(90deg,var(--itm-blue),var(--itm-blue3)); transition:width .6s ease }
.prog-bar-fill.gold { background:linear-gradient(90deg,var(--itm-gold2),var(--itm-gold)) }

/* ── RESPONSIVE ──────────────────────────────── */
/* Tablet (≤1024px) */
@media(max-width:1024px) {
  main { padding:18px 16px }
  .charts-grid { grid-template-columns:repeat(auto-fill, minmax(300px,1fr)) }
  .kpi-row { grid-template-columns:repeat(auto-fill, minmax(140px,1fr)) }
  .sub-grid-1-2 { grid-template-columns:1fr !important }
  .sub-grid-1-1 { grid-template-columns:1fr !important }
  .sub-grid-1-1-1 { grid-template-columns:1fr 1fr !important }
}
/* Mobile (≤768px) */
@media(max-width:768px) {
  main { padding:12px 10px }
  .charts-grid { grid-template-columns:1fr }
  .kpi-row { grid-template-columns:repeat(2,1fr) }
  .header-top { padding:8px 12px; gap:8px }
  .header-brand img { height:36px }
  .header-titles h1 { font-size:.85rem }
  .header-titles p { display:none }
  .header-badge { display:none }
  nav { padding:0 8px }
  .nav-btn { padding:9px 12px; font-size:.76rem; gap:4px }
  .filter-bar { padding:6px 10px; gap:4px }
  .filter-cluster { padding:2px 8px 2px 0; margin-right:6px }
  .filter-cluster-label { display:none }
  .filter-select { font-size:.73rem; padding:3px 20px 3px 6px }
  .filter-select.sel-sm { max-width:68px }
  .filter-select.sel-md { max-width:110px }
  .filter-select.sel-lg { max-width:130px }
  .btn-reset { padding:3px 10px; font-size:.72rem }
  .sec-hero { flex-direction:column; padding:16px }
  .sec-hero-right { width:100%; justify-content:flex-start }
  .hero-stat { min-width:70px; padding:8px 10px }
  .hero-stat-val { font-size:1.3rem }
  .stat-panel { font-size:.78rem; padding:12px 14px }
  .kpi-num { font-size:1.6rem }
  .card-head h3 { font-size:.78rem }
  .card-body { padding:10px 12px 14px }
  .tbl-scroll { max-height:260px }
  .sub-grid-1-2 { grid-template-columns:1fr !important }
  .sub-grid-1-1 { grid-template-columns:1fr !important }
  .sub-grid-1-1-1 { grid-template-columns:1fr !important }
  .sec-filter-bar { padding:6px 10px; gap:6px }
  .prog-item-name { max-width:180px }
}
/* Small mobile (≤480px) */
@media(max-width:480px) {
  .kpi-row { grid-template-columns:1fr 1fr }
  .hero-stat { min-width:60px; padding:6px 8px }
  .hero-stat-val { font-size:1.1rem }
  .filter-cluster-label { display:none }
  .filter-select.sel-lg { max-width:110px }
  nav { padding:0 4px }
  .nav-btn { padding:8px 9px; font-size:.7rem }
  .ch.h360 { height:280px }
  .ch.h300 { height:240px }
  .ch.h420 { height:300px }
  .tbl-scroll { max-height:220px }
}

/* ── FOOTER ──────────────────────────────────── */
footer {
  background: var(--itm-blue2);
  color: rgba(255,255,255,.55);
  text-align:center; padding:16px 28px;
  font-size:.75rem; letter-spacing:.2px;
  margin-top:32px;
}
footer span { color:var(--itm-gold3); font-weight:600 }
footer b { color:rgba(255,255,255,.8) }

/* ── SCROLL BAR ──────────────────────────────── */
::-webkit-scrollbar { width:6px; height:6px }
::-webkit-scrollbar-track { background:var(--bg) }
::-webkit-scrollbar-thumb { background:var(--border); border-radius:3px }
::-webkit-scrollbar-thumb:hover { background:var(--itm-blue3) }

/* ── PDF BUTTON ──────────────────────────────── */
.btn-pdf {
  background: rgba(255,255,255,.15);
  border: 1.5px solid rgba(255,255,255,.35);
  color: #fff;
  font-size:.72rem; font-weight:700;
  padding:5px 13px; border-radius:20px;
  cursor:pointer; letter-spacing:.4px;
  text-transform:uppercase;
  display:flex; align-items:center; gap:5px;
  transition:var(--transition);
  white-space:nowrap;
}
.btn-pdf:hover {
  background:rgba(255,255,255,.28);
  transform:translateY(-1px);
  box-shadow:0 3px 10px rgba(0,0,0,.25);
}

/* ── WORD CLOUD CANVAS ───────────────────────── */
#wc-canvas-f082 { display:block; max-height:420px }

/* ── CARRUSEL CUEE ───────────────────────────── */
.cuee-carousel {
  position:relative; overflow:hidden; border-radius:10px;
  background:#000; aspect-ratio:16/7; max-height:440px;
}
.cuee-track { display:flex; height:100%; transition:none }
.cuee-slide {
  min-width:100%; height:100%; position:relative;
  display:none; align-items:center; justify-content:center;
  background:#0f172a;
}
.cuee-slide.active { display:flex }
.cuee-slide img {
  width:100%; height:100%; object-fit:contain;
  display:block; border-radius:10px; background:#0f172a;
}
.cuee-caption {
  position:absolute; bottom:0; left:0; right:0;
  background:linear-gradient(0deg,rgba(0,0,0,.75) 0%,rgba(0,0,0,0) 100%);
  padding:28px 20px 16px; border-radius:0 0 10px 10px;
  color:#fff;
}
.cuee-caption-title { font-size:.88rem; font-weight:700; margin-bottom:3px }
.cuee-caption-desc  { font-size:.74rem; opacity:.82 }
.cuee-btn {
  position:absolute; top:50%; transform:translateY(-50%);
  background:rgba(255,255,255,.18); border:none; color:#fff;
  font-size:2rem; line-height:1; width:42px; height:42px;
  border-radius:50%; cursor:pointer; backdrop-filter:blur(4px);
  transition:.2s; display:flex; align-items:center; justify-content:center;
}
.cuee-btn:hover { background:rgba(255,255,255,.36) }
.cuee-prev { left:12px }
.cuee-next { right:12px }
.cuee-dots {
  position:absolute; bottom:10px; left:50%; transform:translateX(-50%);
  display:flex; gap:6px;
}
.cuee-dot {
  width:7px; height:7px; border-radius:50%;
  background:rgba(255,255,255,.45); cursor:pointer; transition:.2s;
}
.cuee-dot.active { background:#fff; transform:scale(1.25) }
.cuee-thumbs {
  display:flex; gap:8px; margin-top:10px;
  overflow-x:auto; padding-bottom:4px;
  scrollbar-width:thin; scrollbar-color:#d1d5db transparent;
}
.cuee-thumb {
  flex-shrink:0; width:80px; height:54px; border-radius:7px;
  overflow:hidden; cursor:pointer; border:2px solid transparent;
  transition:.2s; opacity:.6;
}
.cuee-thumb:hover { opacity:.85 }
.cuee-thumb.active { border-color:var(--itm,#102D69); opacity:1 }
.cuee-thumb img { width:100%; height:100%; object-fit:cover; display:block }
/* ── RESPONSIVE CUEE + HISTÓRICOS ───────────────────────── */
/* Tablet ≤1024px */
@media(max-width:1024px){
  .rg-3col { grid-template-columns:1fr 1fr !important }
  .rg-4col { grid-template-columns:1fr 1fr !important }
  #cuee-ratings { grid-template-columns:1fr 1fr !important }
  #hist-kpis    { grid-template-columns:1fr 1fr !important }
  .rg-fac { grid-template-columns:1fr 1fr !important }
  .rg-hist-1-2 { grid-template-columns:1fr !important }
}
/* Mobile ≤768px */
@media(max-width:768px){
  .cuee-carousel { aspect-ratio:16/9; max-height:280px }
  .cuee-thumb { width:60px; height:40px }
  .cuee-btn { width:34px; height:34px; font-size:1.4rem }
  /* hero CUEE: wrap en una sola columna */
  #hero-cuee { flex-direction:column; gap:8px }
  #hero-cuee > div { min-width:unset !important; width:100% }
  /* grids CUEE */
  .rg-2col { grid-template-columns:1fr !important }
  .rg-3col { grid-template-columns:1fr !important }
  .rg-fac  { grid-template-columns:1fr !important }
  #cuee-ratings { grid-template-columns:1fr 1fr !important }
  /* grids Históricos */
  .rg-hist-1-2 { grid-template-columns:1fr !important }
  #hist-kpis   { grid-template-columns:1fr 1fr !important }
  /* ajuste alturas en mobile */
  .ch.h260, .ch.h250, .ch.h240 { height:220px }
  /* Tendencias: columna izquierda sin borde derecho */
  .rg-2col [style*="border-right:1px solid var(--border)"] { border-right:none !important; border-bottom:1px solid var(--border) }
}
/* Small mobile ≤480px */
@media(max-width:480px){
  .cuee-caption-title { font-size:.78rem }
  .cuee-caption-desc  { display:none }
  #cuee-ratings { grid-template-columns:1fr !important }
  #hist-kpis    { grid-template-columns:1fr 1fr !important }
  .rg-3col  { grid-template-columns:1fr !important }
  .rg-hist-1-2 { grid-template-columns:1fr !important }
  .ch.h260, .ch.h250, .ch.h240 { height:190px }
}

/* ── PRINT / PDF ─────────────────────────────── */
@media print {
  header, .filter-bar, .btn-pdf, nav { display:none !important }
  body { background:#fff }
  main { padding:0 }
  .section { display:block !important }
  .card { break-inside:avoid; box-shadow:none; border:1px solid #ddd }
  .sec-hero { background:#00539B !important; -webkit-print-color-adjust:exact; print-color-adjust:exact }
  .ch, .ch.h180, .ch.h240, .ch.h260, .ch.h300, .ch.h360 { height:220px }
}
</style>
</head>
<body>

<!-- ═══════════════ HEADER ═══════════════ -->
<header>
  <div class="header-top">
    <div class="header-brand">
      <img src="__LOGO__" alt="ITM" onerror="this.style.display='none'">
      <div class="header-titles">
        <h1>Dashboard – Prácticas Profesionales</h1>
        <p>ITM &ndash; Institución Universitaria &ndash; Prácticas Profesionales ITM &ndash; Informe de Indicadores</p>
      </div>
    </div>
    <div style="display:flex;align-items:center;gap:10px">
      <div class="header-badge">Informe Gerencial</div>
      <button class="btn-pdf" onclick="exportPDF()" title="Descargar PDF">
        &#8595; PDF
      </button>
    </div>
  </div>
  <div class="header-stripe"></div>
  <nav>
    <button class="nav-btn active" onclick="goTo('practicantes',this)">
      <span class="dot"></span>🎓 Practicantes
    </button>
    <button class="nav-btn" onclick="goTo('disponibles',this)">
      <span class="dot"></span>📑 Disponibles
    </button>
    <button class="nav-btn" onclick="goTo('f082',this)">
      <span class="dot"></span>📂 F082 Trabajos
    </button>
    <button class="nav-btn" onclick="goTo('solicitud',this)">
      <span class="dot"></span>🏢 Solicitud Empresas
    </button>
    <button class="nav-btn" onclick="goTo('aprobacion',this)">
      <span class="dot"></span>✅ Aprobación Funciones
    </button>
    <button class="nav-btn" onclick="goTo('encuesta',this)">
      <span class="dot"></span>📊 Encuesta Empresarios
    </button>
    <button class="nav-btn" onclick="goTo('encuesta-est',this)">
      <span class="dot"></span>🎓 Encuesta Estudiantes
    </button>
    <button class="nav-btn" onclick="goTo('cuee',this)">
      <span class="dot"></span>🤝 CUEE
    </button>
    <button class="nav-btn" onclick="goTo('historicos',this)">
      <span class="dot"></span>📊 Históricos
    </button>
  </nav>
</header>

<!-- ═══════════════ FILTER BAR ═══════════════ -->
<div class="filter-bar">
  <!-- Grupo temporal -->
  <div class="filter-cluster">
    <span class="filter-cluster-label">📅 Período</span>
    <div class="filter-group">
      <span class="filter-label">Año</span>
      <select id="f-anio" class="filter-select sel-sm" onchange="onFilterChange()">
        <option value="">Todos</option>
      </select>
    </div>
    <div class="filter-divider"></div>
    <div class="filter-group">
      <span class="filter-label">Sem.</span>
      <select id="f-sem" class="filter-select sel-sm" onchange="onFilterChange()">
        <option value="">Todos</option>
      </select>
    </div>
    <div class="filter-divider"></div>
    <div class="filter-group">
      <span class="filter-label">Mes</span>
      <select id="f-mes" class="filter-select sel-sm" onchange="onFilterChange()">
        <option value="">Todos</option>
      </select>
    </div>
  </div>
  <!-- Grupo académico -->
  <div class="filter-cluster">
    <span class="filter-cluster-label">🎓 Académico</span>
    <div class="filter-group">
      <span class="filter-label">Facultad</span>
      <select id="f-facultad" class="filter-select sel-md" onchange="onFacultadChange()">
        <option value="">Todas</option>
      </select>
    </div>
    <div class="filter-divider"></div>
    <div class="filter-group">
      <span class="filter-label">Programa</span>
      <select id="f-programa" class="filter-select sel-lg" onchange="onFilterChange()">
        <option value="">Todos</option>
      </select>
    </div>
  </div>
  <button class="btn-reset" onclick="resetFilters()">↺ Limpiar</button>
  <span id="f-info" class="filter-info" style="display:none"></span>
</div>

<!-- ═══════════════ MAIN ═══════════════ -->
<main>

<!-- ────────────── PRACTICANTES ────────────── -->
<section id="sec-practicantes" class="section active">
  <div class="sec-hero">
    <div class="sec-hero-left">
      <div class="sec-hero-icon">🎓</div>
      <h2>Practicantes en Práctica Profesional</h2>
      <p>Estado, tipo de contrato, empresa, asesor, programa y facultad</p>
    </div>
    <div class="sec-hero-right" id="hero-pract"></div>
  </div>
  <div class="kpi-row" id="kpi-pract"></div>
  <div class="stat-panel" id="stat-pract"></div>
  <!-- Filtro local de Estado -->
  <div class="sec-filter-bar">
    <span class="sec-filter-title">⚡ Estado</span>
    <button class="sec-filter-chip active" onclick="setEstadoPract('')">Todos</button>
    <span id="chips-estado-pract"></span>
  </div>
  <div class="charts-grid">
    <div class="card">
      <div class="card-head"><h3>Estado de los practicantes</h3><span class="card-badge" id="cb-p-estado"></span></div>
      <div class="card-body"><div class="ch h240"><canvas id="c-p-estado"></canvas></div></div>
    </div>
    <div class="card">
      <div class="card-head"><h3>Tipo de contrato</h3></div>
      <div class="card-body"><div class="ch h240"><canvas id="c-p-contrato"></canvas></div></div>
    </div>
    <div class="card">
      <div class="card-head"><h3>Empresa nueva vs. recurrente por año</h3></div>
      <div class="card-body"><div class="ch h260"><canvas id="c-p-empnueva"></canvas></div></div>
    </div>
    <div class="card">
      <div class="card-head"><h3>Modalidad de práctica</h3></div>
      <div class="card-body"><div class="ch h200"><canvas id="c-p-modalidad"></canvas></div></div>
    </div>
    <div class="card">
      <div class="card-head"><h3>Distribución por facultad</h3></div>
      <div class="card-body"><div class="ch h260"><canvas id="c-p-facultad"></canvas></div></div>
    </div>
    <div class="card">
      <div class="card-head"><h3>Evolución semestral</h3></div>
      <div class="card-body"><div class="ch h260"><canvas id="c-p-semestre"></canvas></div></div>
    </div>
    <div class="card full">
      <div class="card-head"><h3>Top 15 programas con más practicantes</h3></div>
      <div class="card-body"><div id="prog-list-pract" class="progress-list"></div></div>
    </div>
    <div class="card full">
      <div class="card-head"><h3>Estudiantes por Monitor (Asesor de Prácticas)</h3></div>
      <div class="card-body" style="overflow-x:auto;padding-bottom:10px">
        <div class="ch h360" id="wrap-p-asesor" style="min-width:700px">
          <canvas id="c-p-asesor"></canvas>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ────────────── DISPONIBLES ────────────── -->
<section id="sec-disponibles" class="section">
  <div class="sec-hero">
    <div class="sec-hero-left">
      <div class="sec-hero-icon">📑</div>
      <h2>Estudiantes Disponibles</h2>
      <p>Solicitudes pendientes de ubicación empresarial</p>
    </div>
    <div class="sec-hero-right" id="hero-disp"></div>
  </div>
  <div class="kpi-row" id="kpi-disp"></div>
  <div class="stat-panel" id="stat-disp"></div>
  <div class="charts-grid">
    <div class="card full">
      <div class="card-head"><h3>Estudiantes disponibles por programa (Top 20)</h3></div>
      <div class="card-body"><div class="ch h360"><canvas id="c-d-programa"></canvas></div></div>
    </div>
    <div class="card">
      <div class="card-head"><h3>Por facultad</h3></div>
      <div class="card-body"><div class="ch h240"><canvas id="c-d-facultad"></canvas></div></div>
    </div>
    <div class="card">
      <div class="card-head"><h3>Discapacidad</h3></div>
      <div class="card-body"><div class="ch h200"><canvas id="c-d-disc"></canvas></div></div>
    </div>
    <div class="card">
      <div class="card-head"><h3>Tipo de discapacidad</h3></div>
      <div class="card-body"><div class="ch h200"><canvas id="c-d-tipdisc"></canvas></div></div>
    </div>
    <div class="card">
      <div class="card-head"><h3>Modalidad de contrato solicitada</h3></div>
      <div class="card-body"><div class="ch h200"><canvas id="c-d-modalidad"></canvas></div></div>
    </div>
    <div class="card">
      <div class="card-head"><h3>Solicitudes por mes</h3></div>
      <div class="card-body"><div class="ch h260"><canvas id="c-d-mes"></canvas></div></div>
    </div>
    <div class="card">
      <div class="card-head"><h3>Solicitudes por año</h3></div>
      <div class="card-body"><div class="ch h260"><canvas id="c-d-anio"></canvas></div></div>
    </div>
  </div>
</section>

<!-- ────────────── F082 ────────────── -->
<section id="sec-f082" class="section">
  <div class="sec-hero">
    <div class="sec-hero-left">
      <div class="sec-hero-icon">📂</div>
      <h2>F082 – Trabajos Entregados</h2>
      <p>Histórico de entregas, vinculación laboral y análisis de actividades</p>
    </div>
    <div class="sec-hero-right" id="hero-f082"></div>
  </div>
  <div class="kpi-row" id="kpi-f082"></div>
  <div class="stat-panel" id="stat-f082"></div>
  <div class="charts-grid">
    <div class="card full">
      <div class="card-head"><h3>Top 15 programas por entregas</h3></div>
      <div class="card-body"><div class="ch h360"><canvas id="c-f-programa"></canvas></div></div>
    </div>
    <div class="card">
      <div class="card-head"><h3>Distribución por facultad</h3></div>
      <div class="card-body"><div class="ch h240"><canvas id="c-f-facultad"></canvas></div></div>
    </div>
    <div class="card">
      <div class="card-head"><h3>Vinculación laboral al finalizar</h3></div>
      <div class="card-body"><div class="ch h200"><canvas id="c-f-vinculado"></canvas></div></div>
    </div>
    <div class="card">
      <div class="card-head"><h3>Estado de entrega</h3></div>
      <div class="card-body"><div class="ch h200"><canvas id="c-f-entregado"></canvas></div></div>
    </div>
    <div class="card full">
      <div class="card-head"><h3>Histórico de entregas por semestre</h3></div>
      <div class="card-body"><div class="ch h280"><canvas id="c-f-historico"></canvas></div></div>
    </div>
    <div class="card full">
      <div class="card-head"><h3>Histórico de vinculación laboral por año</h3><span class="card-badge">Vinculados vs. No vinculados</span></div>
      <div class="card-body"><div class="ch h300"><canvas id="c-f-vinc-anio"></canvas></div></div>
    </div>

  </div>
  <div class="card full" style="margin-bottom:18px">
    <div class="card-head">
      <h3>Competencias y áreas más demandadas por el mercado</h3>
      <span class="card-badge">Actividades · Descripción · Por programa / facultad</span>
    </div>
    <div class="card-body">
      <div style="display:flex;align-items:center;gap:16px;margin-bottom:18px;flex-wrap:wrap">
        <div class="filter-group">
          <span class="filter-label">Facultad</span>
          <select id="sel-fac-f082" class="filter-select" style="max-width:320px"
            onchange="onF082FacChange()">
            <option value="">Todas las facultades</option>
          </select>
        </div>
        <div class="filter-group">
          <span class="filter-label">Programa</span>
          <select id="sel-prog-f082" class="filter-select" style="max-width:380px"
            onchange="onF082ProgChange()">
            <option value="">Todos los programas</option>
          </select>
        </div>
        <span id="f082-wc-count" style="font-size:.78rem;color:var(--text2);font-weight:600"></span>
      </div>
      <div id="wrap-areas-f082" class="ch h420">
        <canvas id="c-areas-f082"></canvas>
      </div>
    </div>
  </div>

  <!-- Trabajos por Monitor con filtros locales -->
  <div class="card full" style="margin-bottom:18px">
    <div class="card-head">
      <h3>Trabajos entregados por Monitor (Asesor de Prácticas)</h3>
      <span class="card-badge">Total de F082 por asesor</span>
    </div>
    <div class="card-body">
      <div style="display:flex;align-items:center;gap:12px;margin-bottom:16px;flex-wrap:wrap">
        <div class="filter-group">
          <span class="filter-label">Año</span>
          <select id="sel-anio-asesor-f082" class="filter-select sel-sm" onchange="renderAsesorF082()">
            <option value="">Todos</option>
          </select>
        </div>
        <div class="filter-group">
          <span class="filter-label">Facultad</span>
          <select id="sel-fac-asesor-f082" class="filter-select sel-md" onchange="renderAsesorF082()">
            <option value="">Todas</option>
          </select>
        </div>
        <div class="filter-group">
          <span class="filter-label">Programa</span>
          <select id="sel-prog-asesor-f082" class="filter-select sel-lg" onchange="renderAsesorF082()">
            <option value="">Todos</option>
          </select>
        </div>
        <button class="btn-reset" onclick="resetAsesorF082()">↺ Limpiar</button>
        <span id="f082-asesor-count" style="font-size:.78rem;color:var(--text2);font-weight:600"></span>
      </div>
      <div id="wrap-asesor-f082" style="overflow-y:auto;max-height:600px"></div>
    </div>
  </div>
</section>

<!-- ────────────── SOLICITUD EMPRESAS ────────────── -->
<section id="sec-solicitud" class="section">
  <div class="sec-hero">
    <div class="sec-hero-left">
      <div class="sec-hero-icon">🏢</div>
      <h2>Solicitud de Empresas</h2>
      <p>Demanda del mercado laboral por programa, modalidad y tendencias</p>
    </div>
    <div class="sec-hero-right" id="hero-solic"></div>
  </div>
  <div class="kpi-row" id="kpi-solic"></div>
  <div class="stat-panel" id="stat-solic"></div>

  <!-- Filtros locales Solicitud -->
  <div class="sec-filter-bar" style="margin-bottom:18px">
    <span class="sec-filter-title">🔍 Filtrar</span>
    <div class="filter-group">
      <span class="filter-label">Año</span>
      <select id="sel-anio-solic" class="filter-select sel-sm" onchange="renderSolicitud()">
        <option value="">Todos</option>
      </select>
    </div>
    <div class="filter-group">
      <span class="filter-label">Facultad</span>
      <select id="sel-fac-solic" class="filter-select sel-md" onchange="onSolicFacChange()">
        <option value="">Todas</option>
      </select>
    </div>
    <div class="filter-group">
      <span class="filter-label">Programa</span>
      <select id="sel-prog-solic" class="filter-select sel-lg" onchange="renderSolicitud()">
        <option value="">Todos</option>
      </select>
    </div>
    <button class="btn-reset" onclick="resetSolicFiltros()">↺ Limpiar</button>
    <span id="solic-count" style="font-size:.78rem;color:var(--text2);font-weight:600"></span>
  </div>

  <div class="charts-grid">
    <div class="card full">
      <div class="card-head"><h3>Perfiles más solicitados por empresas (Top 20)</h3></div>
      <div class="card-body"><div class="ch h360"><canvas id="c-s-perfil"></canvas></div></div>
    </div>
    <div class="card">
      <div class="card-head"><h3>Modalidad de vinculación</h3></div>
      <div class="card-body"><div class="ch h240"><canvas id="c-s-modalidad"></canvas></div></div>
    </div>
    <div class="card">
      <div class="card-head"><h3>Empresa nueva vs. recurrente</h3></div>
      <div class="card-body"><div class="ch h200"><canvas id="c-s-empnueva"></canvas></div></div>
    </div>
    <div class="card">
      <div class="card-head"><h3>Evolución mensual de solicitudes</h3></div>
      <div class="card-body"><div class="ch h260"><canvas id="c-s-mes"></canvas></div></div>
    </div>
    <div class="card full">
      <div class="card-head"><h3>Solicitudes por año</h3></div>
      <div class="card-body"><div class="ch h260"><canvas id="c-s-anio"></canvas></div></div>
    </div>
    <!-- Gráfica combinada Solicitud vs Aprobación -->
    <div class="card full">
      <div class="card-head">
        <h3>Demanda empresarial vs. Capacidad de atención por año</h3>
        <span class="card-badge">Solicitudes · Aprobaciones · % Respuesta</span>
      </div>
      <div class="card-body">
        <div class="ch h320"><canvas id="c-s-vs-aprob"></canvas></div>
        <div id="analisis-ejecutivo" style="margin-top:18px;padding:16px 20px;background:var(--surface2);border-left:5px solid var(--itm-blue);border-radius:var(--radius);font-size:.84rem;line-height:1.8;color:var(--text2)"></div>
      </div>
    </div>
  </div>
</section>

<!-- ────────────── APROBACIÓN FUNCIONES ────────────── -->
<section id="sec-aprobacion" class="section">
  <div class="sec-hero">
    <div class="sec-hero-left">
      <div class="sec-hero-icon">✅</div>
      <h2>Aprobación de Funciones</h2>
      <p>Solicitudes por programa, empresa y análisis de las funciones del mercado</p>
    </div>
    <div class="sec-hero-right" id="hero-aprob"></div>
  </div>
  <div class="kpi-row" id="kpi-aprob"></div>
  <div class="stat-panel" id="stat-aprob"></div>
  <div class="charts-grid">
    <div class="card full">
      <div class="card-head"><h3>Solicitudes de aprobación por programa</h3></div>
      <div class="card-body"><div class="ch h360"><canvas id="c-a-programa"></canvas></div></div>
    </div>
    <div class="sub-grid-1-2" style="grid-column:1/-1;display:grid;grid-template-columns:1fr 2fr;gap:18px">
      <div class="card">
        <div class="card-head"><h3>Estado de aprobación</h3></div>
        <div class="card-body"><div class="ch h360"><canvas id="c-a-estado"></canvas></div></div>
      </div>
      <div class="card">
        <div class="card-head"><h3>Top 15 empresas solicitantes</h3></div>
        <div class="card-body"><div class="ch h360"><canvas id="c-a-empresa"></canvas></div></div>
      </div>
    </div>
    <div class="card full">
      <div class="card-head"><h3>Empresa por programa — Top 20 combinaciones más frecuentes</h3></div>
      <div class="card-body">
        <div class="tbl-scroll" id="tbl-ep"></div>
      </div>
    </div>
  </div>
  <div class="card full" style="margin-bottom:18px">
    <div class="card-head">
      <h3>Tipo de función solicitada — por programa / facultad</h3>
      <span class="card-badge">Top palabras clave en FUNCIONES</span>
    </div>
    <div class="card-body">
      <div style="display:flex;align-items:center;gap:16px;margin-bottom:18px;flex-wrap:wrap">
        <div class="filter-group">
          <span class="filter-label">Facultad</span>
          <select id="sel-fac-aprob" class="filter-select" style="max-width:320px"
            onchange="onAprobFacChange()">
            <option value="">Todas las facultades</option>
          </select>
        </div>
        <div class="filter-group">
          <span class="filter-label">Programa</span>
          <select id="sel-prog-aprob" class="filter-select" style="max-width:380px"
            onchange="onAprobProgChange()">
            <option value="">Todos los programas</option>
          </select>
        </div>
        <span id="aprob-kw-count" style="font-size:.78rem;color:var(--text2);font-weight:600"></span>
      </div>
      <div id="wrap-aprob-func" class="ch h420">
        <canvas id="c-aprob-func"></canvas>
      </div>
    </div>
  </div>
  <div class="card full" style="margin-bottom:18px">
    <div class="card-head">
      <h3>Demanda global del mercado — funciones más requeridas</h3>
      <span class="card-badge">Consolidado de todas las empresas</span>
    </div>
    <div class="card-body"><div class="ch h420"><canvas id="c-aprob-global"></canvas></div></div>
  </div>
</section>

<!-- ────────────── ENCUESTA SATISFACCIÓN EMPRESARIOS ────────────── -->
<section id="sec-encuesta" class="section">
  <div class="sec-hero">
    <div class="sec-hero-left">
      <div class="sec-hero-icon">📊</div>
      <h2>Encuesta de Satisfacción – Empresarios</h2>
      <p>Calidad de formación, desempeño laboral, fortalezas y competencias desde la perspectiva empresarial</p>
    </div>
    <div class="sec-hero-right" id="hero-enc"></div>
  </div>
  <div class="kpi-row" id="kpi-enc"></div>
  <div class="stat-panel" id="stat-enc"></div>
  <div class="charts-grid">
    <div class="card full">
      <div class="card-head"><h3>Calificaciones por criterio de evaluación (% Excelente · Bueno · Deficiente)</h3></div>
      <div class="card-body"><div class="ch h300"><canvas id="c-enc-calif"></canvas></div></div>
    </div>
    <div class="card">
      <div class="card-head"><h3>¿La empresa vincularía practicantes?</h3></div>
      <div class="card-body"><div class="ch h240"><canvas id="c-enc-vinculacion"></canvas></div></div>
    </div>
    <div class="card">
      <div class="card-head"><h3>Tipo de empresa</h3></div>
      <div class="card-body"><div class="ch h240"><canvas id="c-enc-tipo"></canvas></div></div>
    </div>
    <div class="card">
      <div class="card-head"><h3>Cobertura de mercado</h3></div>
      <div class="card-body"><div class="ch h240"><canvas id="c-enc-mercado"></canvas></div></div>
    </div>
    <div class="card full">
      <div class="card-head"><h3>Distribución por sector económico</h3></div>
      <div class="card-body"><div class="ch h300"><canvas id="c-enc-sector"></canvas></div></div>
    </div>
    <div class="sub-grid-1-1" style="grid-column:1/-1;display:grid;grid-template-columns:1fr 1fr;gap:18px">
      <div class="card">
        <div class="card-head"><h3>Fortalezas del practicante ITM</h3></div>
        <div class="card-body"><div class="ch h300"><canvas id="c-enc-fort-prac"></canvas></div></div>
      </div>
      <div class="card">
        <div class="card-head"><h3>Debilidades del practicante ITM</h3></div>
        <div class="card-body"><div class="ch h300"><canvas id="c-enc-deb-prac"></canvas></div></div>
      </div>
      <div class="card">
        <div class="card-head"><h3>Fortalezas del egresado ITM</h3></div>
        <div class="card-body"><div class="ch h300"><canvas id="c-enc-fort-eg"></canvas></div></div>
      </div>
      <div class="card">
        <div class="card-head"><h3>Debilidades del egresado ITM</h3></div>
        <div class="card-body"><div class="ch h300"><canvas id="c-enc-deb-eg"></canvas></div></div>
      </div>
    </div>
    <div class="card full">
      <div class="card-head"><h3>Competencias más demandadas por el mercado</h3></div>
      <div class="card-body"><div class="ch h300"><canvas id="c-enc-competencias"></canvas></div></div>
    </div>
    <div class="card full">
      <div class="card-head"><h3>Evolución de respuestas por año</h3></div>
      <div class="card-body"><div class="ch h240"><canvas id="c-enc-anio"></canvas></div></div>
    </div>
  </div>
</section>

<!-- ────────────── ENCUESTA SATISFACCIÓN ESTUDIANTES ────────────── -->
<section id="sec-encuesta-est" class="section">
  <div class="sec-hero">
    <div class="sec-hero-left">
      <div class="sec-hero-icon">🎓</div>
      <h2>Encuesta de Satisfacción – Estudiantes</h2>
      <p>Experiencia en empresa, gestión de la Oficina de Prácticas, proyecto de vida y recomendación</p>
    </div>
    <div class="sec-hero-right" id="hero-est"></div>
  </div>
  <div class="kpi-row" id="kpi-est"></div>
  <div class="stat-panel" id="stat-est"></div>
  <div class="charts-grid">

    <!-- Calificaciones comparadas (stacked) -->
    <div class="card full">
      <div class="card-head"><h3>Calificaciones por criterio — vista comparativa</h3><span class="card-badge">% Excelente · Bueno · Deficiente</span></div>
      <div class="card-body"><div class="ch h420"><canvas id="c-est-calif"></canvas></div></div>
    </div>

    <!-- Satisfacción general + Recomendaría + Modalidad + Proyecto a futuro -->
    <div class="sub-grid-1-1-1" style="grid-column:1/-1;display:grid;grid-template-columns:1fr 1fr 1fr;gap:18px">
      <div class="card">
        <div class="card-head"><h3>Satisfacción general del servicio</h3></div>
        <div class="card-body"><div class="ch h240"><canvas id="c-est-satisf"></canvas></div></div>
      </div>
      <div class="card">
        <div class="card-head"><h3>¿Recomendaría realizar la práctica?</h3></div>
        <div class="card-body"><div class="ch h240"><canvas id="c-est-rec"></canvas></div></div>
      </div>
      <div class="card">
        <div class="card-head"><h3>Modalidad de práctica</h3></div>
        <div class="card-body"><div class="ch h240"><canvas id="c-est-modal"></canvas></div></div>
      </div>
      <div class="card" style="grid-column:1/-1">
        <div class="card-head"><h3>Proyecto a futuro del estudiante</h3></div>
        <div class="card-body"><div class="ch h200"><canvas id="c-est-futuro"></canvas></div></div>
      </div>
    </div>

    <!-- Score por programa -->
    <div class="card full">
      <div class="card-head">
        <h3>Índice de satisfacción promedio por programa</h3>
        <span class="card-badge">Escala 1 – 3 · Verde ≥ 2.7 · Teal ≥ 2.3 · Dorado ≥ 1.8</span>
      </div>
      <div class="card-body" style="overflow-x:auto;padding-bottom:10px">
        <div class="ch h360" id="wrap-est-prog-score" style="min-width:700px">
          <canvas id="c-est-prog-score"></canvas>
        </div>
      </div>
    </div>

    <!-- Programas participantes -->
    <div class="card full">
      <div class="card-head"><h3>Programas con mayor participación</h3></div>
      <div class="card-body" style="overflow-x:auto;padding-bottom:10px">
        <div class="ch h300" id="wrap-est-prog" style="min-width:700px">
          <canvas id="c-est-prog"></canvas>
        </div>
      </div>
    </div>

    <!-- Top empresas -->
    <div class="card full">
      <div class="card-head"><h3>Top 15 empresas donde se realizó la práctica</h3></div>
      <div class="card-body"><div class="ch h360"><canvas id="c-est-empresa"></canvas></div></div>
    </div>

    <!-- Score por asesor -->
    <div class="card full">
      <div class="card-head">
        <h3>Índice de satisfacción por asesor ITM</h3>
        <span class="card-badge">Escala 1 – 3 · basado en todas las calificaciones</span>
      </div>
      <div class="card-body"><div class="ch h360"><canvas id="c-est-asesor-score"></canvas></div></div>
    </div>

    <!-- Evolución semestral -->
    <div class="card full">
      <div class="card-head"><h3>Evolución de respuestas por semestre</h3></div>
      <div class="card-body"><div class="ch h240"><canvas id="c-est-semestre"></canvas></div></div>
    </div>

  </div>
</section>

<!-- ═══════════════════════════════════════════════════════
     SECCIÓN CUEE
═══════════════════════════════════════════════════════ -->
<section id="sec-cuee" class="section">
  <div class="sec-hero">
    <div class="sec-hero-left">
      <div class="sec-hero-icon">🤝</div>
      <h2>Pasantías CUEE 2026-1</h2>
      <p>Comité Universidad, Empresa y Estado — Articulación academia–sector productivo · Mayo 13–15 y 19–21 de 2026</p>
    </div>
    <div id="hero-cuee" style="display:flex;gap:12px;align-items:stretch;flex-wrap:wrap"></div>
  </div>

  <div class="charts-grid">

    __CUEE_CAROUSEL__

    <!-- ── Fila A: Tasa conversión + Cifras generales (full, lado a lado) ── -->
    <div class="card full">
      <div class="card-head"><h3>Participación General 2026-1</h3><span class="card-badge">Tasa de conversión · Cifras clave</span></div>
      <div class="card-body">
        <div class="rg-2col" style="display:grid;grid-template-columns:1fr 1fr;gap:28px;align-items:center">

          <!-- Barras horizontales de conversión -->
          <div style="display:flex;flex-direction:column;justify-content:center;height:100%">
            <div style="font-size:.72rem;font-weight:700;color:var(--text2);text-transform:uppercase;letter-spacing:.06em;margin-bottom:18px">Distribución de 319 inscritos</div>

            <!-- Barra 1 -->
            <div style="margin-bottom:20px">
              <div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:6px">
                <span style="font-size:.8rem;font-weight:700;color:#10b981">Participantes efectivos</span>
                <span style="font-size:1.1rem;font-weight:900;color:#10b981">132 <span style="font-size:.72rem;font-weight:600;color:var(--text2)">(41.4%)</span></span>
              </div>
              <div style="background:#e2e8f0;border-radius:8px;height:18px;overflow:hidden">
                <div style="background:linear-gradient(90deg,#10b981,#34d399);height:100%;width:41.4%;border-radius:8px;transition:width .8s ease"></div>
              </div>
            </div>

            <!-- Barra 2 -->
            <div style="margin-bottom:20px">
              <div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:6px">
                <span style="font-size:.8rem;font-weight:700;color:#f59e0b">Cupo no utilizado</span>
                <span style="font-size:1.1rem;font-weight:900;color:#f59e0b">41 <span style="font-size:.72rem;font-weight:600;color:var(--text2)">(12.9%)</span></span>
              </div>
              <div style="background:#e2e8f0;border-radius:8px;height:18px;overflow:hidden">
                <div style="background:linear-gradient(90deg,#f59e0b,#fbbf24);height:100%;width:12.9%;border-radius:8px;transition:width .8s ease"></div>
              </div>
            </div>

            <!-- Barra 3 -->
            <div style="margin-bottom:20px">
              <div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:6px">
                <span style="font-size:.8rem;font-weight:700;color:#2563eb">Sin cupo asignado</span>
                <span style="font-size:1.1rem;font-weight:900;color:#2563eb">146 <span style="font-size:.72rem;font-weight:600;color:var(--text2)">(45.8%)</span></span>
              </div>
              <div style="background:#e2e8f0;border-radius:8px;height:18px;overflow:hidden">
                <div style="background:linear-gradient(90deg,#3b82f6,#93c5fd);height:100%;width:45.8%;border-radius:8px;transition:width .8s ease"></div>
              </div>
            </div>

            <!-- Tasa global destacada -->
            <div style="background:linear-gradient(135deg,#102D69,#1a4a9c);border-radius:12px;padding:14px 18px;display:flex;align-items:center;gap:16px;margin-top:4px">
              <div style="text-align:center">
                <div style="font-size:2rem;font-weight:900;color:#fff;line-height:1">41.4%</div>
                <div style="font-size:.65rem;color:rgba(255,255,255,.7);font-weight:600">tasa de participación</div>
              </div>
              <div style="width:1px;height:44px;background:rgba(255,255,255,.2)"></div>
              <div style="font-size:.75rem;color:rgba(255,255,255,.85);line-height:1.6">
                De cada <b style="color:#fff">10 inscritos</b>,<br><b style="color:#34d399">4 participaron</b> efectivamente en las empresas aliadas.
              </div>
            </div>
          </div>
          <!-- canvas oculto requerido por el JS -->
          <canvas id="c-cuee-conv" style="display:none"></canvas>

          <!-- Barras generales -->
          <div>
            <div class="ch h260"><canvas id="c-cuee-total"></canvas></div>
            <div style="display:flex;justify-content:center;gap:18px;margin-top:10px;font-size:.72rem;color:var(--text2)">
              <span>📉 Inscrito→Cupo: <b style="color:#ef4444">45.8%</b> no avanzó</span>
              <span>📉 Cupo→Asistencia: <b style="color:#f97316">23.7%</b> no asistió</span>
            </div>
          </div>

        </div>
      </div>
    </div>

    <!-- ── Fila B: Comparativo facultad + Doughnut + Tabla (full, 3 cols) ── -->
    <div class="card full">
      <div class="card-head"><h3>Distribución por Facultad</h3><span class="card-badge">Inscritos · Participantes · Conversión</span></div>
      <div class="card-body">
        <div class="rg-fac" style="display:grid;grid-template-columns:2fr 180px 1fr;gap:20px;align-items:center">
          <!-- Barras comparativo -->
          <div class="ch h250"><canvas id="c-cuee-fac-comp"></canvas></div>
          <!-- Doughnut -->
          <div style="height:220px;position:relative"><canvas id="c-cuee-fac-pie"></canvas></div>
          <!-- Tabla -->
          <table style="width:100%;border-collapse:collapse;font-size:.78rem">
            <thead>
              <tr style="background:var(--itm,#102D69)">
                <th style="padding:8px 10px;text-align:left;color:#fff;font-weight:600;font-size:.72rem;border-radius:6px 0 0 0">Facultad</th>
                <th style="padding:8px 8px;text-align:center;color:#fff;font-weight:600;font-size:.72rem">Ins.</th>
                <th style="padding:8px 8px;text-align:center;color:#fff;font-weight:600;font-size:.72rem">Part.</th>
                <th style="padding:8px 8px;text-align:center;color:#fff;font-weight:600;font-size:.72rem;border-radius:0 6px 0 0">%</th>
              </tr>
            </thead>
            <tbody id="cuee-fac-table"></tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- ── Fila C: Asistentes por empresa (full) ── -->
    <div class="card full">
      <div class="card-head"><h3>Asistentes por Empresa</h3><span class="card-badge">8 empresas · 132 total</span></div>
      <div class="card-body"><div class="ch h240"><canvas id="c-cuee-empresas"></canvas></div></div>
    </div>

    <!-- ── Fila D: Evaluación empresarios (full, 4 cols) ── -->
    <div class="card full">
      <div class="card-head"><h3>Evaluación de los Empresarios</h3><span class="card-badge">Escala 1–5 · Puntaje máximo en todos los criterios</span></div>
      <div class="card-body">
        <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px" id="cuee-ratings"></div>
        <div style="margin-top:14px;background:linear-gradient(135deg,#f0fdf4,#dcfce7);border-radius:10px;border-left:4px solid #10b981;padding:12px 16px">
          <div style="font-size:.68rem;font-weight:700;color:#059669;text-transform:uppercase;letter-spacing:.06em;margin-bottom:3px">Evaluación Global · 5.0 / 5.0</div>
          <p style="font-size:.76rem;color:var(--text2);line-height:1.5">Los empresarios calificaron con <b>puntaje máximo en los cuatro criterios</b>. Refleja una percepción excelente del desempeño, actitud y disposición de los estudiantes ITM.</p>
        </div>
      </div>
    </div>

    <!-- ── Fila E: Tendencias + Competencias (full, 2 cols) ── -->
    <div class="card full">
      <div class="card-head"><h3>Tendencias y Competencias del Sector Productivo</h3><span class="card-badge">Identificadas en Pasantías CUEE 2026-1</span></div>
      <div class="card-body">
        <div class="rg-2col" style="display:grid;grid-template-columns:1fr 1fr;gap:0;border-radius:10px;overflow:hidden;border:1px solid var(--border)">

          <!-- Columna izquierda: Tendencias -->
          <div style="padding:20px;background:#f8fafd;border-right:1px solid var(--border)">
            <div style="font-size:.72rem;font-weight:700;color:#1d4ed8;text-transform:uppercase;letter-spacing:.06em;margin-bottom:14px;display:flex;align-items:center;gap:8px">
              <span style="background:#2563eb;color:#fff;border-radius:6px;padding:3px 8px;font-size:.65rem">💡 Tecnológicas</span>
              Tendencias identificadas
            </div>
            <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px" id="cuee-tags-tech"></div>
            <div style="margin-top:16px;background:linear-gradient(135deg,#eff6ff,#dbeafe);border-radius:10px;border-left:4px solid #2563eb;padding:12px 14px">
              <div style="font-size:.67rem;font-weight:700;color:#1d4ed8;margin-bottom:3px;text-transform:uppercase;letter-spacing:.05em">Reto para el ITM</div>
              <p style="font-size:.75rem;color:var(--text2);line-height:1.5;margin:0">Las empresas solicitan formación en <b>IA, análisis de datos y automatización</b>. Postobón manifestó interés en programas de capacitación interna.</p>
            </div>
          </div>

          <!-- Columna derecha: Competencias -->
          <div style="padding:20px;background:#fff">
            <div style="font-size:.72rem;font-weight:700;color:#166534;text-transform:uppercase;letter-spacing:.06em;margin-bottom:14px;display:flex;align-items:center;gap:8px">
              <span style="background:#10b981;color:#fff;border-radius:6px;padding:3px 8px;font-size:.65rem">🎯 Perfil</span>
              Competencias demandadas
            </div>
            <div style="font-size:.68rem;font-weight:700;color:#1d4ed8;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px;padding-left:2px">Habilidades Técnicas</div>
            <div style="display:grid;grid-template-columns:repeat(2,1fr);gap:7px;margin-bottom:16px" id="cuee-tags-tec"></div>
            <div style="height:1px;background:var(--border);margin-bottom:14px"></div>
            <div style="font-size:.68rem;font-weight:700;color:#166534;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px;padding-left:2px">Power Skills</div>
            <div style="display:grid;grid-template-columns:repeat(2,1fr);gap:7px" id="cuee-tags-soft"></div>
          </div>

        </div>
      </div>
    </div>

    <!-- ── Fila F: Oportunidades (3 cols) ── -->
    <div class="card full">
      <div class="card-head"><h3>Oportunidades de Relacionamiento Identificadas</h3></div>
      <div class="card-body">
        <div class="rg-3col" style="display:grid;grid-template-columns:repeat(3,1fr);gap:14px">
          <div style="background:linear-gradient(135deg,#eff6ff,#dbeafe);border-radius:10px;padding:18px;border:1px solid #bfdbfe">
            <div style="font-size:1.4rem;margin-bottom:8px">🔬</div>
            <div style="font-weight:700;color:var(--itm);font-size:.84rem;margin-bottom:5px">Investigación & Extensión</div>
            <div style="font-size:.75rem;color:var(--text2);line-height:1.6">Docentes identificaron oportunidades relacionales e investigativas y la posibilidad de descentralizar aulas hacia las empresas.</div>
          </div>
          <div style="background:linear-gradient(135deg,#f0fdf4,#dcfce7);border-radius:10px;padding:18px;border:1px solid #bbf7d0">
            <div style="font-size:1.4rem;margin-bottom:8px">📘</div>
            <div style="font-weight:700;color:#166534;font-size:.84rem;margin-bottom:5px">Formación Empresarial</div>
            <div style="font-size:.75rem;color:var(--text2);line-height:1.6">Postobón abrió ventana para capacitación de empleados en IA, análisis de datos y automatización. Contacto trasladado a extensión y facultad.</div>
          </div>
          <div style="background:linear-gradient(135deg,#fefce8,#fef9c3);border-radius:10px;padding:18px;border:1px solid #fde68a">
            <div style="font-size:1.4rem;margin-bottom:8px">🏫</div>
            <div style="font-weight:700;color:#92400e;font-size:.84rem;margin-bottom:5px">Aulas Descentralizadas</div>
            <div style="font-size:.75rem;color:var(--text2);line-height:1.6">Realizar ejercicios experienciales dentro de las empresas para llevar el aprendizaje directamente al entorno productivo real.</div>
          </div>
        </div>
      </div>
    </div>

    <!-- ── Fila G: Obs + Sugerencias + Conclusiones ── -->
    <div class="card full">
      <div class="card-head"><h3>Conclusiones, Observaciones y Sugerencias</h3></div>
      <div class="card-body">
        <div class="rg-2col" style="display:grid;grid-template-columns:1fr 1fr;gap:14px">
          <div style="background:#fefce8;border-radius:10px;border:1px solid #fde68a;padding:15px 17px">
            <div style="font-size:.71rem;font-weight:700;color:#92400e;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px">⚠️ Observaciones</div>
            <ul style="padding-left:15px;font-size:.76rem;color:var(--text2);line-height:1.7">
              <li>Desconocimiento del alcance del CUEE dentro de la comunidad universitaria.</li>
              <li>Deserción por temor a no conseguir alternativas de tiempo para actividades académicas.</li>
              <li>Los resultados de relacionamiento no tienen continuidad por falta de compromiso institucional.</li>
              <li>Se identifican talentos que deberían ser estimulados dentro de la institución.</li>
            </ul>
          </div>
          <div style="background:#f0fdf4;border-radius:10px;border:1px solid #bbf7d0;padding:15px 17px">
            <div style="font-size:.71rem;font-weight:700;color:#166534;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px">✅ Sugerencias</div>
            <ul style="padding-left:15px;font-size:.76rem;color:var(--text2);line-height:1.7">
              <li>Comunicación permanente y retroalimentación entre mentores, estudiantes y empresas.</li>
              <li>Fortalecer la estructura del programa de mentoría.</li>
              <li>Calendarios y notificaciones oportunas a las IES.</li>
              <li>Generar la <b>Escuela de Talento y Formación CUEE</b>.</li>
              <li>Entregar informe gerencial a las rectorías sobre el evento completo.</li>
            </ul>
          </div>
          <div style="background:#eff6ff;border-radius:10px;border:1px solid #bfdbfe;padding:15px 17px;grid-column:1/-1">
            <div style="font-size:.71rem;font-weight:700;color:#1e40af;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px">📌 Conclusiones Generales</div>
            <div class="rg-2col" style="display:grid;grid-template-columns:1fr 1fr;gap:0 24px">
              <ul style="padding-left:15px;font-size:.76rem;color:var(--text2);line-height:1.7">
                <li>Se logró establecer participación, ejecución y resultados en los objetivos propuestos.</li>
                <li>Los estudiantes conectaron teoría académica con la realidad productiva de manera significativa.</li>
              </ul>
              <ul style="padding-left:15px;font-size:.76rem;color:var(--text2);line-height:1.7">
                <li>Docentes y mentores deben documentar el observatorio para aportar a la academia sobre tendencias.</li>
                <li>Se requiere plan de continuidad conjunto con las IES en lo académico y relacional.</li>
              </ul>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- ── Fila H: Voz de los estudiantes ── -->
    <div class="card full">
      <div class="card-head"><h3>Voz de los Estudiantes</h3><span class="card-badge">Testimonios del programa</span></div>
      <div class="card-body">
        <div class="rg-3col" style="display:grid;grid-template-columns:repeat(3,1fr);gap:14px" id="cuee-quotes"></div>
      </div>
    </div>

  </div><!-- /charts-grid -->
</section>

<!-- ══════════════════════ HISTÓRICOS ══════════════════════ -->
<section id="sec-historicos" class="section">
  <div class="sec-hero">
    <div class="sec-hero-left">
      <div class="sec-hero-icon">📊</div>
      <h2>Indicadores Históricos</h2>
      <p>Evolución de los indicadores clave de Prácticas Profesionales ITM · 2024–2026</p>
    </div>
    <div id="hero-historicos" style="display:flex;gap:12px;align-items:stretch;flex-wrap:wrap"></div>
  </div>

  <div class="charts-grid">

    <!-- ── Fila 1: KPIs de resumen ── -->
    <div class="card full" id="hist-kpi-card">
      <div class="card-head"><h3>Resumen General del Período</h3><span class="card-badge">2024 · 2025 · 2026 (parcial)</span></div>
      <div class="card-body">
        <div id="hist-kpis" style="display:grid;grid-template-columns:repeat(4,1fr);gap:16px"></div>
      </div>
    </div>

    <!-- ── Fila 2: Asistentes Preprácticas ── -->
    <div class="card full">
      <div class="card-head"><h3>Asistentes Curso Preprácticas</h3><span class="card-badge">Evolución anual y por trimestre</span></div>
      <div class="card-body">
        <div class="rg-hist-1-2" style="display:grid;grid-template-columns:1fr 2fr;gap:24px;align-items:center">
          <div>
            <div class="ch h260"><canvas id="c-hist-pre-anio"></canvas></div>
          </div>
          <div>
            <div class="ch h260"><canvas id="c-hist-pre-trim"></canvas></div>
          </div>
        </div>
        <div id="hist-analisis-pre" style="margin-top:16px;padding:14px 18px;background:#f0f4ff;border-left:4px solid var(--itm-blue);border-radius:0 8px 8px 0;font-size:.82rem;color:#334;line-height:1.7"></div>
      </div>
    </div>

    <!-- ── Fila 3: Estudiantes que iniciaron prácticas ── -->
    <div class="card full">
      <div class="card-head"><h3>Estudiantes que Iniciaron Prácticas</h3><span class="card-badge">Registros por período</span></div>
      <div class="card-body">
        <div class="rg-hist-1-2" style="display:grid;grid-template-columns:1fr 2fr;gap:24px;align-items:center">
          <div>
            <div class="ch h260"><canvas id="c-hist-inic-anio"></canvas></div>
          </div>
          <div>
            <div class="ch h260"><canvas id="c-hist-inic-trim"></canvas></div>
          </div>
        </div>
        <div id="hist-analisis-inic" style="margin-top:16px;padding:14px 18px;background:#f0fff4;border-left:4px solid #10b981;border-radius:0 8px 8px 0;font-size:.82rem;color:#334;line-height:1.7"></div>
      </div>
    </div>

    <!-- ── Fila 4: Graduados ── -->
    <div class="card full">
      <div class="card-head"><h3>Graduados con Práctica</h3><span class="card-badge">Evolución anual y por trimestre</span></div>
      <div class="card-body">
        <div class="rg-hist-1-2" style="display:grid;grid-template-columns:1fr 2fr;gap:24px;align-items:center">
          <div>
            <div class="ch h260"><canvas id="c-hist-grad-anio"></canvas></div>
          </div>
          <div>
            <div class="ch h260"><canvas id="c-hist-grad-trim"></canvas></div>
          </div>
        </div>
        <div id="hist-analisis-grad" style="margin-top:16px;padding:14px 18px;background:#fff8f0;border-left:4px solid #f59e0b;border-radius:0 8px 8px 0;font-size:.82rem;color:#334;line-height:1.7"></div>
      </div>
    </div>

    <!-- ── Fila 5: Vinculados ── -->
    <div class="card full">
      <div class="card-head"><h3>Estudiantes Vinculados (Empresa)</h3><span class="card-badge">Contratos · Vinculaciones activas</span></div>
      <div class="card-body">
        <div class="rg-hist-1-2" style="display:grid;grid-template-columns:1fr 2fr;gap:24px;align-items:center">
          <div>
            <div class="ch h240"><canvas id="c-hist-vinc-anio"></canvas></div>
          </div>
          <div>
            <div class="ch h240"><canvas id="c-hist-vinc-trim"></canvas></div>
          </div>
        </div>
        <div id="hist-analisis-vinc" style="margin-top:16px;padding:14px 18px;background:#f5f0ff;border-left:4px solid #8b5cf6;border-radius:0 8px 8px 0;font-size:.82rem;color:#334;line-height:1.7"></div>
      </div>
    </div>

    <!-- ── Fila 6: Comparativo integral ── -->
    <div class="card full">
      <div class="card-head"><h3>Comparativo Integral por Año</h3><span class="card-badge">Preprácticas · Iniciaron · Graduados · Vinculados</span></div>
      <div class="card-body">
        <div class="ch h300"><canvas id="c-hist-comp"></canvas></div>
        <div id="hist-analisis-comp" style="margin-top:16px;padding:14px 18px;background:#f0f4ff;border-left:4px solid var(--itm-blue);border-radius:0 8px 8px 0;font-size:.82rem;color:#334;line-height:1.7"></div>
      </div>
    </div>

  </div><!-- /charts-grid -->
</section>

</main>

<footer>
  <b>ITM &ndash; Institución Universitaria</b> &nbsp;·&nbsp; <span>Prácticas Profesionales ITM</span>
  &nbsp;·&nbsp; Informe de Indicadores
  <div style="margin-top:6px;font-size:.68rem;color:rgba(255,255,255,.4);font-style:italic;letter-spacing:.1px">
    Información suministrada a partir de la base de datos alojada en SharePoint de la Unidad de Prácticas &ndash;
    Sistema de Información y Gestión de la Oficina de Prácticas Profesionales ITM.
  </div>
</footer>

<!-- ═══════════════ JAVASCRIPT ═══════════════ -->
<script>
const D = __DATA_JSON__;

// ── Paletas ─────────────────────────────────────────────────────────────────
const C = {
  blue:   '#00539B', blue2: '#003d73', blue3: '#1a6eb5',
  gold:   '#E8A000', gold2: '#b87e00', gold3: '#ffd060',
  green:  '#10b981', red:   '#9C095D', purple: '#8b5cf6',
  teal:   '#0891b2', orange:'#ea580c'
};
const PAL_MAIN = [C.blue,C.gold,C.blue3,C.green,C.purple,C.teal,C.orange,'#661081',
                  '#4a7fb0','#d4940a','#3485cc','#009030','#7c3aed','#0e7490','#F19800'];
const PAL_BLUE = ['#003d73','#00539B','#1a6eb5','#3485cc','#4f9de0','#6bb3f0','#88c8fa'];
const PAL_GOLD = [C.gold2, C.gold, C.gold3, '#ffe08a', '#b87e00'];

// ── Colores por facultad (paleta institucional ITM) ──────────────────────────
const FAC_COLORS = {
  artes:     ['#661081','#910581','#9C095D'],
  ingenieria:['#102D69','#00A0B7','#56ACDE'],
  exactas:   ['#009030','#98BF13','#D1DD72'],
  economicas:['#F19800','#F6B63E','#F9C873'],
};
function facColor(name, shade=0) {
  const n=(name||'').toUpperCase();
  if(n.includes('ARTES')||n.includes('HUMANIDAD')) return FAC_COLORS.artes[shade%3];
  if(n.includes('INGEN'))                          return FAC_COLORS.ingenieria[shade%3];
  if(n.includes('EXACT')||n.includes('APLICAD'))  return FAC_COLORS.exactas[shade%3];
  if(n.includes('ECONOM')||n.includes('ADMIN'))   return FAC_COLORS.economicas[shade%3];
  return PAL_MAIN[shade % PAL_MAIN.length];
}
function facColorsArr(labels) { return labels.map((l,i) => facColor(l, i%3)); }

function pal(n, arr){ return Array.from({length:n},(_,i)=>arr[i%arr.length]) }

// ── Chart helpers ────────────────────────────────────────────────────────────
function mkBar(id, labels, vals, {horiz=false, colors=null, maxVal=null}={}) {
  const c = document.getElementById(id);
  if(!c) return; if(c._ch) c._ch.destroy();
  const maxV = Math.max(...vals, 1);
  c._ch = new Chart(c, {
    type:'bar',
    data:{ labels, datasets:[{
      data:vals,
      backgroundColor: colors||pal(vals.length, PAL_MAIN),
      borderRadius:horiz?4:5, borderSkipped:false, barThickness:'flex', maxBarThickness:32
    }]},
    options:{
      indexAxis: horiz?'y':'x', responsive:true, maintainAspectRatio:false,
      layout:{ padding: horiz ? {right:42} : {top:22} },
      plugins:{
        legend:{display:false},
        tooltip:{callbacks:{label:x=>' '+x.parsed[horiz?'x':'y']}},
        datalabels:{
          anchor: horiz ? 'end' : 'end',
          align:  horiz ? 'right' : 'top',
          offset: 3,
          color: function(ctx){
            if(horiz) return '#1a2540';
            return '#1a2540';
          },
          font:{ size:10, weight:'bold' },
          formatter: v => v,
          clamp: true,
        }
      },
      scales:{
        x:{ grid:{color:'#eef1f7'},
            ticks:{font:{size:10}, maxRotation:horiz?0:38,
              callback:function(v){ const l=this.getLabelForValue(v); return l&&l.length>26?l.slice(0,24)+'…':l }},
            ...(horiz ? {} : { max: maxVal||(maxV*1.18) }) },
        y:{ grid:{color:'#eef1f7'},
            ticks:{font:{size:10},
              callback:function(v){ const l=this.getLabelForValue(v); return l&&l.length>30?l.slice(0,28)+'…':l }},
            ...(horiz ? { max: maxVal||(maxV*1.18) } : {}) }
      }
    },
    plugins:[ChartDataLabels]
  });
}

function mkStacked(id, criterios, datasets) {
  // datasets: [{label:'EXCELENTE', data:[...], bg:'...'}, ...]
  const c = document.getElementById(id);
  if(!c) return; if(c._ch) c._ch.destroy();
  c._ch = new Chart(c, {
    type:'bar',
    data:{ labels:criterios, datasets: datasets.map(d=>({
      label: d.label, data: d.data,
      backgroundColor: d.bg, borderRadius:3, borderSkipped:false
    }))},
    options:{
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      layout:{ padding:{right:10} },
      scales:{
        x:{ stacked:true, grid:{color:'#eef1f7'},
            ticks:{font:{size:10}, callback: v => v+'%'} },
        y:{ stacked:true, grid:{display:false},
            ticks:{font:{size:10},
              callback: function(v){ const l=this.getLabelForValue(v); return l&&l.length>32?l.slice(0,30)+'…':l; }
            }
          }
      },
      plugins:{
        legend:{ position:'top', labels:{font:{size:11}, boxWidth:14} },
        tooltip:{ callbacks:{ label: x => ` ${x.dataset.label}: ${x.parsed.x}%` } },
        datalabels:{
          display: ctx => ctx.dataset.data[ctx.dataIndex] >= 8,
          color:'#fff', font:{size:10,weight:'bold'},
          formatter: v => v+'%', anchor:'center', align:'center'
        }
      }
    },
    plugins:[ChartDataLabels]
  });
}

function mkDoughnut(id, labels, vals, colors) {
  const c = document.getElementById(id);
  if(!c) return; if(c._ch) c._ch.destroy();
  const total = vals.reduce((a,b)=>a+b,0);
  c._ch = new Chart(c, {
    type:'doughnut',
    data:{ labels, datasets:[{ data:vals,
      backgroundColor:colors||pal(vals.length,PAL_MAIN),
      borderWidth:3, borderColor:'#fff', hoverBorderWidth:0
    }]},
    options:{
      responsive:true, maintainAspectRatio:false, cutout:'58%',
      layout:{ padding:10 },
      plugins:{
        legend:{position:'bottom', labels:{font:{size:10}, padding:10, boxWidth:12}},
        tooltip:{callbacks:{label:x=>` ${x.label}: ${x.parsed} (${Math.round(x.parsed/total*100)}%)`}},
        datalabels:{
          color:'#fff',
          font:{ size:11, weight:'bold' },
          textShadowColor:'rgba(0,0,0,.4)',
          textShadowBlur:4,
          formatter:(v, ctx)=>{
            const pct = Math.round(v/total*100);
            return pct >= 5 ? pct+'%\n'+v : '';
          },
          display:(ctx)=>{
            return ctx.dataset.data[ctx.dataIndex]/total > 0.04;
          }
        }
      }
    },
    plugins:[ChartDataLabels]
  });
}

function mkLine(id, labels, vals, color=C.blue) {
  const c = document.getElementById(id);
  if(!c) return; if(c._ch) c._ch.destroy();
  const maxV = Math.max(...vals, 1);
  c._ch = new Chart(c, {
    type:'line',
    data:{ labels, datasets:[{
      data:vals, borderColor:color, backgroundColor:color+'20',
      fill:true, tension:.35, pointRadius:5, pointHoverRadius:7,
      pointBackgroundColor:color, borderWidth:2.5
    }]},
    options:{
      responsive:true, maintainAspectRatio:false,
      layout:{ padding:{ top:24 } },
      plugins:{
        legend:{display:false},
        datalabels:{
          anchor:'top', align:'top', offset:4,
          color: color,
          font:{ size:10, weight:'bold' },
          formatter: v => v,
        }
      },
      scales:{
        x:{ grid:{color:'#eef1f7'}, ticks:{font:{size:10}, maxRotation:35} },
        y:{ grid:{color:'#eef1f7'}, ticks:{font:{size:10}}, beginAtZero:true, max: maxV*1.2 }
      }
    },
    plugins:[ChartDataLabels]
  });
}

// ── Filter state ─────────────────────────────────────────────────────────────
let fAnio='', fSem='', fMes='', fFac='', fProg='';
let fEstadoPract = '';

function setEstadoPract(val) {
  fEstadoPract = val;
  // Actualizar chips activos
  document.querySelectorAll('#chips-estado-pract .sec-filter-chip, .sec-filter-bar .sec-filter-chip').forEach(c => {
    c.classList.toggle('active', c.dataset.val === val || (val==='' && c.dataset.val===undefined));
  });
  renderPracticantes();
}

function initChipsEstado(estados) {
  const container = document.getElementById('chips-estado-pract');
  if(!container) return;
  container.innerHTML = '';
  estados.forEach(e => {
    const b = document.createElement('button');
    b.className = 'sec-filter-chip' + (fEstadoPract===e ? ' active' : '');
    b.textContent = e;
    b.dataset.val = e;
    b.onclick = () => setEstadoPract(e);
    container.appendChild(b);
  });
  // Marcar "Todos" correctamente
  const allBtn = document.querySelector('.sec-filter-bar .sec-filter-chip:not([data-val])');
  if(allBtn) { allBtn.dataset.val=''; allBtn.classList.toggle('active', fEstadoPract===''); }
}

function onFacultadChange() {
  fFac = document.getElementById('f-facultad').value;
  const pSel = document.getElementById('f-programa');
  pSel.innerHTML = '<option value="">Todos</option>';
  if(fFac && D.fac_prog[fFac]) {
    D.fac_prog[fFac].forEach(p=>{
      const o=document.createElement('option'); o.value=p; o.textContent=p; pSel.appendChild(o);
    });
  } else {
    D.all_programas.forEach(p=>{
      const o=document.createElement('option'); o.value=p; o.textContent=p; pSel.appendChild(o);
    });
  }
  fProg = '';
  pSel.value = '';
  onFilterChange();
}

function onFilterChange() {
  fAnio = document.getElementById('f-anio').value;
  fSem  = document.getElementById('f-sem').value;
  fMes  = document.getElementById('f-mes').value;
  fFac  = document.getElementById('f-facultad').value;
  fProg = document.getElementById('f-programa').value;
  const act = document.querySelector('.section.active');
  if(act) render(act.id.replace('sec-',''));
  updateInfo();
}

function resetFilters() {
  ['f-anio','f-sem','f-mes','f-facultad','f-programa'].forEach(id=>{
    document.getElementById(id).value='';
  });
  fAnio=fSem=fMes=fFac=fProg=''; fEstadoPract='';
  const pSel = document.getElementById('f-programa');
  pSel.innerHTML = '<option value="">Todos</option>';
  D.all_programas.forEach(p=>{
    const o=document.createElement('option'); o.value=p; o.textContent=p; pSel.appendChild(o);
  });
  const act = document.querySelector('.section.active');
  if(act) render(act.id.replace('sec-',''));
  updateInfo();
}

function updateInfo() {
  const parts=[];
  if(fAnio) parts.push('Año: '+fAnio);
  if(fSem)  parts.push(fSem);
  if(fMes)  { const m=D.filtros_meses.find(x=>x[0]==parseInt(fMes)); if(m) parts.push(m[1]); }
  if(fFac)  parts.push(fFac.replace('FACULTAD DE ',''));
  if(fProg) parts.push(fProg.length>28?fProg.slice(0,26)+'…':fProg);
  const el=document.getElementById('f-info');
  if(parts.length){ el.textContent='🔍 '+parts.join(' · '); el.style.display='inline-block'; }
  else { el.style.display='none'; }
}

// ── Filter rows ──────────────────────────────────────────────────────────────
function filterRows(rows, {progKey='PROGRAMA', facKey='FACULTAD'}={}) {
  return rows.filter(r=>{
    if(fAnio && String(r.ANIO) !== fAnio) return false;
    if(fSem  && r.SEMESTRE !== fSem)      return false;
    if(fMes  && String(r.MES) !== fMes)   return false;
    if(fFac  && facKey  && r[facKey]  !== fFac)  return false;
    if(fProg && progKey && r[progKey] !== fProg) return false;
    return true;
  });
}

// ── Aggregation helpers ──────────────────────────────────────────────────────
function groupBy(rows, key, n=0) {
  const m={};
  rows.forEach(r=>{
    const k=r[key];
    if(!k || !String(k).trim() || String(k).trim()==='0') return;
    m[k]=(m[k]||0)+1;
  });
  let pairs = Object.entries(m).sort((a,b)=>b[1]-a[1]);
  if(n>0) pairs=pairs.slice(0,n);
  return { labels:pairs.map(p=>p[0]), values:pairs.map(p=>p[1]) };
}

function semSort(rows) {
  const m={};
  rows.filter(r=>r.SEMESTRE&&r.SEMESTRE!=='Sin fecha').forEach(r=>{
    m[r.SEMESTRE]=(m[r.SEMESTRE]||0)+1;
  });
  const p=Object.entries(m).sort((a,b)=>a[0]<b[0]?-1:1);
  return { labels:p.map(x=>x[0]), values:p.map(x=>x[1]) };
}

function mesSort(rows) {
  const mo={Enero:1,Febrero:2,Marzo:3,Abril:4,Mayo:5,Junio:6,
    Julio:7,Agosto:8,Septiembre:9,Octubre:10,Noviembre:11,Diciembre:12};
  const m={};
  rows.filter(r=>r.MES&&r.MES>0).forEach(r=>{ m[r.MES_LABEL]=(m[r.MES_LABEL]||0)+1; });
  const p=Object.entries(m).sort((a,b)=>(mo[a[0]]||99)-(mo[b[0]]||99));
  return { labels:p.map(x=>x[0]), values:p.map(x=>x[1]) };
}

// ── KPI / hero helpers ────────────────────────────────────────────────────────
function heroStat(val, lbl, cls='') {
  return `<div class="hero-stat"><div class="hero-stat-val ${cls}">${val}</div><div class="hero-stat-label">${lbl}</div></div>`;
}
function kpiCard(val, lbl, sub='', cls='') {
  return `<div class="kpi ${cls}"><div class="kpi-num">${val}</div><div class="kpi-lbl">${lbl}</div>${sub?`<div class="kpi-sub">${sub}</div>`:''}</div>`;
}
function progressList(elId, rows, key, n=15) {
  const {labels,values}=groupBy(rows,key,n);
  const max=Math.max(...values,1);
  const html=labels.map((l,i)=>{
    const pct=Math.round(values[i]/max*100);
    const gold=i<3?'gold':'';
    return `<div class="prog-item">
      <div class="prog-item-head">
        <span class="prog-item-name" title="${l}">${l}</span>
        <span class="prog-item-val">${values[i]}</span>
      </div>
      <div class="prog-bar-bg"><div class="prog-bar-fill ${gold}" style="width:${pct}%"></div></div>
    </div>`;
  }).join('');
  document.getElementById(elId).innerHTML=html;
}

// ── Word cloud ────────────────────────────────────────────────────────────────
function renderWC(elId, words) {
  if(!words||!words.length){ document.getElementById(elId).innerHTML='<em style="color:#aaa">Sin datos</em>'; return; }
  const max=Math.max(...words.map(w=>w.count),1);
  const sizes=['xl','lg','md','sm','xs'];
  const html=words.map(w=>{
    const r=w.count/max;
    const cls=r>.7?'xl':r>.5?'lg':r>.3?'md':r>.15?'sm':'xs';
    return `<span class="chip ${cls}" title="${w.count} menciones">${w.word}<sup style="font-size:.6em;margin-left:2px">${w.count}</sup></span>`;
  }).join('');
  document.getElementById(elId).innerHTML=html;
}

function renderPills(pillId, wcId, dataObj) {
  const progs=Object.keys(dataObj);
  document.getElementById(pillId).innerHTML=progs.map((p,i)=>
    `<button class="prog-pill${i===0?' active':''}" onclick="pickPill(this,'${wcId}',${JSON.stringify(dataObj[p])})">${p}</button>`
  ).join('');
  if(progs.length) renderWC(wcId, dataObj[progs[0]]);
}

function pickPill(btn, wcId, words) {
  btn.closest('.prog-pills').querySelectorAll('.prog-pill').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  renderWC(wcId, words);
}

// ── Table EP ─────────────────────────────────────────────────────────────────
function renderTableEP(rows) {
  const ep={};
  rows.forEach(r=>{ if(r.PROGRAMA&&r.EMPRESA){ const k=r.PROGRAMA+'||'+r.EMPRESA; ep[k]=(ep[k]||0)+1; }});
  const arr=Object.entries(ep).map(([k,v])=>{const[p,e]=k.split('||');return{p,e,v};})
    .sort((a,b)=>b.v-a.v).slice(0,20);
  const rankCls=(i)=>i===0?'gold':i===1?'silver':'';
  let h='<table><thead><tr><th>#</th><th>Programa</th><th>Empresa</th><th>Solicitudes</th></tr></thead><tbody>';
  arr.forEach((r,i)=>{ h+=`<tr><td><span class="tbl-rank ${rankCls(i)}">${i+1}</span></td>
    <td>${r.p}</td><td>${r.e}</td>
    <td><span class="badge-val">${r.v}</span></td></tr>`; });
  h+='</tbody></table>';
  document.getElementById('tbl-ep').innerHTML=h;
}

// ── RENDERERS ────────────────────────────────────────────────────────────────
function renderPracticantes() {
  const allRows = filterRows(D.raw_practicantes);

  // Poblar chips de estado con los valores únicos del subconjunto actual
  const estados = [...new Set(allRows.map(r=>r.ESTADO).filter(Boolean))].sort();
  initChipsEstado(estados);

  // Aplicar filtro local de estado
  const rows = fEstadoPract ? allRows.filter(r=>r.ESTADO===fEstadoPract) : allRows;
  const n=rows.length;

  // Fecha de hoy (sin hora) para comparar con FECHA_FIN
  const hoy = new Date(); hoy.setHours(0,0,0,0);
  const fechaFin = r => r.FECHA_FIN_STR ? new Date(r.FECHA_FIN_STR) : null;
  // Fecha límite para rezagados: 6 meses atrás
  const hace6m = new Date(hoy); hace6m.setMonth(hace6m.getMonth() - 6);

  // ACTIVOS: fecha fin aún no ha llegado (o sin fecha)
  const activos     = rows.filter(r=>{ const f=fechaFin(r); return !f || f >= hoy; }).length;
  // FINALIZADOS PENDIENTE: fecha fin pasó pero hace menos de 6 meses
  const finalizados = rows.filter(r=>{ const f=fechaFin(r); return f && f < hoy && f >= hace6m; }).length;
  // REZAGADOS: fecha fin pasó hace más de 6 meses
  const rezagados   = rows.filter(r=>{ const f=fechaFin(r); return f && f < hace6m; }).length;
  // PENDIENTES: sin asesor asignado
  const pendientes  = rows.filter(r=>{ const f=fechaFin(r); return (!r.ASESOR||!r.ASESOR.trim()) && !(f && f < hace6m); }).length;

  const nuevas   = rows.filter(r=>r.EMPRESA_NUEVA&&r.EMPRESA_NUEVA.toLowerCase()==='si').length;
  const asesores = new Set(rows.map(r=>r.ASESOR).filter(Boolean)).size;
  const progs    = new Set(rows.map(r=>r.PROGRAMA).filter(Boolean)).size;
  const pctAct  = n?Math.round(activos/n*100):0;
  const pctFin  = n?Math.round(finalizados/n*100):0;
  const pctRez  = n?Math.round(rezagados/n*100):0;
  const pctPend = n?Math.round(pendientes/n*100):0;

  document.getElementById('hero-pract').innerHTML=
    heroStat(n,'Total','') +
    heroStat(activos,'Activos','green') +
    heroStat(finalizados,'Finaliz. Pendiente','') +
    heroStat(rezagados,'Rezagados','red') +
    heroStat(pendientes,'Pend. Asesor','gold');

  document.getElementById('kpi-pract').innerHTML=
    kpiCard(n,'Total practicantes','Período filtrado','') +
    kpiCard(activos,'Activos',`${pctAct}% · Fecha fin vigente`,'green') +
    kpiCard(finalizados,'Finalizados Pendiente',`${pctFin}% · Finalizaron hace < 6 meses`,'') +
    kpiCard(rezagados,'Rezagados',`${pctRez}% · Finalizaron hace > 6 meses`,'red') +
    kpiCard(pendientes,'Pendientes asignación',`${pctPend}% · Sin asesor asignado`,'gold') +
    kpiCard(nuevas,'Empresas nuevas','Primer convenio','') +
    kpiCard(asesores,'Asesores','Únicos asignados','') +
    kpiCard(progs,'Programas','Con estudiantes','purple');

  const top1Prog = groupBy(rows,'PROGRAMA',1).labels[0]||'N/A';
  const top1Cont = groupBy(rows,'TIPO_CONTRATO',1).labels[0]||'N/A';
  const top1Mod  = groupBy(rows,'MODALIDAD',1).labels[0]||'N/A';
  document.getElementById('stat-pract').innerHTML=
    `Total de practicantes registrados: <strong>${n}</strong>.
     Activos (fecha fin vigente): <strong>${activos} (${pctAct}%)</strong>.
     Finalizados pendiente (< 6 meses): <strong>${finalizados} (${pctFin}%)</strong>.
     Rezagados (> 6 meses desde fin): <strong>${rezagados} (${pctRez}%)</strong>.
     Pendientes por asignar asesor: <strong>${pendientes} (${pctPend}%)</strong>.
     Empresas nuevas vinculadas: <strong>${nuevas}</strong>.
     Programa con más estudiantes: <strong>${top1Prog}</strong>.
     Tipo de contrato predominante: <strong>${top1Cont}</strong>.
     Modalidad más frecuente: <strong>${top1Mod}</strong>.
     Asesores asignados: <strong>${asesores}</strong>.`;

  // Gráfica basada en lógica de fechas (mismos cálculos que los KPIs)
  const estLabels = ['Activos','Rezagados','Finalizados Pendiente','Pendientes asignación'];
  const estValues = [activos, rezagados, finalizados, pendientes];
  const estColors = [C.green, C.orange, C.blue3, C.gold];
  // Filtrar categorías con valor 0
  const estFilt = estLabels.reduce((a,l,i)=>{ if(estValues[i]>0){a.l.push(l);a.v.push(estValues[i]);a.c.push(estColors[i]);} return a; },{l:[],v:[],c:[]});
  document.getElementById('cb-p-estado').textContent=estFilt.l.length+' estados';
  mkDoughnut('c-p-estado', estFilt.l, estFilt.v, estFilt.c);

  const cont=groupBy(rows,'TIPO_CONTRATO');
  mkDoughnut('c-p-contrato', cont.labels, cont.values, PAL_BLUE);

  // Empresa nueva vs recurrente por año (barras agrupadas)
  (function(){
    const byAnio = {};
    rows.filter(r=>r.ANIO&&r.ANIO>0).forEach(r=>{
      const y=String(r.ANIO);
      if(!byAnio[y]) byAnio[y]={nueva:0,recurrente:0};
      if(r.EMPRESA_NUEVA&&r.EMPRESA_NUEVA.toLowerCase()==='si') byAnio[y].nueva++;
      else byAnio[y].recurrente++;
    });
    const anios=Object.keys(byAnio).sort();
    const c=document.getElementById('c-p-empnueva');
    if(!c) return; if(c._ch) c._ch.destroy();
    c._ch=new Chart(c,{
      type:'bar',
      data:{labels:anios, datasets:[
        {label:'Nueva',      data:anios.map(y=>byAnio[y].nueva),      backgroundColor:C.gold,   borderRadius:5, borderSkipped:false},
        {label:'Recurrente', data:anios.map(y=>byAnio[y].recurrente), backgroundColor:C.blue,   borderRadius:5, borderSkipped:false}
      ]},
      options:{
        responsive:true, maintainAspectRatio:false,
        layout:{padding:{top:24}},
        plugins:{
          legend:{position:'top', labels:{font:{size:11}, padding:14, boxWidth:14}},
          tooltip:{callbacks:{label:x=>` ${x.dataset.label}: ${x.parsed.y}`}},
          datalabels:{anchor:'end',align:'top',offset:2,font:{size:10,weight:'bold'},
            color:ctx=>ctx.datasetIndex===0?'#b47c00':'#102D69',
            formatter:v=>v>0?v:''}
        },
        scales:{
          x:{grid:{color:'#eef1f7'}, ticks:{font:{size:11}}},
          y:{grid:{color:'#eef1f7'}, ticks:{font:{size:11}}, beginAtZero:true,
             title:{display:true, text:'Empresas', font:{size:10}, color:'#4b5e7e'}}
        }
      },
      plugins:[ChartDataLabels]
    });
  })();

  const mod=groupBy(rows,'MODALIDAD');
  mkDoughnut('c-p-modalidad', mod.labels, mod.values, [C.blue, C.gold, C.green]);

  progressList('prog-list-pract', rows, 'PROGRAMA', 15);

  const fac=groupBy(rows,'FACULTAD');
  mkDoughnut('c-p-facultad', fac.labels, fac.values, facColorsArr(fac.labels));

  const sem=semSort(rows);
  mkLine('c-p-semestre', sem.labels, sem.values, C.blue);

  // Estudiantes por asesor apilado por estado (lógica de fechas)
  (function(){
    const byAsesor = {};
    rows.forEach(r=>{
      const a = (r.ASESOR&&r.ASESOR.trim()) ? r.ASESOR.trim() : 'Sin asesor';
      if(!byAsesor[a]) byAsesor[a]={activo:0,finalizado:0,rezagado:0};
      const f = r.FECHA_FIN_STR ? new Date(r.FECHA_FIN_STR) : null;
      if(!f || f >= hoy)                    byAsesor[a].activo++;
      else if(f >= hace6m)                  byAsesor[a].finalizado++;
      else                                  byAsesor[a].rezagado++;
    });
    const asLabels = Object.keys(byAsesor).sort((a,b)=>{
      const ta = byAsesor[a].activo+byAsesor[a].finalizado+byAsesor[a].rezagado;
      const tb = byAsesor[b].activo+byAsesor[b].finalizado+byAsesor[b].rezagado;
      return tb - ta;
    });
    const wrap = document.getElementById('wrap-p-asesor');
    if(wrap) wrap.style.width = Math.max(700, asLabels.length*80)+'px';
    const c = document.getElementById('c-p-asesor');
    if(!c) return; if(c._ch) c._ch.destroy();
    const META_ASESOR = 35;
    const totales = asLabels.map(a=>byAsesor[a].activo+byAsesor[a].finalizado+byAsesor[a].rezagado);
    const pcts = totales.map(t=>Math.round(t/META_ASESOR*100));
    const metaLine = asLabels.map(()=>META_ASESOR);
    c._ch = new Chart(c,{
      type:'bar',
      data:{
        labels: asLabels,
        datasets:[
          {label:'Activos',              data:asLabels.map(a=>byAsesor[a].activo),     backgroundColor:C.green,       borderRadius:0, stack:'s',
           datalabels:{anchor:'center', align:'center', font:{size:9,weight:'bold'}, color:'#fff', formatter:v=>v>0?v:null}},
          {label:'Finalizados Pendiente',data:asLabels.map(a=>byAsesor[a].finalizado), backgroundColor:C.blue3,       borderRadius:0, stack:'s',
           datalabels:{anchor:'center', align:'center', font:{size:9,weight:'bold'}, color:'#fff', formatter:v=>v>0?v:null}},
          {label:'Rezagados',            data:asLabels.map(a=>byAsesor[a].rezagado),   backgroundColor:C.orange,      borderRadius:0, stack:'s',
           datalabels:{anchor:'center', align:'center', font:{size:9,weight:'bold'}, color:'#fff', formatter:v=>v>0?v:null}},
          {label:'_total', data:totales, backgroundColor:'rgba(0,0,0,0)', borderWidth:0, stack:'s',
           datalabels:{anchor:'end', align:'top', offset:2, font:{size:10,weight:'bold'}, color:'#102D69',
             formatter:(v,ctx)=>{const p=pcts[ctx.dataIndex]; return v+' ('+p+'%)'}}},
          {label:'Meta (35)', data:metaLine, type:'line', borderColor:'#e74c3c', borderWidth:2,
           borderDash:[5,4], pointRadius:0, fill:false, stack:undefined,
           datalabels:{display:false}}
        ]
      },
      options:{
        responsive:true, maintainAspectRatio:false,
        layout:{padding:{top:32}},
        plugins:{
          legend:{position:'top', labels:{font:{size:11}, padding:14, boxWidth:14,
            filter: item => item.text !== '_total'}},
          tooltip:{callbacks:{label:x=>{
            if(x.dataset.label==='_total') return null;
            if(x.dataset.label==='Meta (35)') return ` Meta: ${META_ASESOR} estudiantes`;
            return ` ${x.dataset.label}: ${x.parsed.y}`;
          }, afterBody:(items)=>{
            const i=items[0]?.dataIndex;
            if(i==null) return [];
            return [`Cumplimiento: ${pcts[i]}% de meta (${META_ASESOR} est.)`];
          }}},
          datalabels:{}
        },
        scales:{
          x:{stacked:true, grid:{color:'#eef1f7'}, ticks:{font:{size:10}, maxRotation:45, minRotation:30}},
          y:{stacked:true, grid:{color:'#eef1f7'}, ticks:{font:{size:11}}, beginAtZero:true,
             title:{display:true, text:'Estudiantes', font:{size:10}, color:'#4b5e7e'}}
        }
      },
      plugins:[ChartDataLabels]
    });
  })();
}

function renderDisponibles() {
  const rows = filterRows(D.raw_disponibles);
  const n=rows.length;
  const disc=rows.filter(r=>r.DISCAPACIDAD&&r.DISCAPACIDAD.toUpperCase()==='SI').length;
  const progs=new Set(rows.map(r=>r.PROGRAMA).filter(Boolean)).size;
  const facs=new Set(rows.map(r=>r.FACULTAD).filter(Boolean)).size;

  document.getElementById('hero-disp').innerHTML=
    heroStat(n,'Disponibles','') +
    heroStat(disc,'Discapacidad','gold') +
    heroStat(progs,'Programas','') +
    heroStat(facs,'Facultades','green');

  document.getElementById('kpi-disp').innerHTML=
    kpiCard(n,'Estudiantes disponibles','Pendientes de ubicación','') +
    kpiCard(disc,'Con discapacidad',`${n?Math.round(disc/n*100):0}% del total`,'red') +
    kpiCard(n-disc,'Sin discapacidad','','green') +
    kpiCard(progs,'Programas','Representados','gold') +
    kpiCard(facs,'Facultades','','');

  const top1P=groupBy(rows,'PROGRAMA',1).labels[0]||'N/A';
  const top1M=groupBy(rows,'MODALIDAD',1).labels[0]||'N/A';
  document.getElementById('stat-disp').innerHTML=
    `Estudiantes en espera de ubicación: <strong>${n}</strong>.
     Programa con mayor demanda disponible: <strong>${top1P}</strong>.
     Estudiantes con algún tipo de discapacidad: <strong>${disc} (${n?Math.round(disc/n*100):0}%)</strong>.
     Modalidad más solicitada: <strong>${top1M}</strong>.
     Todos los registros tienen estado <strong>DISPONIBLE</strong>.`;

  const prog=groupBy(rows,'PROGRAMA',20);
  mkBar('c-d-programa', prog.labels, prog.values, {horiz:true});

  const fac=groupBy(rows,'FACULTAD');
  mkDoughnut('c-d-facultad', fac.labels, fac.values, facColorsArr(fac.labels));

  const dis=groupBy(rows,'DISCAPACIDAD');
  mkDoughnut('c-d-disc', dis.labels, dis.values, [C.blue, C.gold]);

  const drows=rows.filter(r=>r.DISCAPACIDAD&&r.DISCAPACIDAD.toUpperCase()==='SI');
  const td=groupBy(drows,'TIPO_DISCAPACIDAD');
  mkDoughnut('c-d-tipdisc', td.labels, td.values, PAL_GOLD);

  const mod=groupBy(rows,'MODALIDAD');
  mkDoughnut('c-d-modalidad', mod.labels, mod.values, [C.blue, C.green]);

  const mes=mesSort(rows);
  mkLine('c-d-mes', mes.labels, mes.values, C.gold);

  const anio=groupBy(rows.filter(r=>r.ANIO&&r.ANIO>0),'ANIO');
  mkBar('c-d-anio', anio.labels, anio.values, {colors:pal(anio.values.length, PAL_BLUE)});
}

function renderF082() {
  const rows = filterRows(D.raw_f082);
  const n=rows.length;
  const vinc=rows.filter(r=>r.VINCULADO&&r.VINCULADO.toLowerCase()==='si').length;
  const ent=rows.filter(r=>r.ENTREGADO&&r.ENTREGADO.toLowerCase().includes('enviado')).length;
  const proms=rows.map(r=>parseFloat(r.PROMEDIO)).filter(v=>!isNaN(v));
  const prom=proms.length?(proms.reduce((a,b)=>a+b,0)/proms.length).toFixed(2):'N/A';
  const pctV=n?Math.round(vinc/n*100):0;

  document.getElementById('hero-f082').innerHTML=
    heroStat(n,'Trabajos','') +
    heroStat(vinc,'Vinculados','green') +
    heroStat(pctV+'%','% Vinc.','gold') +
    heroStat(prom,'Prom.','');

  document.getElementById('kpi-f082').innerHTML=
    kpiCard(n,'Trabajos registrados','Total histórico','') +
    kpiCard(vinc,'Vinculados laboralmente',`${pctV}% del total`,'green') +
    kpiCard(ent,'Enviados a facultad','Estado: EnviadoFacultad','gold') +
    kpiCard(prom,'Promedio calificación','Escala 0–5','') +
    kpiCard(new Set(rows.map(r=>r.PROGRAMA).filter(Boolean)).size,'Programas','','purple');

  const top1P=groupBy(rows,'PROGRAMA',1).labels[0]||'N/A';
  document.getElementById('stat-f082').innerHTML=
    `Total de trabajos en el sistema: <strong>${n}</strong>.
     Estudiantes vinculados laboralmente al culminar la práctica: <strong>${vinc} (${pctV}%)</strong>.
     Promedio general de calificación: <strong>${prom}/5.0</strong>.
     Trabajos enviados a la facultad: <strong>${ent}</strong>.
     Programa con más entregas: <strong>${top1P}</strong>.`;

  const prog=groupBy(rows,'PROGRAMA',15);
  mkBar('c-f-programa', prog.labels, prog.values, {horiz:true});

  const fac=groupBy(rows,'FACULTAD');
  mkDoughnut('c-f-facultad', fac.labels, fac.values, facColorsArr(fac.labels));

  const vinv=groupBy(rows,'VINCULADO');
  mkDoughnut('c-f-vinculado', vinv.labels, vinv.values, [C.green, C.orange]);

  const entv=groupBy(rows,'ENTREGADO');
  mkDoughnut('c-f-entregado', entv.labels, entv.values, [C.blue, C.gold]);


  const hist=semSort(rows);
  mkLine('c-f-historico', hist.labels, hist.values, C.blue);

  // Histórico vinculación por año (barras agrupadas)
  const vincAnio = {};
  rows.filter(r=>r.ANIO&&r.ANIO>0).forEach(r=>{
    const y = String(r.ANIO);
    if(!vincAnio[y]) vincAnio[y]={si:0,no:0};
    const v = (r.VINCULADO||'').toLowerCase();
    if(v==='si') vincAnio[y].si++; else vincAnio[y].no++;
  });
  const vaYears = Object.keys(vincAnio).sort();
  const vaSi  = vaYears.map(y=>vincAnio[y].si);
  const vaNo  = vaYears.map(y=>vincAnio[y].no);
  (function(){
    const c = document.getElementById('c-f-vinc-anio');
    if(!c) return; if(c._ch) c._ch.destroy();
    const maxVinc = Math.max(...vaSi, ...vaNo, 1);
    c._ch = new Chart(c, {
      type:'bar',
      data:{ labels:vaYears, datasets:[
        { label:'Vinculado',    data:vaSi, backgroundColor:C.green,  borderRadius:5, borderSkipped:false },
        { label:'No vinculado', data:vaNo, backgroundColor:C.blue,   borderRadius:5, borderSkipped:false }
      ]},
      options:{
        responsive:true, maintainAspectRatio:false,
        layout:{ padding:{ top:26 } },
        plugins:{
          legend:{ position:'top', labels:{font:{size:11}, padding:14, boxWidth:14} },
          tooltip:{callbacks:{label:x=>` ${x.dataset.label}: ${x.parsed.y}`}},
          datalabels:{
            anchor:'end', align:'top', offset:2,
            font:{ size:10, weight:'bold' },
            color: ctx => ctx.datasetIndex===0 ? '#059669' : '#102D69',
            formatter: v => v > 0 ? v : '',
          }
        },
        scales:{
          x:{ grid:{color:'#eef1f7'}, ticks:{font:{size:11}} },
          y:{ grid:{color:'#eef1f7'}, ticks:{font:{size:11}}, beginAtZero:true,
              max: maxVinc * 1.2,
              title:{display:true, text:'Cantidad de estudiantes', font:{size:10}, color:'#4b5e7e'} }
        }
      },
      plugins:[ChartDataLabels]
    });
  })();

  // Poblar dropdowns facultad y programa
  const selFac082 = document.getElementById('sel-fac-f082');
  const selProg082 = document.getElementById('sel-prog-f082');
  if(selFac082 && selFac082.options.length <= 1) {
    D.f_facultades_list.forEach(f => {
      const o = document.createElement('option'); o.value = f; o.textContent = f; selFac082.appendChild(o);
    });
  }
  if(selProg082 && selProg082.options.length <= 1) {
    D.f_programas_list.forEach(p => {
      const o = document.createElement('option'); o.value = p; o.textContent = p; selProg082.appendChild(o);
    });
    selProg082.value = '';
  }
  drawAreasF082();

  // Poblar filtros locales del gráfico de asesor
  (function(){
    const selA = document.getElementById('sel-anio-asesor-f082');
    const selF = document.getElementById('sel-fac-asesor-f082');
    const selP = document.getElementById('sel-prog-asesor-f082');
    if(selA && selA.options.length <= 1){
      const anios = [...new Set(D.raw_f082.filter(r=>r.ANIO>0).map(r=>String(r.ANIO)))].sort();
      anios.forEach(y=>{ const o=document.createElement('option'); o.value=y; o.textContent=y; selA.appendChild(o); });
    }
    if(selF && selF.options.length <= 1){
      D.f_facultades_list.forEach(f=>{ const o=document.createElement('option'); o.value=f; o.textContent=f; selF.appendChild(o); });
    }
    if(selP && selP.options.length <= 1){
      D.f_programas_list.forEach(p=>{ const o=document.createElement('option'); o.value=p; o.textContent=p; selP.appendChild(o); });
    }
  })();
  renderAsesorF082();
}

function resetAsesorF082(){
  ['sel-anio-asesor-f082','sel-fac-asesor-f082','sel-prog-asesor-f082'].forEach(id=>{
    const el=document.getElementById(id); if(el) el.value='';
  });
  renderAsesorF082();
}

function renderAsesorF082(){
  const anio = (document.getElementById('sel-anio-asesor-f082')||{}).value||'';
  const fac  = (document.getElementById('sel-fac-asesor-f082') ||{}).value||'';
  const prog = (document.getElementById('sel-prog-asesor-f082')||{}).value||'';

  let rows = filterRows(D.raw_f082);
  if(anio) rows = rows.filter(r=>String(r.ANIO)===anio);
  if(fac)  rows = rows.filter(r=>r.FACULTAD===fac);
  if(prog) rows = rows.filter(r=>r.PROGRAMA===prog);

  const byAsesor = {};
  rows.filter(r=>r.ASESOR&&r.ASESOR.trim()).forEach(r=>{
    const a = r.ASESOR.trim();
    if(!byAsesor[a]) byAsesor[a]=0;
    byAsesor[a]++;
  });
  const labels = Object.keys(byAsesor).sort((a,b)=>byAsesor[b]-byAsesor[a]);
  const values = labels.map(a=>byAsesor[a]);

  const cnt = document.getElementById('f082-asesor-count');
  if(cnt) cnt.textContent = rows.length + ' trabajos · ' + labels.length + ' monitores';

  const wrap = document.getElementById('wrap-asesor-f082');
  if(!wrap) return;

  const maxV = Math.max(...values, 1);
  let html = '<table style="width:100%;border-collapse:collapse;font-size:.84rem">';
  html += '<thead><tr style="background:var(--itm-blue);color:#fff">'
        + '<th style="padding:8px 12px;text-align:center;width:48px;font-size:.72rem;text-transform:uppercase;letter-spacing:.3px">#</th>'
        + '<th style="padding:8px 12px;text-align:left;font-size:.72rem;text-transform:uppercase;letter-spacing:.3px">Monitor (Asesor de Prácticas)</th>'
        + '<th style="padding:8px 12px;text-align:left;font-size:.72rem;text-transform:uppercase;letter-spacing:.3px;min-width:200px">Avance</th>'
        + '<th style="padding:8px 12px;text-align:center;font-size:.72rem;text-transform:uppercase;letter-spacing:.3px">Trabajos</th>'
        + '<th style="padding:8px 12px;text-align:center;font-size:.72rem;text-transform:uppercase;letter-spacing:.3px">%</th>'
        + '</tr></thead><tbody>';

  labels.forEach((name, i) => {
    const v = values[i];
    const pct = Math.round(v / maxV * 100);
    const total = rows.length || 1;
    const pctTotal = (v / total * 100).toFixed(1);
    const bg = i % 2 === 0 ? '#f8fafd' : '#fff';
    const barColor = PAL_MAIN[i % PAL_MAIN.length];
    const rank = i + 1;
    const rankStyle = `font-weight:700;color:var(--itm-blue);font-size:.85rem`;

    html += `<tr style="background:${bg};border-bottom:1px solid #eef1f7">
      <td style="padding:9px 12px;text-align:center"><span style="${rankStyle}">${rank}</span></td>
      <td style="padding:9px 12px;font-weight:600;color:var(--text)">${name}</td>
      <td style="padding:9px 12px">
        <div style="background:#eef1f7;border-radius:6px;height:10px;overflow:hidden">
          <div style="width:${pct}%;height:100%;border-radius:6px;background:${barColor};transition:width .5s ease"></div>
        </div>
      </td>
      <td style="padding:9px 12px;text-align:center;font-weight:800;color:${barColor};font-size:.95rem">${v}</td>
      <td style="padding:9px 12px;text-align:center;font-size:.78rem;color:var(--text2)">${pctTotal}%</td>
    </tr>`;
  });

  html += '</tbody></table>';
  wrap.innerHTML = html;
}

function drawAreasF082() {
  const fac  = (document.getElementById('sel-fac-f082')  || {}).value || '';
  const prog = (document.getElementById('sel-prog-f082') || {}).value || '';

  let areas = [];
  let titulo = '';
  if(prog && D.f_areas_prog[prog]) {
    areas  = D.f_areas_prog[prog];
    titulo = prog;
  } else if(fac && D.f_areas_fac[fac]) {
    areas  = D.f_areas_fac[fac];
    titulo = fac;
  } else {
    // Sin selección: consolidar todas las áreas de todos los programas
    const global = {};
    Object.values(D.f_areas_prog).forEach(arr => {
      arr.forEach(a => { global[a.area] = (global[a.area]||0) + a.count; });
    });
    areas = Object.entries(global).sort((a,b)=>b[1]-a[1]).slice(0,15).map(([area,count])=>({area,count}));
    titulo = 'Todos los programas';
  }

  const labels = areas.map(a => a.area);
  const values = areas.map(a => a.count);

  const wrap = document.getElementById('wrap-areas-f082');
  if(wrap) wrap.style.height = Math.max(300, labels.length * 42 + 60) + 'px';

  // Colores degradados por ranking: primeros más intensos
  const baseColors = [C.blue,'#1a6eb5','#3485cc','#4f9de0',C.gold,C.teal,C.green,C.purple,C.orange,'#661081',
                      '#6bb3f0','#d4940a','#009030','#7c3aed','#F19800'];
  mkBar('c-areas-f082', labels, values, {horiz:true, colors:pal(labels.length, baseColors)});

  const el = document.getElementById('f082-wc-count');
  if(el) el.textContent = labels.length
    ? `${labels.length} competencias identificadas${titulo ? ' · ' + titulo : ''}`
    : 'Sin datos para la selección';
}

function onF082FacChange() {
  const fac     = (document.getElementById('sel-fac-f082') || {}).value || '';
  const selProg = document.getElementById('sel-prog-f082');
  if(!selProg) { drawAreasF082(); return; }
  selProg.innerHTML = '<option value="">Todos los programas</option>';
  const progs = fac ? (D.f_fac_to_prog[fac] || []) : D.f_programas_list;
  progs.forEach(p => {
    const o = document.createElement('option'); o.value = p; o.textContent = p; selProg.appendChild(o);
  });
  if(progs.length) selProg.value = progs[0];
  drawAreasF082();
}

function onF082ProgChange() {
  drawAreasF082();
}

function updateF082Count() { /* reemplazado por drawAreasF082 */ }

function resetSolicFiltros(){
  ['sel-anio-solic','sel-fac-solic','sel-prog-solic'].forEach(id=>{
    const el=document.getElementById(id); if(el) el.value='';
  });
  renderSolicitud();
}

function onSolicFacChange(){
  const fac = (document.getElementById('sel-fac-solic')||{}).value||'';
  const selP = document.getElementById('sel-prog-solic');
  if(selP){
    selP.innerHTML='<option value="">Todos</option>';
    const progs = fac
      ? [...new Set(D.raw_solicitud.filter(r=>r.FACULTAD===fac&&r.PROGRAMA).map(r=>r.PROGRAMA))].sort()
      : [...new Set(D.raw_solicitud.filter(r=>r.PROGRAMA).map(r=>r.PROGRAMA))].sort();
    progs.forEach(p=>{const o=document.createElement('option');o.value=p;o.textContent=p;selP.appendChild(o);});
  }
  renderSolicitud();
}

function getSolicRows(){
  const anio=(document.getElementById('sel-anio-solic')||{}).value||'';
  const fac =(document.getElementById('sel-fac-solic') ||{}).value||'';
  const prog=(document.getElementById('sel-prog-solic')||{}).value||'';
  let rows = filterRows(D.raw_solicitud, {progKey:'PROGRAMA', facKey:'FACULTAD'});
  if(anio) rows=rows.filter(r=>String(r.ANIO)===anio);
  if(fac)  rows=rows.filter(r=>r.FACULTAD===fac);
  if(prog) rows=rows.filter(r=>r.PROGRAMA===prog);
  return rows;
}

function renderSolicitud() {
  // Poblar filtros locales la primera vez
  (function(){
    const selA=document.getElementById('sel-anio-solic');
    const selF=document.getElementById('sel-fac-solic');
    if(selA&&selA.options.length<=1){
      const anios=[...new Set(D.raw_solicitud.filter(r=>r.ANIO>0).map(r=>String(r.ANIO)))].sort();
      anios.forEach(y=>{const o=document.createElement('option');o.value=y;o.textContent=y;selA.appendChild(o);});
    }
    if(selF&&selF.options.length<=1){
      const facs=[...new Set(D.raw_solicitud.filter(r=>r.FACULTAD).map(r=>r.FACULTAD))].sort();
      facs.forEach(f=>{const o=document.createElement('option');o.value=f;o.textContent=f;selF.appendChild(o);});
    }
    const selP=document.getElementById('sel-prog-solic');
    if(selP&&selP.options.length<=1){
      const progs=[...new Set(D.raw_solicitud.filter(r=>r.PROGRAMA).map(r=>r.PROGRAMA))].sort();
      progs.forEach(p=>{const o=document.createElement('option');o.value=p;o.textContent=p;selP.appendChild(o);});
    }
  })();

  const rows = getSolicRows();
  const n=rows.length;
  const nuevas=rows.filter(r=>r.EMPRESA_NUEVA&&r.EMPRESA_NUEVA.toLowerCase()==='si').length;
  const emps=new Set(rows.map(r=>r.EMPRESA).filter(Boolean)).size;
  const progs=new Set(rows.map(r=>r.PROGRAMA).filter(Boolean)).size;
  const facs=new Set(rows.map(r=>r.FACULTAD).filter(Boolean)).size;

  const cnt=document.getElementById('solic-count');
  if(cnt) cnt.textContent=`${n} solicitudes`;

  document.getElementById('hero-solic').innerHTML=
    heroStat(n,'Solicitudes','') +
    heroStat(emps,'Empresas','gold') +
    heroStat(nuevas,'Nuevas','green') +
    heroStat(progs,'Perfiles','');

  document.getElementById('kpi-solic').innerHTML=
    kpiCard(n,'Solicitudes recibidas','Total período','') +
    kpiCard(emps,'Empresas únicas','Que han solicitado','gold') +
    kpiCard(nuevas,'Empresas nuevas',`${n?Math.round(nuevas/n*100):0}% de solicitudes`,'green') +
    kpiCard(progs,'Perfiles solicitados','Programas demandados','purple') +
    kpiCard(facs,'Facultades','Con solicitudes','');

  const top1P=groupBy(rows,'PROGRAMA',1).labels[0]||'N/A';
  const top1M=groupBy(rows,'MODALIDAD',1).labels[0]||'N/A';
  document.getElementById('stat-solic').innerHTML=
    `Solicitudes registradas en el período: <strong>${n}</strong>.
     Empresas únicas: <strong>${emps}</strong>.
     Perfil más solicitado: <strong>${top1P}</strong>.
     Modalidad predominante: <strong>${top1M}</strong>.
     Empresas nuevas en el período: <strong>${nuevas} (${n?Math.round(nuevas/n*100):0}%)</strong>.`;

  const perf=groupBy(rows,'PROGRAMA',20);
  mkBar('c-s-perfil', perf.labels, perf.values, {horiz:true});

  const mod=groupBy(rows,'MODALIDAD');
  mkDoughnut('c-s-modalidad', mod.labels, mod.values, [C.blue, C.gold]);

  const en=groupBy(rows,'EMPRESA_NUEVA');
  mkDoughnut('c-s-empnueva', en.labels, en.values, [C.gold, C.blue]);

  const mes=mesSort(rows);
  mkLine('c-s-mes', mes.labels, mes.values, C.gold);

  const anioSorted=Object.entries(
    rows.filter(r=>r.ANIO&&r.ANIO>0).reduce((m,r)=>{m[r.ANIO]=(m[r.ANIO]||0)+1;return m},{})
  ).sort((a,b)=>Number(a[0])-Number(b[0]));
  mkLine('c-s-anio', anioSorted.map(x=>x[0]), anioSorted.map(x=>x[1]), C.blue);

  // ── Gráfica combinada Solicitud vs Aprobación ──────────────────────────────
  (function(){
    const data = D.sol_vs_aprob || [];
    const labels = data.map(d=>String(d.anio));
    const sol    = data.map(d=>d.solicitudes);
    const apr    = data.map(d=>d.aprobaciones);
    const pct    = data.map(d=>d.pct);

    const c = document.getElementById('c-s-vs-aprob');
    if(!c) return; if(c._ch) c._ch.destroy();
    c._ch = new Chart(c, {
      data:{
        labels,
        datasets:[
          { type:'bar', label:'Solicitudes empresariales', data:sol,
            backgroundColor:'rgba(0,83,155,0.75)', borderRadius:5, borderSkipped:false,
            yAxisID:'yLeft', order:2 },
          { type:'bar', label:'Aprobaciones de funciones', data:apr,
            backgroundColor:'rgba(232,160,0,0.85)', borderRadius:5, borderSkipped:false,
            yAxisID:'yLeft', order:2 },
          { type:'line', label:'% Respuesta institucional', data:pct,
            borderColor:'#059669', backgroundColor:'rgba(5,150,105,0.12)',
            borderWidth:2.5, pointRadius:6, pointBackgroundColor:'#059669',
            fill:true, tension:0.35,
            yAxisID:'yRight', order:1,
            datalabels:{ anchor:'top', align:'top', offset:6,
              font:{size:11,weight:'bold'}, color:'#059669',
              formatter:v=>v+'%' } }
        ]
      },
      options:{
        responsive:true, maintainAspectRatio:false,
        layout:{padding:{top:30}},
        interaction:{mode:'index', intersect:false},
        plugins:{
          legend:{position:'top', labels:{font:{size:11}, padding:14, boxWidth:14}},
          tooltip:{callbacks:{
            label:x=> x.dataset.yAxisID==='yRight'
              ? ` % Respuesta: ${x.parsed.y}%`
              : ` ${x.dataset.label}: ${x.parsed.y.toLocaleString()}`
          }},
          datalabels:{
            display: ctx => ctx.dataset.type==='line',
            anchor:'top', align:'top', offset:4,
            font:{size:11,weight:'bold'}, color:'#059669',
            formatter:v=>v+'%'
          }
        },
        scales:{
          yLeft:{ type:'linear', position:'left', beginAtZero:true,
            grid:{color:'#eef1f7'}, ticks:{font:{size:11}},
            title:{display:true, text:'Cantidad', font:{size:10}, color:'#4b5e7e'} },
          yRight:{ type:'linear', position:'right', beginAtZero:true, max:110,
            grid:{drawOnChartArea:false}, ticks:{font:{size:11}, callback:v=>v+'%'},
            title:{display:true, text:'% Respuesta', font:{size:10}, color:'#059669'} },
          x:{ grid:{color:'#eef1f7'}, ticks:{font:{size:12}} }
        }
      },
      plugins:[ChartDataLabels]
    });

    // Análisis ejecutivo dinámico
    const maxPct = Math.max(...pct);
    const minPct = Math.min(...pct);
    const lastY  = data[data.length-1] || {};
    const firstY = data[0] || {};
    const trend  = pct.length > 1 ? (pct[pct.length-1] > pct[0] ? 'al alza' : 'a la baja') : 'estable';
    const el = document.getElementById('analisis-ejecutivo');
    if(el) el.innerHTML = `
      <strong style="color:var(--itm-blue);font-size:.9rem">📊 Análisis Ejecutivo — Demanda Empresarial y Capacidad de Atención</strong><br><br>
      Durante el período analizado, la Oficina de Prácticas Profesionales ITM recibió un total de
      <strong>${sol.reduce((a,b)=>a+b,0).toLocaleString()} solicitudes empresariales</strong>,
      de las cuales <strong>${apr.reduce((a,b)=>a+b,0).toLocaleString()} llegaron a la etapa de aprobación de funciones</strong>,
      representando una cobertura global del
      <strong>${Math.round(apr.reduce((a,b)=>a+b,0)/sol.reduce((a,b)=>a+b,0)*100)}%</strong>
      frente a la demanda recibida.<br><br>
      ${data.map(d=>`En <strong>${d.anio}</strong>, se registraron <strong>${d.solicitudes.toLocaleString()} solicitudes</strong>
      y <strong>${d.aprobaciones.toLocaleString()} aprobaciones</strong>, con una tasa de respuesta del
      <strong style="color:#059669">${d.pct}%</strong>.`).join(' ')}<br><br>
      La tendencia del porcentaje de respuesta institucional es <strong>${trend}</strong>,
      con un pico de <strong>${maxPct}%</strong> y un mínimo de <strong>${minPct}%</strong>.
      ${trend === 'al alza'
        ? 'Esto refleja una <strong>mejora progresiva en la capacidad de atención</strong> de la Oficina de Prácticas frente a la demanda empresarial.'
        : 'Esto sugiere que la demanda empresarial crece a un ritmo que <strong>supera la capacidad actual de respuesta</strong>, lo que representa una oportunidad de fortalecimiento institucional.'
      }
      La oferta de estudiantes disponibles frente a la demanda activa debe monitorearse para garantizar
      la <strong>cobertura efectiva de vacantes</strong> y el aprovechamiento de los convenios vigentes.
    `;
  })();
}

function renderAprobacion() {
  const rows = filterRows(D.raw_aprobacion, {facKey:null});
  const n=rows.length;
  const aprob=rows.filter(r=>r.ESTADO_APROBACION&&r.ESTADO_APROBACION.toLowerCase()==='aprobado').length + 1;
  const pend=rows.filter(r=>r.ESTADO_APROBACION&&r.ESTADO_APROBACION.toLowerCase()==='pendiente').length;
  const progs=new Set(rows.map(r=>r.PROGRAMA).filter(Boolean)).size;
  const emps=new Set(rows.map(r=>r.EMPRESA).filter(Boolean)).size;
  const pctA=n?Math.round(aprob/n*100):0;

  document.getElementById('hero-aprob').innerHTML=
    heroStat(n,'Solicitudes','') +
    heroStat(aprob,'Aprobadas','green') +
    heroStat(pend,'Pendientes','gold') +
    heroStat(pctA+'%','Tasa Aprob.','');

  document.getElementById('kpi-aprob').innerHTML=
    kpiCard(n,'Solicitudes de funciones','Total período','') +
    kpiCard(aprob,'Aprobadas',`${pctA}% del total`,'green') +
    kpiCard(pend,'Pendientes de revisión','','gold') +
    kpiCard(progs,'Programas','Que han solicitado','') +
    kpiCard(emps,'Empresas','Únicas','purple');

  const top1P=groupBy(rows,'PROGRAMA',1).labels[0]||'N/A';
  const top1E=groupBy(rows,'EMPRESA',1).labels[0]||'N/A';
  document.getElementById('stat-aprob').innerHTML=
    `Solicitudes de aprobación de funciones: <strong>${n}</strong>.
     Solicitudes aprobadas: <strong>${aprob} (${pctA}%)</strong>.
     Pendientes de revisión: <strong>${pend}</strong>.
     Programa con más solicitudes: <strong>${top1P}</strong>.
     Empresa más activa: <strong>${top1E}</strong>.`;

  const prog=groupBy(rows,'PROGRAMA',20);
  mkBar('c-a-programa', prog.labels, prog.values, {horiz:true});

  mkDoughnut('c-a-estado', ['Aprobado','Pendiente'], [aprob, pend], [C.green, C.gold]);

  const emp=groupBy(rows,'EMPRESA',15);
  mkBar('c-a-empresa', emp.labels, emp.values, {horiz:true, colors:pal(emp.labels.length, PAL_GOLD)});

  renderTableEP(rows);
  initAprobFiltros();

  // Demanda global — barra horizontal fija (siempre global)
  const gkw = D.a_funciones_kw;
  const gwrap = document.getElementById('c-aprob-global');
  if(gwrap && gwrap.parentElement)
    gwrap.parentElement.style.height = Math.max(300, gkw.length * 36 + 60) + 'px';
  mkBar('c-aprob-global', gkw.map(w=>w.word), gkw.map(w=>w.count),
    {horiz:true, colors:pal(gkw.length, PAL_BLUE)});
}

function initAprobFiltros() {
  const selFac  = document.getElementById('sel-fac-aprob');
  const selProg = document.getElementById('sel-prog-aprob');
  if(selFac && selFac.options.length <= 1) {
    D.all_facultades.forEach(f=>{
      const o=document.createElement('option'); o.value=f; o.textContent=f; selFac.appendChild(o);
    });
  }
  if(selProg && selProg.options.length <= 1) {
    Object.keys(D.a_funciones_programa).sort().forEach(p=>{
      const o=document.createElement('option'); o.value=p; o.textContent=p; selProg.appendChild(o);
    });
    selProg.value='';
  }
  drawAprobFunciones();
}

function drawAprobFunciones() {
  const fac  = (document.getElementById('sel-fac-aprob')  || {}).value || '';
  const prog = (document.getElementById('sel-prog-aprob') || {}).value || '';

  let words = [], titulo = '';
  if(prog && D.a_funciones_programa[prog]) {
    words  = D.a_funciones_programa[prog];
    titulo = prog;
  } else if(fac && D.a_funciones_fac && D.a_funciones_fac[fac]) {
    words  = D.a_funciones_fac[fac];
    titulo = fac;
  } else {
    words  = D.a_funciones_kw;
    titulo = 'Todos los programas';
  }

  const labels = words.map(w=>w.word);
  const values = words.map(w=>w.count);
  const wrap   = document.getElementById('wrap-aprob-func');
  if(wrap) wrap.style.height = Math.max(300, labels.length * 42 + 60) + 'px';
  mkBar('c-aprob-func', labels, values, {horiz:true, colors:pal(labels.length, PAL_BLUE)});

  const el = document.getElementById('aprob-kw-count');
  if(el) el.textContent = labels.length
    ? `${labels.length} términos clave · ${titulo.length>50?titulo.slice(0,48)+'…':titulo}`
    : 'Sin datos para la selección';
}

function onAprobFacChange() {
  const fac     = (document.getElementById('sel-fac-aprob') || {}).value || '';
  const selProg = document.getElementById('sel-prog-aprob');
  if(!selProg) { drawAprobFunciones(); return; }
  selProg.innerHTML = '<option value="">Todos los programas</option>';
  const progs = fac
    ? (D.fac_prog[fac] || []).filter(p => D.a_funciones_programa[p])
    : Object.keys(D.a_funciones_programa).sort();
  progs.forEach(p=>{
    const o=document.createElement('option'); o.value=p; o.textContent=p; selProg.appendChild(o);
  });
  selProg.value='';
  drawAprobFunciones();
}

function onAprobProgChange() { drawAprobFunciones(); }

// ── Encuesta helpers ──────────────────────────────────────────────────────────
function cleanLabel(s) {
  s = String(s).trim().replace(/^[A-Za-z]\.\s+/, '');
  if (!s) return s;
  return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();
}
function parseMulti(rows, key, top=12) {
  const c={};
  rows.forEach(r=>{
    if(!r[key]) return;
    r[key].split(';').forEach(p=>{
      p=cleanLabel(p.trim());
      if(p && p.toLowerCase()!=='nan') c[p]=(c[p]||0)+1;
    });
  });
  let pairs=Object.entries(c).sort((a,b)=>b[1]-a[1]);
  if(top>0) pairs=pairs.slice(0,top);
  return {labels:pairs.map(p=>p[0]), values:pairs.map(p=>p[1])};
}

function countCalif(rows, key) {
  const order=['EXCELENTE','BUENO','REGULAR','DEFICIENTE'];
  const c={};
  rows.forEach(r=>{
    const v=(r[key]||'').toUpperCase().trim();
    if(v) c[v]=(c[v]||0)+1;
  });
  return {order: order.filter(o=>c[o]), counts: order.reduce((a,o)=>{a[o]=c[o]||0;return a;},{})};
}

function buildCalifDatasets(rows, criterioKeys, criterioLabels) {
  const order=['EXCELENTE','BUENO','DEFICIENTE'];
  const colores={EXCELENTE:'#10b981', BUENO:'#00539B', REGULAR:'#E8A000', DEFICIENTE:'#ef4444'};
  const totales=criterioKeys.map(k=>{
    let t=0; rows.forEach(r=>{ if(r[k] && r[k].toUpperCase().trim()) t++; }); return t||1;
  });
  return order.map(cat=>({
    label: cat,
    bg: colores[cat],
    data: criterioKeys.map((k,i)=>{
      let cnt=0; rows.forEach(r=>{ if((r[k]||'').toUpperCase().trim()===cat) cnt++; });
      return Math.round(cnt/totales[i]*100);
    })
  }));
}

function renderEncuesta() {
  const rows = filterRows(D.enc_raw, {progKey:'PROGRAMA', facKey:null});
  const n = rows.length;
  const emps = new Set(rows.map(r=>r.empresa).filter(Boolean)).size;
  const vinSi = rows.filter(r=>(r.vinculacionpracticantes||'').toUpperCase()==='SI').length;
  const pctVin = n ? Math.round(vinSi/n*100) : 0;
  const sectores = new Set(rows.map(r=>r.sector).filter(Boolean)).size;

  document.getElementById('hero-enc').innerHTML=
    heroStat(n,'Encuestas','') +
    heroStat(emps,'Empresas','') +
    heroStat(pctVin+'%','Vincularían','green') +
    heroStat(sectores,'Sectores','gold');

  document.getElementById('kpi-enc').innerHTML=
    kpiCard(n,'Encuestas respondidas','Total período','') +
    kpiCard(emps,'Empresas únicas','Participantes','purple') +
    kpiCard(vinSi,'Vincularían practicantes',`${pctVin}% dispuestas`,'green') +
    kpiCard(n-vinSi,'No vincularían',`${100-pctVin}% del total`,'red') +
    kpiCard(sectores,'Sectores económicos','Representados','gold');

  const top1Sec = groupBy(rows,'sector',1).labels[0]||'N/A';
  const top1Prog= groupBy(rows,'programa',1).labels[0]||'N/A';
  document.getElementById('stat-enc').innerHTML=
    `Se analizaron <strong>${n}</strong> encuestas de <strong>${emps}</strong> empresas únicas.
     El <strong>${pctVin}%</strong> manifestó disposición para vincular practicantes.
     Sector más representado: <strong>${top1Sec}</strong>.
     Programa con mayor participación: <strong>${top1Prog}</strong>.`;

  // Calificaciones apiladas (stacked bar %)
  const criterioKeys  = ['IMPACTO_SOCIAL','CALIF_FORMACION_EG','CALIF_DESEMPENO_EG','CALIF_FORMACION_PRAC','CALIF_DESEMPENO_PRAC'];
  const criterioLabels= ['Impacto social ITM','Formación egresado','Desempeño egresado','Formación practicante','Desempeño practicante'];
  const datasets = buildCalifDatasets(rows, criterioKeys, criterioLabels);
  mkStacked('c-enc-calif', criterioLabels, datasets);

  // Vinculación
  const vinc = groupBy(rows,'vinculacionpracticantes');
  mkDoughnut('c-enc-vinculacion', vinc.labels, vinc.values, [C.green, C.orange, C.gold]);

  // Tipo empresa
  const tipo = groupBy(rows,'tipo');
  mkDoughnut('c-enc-tipo', tipo.labels, tipo.values, [C.blue, C.gold, C.teal, C.purple]);

  // Mercado
  const merc = groupBy(rows,'mercado');
  mkDoughnut('c-enc-mercado', merc.labels, merc.values, [C.blue3, C.gold, C.green]);

  // Sector
  const sec = groupBy(rows,'sector',15);
  mkBar('c-enc-sector', sec.labels, sec.values, {horiz:true, colors:pal(sec.labels.length, PAL_MAIN)});

  // Fortalezas y debilidades practicante
  const fp = parseMulti(rows,'FORTALEZAS_PRAC');
  mkBar('c-enc-fort-prac', fp.labels, fp.values, {horiz:true, colors:pal(fp.labels.length,PAL_BLUE)});

  const dp = parseMulti(rows,'DEBILIDADES_PRAC');
  mkBar('c-enc-deb-prac', dp.labels, dp.values, {horiz:true, colors:pal(dp.labels.length,['#475569','#64748b','#334155','#94a3b8','#1e293b'])});

  // Fortalezas y debilidades egresado
  const fe = parseMulti(rows,'FORTALEZAS_EG');
  mkBar('c-enc-fort-eg', fe.labels, fe.values, {horiz:true, colors:pal(fe.labels.length,PAL_BLUE)});

  const de = parseMulti(rows,'DEBILIDADES_EG');
  mkBar('c-enc-deb-eg', de.labels, de.values, {horiz:true, colors:pal(de.labels.length,['#475569','#64748b','#334155','#94a3b8','#1e293b'])});

  // Competencias demandadas
  const comp = parseMulti(rows,'COMPETENCIAS');
  mkBar('c-enc-competencias', comp.labels, comp.values, {horiz:true, colors:pal(comp.labels.length,PAL_GOLD)});

  // Evolución por año
  const anios = groupBy(rows,'ANIO');
  const aniosSorted = [...anios.labels.map((l,i)=>({l,v:anios.values[i]}))].sort((a,b)=>a.l-b.l);
  mkLine('c-enc-anio', aniosSorted.map(x=>x.l), aniosSorted.map(x=>x.v), C.gold);
}

// ── Encuesta Estudiantes ──────────────────────────────────────────────────────
const EST_CALIF_KEYS   = ['CALIF_FORMACION','CALIF_ORIENTACION','CALIF_PERTINENCIA',
                           'CALIF_TUTOR','CALIF_DESEMPENO','CALIF_PUESTO',
                           'CALIF_ASESOR_ITM','CALIF_SEGUIMIENTO','CALIF_INFORMACION',
                           'SATISFACCION_GENERAL'];
const EST_CALIF_LABELS = ['Formación académica','Orientación profesional','Pertinencia práctica',
                           'Tutor empresarial','Desempeño de tareas','Puesto de trabajo',
                           'Asesor Oficina Prácticas','Seguimiento Oficina','Información normatividad',
                           'Satisfacción general'];
const EST_SCORE_MAP = {excelente:3,bueno:2,deficiente:1};

function estScore(rows, keys) {
  let t=0, n=0;
  rows.forEach(r=>keys.forEach(k=>{
    const v=EST_SCORE_MAP[(r[k]||'').toLowerCase().trim()];
    if(v){t+=v;n++;}
  }));
  return n ? t/n : 0;
}

function scoreBarColor(s) {
  if(s>=2.7) return C.green;
  if(s>=2.3) return C.teal;
  if(s>=1.8) return C.gold;
  return C.orange;
}

function mkBarScore(id, labels, scores, counts) {
  const c=document.getElementById(id);
  if(!c) return; if(c._ch) c._ch.destroy();
  const maxV=Math.max(...scores,1);
  c._ch=new Chart(c,{
    type:'bar',
    data:{labels, datasets:[{
      data:scores,
      backgroundColor: scores.map(s=>scoreBarColor(s)),
      borderRadius:4, borderSkipped:false, barThickness:'flex', maxBarThickness:28
    }]},
    options:{
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      layout:{padding:{right:56}},
      scales:{
        x:{min:1,max:3,grid:{color:'#eef1f7'},
           ticks:{font:{size:10}, callback:v=>v.toFixed(1)},
           title:{display:true,text:'Índice (1=Deficiente · 2=Bueno · 3=Excelente)',font:{size:9},color:'#4b5e7e'}},
        y:{grid:{display:false},ticks:{font:{size:10},
           callback:function(v){const l=this.getLabelForValue(v);return l&&l.length>32?l.slice(0,30)+'…':l;}}}
      },
      plugins:{
        legend:{display:false},
        tooltip:{callbacks:{label:x=>{
          const i=x.dataIndex;
          return ` Score: ${x.parsed.x.toFixed(2)}  |  n=${counts[i]} resp.`;
        }}},
        datalabels:{anchor:'end',align:'right',offset:4,
          color:'#1a2540',font:{size:10,weight:'bold'},
          formatter:(v,ctx)=>{
            const i=ctx.dataIndex;
            return v.toFixed(2)+(counts&&counts[i]?' ('+counts[i]+')'  :'');
          },clamp:true}
      }
    },
    plugins:[ChartDataLabels]
  });
}

function renderEncuestaEst() {
  const rows = filterRows(D.est_raw, {facKey:null});
  const n = rows.length;
  const recSi = rows.filter(r=>(r.RECOMIENDA||'').toUpperCase()==='SI').length;
  const pctRec = n ? Math.round(recSi/n*100) : 0;
  const excGen = rows.filter(r=>(r.SATISFACCION_GENERAL||'').toLowerCase()==='excelente').length;
  const pctExc = n ? Math.round(excGen/n*100) : 0;
  const progs  = new Set(rows.map(r=>r.PROGRAMA).filter(Boolean)).size;

  document.getElementById('hero-est').innerHTML=
    heroStat(n,'Respuestas','') +
    heroStat(pctRec+'%','Recomendarían','green') +
    heroStat(pctExc+'%','Satisf. Excelente','gold') +
    heroStat(progs,'Programas','');

  document.getElementById('kpi-est').innerHTML=
    kpiCard(n,'Encuestas respondidas','Total período','') +
    kpiCard(recSi,'Recomendarían la práctica',`${pctRec}% del total`,'green') +
    kpiCard(excGen,'Satisfacción general Excelente',`${pctExc}% de respuestas`,'gold') +
    kpiCard(n-recSi,'No recomendarían',`${100-pctRec}% del total`,'') +
    kpiCard(progs,'Programas participantes','Representados','purple');

  const top1Prog = groupBy(rows,'PROGRAMA',1).labels[0]||'N/A';
  const top1Mod  = groupBy(rows,'MODALIDAD',1).labels[0]||'N/A';
  const scoreGlobal = estScore(rows, EST_CALIF_KEYS);
  document.getElementById('stat-est').innerHTML=
    `Se analizaron <strong>${n}</strong> encuestas de estudiantes.
     El <strong>${pctRec}%</strong> recomendaría realizar la práctica profesional.
     El índice de satisfacción global es <strong>${scoreGlobal.toFixed(2)}/3.00</strong>.
     Programa con más respuestas: <strong>${top1Prog}</strong>.
     Modalidad predominante: <strong>${top1Mod}</strong>.`;

  // ── Calificaciones comparadas (stacked %)
  const datasets = buildCalifDatasets(rows, EST_CALIF_KEYS, EST_CALIF_LABELS);
  mkStacked('c-est-calif', EST_CALIF_LABELS, datasets);

  // ── Satisfacción general
  const satisf = groupBy(rows,'SATISFACCION_GENERAL');
  const satisfOrder = ['Excelente','Bueno','Deficiente'];
  const satisfSorted = satisfOrder.filter(o=>satisf.labels.includes(o));
  mkDoughnut('c-est-satisf', satisfSorted,
    satisfSorted.map(o=>satisf.values[satisf.labels.indexOf(o)]),
    [C.green, C.blue, C.orange]);

  // ── Recomendaría
  const rec = groupBy(rows,'RECOMIENDA');
  mkDoughnut('c-est-rec', rec.labels, rec.values, [C.green, C.orange]);

  // ── Modalidad
  const mod = groupBy(rows,'MODALIDAD');
  mkDoughnut('c-est-modal', mod.labels, mod.values, [C.blue, C.gold, C.teal]);

  // ── Proyecto a futuro
  const fut = groupBy(rows,'PROYECTO_FUTURO');
  mkBar('c-est-futuro', fut.labels, fut.values,
    {horiz:true, colors:pal(fut.labels.length,[C.blue,C.gold,C.green,C.purple,C.teal])});

  // ── Score por programa (dinámico desde rows)
  const progMap={};
  rows.forEach(r=>{
    const p=r.PROGRAMA; if(!p) return;
    if(!progMap[p]) progMap[p]={t:0,n:0,cnt:0};
    EST_CALIF_KEYS.forEach(k=>{
      const v=EST_SCORE_MAP[(r[k]||'').toLowerCase().trim()];
      if(v){progMap[p].t+=v;progMap[p].n++;}
    });
    progMap[p].cnt++;
  });
  const progScores=Object.entries(progMap)
    .map(([p,d])=>({p,s:d.n?d.t/d.n:0,cnt:d.cnt}))
    .filter(x=>x.cnt>=3)
    .sort((a,b)=>b.s-a.s);
  const psLabels=progScores.map(x=>x.p);
  const psValues=progScores.map(x=>parseFloat(x.s.toFixed(2)));
  const psCounts=progScores.map(x=>x.cnt);
  const wrapPS=document.getElementById('wrap-est-prog-score');
  if(wrapPS) wrapPS.style.height=Math.max(300,psLabels.length*36+60)+'px';
  mkBarScore('c-est-prog-score', psLabels, psValues, psCounts);

  // ── Programas con más respuestas
  const prog=groupBy(rows,'PROGRAMA',20);
  const wrapP=document.getElementById('wrap-est-prog');
  if(wrapP) wrapP.style.width=Math.max(700,prog.labels.length*62)+'px';
  mkBar('c-est-prog', prog.labels, prog.values, {horiz:false});

  // ── Top 15 empresas
  const emp=groupBy(rows,'EMPRESA',15);
  mkBar('c-est-empresa', emp.labels, emp.values,
    {horiz:true, colors:pal(emp.labels.length, PAL_MAIN)});

  // ── Score por asesor (dinámico desde rows)
  const asMap={};
  rows.forEach(r=>{
    const a=r.ASESOR_NOMBRE; if(!a||a.length<4) return;
    if(!asMap[a]) asMap[a]={t:0,n:0,cnt:0};
    EST_CALIF_KEYS.forEach(k=>{
      const v=EST_SCORE_MAP[(r[k]||'').toLowerCase().trim()];
      if(v){asMap[a].t+=v;asMap[a].n++;}
    });
    asMap[a].cnt++;
  });
  const asScores=Object.entries(asMap)
    .map(([a,d])=>({a,s:d.n?d.t/d.n:0,cnt:d.cnt}))
    .filter(x=>x.cnt>=2)
    .sort((a,b)=>b.s-a.s)
    .slice(0,20);
  const asLabels=asScores.map(x=>x.a);
  const asValues=asScores.map(x=>parseFloat(x.s.toFixed(2)));
  const asCounts=asScores.map(x=>x.cnt);
  const wrapAS=document.getElementById('c-est-asesor-score');
  if(wrapAS&&wrapAS.parentElement)
    wrapAS.parentElement.style.height=Math.max(300,asLabels.length*36+60)+'px';
  mkBarScore('c-est-asesor-score', asLabels, asValues, asCounts);

  // ── Evolución semestral
  const sem=semSort(rows);
  mkLine('c-est-semestre', sem.labels, sem.values, C.blue);
}

// ── Navigation ────────────────────────────────────────────────────────────────
function goTo(name, btn) {
  document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(b=>b.classList.remove('active'));
  document.getElementById('sec-'+name).classList.add('active');
  if(btn) btn.classList.add('active');
  render(name);
}

function render(name) {
  if(name==='practicantes') renderPracticantes();
  else if(name==='disponibles') renderDisponibles();
  else if(name==='f082') renderF082();
  else if(name==='solicitud') renderSolicitud();
  else if(name==='aprobacion') renderAprobacion();
  else if(name==='encuesta') renderEncuesta();
  else if(name==='encuesta-est') renderEncuestaEst();
  else if(name==='cuee') renderCuee();
  else if(name==='historicos') renderHistoricos();
}

// ── CUEE ──────────────────────────────────────────────────────────────────────
function renderCuee() {
  // ── Datos estáticos 2026-1 ──
  const INSCR=319, CUPOS=173, PARTIC=132, EMPRESAS=8;
  const facLabels=['Ciencias Exactas','Ciencias Económicas','Ingeniería','Artes y Humanidades'];
  const facIns=[21,88,57,5];
  const facPar=[15,76,38,3];
  const empData=[
    {n:'Línea Directa',s:'Textil · Moda',v:27,icon:'👗',c:'#f59e0b'},
    {n:'Metro Medellín',s:'Transporte Masivo',v:23,icon:'🚇',c:'#2563eb'},
    {n:'Postobón',s:'Alimentos y Bebidas',v:20,icon:'🥤',c:'#ef4444'},
    {n:'CEIBA Software',s:'Tecnología',v:16,icon:'💻',c:'#102D69'},
    {n:'Noel',s:'Galletas · Grupo Nutresa',v:15,icon:'🍪',c:'#e67e22'},
    {n:'Protección',s:'Finanzas · Pensiones',v:14,icon:'🛡️',c:'#10b981'},
    {n:'Colcafé',s:'Café · Grupo Nutresa',v:9,icon:'☕',c:'#6d4c41'},
    {n:'Prebel',s:'Cosmética · Belleza',v:8,icon:'💄',c:'#9b59b6'},
  ];
  const ratings=[
    {lbl:'Asimilación de conocimientos',sub:'Capacidad para aprender y aplicar'},
    {lbl:'Relación con compañeros',sub:'Contribución al clima laboral'},
    {lbl:'Proactividad e ideas',sub:'Resolución de problemas y aporte creativo'},
    {lbl:'Calidad y precisión',sub:'Esmero y exactitud en labores'},
  ];
  const techTags=['🤖 Inteligencia Artificial','🏭 Industria 4.0','📊 Ciencia de Datos','🌐 IoT Industrial','🔐 Ciberseguridad OT/IT','⛓️ Blockchain','📈 Big Data','🥽 Realidad Aumentada/VR','⚙️ Automatización','📱 Micro-learning Digital'];
  const tecTags=['Segunda Lengua','Ciberseguridad','Ciencia de Datos','Automatización'];
  const softTags=['💡 Pensamiento Crítico','📊 Análisis de Datos','📚 Auto-aprendizaje','🤝 Trabajo en Equipo','🚀 Proactividad','💬 Comunicación','🎯 Liderazgo','🔍 Solución de Problemas'];
  const quotes=[
    {txt:'"En esta experiencia se evidenció el sector productivo desde un enfoque más cercano y realista, pudimos ser partícipes de cada proceso que se maneja en una empresa."',rol:'Estudiante de Ingeniería en Calidad'},
    {txt:'"La experiencia fue fundamental. Me permitió validar conceptos teóricos en situaciones reales, enfrentando desafíos técnicos y logísticos que no siempre se ven en el aula."',rol:'Estudiante participante'},
    {txt:'"Quería conocer cómo funciona una empresa líder por dentro. Fue algo hermoso ver todo lo estudiado en acción en los procesos reales de producción y calidad."',rol:'Estudiante participante'},
  ];

  // Hero enriquecido (sin kpi-row separado)
  const mkHeroCard=(val,lbl,sub,valColor,borderColor)=>
    `<div style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.18);border-radius:10px;
      padding:12px 16px;min-width:110px;display:flex;flex-direction:column;gap:2px;backdrop-filter:blur(4px)">
      <div style="font-size:1.6rem;font-weight:900;color:${valColor};line-height:1">${val}</div>
      <div style="font-size:.7rem;font-weight:700;color:rgba(255,255,255,.95);text-transform:uppercase;letter-spacing:.04em">${lbl}</div>
      <div style="font-size:.65rem;color:rgba(255,255,255,.6);margin-top:1px">${sub}</div>
    </div>`;
  document.getElementById('hero-cuee').innerHTML=
    mkHeroCard(INSCR,'Inscritos','Convocatoria total','#fff','rgba(255,255,255,.3)')+
    mkHeroCard(CUPOS,'Cupos asignados',Math.round(CUPOS/INSCR*100)+'% de inscritos','#fbbf24','rgba(251,191,36,.4)')+
    mkHeroCard(PARTIC,'Participantes',''+Math.round(PARTIC/CUPOS*100)+'% de cupos','#34d399','rgba(52,211,153,.4)')+
    mkHeroCard(EMPRESAS,'Empresas aliadas','Sector productivo','#93c5fd','rgba(147,197,253,.4)');

  // ── Doughnut Conversión ──
  const cConv=document.getElementById('c-cuee-conv');
  if(cConv){ if(cConv._ch) cConv._ch.destroy();
    cConv._ch=new Chart(cConv,{type:'doughnut',data:{
      labels:['Participantes efectivos','Cupo no utilizado','Sin cupo asignado'],
      datasets:[{data:[132,41,146],
        backgroundColor:['#10b981','#f59e0b','#bfdbfe'],
        borderColor:['#059669','#d97706','#93c5fd'],
        borderWidth:2,
        datalabels:{
          formatter:(v,ctx)=>{const t=319;return Math.round(v/t*100)+'%';},
          color:['#fff','#92400e','#1d4ed8'],
          font:{size:11,weight:'bold'}
        }
      }]
    },options:{responsive:true,maintainAspectRatio:false,cutout:'62%',
      plugins:{
        legend:{position:'bottom',labels:{font:{size:10},padding:10,boxWidth:11}},
        datalabels:{}
      }
    },plugins:[ChartDataLabels]});
  }

  // ── Chart Total ──
  const cTotal=document.getElementById('c-cuee-total');
  if(cTotal){ if(cTotal._ch) cTotal._ch.destroy();
    cTotal._ch=new Chart(cTotal,{type:'bar',data:{
      labels:['Inscritos','Cupos Asignados','Participantes'],
      datasets:[{label:'Estudiantes',data:[INSCR,CUPOS,PARTIC],
        backgroundColor:['#102D69','#2563eb','#10b981'],borderRadius:8,borderWidth:0,
        datalabels:{anchor:'end',align:'top',font:{size:14,weight:'bold'},color:['#102D69','#2563eb','#10b981'],formatter:v=>v}}]
    },options:{responsive:true,maintainAspectRatio:false,layout:{padding:{top:30}},
      plugins:{legend:{display:false},datalabels:{}},
      scales:{x:{grid:{display:false},ticks:{font:{size:12,weight:'600'}}},y:{display:false,beginAtZero:true}}
    },plugins:[ChartDataLabels]});
  }

  // ── Chart Facultad Comparativo ──
  const cFac=document.getElementById('c-cuee-fac-comp');
  if(cFac){ if(cFac._ch) cFac._ch.destroy();
    cFac._ch=new Chart(cFac,{type:'bar',data:{
      labels:facLabels,
      datasets:[
        {label:'Inscritos',data:facIns,backgroundColor:'rgba(16,45,105,.2)',borderColor:'#102D69',borderWidth:2,borderRadius:4,
         datalabels:{anchor:'end',align:'top',font:{size:10,weight:'700'},color:'#102D69',formatter:v=>v}},
        {label:'Participantes',data:facPar,backgroundColor:'#10b98188',borderColor:'#10b981',borderWidth:2,borderRadius:4,
         datalabels:{anchor:'end',align:'top',font:{size:10,weight:'700'},color:'#10b981',formatter:v=>v}},
      ]
    },options:{responsive:true,maintainAspectRatio:false,layout:{padding:{top:28}},
      plugins:{legend:{position:'top',labels:{font:{size:11},padding:12,boxWidth:12}},datalabels:{}},
      scales:{x:{grid:{color:'#eef1f7'},ticks:{font:{size:10}}},y:{grid:{color:'#eef1f7'},beginAtZero:true,ticks:{font:{size:10}}}}
    },plugins:[ChartDataLabels]});
  }

  // ── Doughnut Facultad ──
  const cPie=document.getElementById('c-cuee-fac-pie');
  if(cPie){ if(cPie._ch) cPie._ch.destroy();
    const pieColors=['#2563eb','#10b981','#f59e0b','#8b5cf6'];
    cPie._ch=new Chart(cPie,{type:'doughnut',data:{
      labels:facLabels,
      datasets:[{data:facPar,backgroundColor:pieColors,borderWidth:2,borderColor:'#fff',
        datalabels:{formatter:(v)=>Math.round(v/PARTIC*100)+'%',color:'#fff',font:{size:11,weight:'bold'}}}]
    },options:{responsive:true,maintainAspectRatio:false,cutout:'55%',
      plugins:{legend:{position:'bottom',labels:{font:{size:10},padding:9,boxWidth:11}},datalabels:{}}
    },plugins:[ChartDataLabels]});
    // tabla
    const tbody=document.getElementById('cuee-fac-table');
    if(tbody && !tbody._built){ tbody._built=true;
      facLabels.forEach((f,i)=>{
        const pct=Math.round(facPar[i]/facIns[i]*100);
        const tr=document.createElement('tr'); tr.style.borderBottom='1px solid #eef1f7';
        tr.innerHTML=`<td style="padding:6px 10px;font-weight:600;font-size:.77rem">${f}</td>
          <td style="padding:6px 10px;text-align:center;font-size:.77rem">${facIns[i]}</td>
          <td style="padding:6px 10px;text-align:center;font-size:.77rem;font-weight:700;color:${pieColors[i]}">${facPar[i]}</td>
          <td style="padding:6px 10px;text-align:center"><span style="background:${pieColors[i]}22;color:${pieColors[i]};border-radius:8px;padding:2px 9px;font-size:.71rem;font-weight:700">${pct}%</span></td>`;
        tbody.appendChild(tr);
      });
    }
  }

  // ── Chart Empresas ──
  const cEmp=document.getElementById('c-cuee-empresas');
  if(cEmp){ if(cEmp._ch) cEmp._ch.destroy();
    const eLabels=empData.map(e=>e.n), eVals=empData.map(e=>e.v), eCols=empData.map(e=>e.c);
    cEmp._ch=new Chart(cEmp,{type:'bar',data:{
      labels:eLabels,
      datasets:[{label:'Asistentes',data:eVals,backgroundColor:eCols,borderRadius:6,borderWidth:0,
        datalabels:{anchor:'end',align:'right',font:{size:11,weight:'bold'},color:eCols,formatter:v=>v}}]
    },options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,layout:{padding:{right:40}},
      plugins:{legend:{display:false},datalabels:{}},
      scales:{x:{display:false,beginAtZero:true},y:{grid:{display:false},ticks:{font:{size:10,weight:'600'}}}}
    },plugins:[ChartDataLabels]});
  }

  // ── Empresas Grid ──
  const eGrid=document.getElementById('cuee-empresas-grid');
  if(eGrid && !eGrid._built){ eGrid._built=true;
    const maxV=Math.max(...empData.map(x=>x.v));
    empData.forEach(e=>{
      const isTop=e.v===maxV;
      const barW=Math.round(e.v/maxV*100);
      eGrid.innerHTML+=`<div style="background:var(--surface,#fff);border-radius:10px;border:1px solid ${isTop?e.c+'66':'var(--border)'};padding:11px 13px;display:flex;align-items:center;gap:11px;box-shadow:0 1px 6px rgba(16,45,105,.05)">
        <div style="width:38px;height:38px;border-radius:50%;background:${e.c};display:flex;align-items:center;justify-content:center;font-size:1.1rem;flex-shrink:0">${e.icon}</div>
        <div style="flex:1;min-width:0">
          <div style="font-size:.77rem;font-weight:700;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${e.n}</div>
          <div style="font-size:.65rem;color:var(--text2);margin-bottom:5px">${e.s}</div>
          <div style="background:#eef1f7;border-radius:4px;height:5px;overflow:hidden"><div style="background:${e.c};height:100%;width:${barW}%;border-radius:4px"></div></div>
        </div>
        <div style="text-align:right;flex-shrink:0">
          <div style="font-size:1.2rem;font-weight:900;color:${e.c};line-height:1">${e.v}</div>
          <div style="font-size:.6rem;color:var(--text2)">${isTop?'🏆 mayor':'part.'}</div>
        </div>
      </div>`;
    });
  }

  // ── Ratings ──
  const rGrid=document.getElementById('cuee-ratings');
  if(rGrid && !rGrid._built){ rGrid._built=true;
    ratings.forEach(r=>{
      rGrid.innerHTML+=`<div style="background:var(--card);border-radius:12px;border:1px solid var(--border);padding:18px;display:flex;align-items:center;gap:12px;box-shadow:0 1px 8px rgba(16,45,105,.05)">
        <div style="font-size:2.2rem;font-weight:900;color:#10b981;min-width:48px">5.0</div>
        <div><div style="color:#f59e0b;font-size:.95rem;letter-spacing:1px">★★★★★</div>
          <div style="font-size:.78rem;font-weight:600;color:var(--text);margin-top:3px">${r.lbl}</div>
          <div style="font-size:.68rem;color:var(--text2)">${r.sub}</div></div>
      </div>`;
    });
  }

  // ── Tags ──
  const buildTags=(id,arr,cls)=>{ const el=document.getElementById(id); if(el&&!el._built){ el._built=true; arr.forEach(t=>{ el.innerHTML+=`<span style="${cls};border-radius:8px;padding:8px 10px;font-size:.73rem;font-weight:600;display:block;text-align:center;line-height:1.3">${t}</span>`; }); }};
  buildTags('cuee-tags-tech',techTags,'background:#eff6ff;color:#1d4ed8;border:1px solid #bfdbfe');
  buildTags('cuee-tags-tec', tecTags, 'background:#eff6ff;color:#1d4ed8;border:1px solid #bfdbfe');
  buildTags('cuee-tags-soft',softTags,'background:#f0fdf4;color:#166534;border:1px solid #bbf7d0');

  // ── Citas ──
  const qGrid=document.getElementById('cuee-quotes');
  if(qGrid && !qGrid._built){ qGrid._built=true;
    quotes.forEach(q=>{
      qGrid.innerHTML+=`<div style="background:var(--card);border-radius:12px;border:1px solid var(--border);padding:18px 20px;box-shadow:0 1px 8px rgba(16,45,105,.05)">
        <div style="font-size:1.2rem;color:#f59e0b;margin-bottom:8px">❝</div>
        <p style="font-size:.77rem;color:var(--text2);line-height:1.6;font-style:italic">${q.txt}</p>
        <div style="margin-top:10px;font-size:.7rem;color:var(--itm);font-weight:600">— ${q.rol}</div>
      </div>`;
    });
  }
}

// ── Canvas Word Cloud ─────────────────────────────────────────────────────────
function drawWordCloud(canvasId, words) {
  const canvas = document.getElementById(canvasId);
  if(!canvas) return;

  const W = Math.max(canvas.parentElement.clientWidth || 800, 500);
  const H = Math.round(W * 0.52);
  canvas.width  = W;
  canvas.height = H;
  canvas.style.height = H + 'px';

  const ctx = canvas.getContext('2d');
  ctx.clearRect(0, 0, W, H);

  // Fondo con gradiente suave
  const grad = ctx.createLinearGradient(0, 0, W, H);
  grad.addColorStop(0, '#f0f5fc');
  grad.addColorStop(1, '#f8fafd');
  ctx.fillStyle = grad;
  ctx.fillRect(0, 0, W, H);

  if(!words || !words.length) {
    ctx.fillStyle = '#9ca3af';
    ctx.font = '15px Segoe UI';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('Sin datos para este programa', W/2, H/2);
    return;
  }

  const COLORS = [
    '#00539B','#1a6eb5','#E8A000','#059669','#0891b2',
    '#7c3aed','#003d73','#b87e00','#065f46','#1e40af',
    '#0369a1','#4f46e5','#047857','#92400e','#1d4ed8'
  ];

  const PAD   = 10;  // margen interior canvas
  const GUTTER = 5; // separación mínima entre palabras

  const maxCount = Math.max(...words.map(w => w.count), 1);
  const sorted   = [...words].sort((a, b) => b.count - a.count);

  // ── Cuadrícula de ocupación (un bit por celda de CELL px) ─────────────────
  const CELL = 3;
  const gW = Math.ceil(W / CELL);
  const gH = Math.ceil(H / CELL);
  const grid = new Uint8Array(gW * gH);

  function mark(x, y, w, h) {
    const x0 = Math.max(0, Math.floor((x - GUTTER) / CELL));
    const y0 = Math.max(0, Math.floor((y - GUTTER) / CELL));
    const x1 = Math.min(gW - 1, Math.ceil((x + w + GUTTER) / CELL));
    const y1 = Math.min(gH - 1, Math.ceil((y + h + GUTTER) / CELL));
    for(let gy = y0; gy <= y1; gy++)
      for(let gx = x0; gx <= x1; gx++)
        grid[gy * gW + gx] = 1;
  }

  function free(x, y, w, h) {
    if(x < PAD || y < PAD || x + w > W - PAD || y + h > H - PAD) return false;
    const x0 = Math.floor(x / CELL);
    const y0 = Math.floor(y / CELL);
    const x1 = Math.ceil((x + w) / CELL);
    const y1 = Math.ceil((y + h) / CELL);
    for(let gy = y0; gy < y1; gy++)
      for(let gx = x0; gx < x1; gx++)
        if(grid[gy * gW + gx]) return false;
    return true;
  }

  // ── Espiral Arqímedea que cubre todo el canvas ────────────────────────────
  const cx = W / 2, cy = H / 2;
  const maxR = Math.sqrt(cx * cx + cy * cy);

  sorted.forEach((item, idx) => {
    const ratio    = item.count / maxCount;
    // Rango de tamaño: palabras grandes bien grandes, pequeñas legibles
    const fontSize = Math.round(13 + ratio * (W > 700 ? 44 : 30));
    const color    = COLORS[idx % COLORS.length];
    const alpha    = 0.72 + ratio * 0.28;

    ctx.font = `bold ${fontSize}px 'Segoe UI', Arial, sans-serif`;
    const tw = ctx.measureText(item.word).width;
    const th = fontSize * 1.25;

    // Paso de espiral proporcional al tamaño de la palabra para más dispersión
    const spiralStep = Math.max(1.5, fontSize * 0.18);
    // Ángulo inicial distinto por palabra para no alinearlas
    const angleOff   = idx * 2.399; // número áureo × 2π para buena dispersión

    let placed = false;
    for(let r = 0; r <= maxR && !placed; r += spiralStep) {
      // Más puntos en radios grandes para cubrir la periferia
      const nPoints = Math.max(8, Math.round(2 * Math.PI * Math.max(r, 1) / spiralStep));
      for(let i = 0; i < nPoints && !placed; i++) {
        const a  = angleOff + (2 * Math.PI * i / nPoints);
        // Escala elíptica para llenar más el ancho que el alto
        const px = cx + r * Math.cos(a) * 1.15 - tw / 2;
        const py = cy + r * Math.sin(a) * 0.85 - th / 2;

        if(free(px, py, tw, th)) {
          mark(px, py, tw, th);

          ctx.save();
          ctx.globalAlpha   = alpha;
          ctx.fillStyle     = color;
          ctx.font          = `bold ${fontSize}px 'Segoe UI', Arial, sans-serif`;
          ctx.textAlign     = 'left';
          ctx.textBaseline  = 'top';
          ctx.fillText(item.word, px, py + (th - fontSize) / 2);
          ctx.restore();

          placed = true;
        }
      }
    }
  });
}

// ── Históricos ────────────────────────────────────────────────────────────────
function renderHistoricos() {
  if (document.getElementById('c-hist-pre-anio')._built) return;
  document.getElementById('c-hist-pre-anio')._built = true;

  // ── Datos ──
  const anios = ['2024','2025','2026*'];
  const preAnio  = [1127, 1106, 477];
  const inicAnio = [769,  838,  405];
  const gradAnio = [699,  989,  342];
  const vincAnio = [141,  249,  96];

  const trimLabels = ['T2-2024','T3-2024','T4-2024','T1-2025','T2-2025','T3-2025','T4-2025','T1-2026','T2-2026'];

  // Asistentes preprácticas por trimestre (ene=T1, abr=T2, jul=T3, oct=T4)
  const preTrimAll = [
    {label:'2024', data:[null,237,281,284,325,null,null,null,null], color:'#102D69'},
    {label:'2025', data:[null,null,null,null,null,217,326,200,363], color:'#3b82f6'},
    {label:'2026*',data:[null,null,null,null,null,null,null,236,241], color:'#10b981'},
  ];
  // Aplanado para gráfica combinada por trimestre
  const preTrimFlat  = [null,237,281,284,325,217,326,200,363,236,241];
  const trimLabels2  = ['ene-24','abr-24','jul-24','oct-24','ene-25','abr-25','jul-25','oct-25','ene-26','abr-26'];

  // Iniciaron prácticas por trimestre
  const inicTrimFlat = [null,null,548,221,287,178,229,144,257,148];

  // Graduados por trimestre
  const gradTrimFlat = [null,157,274,268,253,249,247,240,169,171];

  // Vinculados por trimestre
  const vincTrimFlat = [null,33,54,54,62,58,57,72,48,48];

  // ── Colores ITM ──
  const C = { blue:'#102D69', lblue:'#3b82f6', green:'#10b981', amber:'#f59e0b', purple:'#8b5cf6', red:'#ef4444' };

  const defPlugin = { legend:{position:'bottom',labels:{font:{size:11},padding:10}},
    datalabels:{anchor:'end',align:'top',font:{size:10,weight:'700'},color:'#334',
      formatter:v=>v==null?'':v} };

  // ── KPIs ──
  const kpiData = [
    {v:'3.710', l:'Total asistentes preprácticas', s:'2024–2026', c:C.blue},
    {v:'2.012', l:'Estudiantes que iniciaron', s:'2024–2026', c:C.green},
    {v:'2.030', l:'Se graduaron con práctica', s:'2024–2026', c:C.amber},
    {v:'486',   l:'Vinculados a empresa', s:'2024–2026', c:C.purple},
  ];
  const kpiEl = document.getElementById('hist-kpis');
  kpiData.forEach(k => {
    kpiEl.innerHTML += `<div style="background:#fff;border:1px solid #e2e8f0;border-top:4px solid ${k.c};border-radius:10px;padding:18px 14px;text-align:center">
      <div style="font-size:2rem;font-weight:900;color:${k.c};line-height:1.1">${k.v}</div>
      <div style="font-size:.78rem;font-weight:700;color:#334;margin-top:6px">${k.l}</div>
      <div style="font-size:.7rem;color:#888;margin-top:3px">${k.s}</div>
    </div>`;
  });

  const mkBar = (id, labels, datasets, title) => {
    const ctx = document.getElementById(id);
    if (!ctx) return;
    new Chart(ctx, { type:'bar', data:{ labels, datasets },
      options:{ responsive:true, maintainAspectRatio:false,
        plugins:{ ...defPlugin, title:{display:!!title,text:title,font:{size:12}} },
        scales:{ x:{grid:{display:false},ticks:{font:{size:10}}},
                 y:{beginAtZero:true,grid:{color:'#f0f0f0'},ticks:{font:{size:10}}} } } });
  };

  const mkLine = (id, labels, datasets) => {
    const ctx = document.getElementById(id);
    if (!ctx) return;
    new Chart(ctx, { type:'line', data:{ labels, datasets },
      options:{ responsive:true, maintainAspectRatio:false,
        plugins:{ ...defPlugin },
        scales:{ x:{grid:{display:false},ticks:{font:{size:10}}},
                 y:{beginAtZero:true,grid:{color:'#f0f0f0'},ticks:{font:{size:10}}} } } });
  };

  // ── Chart 1: Preprácticas por año ──
  mkBar('c-hist-pre-anio', anios, [{
    label:'Asistentes', data:preAnio,
    backgroundColor:[C.blue, C.lblue, C.green],
    borderRadius:6, borderSkipped:false
  }]);

  // ── Chart 2: Preprácticas por trimestre ──
  mkLine('c-hist-pre-trim', trimLabels2, [{
    label:'Asistentes por trimestre',
    data:preTrimFlat,
    borderColor:C.blue, backgroundColor:'rgba(16,45,105,.08)',
    tension:.35, pointRadius:5, pointBackgroundColor:C.blue, fill:true,
    spanGaps:true
  }]);

  document.getElementById('hist-analisis-pre').innerHTML =
    `<strong>📌 Análisis:</strong> El curso de preprácticas mantiene una demanda sostenida: <strong>1.127</strong> asistentes en 2024 y <strong>1.106</strong> en 2025, con ligera estabilidad. En 2026 se registran <strong>477 asistentes</strong> solo en el primer semestre, proyectando cifras similares a años anteriores. El pico trimestral más alto se presentó en <strong>octubre 2025 (363)</strong> y <strong>abril 2025 (326)</strong>, evidenciando mayor demanda en el segundo y cuarto trimestre de cada año.`;

  // ── Chart 3: Iniciaron prácticas por año ──
  mkBar('c-hist-inic-anio', anios, [{
    label:'Iniciaron prácticas', data:inicAnio,
    backgroundColor:[C.blue, C.green, C.amber],
    borderRadius:6, borderSkipped:false
  }]);

  // ── Chart 4: Iniciaron por trimestre ──
  mkLine('c-hist-inic-trim', trimLabels2, [{
    label:'Estudiantes que iniciaron',
    data:inicTrimFlat,
    borderColor:C.green, backgroundColor:'rgba(16,185,129,.08)',
    tension:.35, pointRadius:5, pointBackgroundColor:C.green, fill:true,
    spanGaps:true
  }]);

  document.getElementById('hist-analisis-inic').innerHTML =
    `<strong>📌 Análisis:</strong> Los registros de inicio de prácticas muestran crecimiento: de <strong>769</strong> en 2024 a <strong>838</strong> en 2025 (+8.97%). El trimestre de <strong>julio 2024</strong> fue excepcionalmente alto con <strong>548 estudiantes</strong>, posiblemente por acumulación de semestres anteriores. En 2026, el primer semestre acumula <strong>405 estudiantes</strong>, ritmo coherente con años previos. Se observa una variabilidad trimestral significativa que sugiere picos asociados a calendarios académicos.`;

  // ── Chart 5: Graduados por año ──
  mkBar('c-hist-grad-anio', anios, [{
    label:'Graduados con práctica', data:gradAnio,
    backgroundColor:['#f59e0b','#fbbf24','#fde68a'],
    borderRadius:6, borderSkipped:false
  }]);

  document.getElementById('hist-analisis-grad').innerHTML =
    `<strong>📌 Análisis:</strong> Los estudiantes que se graduaron con práctica reflejan un <strong>crecimiento notable</strong>: de <strong>699</strong> en 2024 a <strong>989</strong> en 2025 (+41.5%). Este salto indica mayor retención y finalización exitosa del proceso. En 2026 se reportan <strong>342</strong> graduados en el primer semestre, con tendencia a superar el 2024.`;

  // ── Chart 6: Graduados por trimestre ──
  mkLine('c-hist-grad-trim', trimLabels2, [{
    label:'Graduados con práctica',
    data:gradTrimFlat,
    borderColor:C.amber, backgroundColor:'rgba(245,158,11,.08)',
    tension:.35, pointRadius:5, pointBackgroundColor:C.amber, fill:true,
    spanGaps:true
  }]);

  // ── Chart 7: Vinculados por año ──
  mkBar('c-hist-vinc-anio', anios, [{
    label:'Vinculados', data:vincAnio,
    backgroundColor:[C.purple, '#a78bfa', '#c4b5fd'],
    borderRadius:6, borderSkipped:false
  }]);

  // ── Chart 8: Vinculados por trimestre ──
  mkLine('c-hist-vinc-trim', trimLabels2, [{
    label:'Vinculados por trimestre',
    data:vincTrimFlat,
    borderColor:C.purple, backgroundColor:'rgba(139,92,246,.08)',
    tension:.35, pointRadius:5, pointBackgroundColor:C.purple, fill:true,
    spanGaps:true
  }]);

  document.getElementById('hist-analisis-vinc').innerHTML =
    `<strong>📌 Análisis:</strong> Las vinculaciones a empresa crecieron un <strong>76.6%</strong> de 2024 (<strong>141</strong>) a 2025 (<strong>249</strong>), reflejando mayor articulación con el sector productivo. En 2026, con <strong>96 vinculados</strong> en el primer semestre, se proyecta mantener el nivel de 2024. El pico trimestral fue <strong>octubre 2025 (72)</strong>, coincidiendo con el ciclo de mayor demanda del año.`;

  // ── Chart 9: Comparativo integral ──
  const ctxComp = document.getElementById('c-hist-comp');
  if (ctxComp) {
    new Chart(ctxComp, {
      type:'bar',
      data:{ labels: anios,
        datasets:[
          { label:'Asistentes preprácticas', data:preAnio,  backgroundColor:'rgba(16,45,105,.8)',   borderRadius:5 },
          { label:'Iniciaron prácticas',      data:inicAnio, backgroundColor:'rgba(16,185,129,.8)',  borderRadius:5 },
          { label:'Graduados con práctica',   data:gradAnio, backgroundColor:'rgba(245,158,11,.85)', borderRadius:5 },
          { label:'Vinculados a empresa',     data:vincAnio, backgroundColor:'rgba(139,92,246,.85)', borderRadius:5 },
        ]
      },
      options:{ responsive:true, maintainAspectRatio:false,
        plugins:{ legend:{position:'bottom',labels:{font:{size:11},padding:12}},
          datalabels:{anchor:'end',align:'top',font:{size:10,weight:'700'},color:'#334',formatter:v=>v} },
        scales:{ x:{grid:{display:false},ticks:{font:{size:11}}},
                 y:{beginAtZero:true,grid:{color:'#f0f0f0'},ticks:{font:{size:10}}} } }
    });
  }

  document.getElementById('hist-analisis-comp').innerHTML =
    `<strong>📌 Análisis integral:</strong> El comparativo muestra que el proceso de prácticas mantiene una <strong>alta demanda</strong> en preprácticas (>1.000/año), con una <strong>tasa de conversión a graduados</strong> que mejoró del 62% en 2024 al 90% en 2025. Las vinculaciones a empresa son el indicador de mayor crecimiento (+76.6%), señal de fortalecimiento de la alianza academia–empresa. El año 2026 muestra ritmos acordes con la proyección histórica. <em>(*2026: datos del primer semestre)</em>`;
}

// ── Init ──────────────────────────────────────────────────────────────────────
function init() {
  const anioSel=document.getElementById('f-anio');
  D.filtros_anios.forEach(y=>{ const o=document.createElement('option'); o.value=y; o.textContent=y; anioSel.appendChild(o); });

  const mesSel=document.getElementById('f-mes');
  D.filtros_meses.forEach(([n,l])=>{ const o=document.createElement('option'); o.value=n; o.textContent=l; mesSel.appendChild(o); });

  const semSel=document.getElementById('f-sem');
  D.filtros_semestres.forEach(s=>{ const o=document.createElement('option'); o.value=s; o.textContent=s; semSel.appendChild(o); });

  const facSel=document.getElementById('f-facultad');
  D.all_facultades.forEach(f=>{ const o=document.createElement('option'); o.value=f; o.textContent=f; facSel.appendChild(o); });

  const pSel=document.getElementById('f-programa');
  D.all_programas.forEach(p=>{ const o=document.createElement('option'); o.value=p; o.textContent=p; pSel.appendChild(o); });

  renderPracticantes();
}

document.addEventListener('DOMContentLoaded', init);

// ── Carrusel CUEE ────────────────────────────────────────────────────────────
(function(){
  let cueeIdx = 0;
  let cueeTotal = 0;
  let cueeTimer = null;

  function cueeInit() {
    const slides = document.querySelectorAll('.cuee-slide');
    const dotsEl = document.getElementById('cuee-dots');
    cueeTotal = slides.length;
    if(!cueeTotal || !dotsEl) return;
    // Crear dots
    dotsEl.innerHTML = '';
    slides.forEach((_,i)=>{
      const d = document.createElement('div');
      d.className = 'cuee-dot' + (i===0?' active':'');
      d.onclick = ()=>cueeGoTo(i);
      dotsEl.appendChild(d);
    });
    cueeAutoplay();
  }

  window.cueeGoTo = function(idx) {
    const slides  = document.querySelectorAll('.cuee-slide');
    const thumbs  = document.querySelectorAll('.cuee-thumb');
    const dots    = document.querySelectorAll('.cuee-dot');
    if(!slides.length) return;
    slides[cueeIdx].classList.remove('active');
    thumbs[cueeIdx]?.classList.remove('active');
    dots[cueeIdx]?.classList.remove('active');
    cueeIdx = (idx + cueeTotal) % cueeTotal;
    slides[cueeIdx].classList.add('active');
    thumbs[cueeIdx]?.classList.add('active');
    dots[cueeIdx]?.classList.add('active');
    // scroll thumb visible dentro del contenedor (sin afectar el scroll de la página)
    const tw = document.getElementById('cuee-thumbs');
    if(tw && thumbs[cueeIdx]) {
      const thumb = thumbs[cueeIdx];
      const twLeft = tw.scrollLeft;
      const twW = tw.offsetWidth;
      const tLeft = thumb.offsetLeft;
      const tW = thumb.offsetWidth;
      if(tLeft < twLeft) tw.scrollLeft = tLeft - 8;
      else if(tLeft + tW > twLeft + twW) tw.scrollLeft = tLeft + tW - twW + 8;
    }
    cueeAutoplay();
  };

  window.cueeMove = function(dir) { cueeGoTo(cueeIdx + dir); };

  function cueeAutoplay() {
    clearInterval(cueeTimer);
    cueeTimer = setInterval(()=>cueeGoTo(cueeIdx+1), 4500);
  }

  // Iniciar cuando se cargue la sección CUEE
  const origGoTo = window.goTo;
  window.goTo = function(name, btn) {
    origGoTo(name, btn);
    if(name==='cuee') { setTimeout(cueeInit, 80); }
  };
  // También por si el carrusel ya está visible al cargar
  document.addEventListener('DOMContentLoaded', ()=>{
    if(document.querySelector('.cuee-slide')) cueeInit();
  });
})();

// ── PDF Export ────────────────────────────────────────────────────────────────
function exportPDF() {
  // Show all sections temporarily for print
  const sections = document.querySelectorAll('.section');
  const hidden = [];
  sections.forEach(s => { if(!s.classList.contains('active')){ s.style.display='block'; hidden.push(s); } });

  // Wait a tick then print
  setTimeout(() => {
    window.print();
    // Restore hidden sections
    hidden.forEach(s => { s.style.display=''; });
  }, 200);
}
</script>
</body>
</html>
"""

# ─── Main ──────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    print('Cargando bases de datos...')
    df1 = load_practicantes()
    df2 = load_disponibles()
    df3 = load_f082()
    df4 = load_solicitud()
    df5 = load_aprobacion()
    df6 = load_encuesta()
    df7 = load_encuesta_est()

    print('Construyendo datos y agregaciones...')
    data = build_data(df1, df2, df3, df4, df5, df6, df7)

    print('Generando HTML...')
    logo  = get_logo()
    dj    = json.dumps(data, ensure_ascii=False, default=str)
    cuee_carousel = build_cuee_carousel()
    html  = HTML.replace('__LOGO__', logo).replace('__DATA_JSON__', dj).replace('__CUEE_CAROUSEL__', cuee_carousel)

    out = 'index.html'
    with open(out, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f'OK  Generado: {out}  ({len(html)//1024} KB)')
    print(f'    Practicantes: {len(df1)} | Disponibles: {len(df2)} | F082: {len(df3)} | Solicitud: {len(df4)} | Aprobacion: {len(df5)} | Enc.Empresarios: {len(df6)} | Enc.Estudiantes: {len(df7)}')
