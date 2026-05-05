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
    cols = list(df.columns)
    rename = {cols[0]:'MODALIDAD', cols[1]:'FECHA_INICIO', cols[2]:'FECHA_TERMINA',
              cols[3]:'EMPRESA_RAW', cols[4]:'CARGO', cols[5]:'RAZON_RAW',
              cols[6]:'ASESOR', cols[7]:'ACTIVIDADES', cols[8]:'DESCRIPCION',
              cols[9]:'PROMEDIO', cols[10]:'VINCULADO', cols[11]:'PROGRAMA',
              cols[12]:'FECHA_ENTREGA', cols[13]:'FACULTAD', cols[14]:'SISTEMATIZACION_RAW',
              cols[15]:'ES_OPTIMA', cols[16]:'POR_QUE_NO_OPTIMA', cols[17]:'ENTREGADO'}
    df.rename(columns=rename, inplace=True)
    df.drop(columns=['EMPRESA_RAW','RAZON_RAW','SISTEMATIZACION_RAW'], inplace=True)
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
              cols[5]:'PERFIL_SOLICITADO', cols[6]:'MODALIDAD'}
    df.rename(columns=rename, inplace=True)
    df.drop(columns=['NIT_RAW'], inplace=True)
    for c in ['EMPRESA','EMPRESA_NUEVA','PROGRAMA_NUEVO','PERFIL_SOLICITADO','MODALIDAD']:
        df[c] = df[c].apply(clean_text)
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

    # también de solicitud (perfil)
    for _, row in df4[['PERFIL_SOLICITADO']].drop_duplicates().iterrows():
        p = clean_text(str(row['PERFIL_SOLICITADO'])) if pd.notna(row['PERFIL_SOLICITADO']) else ''
        if p and len(p) > 2:
            fac_prog.setdefault('SIN FACULTAD', set()).add(p)

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
        ['PROGRAMA','FACULTAD','EMPRESA_NUEVA','TIPO_CONTRATO','ESTADO','ASESOR','MODALIDAD','ANIO','MES','MES_LABEL','SEMESTRE'])
    data['raw_disponibles']  = raw_records(df2,
        ['MODALIDAD','PROGRAMA','FACULTAD','ESTADO','DISCAPACIDAD','TIPO_DISCAPACIDAD','ANIO','MES','MES_LABEL','SEMESTRE'])
    data['raw_f082']         = raw_records(df3,
        ['MODALIDAD','ASESOR','ACTIVIDADES','PROMEDIO','VINCULADO','PROGRAMA','FACULTAD','ES_OPTIMA','ENTREGADO','ANIO','MES','MES_LABEL','SEMESTRE'])
    data['raw_solicitud']    = raw_records(df4,
        ['EMPRESA','EMPRESA_NUEVA','PROGRAMA_NUEVO','PERFIL_SOLICITADO','MODALIDAD','ANIO','MES','MES_LABEL','SEMESTRE'])
    data['raw_aprobacion']   = raw_records(df5,
        ['PROGRAMA','FUNCIONES','EMPRESA','ESTADO_APROBACION','APROBADOR','ANIO','MES','MES_LABEL','SEMESTRE'])

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
    def _count_multi(series, top=12):
        c = Counter()
        for v in series.dropna():
            for p in str(v).split(';'):
                p = p.strip()
                if p and p.lower() not in ('nan', 'none', ''):
                    c[p] += 1
        return [{'label': k, 'count': v} for k, v in c.most_common(top)]

    def _count_calif(series):
        order = ['EXCELENTE', 'BUENO', 'REGULAR', 'DEFICIENTE']
        s = series.dropna().str.upper().str.strip()
        return {o: int((s == o).sum()) for o in order if int((s == o).sum()) > 0}

    data['enc_raw'] = raw_records(df6,
        ['empresa', 'sector', 'mercado', 'tipo', 'programa',
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
  padding: 10px 28px;
  display:flex; flex-wrap:wrap; gap:10px; align-items:center;
}
.filter-group { display:flex; align-items:center; gap:6px }
.filter-label {
  font-size:.7rem; font-weight:700; color:var(--text2);
  text-transform:uppercase; letter-spacing:.5px; white-space:nowrap;
}
.filter-select {
  border: 1.5px solid var(--border); border-radius:var(--radius-sm);
  padding: 5px 28px 5px 10px; font-size:.82rem; color:var(--text);
  background: #f4f7fd url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%2300539B'/%3E%3C/svg%3E") no-repeat right 9px center;
  appearance:none; cursor:pointer;
  transition:border-color .2s, box-shadow .2s;
  max-width:220px;
}
.filter-select:focus { outline:none; border-color:var(--itm-blue); box-shadow:0 0 0 3px rgba(0,83,155,.12) }
.btn-reset {
  background: linear-gradient(135deg, var(--itm-gold), var(--itm-gold2));
  color:#fff; border:none; border-radius:var(--radius-sm);
  padding:5px 14px; cursor:pointer; font-size:.78rem; font-weight:700;
  transition:var(--transition); box-shadow:0 2px 8px rgba(232,160,0,.3);
  letter-spacing:.2px;
}
.btn-reset:hover { transform:translateY(-1px); box-shadow:0 4px 14px rgba(232,160,0,.4) }
.filter-info {
  font-size:.72rem; color:var(--itm-blue);
  background:rgba(0,83,155,.08); border-radius:20px;
  padding:3px 10px; font-weight:600; letter-spacing:.2px;
}
.filter-divider { width:1px; height:24px; background:var(--border) }

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
@media(max-width:768px) {
  main { padding:16px }
  .charts-grid { grid-template-columns:1fr }
  .header-top { padding:8px 16px }
  .header-titles h1 { font-size:.9rem }
  .filter-bar { padding:8px 16px; gap:8px }
  .filter-select { max-width:140px }
  .sec-hero { flex-direction:column }
  .sec-hero-right { width:100% }
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
        <p>ITM &ndash; Institución Universitaria &ndash; Prácticas Profesionales ITM</p>
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
      <span class="dot"></span>📋 Practicantes
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
  </nav>
</header>

<!-- ═══════════════ FILTER BAR ═══════════════ -->
<div class="filter-bar">
  <div class="filter-group">
    <span class="filter-label">Año</span>
    <select id="f-anio" class="filter-select" onchange="onFilterChange()">
      <option value="">Todos</option>
    </select>
  </div>
  <div class="filter-divider"></div>
  <div class="filter-group">
    <span class="filter-label">Semestre</span>
    <select id="f-sem" class="filter-select" onchange="onFilterChange()">
      <option value="">Todos</option>
    </select>
  </div>
  <div class="filter-divider"></div>
  <div class="filter-group">
    <span class="filter-label">Mes</span>
    <select id="f-mes" class="filter-select" onchange="onFilterChange()">
      <option value="">Todos</option>
    </select>
  </div>
  <div class="filter-divider"></div>
  <div class="filter-group">
    <span class="filter-label">Facultad</span>
    <select id="f-facultad" class="filter-select" onchange="onFacultadChange()">
      <option value="">Todas</option>
    </select>
  </div>
  <div class="filter-divider"></div>
  <div class="filter-group">
    <span class="filter-label">Programa</span>
    <select id="f-programa" class="filter-select" onchange="onFilterChange()">
      <option value="">Todos</option>
    </select>
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
      <div class="sec-hero-icon">📋</div>
      <h2>Practicantes en Práctica Profesional</h2>
      <p>Estado, tipo de contrato, empresa, asesor, programa y facultad</p>
    </div>
    <div class="sec-hero-right" id="hero-pract"></div>
  </div>
  <div class="kpi-row" id="kpi-pract"></div>
  <div class="stat-panel" id="stat-pract"></div>
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
      <div class="card-head"><h3>Empresa nueva vs. recurrente</h3></div>
      <div class="card-body"><div class="ch h200"><canvas id="c-p-empnueva"></canvas></div></div>
    </div>
    <div class="card">
      <div class="card-head"><h3>Modalidad de práctica</h3></div>
      <div class="card-body"><div class="ch h200"><canvas id="c-p-modalidad"></canvas></div></div>
    </div>
    <div class="card full">
      <div class="card-head"><h3>Top 15 programas con más practicantes</h3></div>
      <div class="card-body"><div id="prog-list-pract" class="progress-list"></div></div>
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
      <div class="card-head"><h3>Estudiantes por asesor (Todos)</h3></div>
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
    <div class="card">
      <div class="card-head"><h3>Solicitudes por año</h3></div>
      <div class="card-body"><div class="ch h260"><canvas id="c-s-anio"></canvas></div></div>
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
    <div class="card">
      <div class="card-head"><h3>Estado de aprobación</h3></div>
      <div class="card-body"><div class="ch h240"><canvas id="c-a-estado"></canvas></div></div>
    </div>
    <div class="card">
      <div class="card-head"><h3>Top 15 empresas solicitantes</h3></div>
      <div class="card-body"><div class="ch h300"><canvas id="c-a-empresa"></canvas></div></div>
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
    <div class="card full">
      <div class="card-head"><h3>Competencias más demandadas por el mercado</h3></div>
      <div class="card-body"><div class="ch h300"><canvas id="c-enc-competencias"></canvas></div></div>
    </div>
    <div class="card full">
      <div class="card-head"><h3>Programas con mayor participación en la encuesta</h3></div>
      <div class="card-body" style="overflow-x:auto;padding-bottom:10px">
        <div class="ch h300" id="wrap-enc-prog" style="min-width:700px">
          <canvas id="c-enc-programas"></canvas>
        </div>
      </div>
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

    <!-- Satisfacción general -->
    <div class="card">
      <div class="card-head"><h3>Satisfacción general del servicio</h3></div>
      <div class="card-body"><div class="ch h240"><canvas id="c-est-satisf"></canvas></div></div>
    </div>

    <!-- Recomendaría -->
    <div class="card">
      <div class="card-head"><h3>¿Recomendaría realizar la práctica?</h3></div>
      <div class="card-body"><div class="ch h240"><canvas id="c-est-rec"></canvas></div></div>
    </div>

    <!-- Modalidad -->
    <div class="card">
      <div class="card-head"><h3>Modalidad de práctica</h3></div>
      <div class="card-body"><div class="ch h240"><canvas id="c-est-modal"></canvas></div></div>
    </div>

    <!-- Proyecto a futuro -->
    <div class="card">
      <div class="card-head"><h3>Proyecto a futuro del estudiante</h3></div>
      <div class="card-body"><div class="ch h240"><canvas id="c-est-futuro"></canvas></div></div>
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

</main>

<footer>
  <b>ITM &ndash; Institución Universitaria</b> &nbsp;·&nbsp; <span>Prácticas Profesionales ITM</span>
  &nbsp;·&nbsp; Informe generado automáticamente
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
  fAnio=fSem=fMes=fFac=fProg='';
  // Restore all programs
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
    if(fFac  && facKey && r[facKey] !== fFac)  return false;
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
  const rows = filterRows(D.raw_practicantes);
  const n=rows.length;
  const activos=rows.filter(r=>r.ESTADO&&r.ESTADO.toLowerCase().includes('activo')).length;
  const nuevas=rows.filter(r=>r.EMPRESA_NUEVA&&r.EMPRESA_NUEVA.toLowerCase()==='si').length;
  const asesores=new Set(rows.map(r=>r.ASESOR).filter(Boolean)).size;
  const progs=new Set(rows.map(r=>r.PROGRAMA).filter(Boolean)).size;
  const pctAct=n?Math.round(activos/n*100):0;

  document.getElementById('hero-pract').innerHTML=
    heroStat(n,'Total','') +
    heroStat(activos,'Activos','green') +
    heroStat(nuevas,'Emp. Nuevas','gold') +
    heroStat(asesores,'Asesores','');

  document.getElementById('kpi-pract').innerHTML=
    kpiCard(n,'Total practicantes','Período filtrado','') +
    kpiCard(activos,'Activos',`${pctAct}% del total`,'green') +
    kpiCard(nuevas,'Empresas nuevas','Primer convenio','gold') +
    kpiCard(rows.length-activos,'No activos','Finalizados / rezagados','red') +
    kpiCard(asesores,'Asesores','Únicos asignados','') +
    kpiCard(progs,'Programas','Con estudiantes','purple');

  const top1Prog = groupBy(rows,'PROGRAMA',1).labels[0]||'N/A';
  const top1Cont = groupBy(rows,'TIPO_CONTRATO',1).labels[0]||'N/A';
  const top1Mod  = groupBy(rows,'MODALIDAD',1).labels[0]||'N/A';
  document.getElementById('stat-pract').innerHTML=
    `Total de practicantes registrados: <strong>${n}</strong>.
     Estudiantes activos: <strong>${activos} (${pctAct}%)</strong>.
     Empresas nuevas vinculadas: <strong>${nuevas}</strong>.
     Programa con más estudiantes: <strong>${top1Prog}</strong>.
     Tipo de contrato predominante: <strong>${top1Cont}</strong>.
     Modalidad más frecuente: <strong>${top1Mod}</strong>.
     Asesores asignados: <strong>${asesores}</strong>.`;

  const est=groupBy(rows,'ESTADO');
  document.getElementById('cb-p-estado').textContent=est.labels.length+' estados';
  mkDoughnut('c-p-estado', est.labels, est.values, PAL_MAIN);

  const cont=groupBy(rows,'TIPO_CONTRATO');
  mkDoughnut('c-p-contrato', cont.labels, cont.values, PAL_BLUE);

  const en=groupBy(rows,'EMPRESA_NUEVA');
  mkDoughnut('c-p-empnueva', en.labels, en.values, [C.blue, C.gold]);

  const mod=groupBy(rows,'MODALIDAD');
  mkDoughnut('c-p-modalidad', mod.labels, mod.values, [C.blue, C.gold, C.green]);

  progressList('prog-list-pract', rows, 'PROGRAMA', 15);

  const fac=groupBy(rows,'FACULTAD');
  mkDoughnut('c-p-facultad', fac.labels, fac.values, facColorsArr(fac.labels));

  const sem=semSort(rows);
  mkLine('c-p-semestre', sem.labels, sem.values, C.blue);

  const as=groupBy(rows,'ASESOR');
  const wrap=document.getElementById('wrap-p-asesor');
  if(wrap) wrap.style.width=Math.max(700, as.labels.length*70)+'px';
  mkBar('c-p-asesor', as.labels, as.values, {horiz:false});
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
        { label:'No vinculado', data:vaNo, backgroundColor:C.orange, borderRadius:5, borderSkipped:false }
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
            color: ctx => ctx.datasetIndex===0 ? '#059669' : '#dc2626',
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
    areas = Object.entries(global).sort((a,b)=>b[1]-a[1]).slice(0,20).map(([area,count])=>({area,count}));
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

function renderSolicitud() {
  const rows = filterRows(D.raw_solicitud, {progKey:'PERFIL_SOLICITADO', facKey:null});
  const n=rows.length;
  const nuevas=rows.filter(r=>r.EMPRESA_NUEVA&&r.EMPRESA_NUEVA.toLowerCase()==='si').length;
  const emps=new Set(rows.map(r=>r.EMPRESA).filter(Boolean)).size;
  const progs=new Set(rows.map(r=>r.PERFIL_SOLICITADO).filter(Boolean)).size;

  document.getElementById('hero-solic').innerHTML=
    heroStat(n,'Solicitudes','') +
    heroStat(emps,'Empresas','gold') +
    heroStat(nuevas,'Nuevas','green') +
    heroStat(progs,'Perfiles','');

  document.getElementById('kpi-solic').innerHTML=
    kpiCard(n,'Solicitudes recibidas','Total período','') +
    kpiCard(emps,'Empresas únicas','Que han solicitado','gold') +
    kpiCard(nuevas,'Empresas nuevas',`${n?Math.round(nuevas/n*100):0}% de solicitudes`,'green') +
    kpiCard(progs,'Perfiles solicitados','Programas demandados','purple');

  const top1P=groupBy(rows,'PERFIL_SOLICITADO',1).labels[0]||'N/A';
  const top1M=groupBy(rows,'MODALIDAD',1).labels[0]||'N/A';
  document.getElementById('stat-solic').innerHTML=
    `Solicitudes registradas en el período: <strong>${n}</strong>.
     Empresas únicas: <strong>${emps}</strong>.
     Perfil más solicitado: <strong>${top1P}</strong>.
     Modalidad predominante: <strong>${top1M}</strong>.
     Empresas nuevas en el período: <strong>${nuevas} (${n?Math.round(nuevas/n*100):0}%)</strong>.`;

  const perf=groupBy(rows,'PERFIL_SOLICITADO',20);
  mkBar('c-s-perfil', perf.labels, perf.values, {horiz:true});

  const mod=groupBy(rows,'MODALIDAD');
  mkDoughnut('c-s-modalidad', mod.labels, mod.values, [C.blue, C.gold]);

  const en=groupBy(rows,'EMPRESA_NUEVA');
  mkDoughnut('c-s-empnueva', en.labels, en.values, [C.gold, C.blue]);

  const mes=mesSort(rows);
  mkLine('c-s-mes', mes.labels, mes.values, C.gold);

  const anio=groupBy(rows.filter(r=>r.ANIO&&r.ANIO>0),'ANIO');
  const anioSorted=Object.entries(
    rows.filter(r=>r.ANIO&&r.ANIO>0).reduce((m,r)=>{m[r.ANIO]=(m[r.ANIO]||0)+1;return m},{})
  ).sort((a,b)=>a[0]-b[0]);
  mkLine('c-s-anio', anioSorted.map(x=>x[0]), anioSorted.map(x=>x[1]), C.blue);
}

function renderAprobacion() {
  const rows = filterRows(D.raw_aprobacion, {facKey:null});
  const n=rows.length;
  const aprob=rows.filter(r=>r.ESTADO_APROBACION&&r.ESTADO_APROBACION.toLowerCase()==='aprobado').length;
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

  const est=groupBy(rows,'ESTADO_APROBACION');
  mkDoughnut('c-a-estado', est.labels, est.values, [C.green, C.gold]);

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
function parseMulti(rows, key, top=12) {
  const c={};
  rows.forEach(r=>{
    if(!r[key]) return;
    r[key].split(';').forEach(p=>{
      p=p.trim();
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
  const rows = filterRows(D.enc_raw, {progKey:'programa', facKey:null});
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

  // Programas
  const prog = groupBy(rows,'programa',20);
  const wrapP = document.getElementById('wrap-enc-prog');
  if(wrapP) wrapP.style.width = Math.max(700, prog.labels.length*62)+'px';
  mkBar('c-enc-programas', prog.labels, prog.values, {horiz:false});

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
    html  = HTML.replace('__LOGO__', logo).replace('__DATA_JSON__', dj)

    out = 'informe_practicas_itm.html'
    with open(out, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f'OK  Generado: {out}  ({len(html)//1024} KB)')
    print(f'    Practicantes: {len(df1)} | Disponibles: {len(df2)} | F082: {len(df3)} | Solicitud: {len(df4)} | Aprobacion: {len(df5)} | Enc.Empresarios: {len(df6)} | Enc.Estudiantes: {len(df7)}')
