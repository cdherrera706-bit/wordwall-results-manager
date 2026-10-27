"""Wordwall Results Manager — run with: streamlit run wordwall_app.py"""
import io
import re
import unicodedata
from collections import defaultdict
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title='Wordwall Results Manager', page_icon='📚', layout='wide')
st.title('📚 Wordwall Results Manager')
st.caption('Consolida actividades, clasifica por grupo y promedia hasta los tres primeros intentos por actividad.')

NAME_RE = re.compile(r'^\s*(\d+[A-Za-z]?)\s+(.+?)\s*$')
CORRECT = {'✔', '✓', 'true', 'correct', 'correcto', '1', 'yes', 'sí', 'si'}
INCORRECT = {'✖', '✘', '✗', 'false', 'incorrect', 'incorrecto', '0', 'no'}

def normalize_name(value):
    text = unicodedata.normalize('NFKD', str(value))
    text = ''.join(c for c in text if not unicodedata.combining(c))
    return ' '.join(text.casefold().split())

def parse_detailed(df, filename):
    """Read Wordwall detailed answer sheets (A=questions/names, C=correct)."""
    rows = df.fillna('').astype(str).values.tolist()
    activity = str(rows[0][0]).strip() if rows and rows[0] else Path(filename).stem
    if not activity: activity = Path(filename).stem
    results = []
    person = None
    correct = total = 0
    def finish():
        if person is not None and total:
            group, name = person
            results.append({'Grupo':group,'Alumno':name,'Actividad':activity,
                            'Archivo':filename,'Aciertos':correct,'Preguntas':total,
                            'Calificación':round(100*correct/total, 6)})
    for row in rows[1:]:
        a = str(row[0]).strip() if len(row)>0 else ''
        b = str(row[1]).strip() if len(row)>1 else ''
        c = str(row[2]).strip() if len(row)>2 else ''
        m = NAME_RE.fullmatch(a) if not b and not c else None
        if m:
            finish()
            person = (m.group(1), m.group(2).strip())
            correct = total = 0
            continue
        if person is None: continue
        if a.casefold() == 'question' and b.casefold() == 'answer': continue
        mark = c.casefold()
        if mark in CORRECT or mark in INCORRECT:
            total += 1
            correct += int(mark in CORRECT)
    finish()
    return results

def read_upload(upload):
    data = upload.getvalue()
    if upload.name.lower().endswith('.csv'):
        for encoding in ('utf-8-sig','cp1252','latin-1'):
            try:
                frame = pd.read_csv(io.BytesIO(data), header=None, encoding=encoding, sep=None, engine='python', dtype=str)
                return parse_detailed(frame, upload.name)
            except UnicodeError: continue
    else:
        frame = pd.read_excel(io.BytesIO(data), header=None, dtype=str, engine='openpyxl')
        return parse_detailed(frame, upload.name)
    raise ValueError('No se pudo leer el CSV')

uploads = st.file_uploader('Sube tus resultados de Wordwall (Excel o CSV)', type=['xlsx','csv'], accept_multiple_files=True)
period = st.text_input('Mes del reporte (etiqueta)', value='Octubre 2026')
st.info('Los intentos se ordenan por su aparición en los archivos cargados. Si un archivo no incluye fecha/hora por intento, este orden no necesariamente representa el orden cronológico real.')

if uploads:
    attempts = []
    errors = []
    for upload in uploads:
        try:
            found = read_upload(upload)
            if not found: errors.append(f'{upload.name}: no se identificaron intentos. Revisa que sea un reporte detallado con columnas Question, Answer y Correct.')
            attempts.extend(found)
        except Exception as exc:
            errors.append(f'{upload.name}: {exc}')
    for error in errors: st.warning(error)
    if attempts:
        raw = pd.DataFrame(attempts)
        # Names are normalized only for grouping; preserve the first display spelling.
        raw['_key'] = raw['Alumno'].map(normalize_name)
        raw['Intento'] = raw.groupby(['Grupo','_key','Actividad']).cumcount()+1
        raw['Incluido'] = raw['Intento'] <= 3
        used = raw[raw['Incluido']]
        scores = (used.groupby(['Grupo','_key','Actividad'], sort=False)
                  .agg(Calificación=('Calificación','mean'), Intentos=('Intento','count'), Alumno=('Alumno','first'))
                  .reset_index())
        scores['Calificación'] = scores['Calificación'].round(2)
        st.subheader('Resumen de resultados')
        c1,c2,c3 = st.columns(3)
        c1.metric('Grupos', raw['Grupo'].nunique())
        c2.metric('Actividades', raw['Actividad'].nunique())
        c3.metric('Alumnos', raw[['Grupo','_key']].drop_duplicates().shape[0])
        groups = sorted(raw['Grupo'].unique(), key=lambda x:(not x.isdigit(), int(x) if x.isdigit() else x))
        selected = st.selectbox('Ver grupo', groups)
        def group_table(group):
            sub = scores[scores['Grupo']==group]
            pivot = sub.pivot(index='_key', columns='Actividad', values='Calificación')
            names = sub.groupby('_key',sort=False)['Alumno'].first()
            pivot.insert(0, 'Alumno', names.reindex(pivot.index))
            cols = [x for x in pivot.columns if x!='Alumno']
            pivot['Promedio mensual'] = pivot[cols].mean(axis=1).round(2)
            pivot['Actividades realizadas'] = pivot[cols].notna().sum(axis=1)
            return pivot.reset_index(drop=True).sort_values('Alumno',key=lambda s:s.map(normalize_name))
        st.dataframe(group_table(selected), hide_index=True, use_container_width=True)
        with st.expander('Historial de intentos'):
            st.dataframe(raw.drop(columns=['_key']), hide_index=True, use_container_width=True)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            summary = pd.concat([group_table(g).assign(Grupo=g) for g in groups], ignore_index=True)
            summary.to_excel(writer, sheet_name='Resumen mensual', index=False)
            for g in groups:
                group_table(g).to_excel(writer, sheet_name=f'Grupo {g}'[:31], index=False)
            raw.drop(columns=['_key']).to_excel(writer, sheet_name='Historial intentos', index=False)
            scores.drop(columns=['_key']).to_excel(writer, sheet_name='Calificaciones', index=False)
            activity_summary = scores.groupby('Actividad').agg(Promedio=('Calificación','mean'),Alumnos=('Alumno','count')).reset_index()
            activity_summary.to_excel(writer, sheet_name='Actividades', index=False)
            for sheet in writer.book.worksheets:
                sheet.freeze_panes = 'B2'
                sheet.auto_filter.ref = sheet.dimensions
                from openpyxl.styles import Font, PatternFill, Alignment
                for cell in sheet[1]:
                    cell.font = Font(color='FFFFFF', bold=True)
                    cell.fill = PatternFill('solid', fgColor='174A69')
                    cell.alignment = Alignment(wrap_text=True)
                for col in sheet.columns:
                    letter = col[0].column_letter
                    width = max(len(str(c.value or '')) for c in list(col)[:250]) + 2
                    sheet.column_dimensions[letter].width = min(55, max(13,width))
        st.download_button('⬇️ Descargar reporte mensual de Excel', output.getvalue(),
                           file_name='Wordwall_'+re.sub(r'[^A-Za-z0-9_-]+','_',period)+'.xlsx',
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', type='primary')
    else: st.error('No se encontraron resultados compatibles.')
else:
    st.write('Carga los tres archivos de ejemplo para comenzar. Se aceptan varios archivos simultáneamente.')
