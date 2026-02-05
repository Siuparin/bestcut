# bestcut_webapp.py
# Installa prima: pip install streamlit openpyxl

import streamlit as st
from dataclasses import dataclass
from typing import List, Tuple
import copy
from itertools import combinations
from datetime import datetime
import pandas as pd
from io import BytesIO

# Per Excel
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    EXCEL_DISPONIBILE = True
except ImportError:
    EXCEL_DISPONIBILE = False

# Configurazione pagina Streamlit
st.set_page_config(
    page_title="BestCut - Ottimizzatore Taglio Tubi",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personalizzato per stile moderno
st.markdown("""
<style>
    .main-header {
        font-size: 3rem;
        font-weight: bold;
        color: #2196F3;
        text-align: center;
        margin-bottom: 0;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        background-color: #d4edda;
        color: #155724;
        padding: 1rem;
        border-radius: 10px;
        border-left: 5px solid #28a745;
    }
    .warning-box {
        background-color: #fff3cd;
        color: #856404;
        padding: 1rem;
        border-radius: 10px;
        border-left: 5px solid #ffc107;
    }
    .error-box {
        background-color: #f8d7da;
        color: #721c24;
        padding: 1rem;
        border-radius: 10px;
        border-left: 5px solid #dc3545;
    }
    .stButton>button {
        background-color: #2196F3;
        color: white;
        font-weight: bold;
        border-radius: 10px;
        padding: 0.5rem 2rem;
    }
    .stButton>button:hover {
        background-color: #1976D2;
    }
    .metric-card {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 10px;
        border: 1px solid #dee2e6;
    }
</style>
""", unsafe_allow_html=True)

@dataclass
class Spezzone:
    lunghezza: float
    id: int
    
@dataclass
class TaglioRichiesto:
    lunghezza: float
    quantita: int
    
@dataclass
class PianoTaglio:
    spezzone_id: int
    spezzone_lunghezza: float
    tagli: List[float]
    scarto: float

class OttimizzatoreTagli:
    def __init__(self, soglia_scarto: float = 0.3):
        self.soglia_scarto = soglia_scarto
        
    def calcola_ottimale(self, spezzoni: List[Spezzone], richieste: List[TaglioRichiesto]) -> Tuple[List[PianoTaglio], float, bool]:
        tagli_necessari = []
        for richiesta in richieste:
            for _ in range(richiesta.quantita):
                tagli_necessari.append(richiesta.lunghezza)
        
        tagli_necessari.sort(reverse=True)
        totale_spezzoni = sum(s.lunghezza for s in spezzoni)
        totale_tagli = sum(tagli_necessari)
        
        if totale_tagli > totale_spezzoni:
            return [], totale_spezzoni - totale_tagli, False
        
        spezzoni_work = copy.deepcopy(spezzoni)
        tagli_rimanenti = tagli_necessari.copy()
        piani = []
        
        while tagli_rimanenti:
            migliore_idx = -1
            migliore_combo = None
            migliore_scarto = float('inf')
            
            for idx, spezzone in enumerate(spezzoni_work):
                if spezzone.lunghezza <= 0:
                    continue
                    
                combo, scarto = self._trova_combinazione(spezzone.lunghezza, tagli_rimanenti)
                
                if combo and scarto < migliore_scarto:
                    migliore_scarto = scarto
                    migliore_idx = idx
                    migliore_combo = combo
            
            if migliore_idx == -1:
                return None, 0, False
            
            spezzone_usato = spezzoni_work[migliore_idx]
            for taglio in migliore_combo:
                tagli_rimanenti.remove(taglio)
            
            nuova_lunghezza = spezzone_usato.lunghezza - sum(migliore_combo)
            if nuova_lunghezza > 0.01:
                spezzoni_work[migliore_idx] = Spezzone(nuova_lunghezza, spezzone_usato.id)
            else:
                spezzoni_work.pop(migliore_idx)
            
            piani.append(PianoTaglio(
                spezzone_id=spezzone_usato.id,
                spezzone_lunghezza=spezzone_usato.lunghezza,
                tagli=migliore_combo,
                scarto=migliore_scarto
            ))
        
        scarto_totale = sum(p.scarto for p in piani)
        return piani, scarto_totale, True
    
    def _trova_combinazione(self, max_lunghezza: float, tagli_disp: List[float]) -> Tuple[List[float], float]:
        n = len(tagli_disp)
        best_combo = []
        best_scarto = max_lunghezza
        
        for r in range(1, min(n + 1, 6)):
            for combo_indices in combinations(range(n), r):
                somma = sum(tagli_disp[i] for i in combo_indices)
                if somma <= max_lunghezza:
                    scarto = max_lunghezza - somma
                    if scarto < best_scarto:
                        best_scarto = scarto
                        best_combo = [tagli_disp[i] for i in combo_indices]
        
        if not best_combo:
            return None, float('inf')
        
        return best_combo, best_scarto

def crea_excel_download(spezzoni, richieste, piani, soglia):
    """Crea file Excel in memoria per il download"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Piano Taglio"
    
    header_font = Font(name='Calibri', size=14, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2196F3", end_color="2196F3", fill_type="solid")
    subheader_font = Font(name='Calibri', size=12, bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                  top=Side(style='thin'), bottom=Side(style='thin'))
    
    ws.merge_cells('A1:E1')
    ws['A1'] = "PIANO DI TAGLIO TUBI"
    ws['A1'].font = Font(size=18, bold=True, color="2196F3")
    ws['A1'].alignment = Alignment(horizontal='center')
    ws.row_dimensions[1].height = 30
    
    ws['A2'] = f"Generato: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws['A2'].font = Font(italic=True)
    
    row = 4
    ws.merge_cells(f'A{row}:E{row}')
    ws[f'A{row}'] = "SPEZZONI DISPONIBILI"
    ws[f'A{row}'].font = header_font
    ws[f'A{row}'].fill = header_fill
    ws[f'A{row}'].alignment = Alignment(horizontal='center')
    
    row += 1
    for col, title in [('A', 'ID'), ('B', 'Lunghezza (m)'), ('C', 'Lunghezza (cm)')]:
        ws[f'{col}{row}'] = title
        ws[f'{col}{row}'].font = subheader_font
        ws[f'{col}{row}'].border = border
    
    for spezzone in spezzoni:
        row += 1
        ws[f'A{row}'] = spezzone.id
        ws[f'B{row}'] = spezzone.lunghezza
        ws[f'C{row}'] = spezzone.lunghezza * 100
        for col in ['A', 'B', 'C']:
            ws[f'{col}{row}'].border = border
    
    row += 2
    ws.merge_cells(f'A{row}:E{row}')
    ws[f'A{row}'] = "TAGLI RICHIESTI"
    ws[f'A{row}'].font = header_font
    ws[f'A{row}'].fill = header_fill
    ws[f'A{row}'].alignment = Alignment(horizontal='center')
    
    row += 1
    headers = [('A', 'Misura (m)'), ('B', 'Misura (cm)'), ('C', 'Quantità'), ('D', 'Totale (m)')]
    for col, title in headers:
        ws[f'{col}{row}'] = title
        ws[f'{col}{row}'].font = subheader_font
        ws[f'{col}{row}'].border = border
    
    for richiesta in richieste:
        row += 1
        ws[f'A{row}'] = richiesta.lunghezza
        ws[f'B{row}'] = richiesta.lunghezza * 100
        ws[f'C{row}'] = richiesta.quantita
        ws[f'D{row}'] = richiesta.lunghezza * richiesta.quantita
        for col in ['A', 'B', 'C', 'D']:
            ws[f'{col}{row}'].border = border
    
    row += 2
    ws.merge_cells(f'A{row}:E{row}')
    ws[f'A{row}'] = "PIANO DI TAGLIO DETTAGLIATO"
    ws[f'A{row}'].font = header_font
    ws[f'A{row}'].fill = header_fill
    ws[f'A{row}'].alignment = Alignment(horizontal='center')
    
    for piano in piani:
        row += 1
        ws.merge_cells(f'A{row}:E{row}')
        ws[f'A{row}'] = f"Spezzone #{piano.spezzone_id} ({piano.spezzone_lunghezza:.3f}m)"
        ws[f'A{row}'].font = subheader_font
        ws[f'A{row}'].fill = PatternFill(start_color="E3F2FD", fill_type="solid")
        
        row += 1
        headers = [('A', 'N°'), ('B', 'Misura (m)'), ('C', 'Misura (cm)'), 
                  ('D', 'Inizio (m)'), ('E', 'Fine (m)')]
        for col, title in headers:
            ws[f'{col}{row}'] = title
            ws[f'{col}{row}'].font = Font(bold=True)
            ws[f'{col}{row}'].border = border
            ws[f'{col}{row}'].fill = PatternFill(start_color="F5F5F5", fill_type="solid")
        
        posizione = 0.0
        for i, taglio in enumerate(piano.tagli, 1):
            row += 1
            ws[f'A{row}'] = i
            ws[f'B{row}'] = taglio
            ws[f'C{row}'] = taglio * 100
            ws[f'D{row}'] = posizione
            ws[f'E{row}'] = posizione + taglio
            for col in ['A', 'B', 'C', 'D', 'E']:
                ws[f'{col}{row}'].border = border
            posizione += taglio
        
        row += 1
        ws[f'A{row}'] = "SCARTO"
        color = "4CAF50" if piano.scarto <= soglia else "f44336"
        ws[f'A{row}'].font = Font(bold=True, color=color)
        ws.merge_cells(f'B{row}:C{row}')
        ws[f'B{row}'] = piano.scarto
        ws[f'B{row}'].font = Font(bold=True)
        ws[f'D{row}'] = "OTTIMALE" if piano.scarto <= soglia else "DA RIUTILIZZARE"
        for col in ['A', 'B', 'C', 'D']:
            ws[f'{col}{row}'].border = border
        row += 1
    
    for i, w in enumerate([15, 15, 15, 18, 18], 1):
        ws.column_dimensions[chr(64+i)].width = w
    
    # Salva in memoria
    excel_buffer = BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    return excel_buffer

def main():
    # Header
    st.markdown('<p class="main-header">🔧 BestCut</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Ottimizzatore intelligente per il taglio dei tubi</p>', unsafe_allow_html=True)
    
    # Inizializza session state
    if 'spezzoni' not in st.session_state:
        st.session_state.spezzoni = []
        st.session_state.prossimo_id = 1
        st.session_state.piani = None
        st.session_state.richieste = None
        st.session_state.soglia = 0.3
    
    # Layout a colonne
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("📦 Spezzoni Disponibili")
        
        # Input nuovo spezzone
        nuovo_spezzone = st.number_input(
            "Lunghezza spezzone (metri)",
            min_value=0.0,
            value=6.0,
            step=0.1,
            format="%.2f",
            key="input_spezzone"
        )
        
        if st.button("➕ Aggiungi Spezzone", use_container_width=True):
            if nuovo_spezzone > 0:
                st.session_state.spezzoni.append(
                    Spezzone(nuovo_spezzone, st.session_state.prossimo_id)
                )
                st.session_state.prossimo_id += 1
                st.success(f"✅ Aggiunto spezzone da {nuovo_spezzone:.2f}m")
                st.rerun()
            else:
                st.error("❌ Inserisci una lunghezza valida")
        
        # Mostra spezzoni in una tabella
        if st.session_state.spezzoni:
            data = [{"ID": s.id, "Lunghezza (m)": f"{s.lunghezza:.2f}", "Lunghezza (cm)": f"{s.lunghezza*100:.0f}"} 
                   for s in st.session_state.spezzoni]
            df = pd.DataFrame(data)
            st.dataframe(df, use_container_width=True, hide_index=True)
            
            # Selezione per eliminare
            id_da_rimuovere = st.selectbox(
                "Seleziona spezzone da rimuovere",
                options=[s.id for s in st.session_state.spezzoni],
                format_func=lambda x: f"ID {x} - {next(s.lunghezza for s in st.session_state.spezzoni if s.id == x):.2f}m"
            )
            
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("🗑️ Rimuovi", use_container_width=True):
                    st.session_state.spezzoni = [s for s in st.session_state.spezzoni if s.id != id_da_rimuovere]
                    st.success("✅ Rimosso!")
                    st.rerun()
            with col_btn2:
                if st.button("🗑️🗑️ Rimuovi Tutti", use_container_width=True):
                    st.session_state.spezzoni = []
                    st.session_state.prossimo_id = 1
                    st.success("✅ Tutti rimossi!")
                    st.rerun()
        else:
            st.info("💡 Nessuno spezzone inserito. Aggiungi i tubi disponibili.")
    
    with col2:
        st.subheader("✂️ Tagli Richiesti")
        
        # Soglia scarto
        st.session_state.soglia = st.number_input(
            "Soglia scarto accettabile (metri)",
            min_value=0.0,
            value=0.3,
            step=0.05,
            format="%.2f",
            help="Scarti inferiori a questo valore sono considerati ottimali"
        )
        st.caption(f"= {st.session_state.soglia*100:.0f} centimetri")
        
        # 5 campi per i tagli
        st.markdown("---")
        richieste = []
        
        for i in range(5):
            cols = st.columns([1, 2, 2])
            with cols[0]:
                st.markdown(f"**#{i+1}**")
            with cols[1]:
                misura = st.number_input(
                    f"Misura {i+1} (m)",
                    min_value=0.0,
                    value=3.2 if i == 0 else (0.5 if i == 1 else 0.0),
                    step=0.1,
                    format="%.2f",
                    key=f"misura_{i}",
                    label_visibility="collapsed"
                )
            with cols[2]:
                qty = st.number_input(
                    f"Qty {i+1}",
                    min_value=0,
                    value=1 if i == 0 else (5 if i == 1 else 0),
                    step=1,
                    key=f"qty_{i}",
                    label_visibility="collapsed"
                )
            
            if misura > 0 and qty > 0:
                richieste.append(TaglioRichiesto(misura, qty))
        
        st.markdown("---")
        st.write(f"**{len(richieste)} tipi di tagli configurati**")
    
    # Bottone calcola centrato
    st.markdown("---")
    col_center = st.columns([1, 2, 1])
    with col_center[1]:
        if st.button("🚀 CALCOLA OTTIMIZZAZIONE", use_container_width=True, type="primary"):
            if not st.session_state.spezzoni:
                st.error("❌ Aggiungi almeno uno spezzone!")
            elif not richieste:
                st.error("❌ Inserisci almeno un taglio!")
            else:
                with st.spinner("⏳ Calcolo in corso..."):
                    ottim = OttimizzatoreTagli(st.session_state.soglia)
                    piani, scarto_tot, ok = ottim.calcola_ottimale(
                        copy.deepcopy(st.session_state.spezzoni), 
                        richieste
                    )
                    st.session_state.piani = piani
                    st.session_state.richieste = richieste
                
                if not ok:
                    st.markdown('<div class="error-box">❌ IMPOSSIBILE! Spezzoni insufficienti.</div>', unsafe_allow_html=True)
                    col_a, col_b = st.columns(2)
                    with col_a:
                        st.metric("Disponibile", f"{sum(s.lunghezza for s in st.session_state.spezzoni):.2f}m")
                    with col_b:
                        st.metric("Richiesto", f"{sum(r.lunghezza*r.quantita for r in richieste):.2f}m")
                else:
                    st.success("✅ Ottimizzazione completata!")
    
    # Risultati
    if st.session_state.piani:
        st.markdown("---")
        st.subheader("📋 Risultati")
        
        # Metriche
        piani = st.session_state.piani
        richieste = st.session_state.richieste
        scarto_tot = sum(p.scarto for p in piani)
        efficienza = (1 - scarto_tot/sum(p.spezzone_lunghezza for p in piani))*100
        
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        with col_m1:
            st.metric("Spezzoni usati", len(piani))
        with col_m2:
            st.metric("Scarto totale", f"{scarto_tot:.3f}m")
        with col_m3:
            st.metric("Scarto", f"{scarto_tot*100:.1f}cm")
        with col_m4:
            st.metric("Efficienza", f"{efficienza:.1f}%")
        
        # Dettaglio per spezzone
        for piano in piani:
            with st.expander(f"🔧 Spezzone #{piano.spezzone_id} ({piano.spezzone_lunghezza:.3f}m)"):
                # Tabella tagli
                data_tagli = []
                pos = 0.0
                for i, taglio in enumerate(piano.tagli, 1):
                    data_tagli.append({
                        "N°": i,
                        "Misura (m)": f"{taglio:.3f}",
                        "Misura (cm)": f"{taglio*100:.1f}",
                        "Inizio (m)": f"{pos:.3f}",
                        "Fine (m)": f"{pos+taglio:.3f}"
                    })
                    pos += taglio
                
                df_tagli = pd.DataFrame(data_tagli)
                st.dataframe(df_tagli, use_container_width=True, hide_index=True)
                
                # Scarto
                if piano.scarto <= st.session_state.soglia:
                    st.success(f"✅ Scarto: {piano.scarto:.3f}m ({piano.scarto*100:.1f}cm) - OTTIMALE")
                else:
                    st.warning(f"⚠️ Scarto: {piano.scarto:.3f}m ({piano.scarto*100:.1f}cm) - DA RIUTILIZZARE")
        
        # Riepilogo
        st.markdown("---")
        st.subheader("📊 Riepilogo Tagli")
        
        tagli_fatti = {}
        for p in piani:
            for t in p.tagli:
                tagli_fatti[t] = tagli_fatti.get(t, 0) + 1
        
        data_riep = []
        for rich in richieste:
            fatti = tagli_fatti.get(rich.lunghezza, 0)
            data_riep.append({
                "Misura": f"{rich.lunghezza:.2f}m",
                "Richiesti": rich.quantita,
                "Ottenuti": fatti,
                "Stato": "✅ Completato" if fatti >= rich.quantita else "⚠️ Parziale"
            })
        
        df_riep = pd.DataFrame(data_riep)
        st.dataframe(df_riep, use_container_width=True, hide_index=True)
        
        # Download Excel
        if EXCEL_DISPONIBILE:
            st.markdown("---")
            excel_buffer = crea_excel_download(
                st.session_state.spezzoni,
                richieste,
                piani,
                st.session_state.soglia
            )
            
            st.download_button(
                label="📥 Scarica Report Excel",
                data=excel_buffer,
                file_name=f"BestCut_Piano_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

if __name__ == "__main__":
    main()
