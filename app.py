"""
Ferramenta de Validação de Faturamento Excel
Aplicação Streamlit para upload e processamento de arquivos Excel
"""

import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# ========================================
# Configuração da Página
# ========================================

st.set_page_config(
    page_title="Validação de Faturamento Excel",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ========================================
# Título Principal
# ========================================

st.title("📊 Ferramenta de Validação de Faturamento")
st.markdown("---")

# ========================================
# Sidebar - Configurações
# ========================================

st.sidebar.header("⚙️ Configurações")
st.sidebar.markdown("### 📅 Período de Análise")

# Selectbox para Mês
meses = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 
         'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']
mes_selecionado = st.sidebar.selectbox(
    "Mês",
    options=meses,
    index=0
)

# Selectbox para Ano
anos = ['24', '25', '26']
ano_selecionado = st.sidebar.selectbox(
    "Ano",
    options=anos,
    index=1  # Default para '25'
)

# Concatenar para formar target_month no formato MMM.YY
target_month = f"{mes_selecionado}.{ano_selecionado}"

# Exibir o período selecionado
st.sidebar.success(f"**Período Selecionado:** {target_month}")
st.sidebar.markdown("---")

# ========================================
# Área Principal - Upload de Arquivos
# ========================================

st.header("📁 Upload de Arquivos")

col1, col2 = st.columns(2)

with col1:
    st.subheader("Arquivo PARCEIRO")
    arquivo_parceiro = st.file_uploader(
        "Selecione o arquivo PARCEIRO (.xlsx)",
        type=['xlsx'],
        key='parceiro',
        help="Arquivo Excel com dados do parceiro"
    )
    
    if arquivo_parceiro:
        st.success(f"✅ {arquivo_parceiro.name}")
        st.info(f"Tamanho: {arquivo_parceiro.size / 1024:.2f} KB")

with col2:
    st.subheader("Arquivo BASE")
    arquivo_base = st.file_uploader(
        "Selecione o arquivo BASE (.xlsx ou .xlsm)",
        type=['xlsx', 'xlsm'],
        key='base',
        help="Arquivo Excel base (fórmulas serão preservadas)"
    )
    
    if arquivo_base:
        st.success(f"✅ {arquivo_base.name}")
        st.info(f"Tamanho: {arquivo_base.size / 1024:.2f} KB")

st.markdown("---")

# ========================================
# Botão de Processamento
# ========================================

st.header("🚀 Processamento")

# Verificar se ambos os arquivos foram carregados
arquivos_prontos = arquivo_parceiro is not None and arquivo_base is not None

if not arquivos_prontos:
    st.warning("⚠️ Por favor, faça upload dos dois arquivos para continuar.")

# Botão de processamento
processar = st.button(
    "🔄 Iniciar Processamento",
    type="primary",
    disabled=not arquivos_prontos,
    use_container_width=True
)

# ========================================
# Lógica de Processamento e Session State
# ========================================

if processar and arquivos_prontos:
    try:
        with st.spinner("Processando arquivos..."):
            
            # Armazenar target_month no session_state
            st.session_state['target_month'] = target_month
            
            # ==========================================
            # Processar Arquivo PARCEIRO
            # ==========================================
            st.info("📄 Carregando arquivo PARCEIRO...")
            
            # Carregar arquivo PARCEIRO com pandas
            parceiro_data = pd.read_excel(arquivo_parceiro)
            st.session_state['parceiro_data'] = parceiro_data
            st.session_state['parceiro_filename'] = arquivo_parceiro.name
            
            # ==========================================
            # Processar Arquivo BASE
            # ==========================================
            st.info("📄 Carregando arquivo BASE (preservando fórmulas)...")
            
            # Carregar arquivo BASE com openpyxl (data_only=False para preservar fórmulas)
            base_workbook = openpyxl.load_workbook(
                BytesIO(arquivo_base.read()),
                data_only=False
            )
            st.session_state['base_workbook'] = base_workbook
            st.session_state['base_filename'] = arquivo_base.name
            
            # Converter primeira aba para DataFrame para preview
            primeira_aba = base_workbook.sheetnames[0]
            ws = base_workbook[primeira_aba]
            
            # Extrair dados para DataFrame
            data = ws.values
            cols = next(data)
            base_data = pd.DataFrame(data, columns=cols)
            st.session_state['base_data'] = base_data
            st.session_state['base_sheetnames'] = base_workbook.sheetnames
            
        # Mensagem de sucesso
        st.success("✅ Arquivos processados com sucesso!")
        st.balloons()
        
        # Flag para indicar que o processamento foi concluído
        st.session_state['processado'] = True
        
    except Exception as e:
        st.error(f"❌ Erro ao processar arquivos: {str(e)}")
        st.exception(e)

# ========================================
# Exibir Preview dos Dados (se processados)
# ========================================

if st.session_state.get('processado', False):
    st.markdown("---")
    st.header("👁️ Preview dos Dados")
    
    # Informações gerais
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Período", st.session_state['target_month'])
    
    with col2:
        st.metric("Linhas PARCEIRO", len(st.session_state['parceiro_data']))
    
    with col3:
        st.metric("Abas BASE", len(st.session_state.get('base_sheetnames', [])))
    
    st.markdown("---")
    
    # Preview do arquivo PARCEIRO
    st.subheader(f"📊 Arquivo PARCEIRO: {st.session_state.get('parceiro_filename', '')}")
    parceiro_df = st.session_state['parceiro_data']
    
    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**Dimensões:** {parceiro_df.shape[0]} linhas × {parceiro_df.shape[1]} colunas")
    with col2:
        st.write(f"**Colunas:** {', '.join(parceiro_df.columns.astype(str).tolist()[:5])}{'...' if len(parceiro_df.columns) > 5 else ''}")
    
    st.dataframe(parceiro_df.head(10), use_container_width=True)
    
    st.markdown("---")
    
    # Preview do arquivo BASE
    st.subheader(f"📊 Arquivo BASE: {st.session_state.get('base_filename', '')}")
    base_df = st.session_state['base_data']
    
    # Informações sobre as abas
    abas = st.session_state.get('base_sheetnames', [])
    st.write(f"**Abas disponíveis:** {', '.join(abas)}")
    
    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**Dimensões (1ª aba):** {base_df.shape[0]} linhas × {base_df.shape[1]} colunas")
    with col2:
        st.write(f"**Colunas:** {', '.join(base_df.columns.astype(str).tolist()[:5])}{'...' if len(base_df.columns) > 5 else ''}")
    
    st.dataframe(base_df.head(10), use_container_width=True)
    
    st.info("💡 **Nota:** As fórmulas do arquivo BASE foram preservadas no objeto openpyxl armazenado no session_state.")
    
    st.markdown("---")
    st.success("✅ Sistema pronto para próximas etapas de validação!")

# ========================================
# Footer
# ========================================

st.sidebar.markdown("---")
st.sidebar.markdown("### 📌 Instruções")
st.sidebar.markdown("""
1. Selecione o **mês** e **ano**
2. Faça upload do arquivo **PARCEIRO**
3. Faça upload do arquivo **BASE**
4. Clique em **Iniciar Processamento**
5. Visualize o preview dos dados
""")

st.sidebar.markdown("---")
st.sidebar.caption("Ferramenta de Validação de Faturamento v1.0")
