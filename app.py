import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime
import io
import zipfile
from openpyxl.styles import Font, PatternFill, Alignment

def extrair_dados_nfse(xml_file):
    
    all_dados = []

    if xml_file.name.endswith('.zip'):
        with zipfile.ZipFile(xml_file, 'r') as zip_ref:
            for xml_filename in zip_ref.namelist():
                if xml_filename.endswith('.xml'):
                    with zip_ref.open(xml_filename) as single_xml_file:
                        all_dados.extend(processar_xml(io.BytesIO(single_xml_file.read())))
    elif xml_file.name.endswith('.xml'):
        all_dados.extend(processar_xml(xml_file))

    return pd.DataFrame(all_dados)

def processar_xml(xml_content):
    try:
        tree = ET.parse(xml_content)
        root = tree.getroot()
    except ET.ParseError:
        return [] # Retorna lista vazia se o XML for inválido

    ns = {'ns': 'http://www.abrasf.org.br/nfse.xsd'}
    
    dados = []
    
    for nfse in root.findall('.//ns:InfNfse', ns):
        
        # Dados do Tomador (CPF ou CNPJ)
        cpf_cnpj_tomador_element = nfse.find('.//ns:IdentificacaoTomador/ns:CpfCnpj', ns)
        if cpf_cnpj_tomador_element is not None:
            cpf_tomador = cpf_cnpj_tomador_element.find('.//ns:Cpf', ns)
            cnpj_tomador = cpf_cnpj_tomador_element.find('.//ns:Cnpj', ns)
            cpf_cnpj_valor = cpf_tomador.text if cpf_tomador is not None else cnpj_tomador.text if cnpj_tomador is not None else 'N/A'
        else:
            cpf_cnpj_valor = 'N/A'

        razao_social_tomador_element = nfse.find('.//ns:Tomador/ns:RazaoSocial', ns)
        razao_social_tomador = razao_social_tomador_element.text if razao_social_tomador_element is not None else 'N/A'
        
        # Dados da NFS-e
        numero_nfse_element = nfse.find('.//ns:Numero', ns)
        numero_nfse = numero_nfse_element.text if numero_nfse_element is not None else 'N/A'
        
        valor_servicos_element = nfse.find('.//ns:Servico/ns:Valores/ns:ValorServicos', ns)
        valor_servicos = valor_servicos_element.text if valor_servicos_element is not None else '0'
        
        valor_iss_element = nfse.find('.//ns:ValoresNfse/ns:ValorIss', ns)
        valor_iss = valor_iss_element.text if valor_iss_element is not None else 'N/A'
        
        iss_retido_element = nfse.find('.//ns:Servico/ns:IssRetido', ns)
        iss_retido = iss_retido_element.text if iss_retido_element is not None else '2'
        iss_retido_texto = "Sim" if iss_retido == "1" else "Não"
        
        data_emissao_element = nfse.find('.//ns:DataEmissao', ns)
        if data_emissao_element is not None and data_emissao_element.text:
            try:
                data_emissao = datetime.fromisoformat(data_emissao_element.text).strftime('%d/%m/%Y')
            except:
                data_emissao = 'N/A'
        else:
            data_emissao = 'N/A'
        
        item_lista_servico_element = nfse.find('.//ns:Servico/ns:ItemListaServico', ns)
        item_lista_servico = item_lista_servico_element.text if item_lista_servico_element is not None else 'N/A'
        
        codigo_nbs_element = nfse.find('.//ns:Servico/ns:CodigoNbs', ns)
        codigo_nbs = codigo_nbs_element.text if codigo_nbs_element is not None else 'N/A'
        
        codigo_cnae_element = nfse.find('.//ns:Servico/ns:CodigoCnae', ns)
        codigo_cnae = codigo_cnae_element.text if codigo_cnae_element is not None else 'N/A'
        
        # Base de Cálculo IBSCBS
        vBC_element = nfse.find('.//ns:IBSCBS/ns:valores/ns:vBC', ns)
        vBC = vBC_element.text if vBC_element is not None else '0'
        
        # Dados IBSCBS - tratando elementos que podem não existir
        pIBSUF_element = nfse.find('.//ns:IBSCBS/ns:valores/ns:uf/ns:pIBSUF', ns)
        pIBSUF = pIBSUF_element.text if pIBSUF_element is not None else '0'
        
        pRedAliqUF_element = nfse.find('.//ns:IBSCBS/ns:valores/ns:uf/ns:pRedAliqUF', ns)
        pRedAliqUF = pRedAliqUF_element.text if pRedAliqUF_element is not None else '0'
        
        pAliqEfetUF_element = nfse.find('.//ns:IBSCBS/ns:valores/ns:uf/ns:pAliqEfetUF', ns)
        pAliqEfetUF = pAliqEfetUF_element.text if pAliqEfetUF_element is not None else '0'
        
        pRedAliqMun_element = nfse.find('.//ns:IBSCBS/ns:valores/ns:mun/ns:pRedAliqMun', ns)
        pRedAliqMun = pRedAliqMun_element.text if pRedAliqMun_element is not None else '0'
        
        pCBS_element = nfse.find('.//ns:IBSCBS/ns:valores/ns:fed/ns:pCBS', ns)
        pCBS = pCBS_element.text if pCBS_element is not None else '0'
        
        pRedAliqCBS_element = nfse.find('.//ns:IBSCBS/ns:valores/ns:fed/ns:pRedAliqCBS', ns)
        pRedAliqCBS = pRedAliqCBS_element.text if pRedAliqCBS_element is not None else '0'
        
        pAliqEfetCBS_element = nfse.find('.//ns:IBSCBS/ns:valores/ns:fed/ns:pAliqEfetCBS', ns)
        pAliqEfetCBS = pAliqEfetCBS_element.text if pAliqEfetCBS_element is not None else '0'
        
        vIBSUF_element = nfse.find('.//ns:IBSCBS/ns:totCIBS/ns:gIBS/ns:gIBSUFTot/ns:vIBSUF', ns)
        vIBSUF = vIBSUF_element.text if vIBSUF_element is not None else '0'
        
        vCBS_element = nfse.find('.//ns:IBSCBS/ns:totCIBS/ns:gCBS/ns:vCBS', ns)
        vCBS = vCBS_element.text if vCBS_element is not None else '0'
        
        # Descrição do Serviço
        discriminacao_element = nfse.find('.//ns:Servico/ns:Discriminacao', ns)
        discriminacao = discriminacao_element.text if discriminacao_element is not None else 'N/A'
        
        dados.append({
            "CPF/CNPJ Tomador": cpf_cnpj_valor,
            "Razão Social Tomador": razao_social_tomador,
            "Número NFS-e": numero_nfse,
            "Valor do Serviço": float(valor_servicos),
            "Alíquota": "3%",
            "ISS": float(valor_iss) if valor_iss != 'N/A' else 0,
            "ISS Retido": iss_retido_texto,
            "Data de Emissão": data_emissao,
            "Item": item_lista_servico,
            "Código NBS": codigo_nbs,
            "Código CNAE": codigo_cnae,
            "Base de Cálculo IBSCBS": float(vBC),
            "pIBSUF": float(pIBSUF),
            "pRedAliqUF": float(pRedAliqUF),
            "pAliqEfetUF": float(pAliqEfetUF),
            "pRedAliqMun": float(pRedAliqMun),
            "pCBS": float(pCBS),
            "pRedAliqCBS": float(pRedAliqCBS),
            "pAliqEfetCBS": float(pAliqEfetCBS),
            "vIBSUF": float(vIBSUF),
            "vCBS": float(vCBS),
            "Descrição do Serviço": discriminacao
        })
        
    return dados

def format_cpf_cnpj(value):
    if value == 'N/A' or value is None:
        return 'N/A'
    
    cleaned_value = ''.join(filter(str.isdigit, str(value)))
    
    if len(cleaned_value) == 11:
        return f'{cleaned_value[:3]}.{cleaned_value[3:6]}.{cleaned_value[6:9]}-{cleaned_value[9:]}'
    elif len(cleaned_value) == 14:
        return f'{cleaned_value[:2]}.{cleaned_value[2:5]}.{cleaned_value[5:8]}/{cleaned_value[8:12]}-{cleaned_value[12:]}'
    else:
        return value

def format_brazilian_currency(value):
    try:
        float_value = float(value)
        # Formata para o padrão brasileiro: R$ 1.234,56
        return f'R$ {float_value:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    except (ValueError, TypeError):
        return f'R$ 0,00'

st.set_page_config(layout="wide")
st.title("Relatório  NFS-e")

uploaded_file = st.file_uploader("Escolha um arquivo XML ou ZIP de NFS-e", type=["xml", "zip"])

if uploaded_file is not None:
    df = extrair_dados_nfse(uploaded_file)
    
    if not df.empty:
        # Quadro de Resumo
        st.subheader("📊 Resumo dos Valores")
        
        total_ibs = df['vIBSUF'].sum()
        total_cbs = df['vCBS'].sum()
        total_servicos = df['Valor do Serviço'].sum()
        total_iss = df['ISS'].sum()
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(label="Total IBS (UF)", value=format_brazilian_currency(total_ibs))
        
        with col2:
            st.metric(label="Total CBS", value=format_brazilian_currency(total_cbs))
        
        with col3:
            st.metric(label="Total de Serviços", value=format_brazilian_currency(total_servicos))
        
        with col4:
            st.metric(label="Total ISS", value=format_brazilian_currency(total_iss))
        
        st.divider()
        
        # Excluir as colunas solicitadas
        df = df.drop(columns=['pIBSMun', 'pAliqEfetMun', 'vIBSMun'], errors='ignore')
        
        if 'CPF/CNPJ Tomador' in df.columns:
            df['CPF/CNPJ Tomador'] = df['CPF/CNPJ Tomador'].apply(format_cpf_cnpj)

        # Colunas sem formatação percentual (valores numéricos simples)
        cols_numeric = ["pIBSUF", "pRedAliqUF", "pAliqEfetUF", "pRedAliqMun", "pCBS", "pRedAliqCBS", "pAliqEfetCBS"]

        cols_to_format_currency = [
            "Valor do Serviço", "ISS", "Base de Cálculo IBSCBS",
            "vIBSUF", "vCBS"
        ]
        
        all_numeric_cols = cols_numeric + cols_to_format_currency
        for col in all_numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # Estilizando o DataFrame
        formatters = {col: '{:.2f}'.format for col in cols_numeric}
        formatters.update({col: format_brazilian_currency for col in cols_to_format_currency})
        
        st.dataframe(df.style.format(formatters)
                           .set_properties(**{'text-align': 'center'})
                           .set_table_styles([
                               {'selector': 'thead th', 'props': [('background-color', '#1f77b4'), ('color', 'white'), ('text-align', 'center'), ('font-weight', 'bold')]},
                               {'selector': 'tbody tr', 'props': [('text-align', 'center')]}
                           ]))
        
        # Download Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Dados NFS-e')
            
            # Aplica a formatação na planilha do Excel
            workbook = writer.book
            worksheet = writer.sheets['Dados NFS-e']
            
            # Formatação de moeda
            currency_format = 'R$ #,##0.00'
            number_format = '0.00'
            
            # Estilo do cabeçalho
            header_fill = PatternFill(start_color='1F77B4', end_color='1F77B4', fill_type='solid')
            header_font = Font(bold=True, color='FFFFFF', size=11)
            header_alignment = Alignment(horizontal='center', vertical='center')
            
            # Alinhamento centralizado para células de dados
            center_alignment = Alignment(horizontal='center', vertical='center')
            
            # Aplicar estilo ao cabeçalho
            for cell in worksheet[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_alignment
            
            # Encontra as colunas para formatar e aplica o estilo
            for col_idx, col_name in enumerate(df.columns, 1):
                col_letter = chr(ord('A') + col_idx - 1) if col_idx <= 26 else chr(ord('A') + (col_idx - 1) // 26 - 1) + chr(ord('A') + (col_idx - 1) % 26)
                
                if col_name in cols_to_format_currency:
                    for cell in worksheet[col_letter][1:]: # Pula o cabeçalho
                        cell.number_format = currency_format
                        cell.alignment = center_alignment
                elif col_name in cols_numeric:
                    for cell in worksheet[col_letter][1:]: # Pula o cabeçalho
                        cell.number_format = number_format
                        cell.alignment = center_alignment
                else:
                    for cell in worksheet[col_letter][1:]: # Pula o cabeçalho
                        cell.alignment = center_alignment
                
                # Auto-ajustar largura das colunas
                max_length = 0
                column = worksheet[col_letter]
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)  # Limita a largura máxima em 50
                worksheet.column_dimensions[col_letter].width = adjusted_width

        st.download_button(
            label="Baixar Planilha Excel",
            data=output.getvalue(),
            file_name="dados_nfse.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Nenhum dado de NFS-e foi encontrado nos arquivos fornecidos.")