import os
import re
import docx
import openpyxl

# MÓDULO DE INTELIGÊNCIA DE ENGENHARIA 
#______________________________________________________________________
def convert_to_mm(dim_str):
    """Converte dimensões em polegadas (ex: "1.1/2"") para mm."""
    dim_str = dim_str.strip().replace(',', '.')
    total_mm = 0.0
    try:
        if '"' in dim_str:
            dim_str = dim_str.replace('"', '')
            parts = dim_str.split('.')
            if parts[0] and '/' in parts[0]:
                num, den = map(float, parts[0].split('/'))
                total_mm += (num / den) * 25.4
            elif parts[0]:
                total_mm += float(parts[0]) * 25.4
            if len(parts) > 1 and '/' in parts[1]:
                num, den = map(float, parts[1].split('/'))
                total_mm += (num / den) * 25.4
        else:
            total_mm = float(dim_str)
    except (ValueError, ZeroDivisionError): return 0.0
    return total_mm

# ==============================================================================
# Função para identificar o tipo de perfil e retornar o código correto
# ==============================================================================
def classificar_e_mapear_perfil(desc):
    """Identifica o TIPO de perfil e retorna o código e uma chave de classificação."""
    desc_upper = desc.upper()
    if '[' in desc_upper or '][' in desc_upper: return 'U.s', 'PERFIL_U'
    if 'UENR' in desc_upper or 'IENR' in desc_upper or 'CART' in desc_upper or 'CA ' in desc_upper: return 'U.e', 'TERCA'
    if 'L DOBRADO' in desc_upper or desc_upper.startswith('L '): return 'L DOBRADO', 'CANTONEIRA'
    
    
    # Se encontrar 'RED', mapeia para o código exato da planilha.
    if 'RED' in desc_upper:
        return 'FERRO MECANICO RED.', 'TUBO' # Mapeamento corrigido
        
    if 'TUBO' in desc_upper: return 'TUBO', 'TUBO' # Mantém genérico se for outra coisa
    
    return 'N/D', 'OUTROS'


# ==============================================================================
# Extrai as 4 medidas principais de uma descrição de perfil
# ==============================================================================
def parse_dimensoes_inteligente(desc, tipo_perfil):
    """Aplica regras de extração de dimensões e retorna as 4 medidas principais."""
    a, b, c, esp = 0.0, 0.0, 0.0, 0.0
    numeros_str_list = re.findall(r'[\d\./"]+', desc)
    
    if tipo_perfil in ['PERFIL_U']:
        if len(numeros_str_list) >= 3:
            a = convert_to_mm(numeros_str_list[0])
            b = convert_to_mm(numeros_str_list[1])
            esp = convert_to_mm(numeros_str_list[2])
    elif tipo_perfil == 'TERCA':
        if len(numeros_str_list) >= 4:
            a = convert_to_mm(numeros_str_list[0])
            b = convert_to_mm(numeros_str_list[1])
            c = convert_to_mm(numeros_str_list[2])
            esp = convert_to_mm(numeros_str_list[3])
    elif tipo_perfil == 'CANTONEIRA':
        if len(numeros_str_list) == 2:
            aba = convert_to_mm(numeros_str_list[0])
            a, b = aba, aba
            esp = convert_to_mm(numeros_str_list[1])
        elif len(numeros_str_list) >= 3:
            a = convert_to_mm(numeros_str_list[0])
            b = convert_to_mm(numeros_str_list[1])
            esp = convert_to_mm(numeros_str_list[2])
    
    # --- CORREÇÃO APLICADA AQUI ---
    # Lógica específica para Tubo/RED
    elif tipo_perfil == 'TUBO':
        # Para RED 12.7, a única medida é a espessura/diâmetro.
        if len(numeros_str_list) >= 1:
            esp = convert_to_mm(numeros_str_list[0]) # Coloca o valor na variável 'esp'
            
    return a, b, c, esp

# ==============================================================================
# MÓDULO PRINCIPAL DO SCRIPT (Leitura e Preenchimento Não-Destrutivo)
# ==============================================================================

def extrair_dados_word(caminho_arquivo_word):
    """Função robusta para extrair dados do Word."""
    try:
        documento = docx.Document(caminho_arquivo_word) #<----------------- Abrir o arquivo Word como não foi atribuido um camiho, utiliza o diretório atual (Pasta do script)
        tabela = documento.tables[0]
        if len(tabela.rows) < 2: return None #<----------------- Verifica se há pelo menos uma linha de dados além do cabeçalho
        perfils_str = tabela.cell(1, 0).text; acos_str = tabela.cell(1, 1).text
        ltotais_str = tabela.cell(1, 2).text; pesos_str = tabela.cell(1, 3).text
        lista_perfis = list(filter(None, perfils_str.strip().split('\n')))
        lista_acos = list(filter(None, acos_str.strip().split('\n')))
        lista_ltotais = list(filter(None, ltotais_str.strip().split('\n')))
        lista_pesos = list(filter(None, pesos_str.strip().split('\n')))
        num_perfis = len(lista_perfis)
        if not (num_perfis == len(lista_ltotais) == len(lista_pesos)): return None
        if num_perfis == 0: return None
        print(f"\nSUCESSO NA LEITURA: {num_perfis} linhas de dados encontradas. Processando...\n")
        dados_finais = []
        for i in range(num_perfis):
            perfil, aco = lista_perfis[i].strip(), lista_acos[i].strip() if i < len(lista_acos) else lista_acos[0].strip()
            l_total_str, peso_str = lista_ltotais[i].strip().replace(',', '.'), lista_pesos[i].strip().replace(',', '.')
            try:
                l_total_m = float(l_total_str) / 100 if l_total_str else 0.0
                peso_final = float(peso_str) if peso_str else 0.0
                dados_finais.append([perfil, aco, l_total_m, peso_final])
            except ValueError: continue
        return dados_finais
    except Exception as e:
        print(f"Erro crítico ao ler o arquivo Word: {e}")
        return None

def encontrar_proxima_linha_vazia(sheet, codigo_secao, linha_inicio_busca):
    """
    Encontra a primeira linha vazia para uma seção, aceitando placeholders como 'X' ou 0.
    """
    for row in range(linha_inicio_busca, sheet.max_row + 2):
        celula_codigo = sheet.cell(row=row, column=1)
        # --- LÓGICA CORRIGIDA ---
        # A referência agora é a Coluna B, a primeira que vamos preencher.
        celula_dado_ref = sheet.cell(row=row, column=2) 
        
        # Considera a linha vazia se o código da seção bate E a célula de dados é um placeholder.
        if celula_codigo.value == codigo_secao and celula_dado_ref.value in [None, 0, 'X', '']:
            return row
    return None

def preencher_planilha_excel(caminho_planilha, dados_materiais):
    """Preenche a planilha de forma não-destrutiva, seguindo a estrutura de colunas exata."""
    try:
        workbook = openpyxl.load_workbook(caminho_planilha)
        sheet = workbook.active
        
        dados_agrupados = {}
        for item in dados_materiais:
            codigo_excel, _ = classificar_e_mapear_perfil(item[0])
            if codigo_excel not in dados_agrupados: dados_agrupados[codigo_excel] = []
            dados_agrupados[codigo_excel].append(item)
            
        for codigo_secao, itens_da_secao in dados_agrupados.items():
            print(f"Processando seção '{codigo_secao}' com {len(itens_da_secao)} itens...")
            
            linha_de_busca_secao = 4
            for item in itens_da_secao:
                linha_alvo = encontrar_proxima_linha_vazia(sheet, codigo_secao, linha_de_busca_secao)
                
                if linha_alvo is None:
                    print(f"  AVISO: Não há mais espaço na planilha para a seção '{codigo_secao}'. Item '{item[0]}' não inserido.")
                    continue

                perfil_desc, aco_tipo, l_total_m, peso_total = item
                _, tipo_perfil = classificar_e_mapear_perfil(perfil_desc)
                dim_a, dim_b, dim_c, dim_esp = parse_dimensoes_inteligente(perfil_desc, tipo_perfil)

                #PARA PERFIS NOVOS UTILIZAR ESSE BLOCO PARA CADASTRAR NOVOS PERFIS
                if tipo_perfil in ['PERFIL_U', 'TERCA']:
                    sheet.cell(row=linha_alvo, column=2).value = dim_a         # Coluna B -> Medida A
                    sheet.cell(row=linha_alvo, column=4).value = dim_b         # Coluna D -> Medida B
                    sheet.cell(row=linha_alvo, column=6).value = dim_c         # Coluna F -> Medida C

                elif tipo_perfil == 'CANTONEIRA':                              # Preenche APENAS as colunas D e F para Cantoneiras
                    sheet.cell(row=linha_alvo, column=4).value = dim_a         # Coluna D -> Medida A (CANTONEIRA)
                    sheet.cell(row=linha_alvo, column=6).value = dim_b         # Coluna F -> Medida B (CANTONEIRA)

                # --- MAPEAMENTO DE COLUNAS ---
                
                sheet.cell(row=linha_alvo, column=8).value = dim_esp       # Coluna H -> esp.
                sheet.cell(row=linha_alvo, column=9).value = aco_tipo      # Coluna I -> Tipo de Material
                sheet.cell(row=linha_alvo, column=10).value = l_total_m    # Coluna J -> Quantidade (m)
                sheet.cell(row=linha_alvo, column=17).value = peso_total   # Coluna Q -> Total Kg
                # -----------------------------
                linha_de_busca_secao = linha_alvo + 1

        workbook.save(caminho_planilha) #<----------------- Salva as alterações na planilha (utiliza o diretório atual se não for atribuido um caminho)
        print(f"\nSUCESSO FINAL: Planilha atualizada com os dados preenchidos em suas seções.")

    except PermissionError:
        print("\nERRO: A planilha Excel está aberta. Por favor, feche-a e tente novamente.")
    except Exception as e:
        print(f"Erro ao preencher a planilha Excel: {e}")

# ==============================================================================
# PONTO DE PARTIDA DO SCRIPT
# ==============================================================================
if __name__ == "__main__":
    arquivo_word = "lista-material.docx" #<----------------- Nome do arquivo Word (deve estar na mesma pasta do script ou fornecer o caminho completo)
    planilha_excel = "TABELA-DE-AÇO R8.xlsx" #<----------------- Nome da planilha Excel (deve estar na mesma pasta do script ou fornecer o caminho completo)
    
    try:
        with open(planilha_excel, "a"): pass
    except IOError:
        print(f"\nERRO: A planilha '{planilha_excel}' está aberta ou não pôde ser acessada.")
        exit()

    dados_extraidos = extrair_dados_word(arquivo_word)
    if dados_extraidos:
        preencher_planilha_excel(planilha_excel, dados_extraidos)
    else:
        print("Nenhum dado foi extraído do arquivo Word. Verifique o arquivo.")