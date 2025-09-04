import pandas as pd
import requests
import re
import time
import os
import urllib3

# Desabilitar avisos de requisição insegura (não recomendado para produção)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- PARÂMETROS DE TEMPO AJUSTÁVEIS ---
# Pausa em segundos entre a consulta de cada CNPJ da lista.
PAUSA_ENTRE_CNPJS = 1
# Pausa base em segundos para as tentativas em caso de falha (será multiplicada a cada tentativa).
PAUSA_POR_TENTATIVA = 3
# -----------------------------------------

# --- MAPEAMENTOS PARA DERIVAÇÃO DE DADOS ---

def obter_regiao_por_uf(uf):
    """Retorna a região do Brasil com base na sigla da Unidade Federativa."""
    mapa_uf_regiao = {
        'AC': 'Norte', 'AP': 'Norte', 'AM': 'Norte', 'PA': 'Norte', 'RO': 'Norte', 'RR': 'Norte', 'TO': 'Norte',
        'AL': 'Nordeste', 'BA': 'Nordeste', 'CE': 'Nordeste', 'MA': 'Nordeste', 'PB': 'Nordeste', 'PE': 'Nordeste', 'PI': 'Nordeste', 'RN': 'Nordeste', 'SE': 'Nordeste',
        'DF': 'Centro-Oeste', 'GO': 'Centro-Oeste', 'MT': 'Centro-Oeste', 'MS': 'Centro-Oeste',
        'ES': 'Sudeste', 'MG': 'Sudeste', 'RJ': 'Sudeste', 'SP': 'Sudeste',
        'PR': 'Sul', 'RS': 'Sul', 'SC': 'Sul'
    }
    return mapa_uf_regiao.get(uf, 'Não identificado')

def obter_secao_cnae(codigo_cnae):
    """Retorna a Seção (grande grupo) e uma classificação simplificada (indústria) com base no código CNAE."""
    if not codigo_cnae or not isinstance(codigo_cnae, (str, int)):
        return 'Não identificado', 'Não identificado'

    divisao = int(str(codigo_cnae)[:2])
    
    mapa_secao = {
        'A': (range(1, 4), 'Agricultura, Pecuária, Produção Florestal, Pesca e Aquicultura', 'Agronegócio'),
        'B': (range(5, 10), 'Indústrias Extrativas', 'Indústria'),
        'C': (range(10, 34), 'Indústrias de Transformação', 'Indústria'),
        'D': (range(35, 36), 'Eletricidade e Gás', 'Infraestrutura'),
        'E': (range(36, 40), 'Água, Esgoto, Atividades de Gestão de Resíduos e Descontaminação', 'Infraestrutura'),
        'F': (range(41, 44), 'Construção', 'Construção'),
        'G': (range(45, 48), 'Comércio; Reparação de Veículos Automotores e Motocicletas', 'Comércio'),
        'H': (range(49, 54), 'Transporte, Armazenagem e Correio', 'Serviços'),
        'I': (range(55, 57), 'Alojamento e Alimentação', 'Serviços'),
        'J': (range(58, 64), 'Informação e Comunicação', 'Serviços'),
        'K': (range(64, 67), 'Atividades Financeiras, de Seguros e Serviços Relacionados', 'Serviços Financeiros'),
        'L': (range(68, 69), 'Atividades Imobiliárias', 'Serviços'),
        'M': (range(69, 76), 'Atividades Profissionais, Científicas e Técnicas', 'Serviços'),
        'N': (range(77, 83), 'Atividades Administrativas e Serviços Complementares', 'Serviços'),
        'O': (range(84, 85), 'Administração Pública, Defesa e Seguridade Social', 'Serviços Públicos'),
        'P': (range(85, 86), 'Educação', 'Serviços'),
        'Q': (range(86, 89), 'Saúde Humana e Serviços Sociais', 'Saúde'),
        'R': (range(90, 94), 'Artes, Cultura, Esporte e Recreação', 'Serviços'),
        'S': (range(94, 97), 'Outras Atividades de Serviços', 'Serviços'),
        'T': (range(97, 98), 'Serviços Domésticos', 'Serviços'),
        'U': (range(99, 100), 'Organismos Internacionais e Outras Instituições Extraterritoriais', 'Outros')
    }

    for _, (intervalo, nome_secao, industria) in mapa_secao.items():
        if divisao in intervalo:
            return nome_secao, industria
            
    return 'Não identificado', 'Não identificado'

def limpar_cnpj(cnpj):
    """Remove caracteres não numéricos de uma string de CNPJ."""
    if isinstance(cnpj, str):
        return re.sub(r'\D', '', cnpj)
    return ''

def consultar_cnpj_api(cnpj):
    """Consulta um CNPJ na BrasilAPI com lógica de tentativas."""
    cnpj_limpo = limpar_cnpj(cnpj)
    if not cnpj_limpo or len(cnpj_limpo) != 14:
        print(f"  -> CNPJ '{cnpj}' inválido. Pulando.")
        return None, 'CNPJ Inválido'
        
    url = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj_limpo}"
    max_tentativas = 3
    
    for tentativa in range(max_tentativas):
        try:
            response = requests.get(url, timeout=20, verify=False)
            if response.status_code == 200:
                return response.json(), 'Sucesso'
            elif response.status_code == 404:
                print(f"  -> CNPJ {cnpj_limpo} não encontrado na base de dados.")
                return None, 'Não Encontrado'
            else:
                print(f"  -> Tentativa {tentativa + 1}/{max_tentativas} falhou com status {response.status_code}.")
        except requests.exceptions.RequestException as e:
            print(f"  -> Tentativa {tentativa + 1}/{max_tentativas} falhou com erro de conexão: {e}")
        
        if tentativa < max_tentativas - 1:
            pausa = (tentativa + 1) * PAUSA_POR_TENTATIVA
            print(f"     Aguardando {pausa} segundos antes de tentar novamente...")
            time.sleep(pausa)
            
    return None, 'Falha na Consulta'

def enriquecer_planilha(caminho_arquivo):
    """Função principal para ler, processar e salvar a planilha enriquecida."""
    print("--- INICIANDO SCRIPT DE ENRIQUECIMENTO DE CNPJ ---")
    
    try:
        # Tenta ler como CSV, se falhar, tenta como Excel
        if caminho_arquivo.endswith('.csv'):
            df = pd.read_csv(caminho_arquivo)
        else:
            df = pd.read_excel(caminho_arquivo)
        print(f"\n[PASSO 1] Arquivo '{os.path.basename(caminho_arquivo)}' lido com sucesso.")
    except Exception as e:
        print(f"\n[ERRO FATAL] Não foi possível ler o arquivo. Verifique o caminho e o formato. Erro: {e}")
        return

    if 'cnpj' not in df.columns:
        print("\n[ERRO FATAL] A coluna 'cnpj' não foi encontrada na planilha.")
        return

    # Garante que as colunas desejadas existam no DataFrame
    colunas_desejadas = [
        'porte_da_empresa', 'matriz_filial', 'situacao_especial', 'data_situacao_especial',
        'regiao', 'grande_grupo', 'industria', 'regime_tributario', 'origem'
    ]
    for col in colunas_desejadas:
        if col not in df.columns:
            df[col] = None

    lista_linhas_atualizadas = []
    total_cnpjs = len(df)
    
    print(f"\n[PASSO 2] Iniciando consulta à API para {total_cnpjs} CNPJs...")
    
    for index, linha in df.iterrows():
        cnpj_original = str(linha.get('cnpj', ''))
        
        print(f"\nProcessando linha {index + 1}/{total_cnpjs}: CNPJ '{cnpj_original}'")
        
        dados_api, status = consultar_cnpj_api(cnpj_original)
        
        linha_atualizada = linha.to_dict()
        linha_atualizada['status_consulta'] = status
        
        if dados_api:
            print("  -> SUCESSO: Dados obtidos da API. Preenchendo colunas...")
            
            # Mapeamento e preenchimento dos campos
            linha_atualizada.update({
                'razao_social': dados_api.get('razao_social'),
                'nome_fantasia': dados_api.get('nome_fantasia'),
                'data_inicio_atividade': dados_api.get('data_inicio_atividade'),
                'situacao_cadastrar': dados_api.get('descricao_situacao_cadastral'),
                'data_situacao_cadastral': dados_api.get('data_situacao_cadastral'),
                'motivo_da_situacao_cadastral': dados_api.get('descricao_motivo_situacao_cadastral'),
                'natureza_juridica': dados_api.get('natureza_juridica'),
                'capital_social': dados_api.get('capital_social'),
                'logradouro': f"{dados_api.get('descricao_tipo_de_logradouro', '')} {dados_api.get('logradouro', '')}".strip(),
                'numero': dados_api.get('numero'),
                'complemento': dados_api.get('complemento'),
                'cep': dados_api.get('cep'),
                'bairro': dados_api.get('bairro'),
                'municipio': dados_api.get('municipio'),
                'unidade_federativa': dados_api.get('uf'),
                'email': dados_api.get('email'),
                'ddd_1': dados_api.get('ddd_telefone_1'),
                'telefone_1': dados_api.get('ddd_telefone_1'), # Ajuste conforme necessidade
                
                # Novos campos solicitados
                'porte_da_empresa': dados_api.get('porte'),
                'matriz_filial': dados_api.get('descricao_identificador_matriz_filial'),
                'situacao_especial': dados_api.get('descricao_situacao_especial'),
                'data_situacao_especial': dados_api.get('data_situacao_especial'),
                
                # Campos derivados
                'regiao': obter_regiao_por_uf(dados_api.get('uf')),
                'regime_tributario': 'MEI' if dados_api.get('opcao_pelo_mei') else ('Simples Nacional' if dados_api.get('opcao_pelo_simples') else 'Outros/Normal'),
                
                # Origem dos dados
                'origem': 'BrasilAPI'
            })
            
            # Campos derivados do CNAE
            cnae_principal = dados_api.get('cnae_fiscal')
            grande_grupo, industria = obter_secao_cnae(cnae_principal)
            linha_atualizada['grande_grupo'] = grande_grupo
            linha_atualizada['industria'] = industria
            
        else:
            print(f"  -> FALHA: {status} para o CNPJ '{cnpj_original}'.")

        lista_linhas_atualizadas.append(linha_atualizada)
        time.sleep(PAUSA_ENTRE_CNPJS)
        
    df_enriquecido = pd.DataFrame(lista_linhas_atualizadas)
    
    # Reordenar colunas para colocar 'status_consulta' no início
    if 'status_consulta' in df_enriquecido.columns:
        cols = ['cnpj', 'status_consulta'] + [c for c in df_enriquecido.columns if c not in ['cnpj', 'status_consulta']]
        df_enriquecido = df_enriquecido[cols]

    arquivo_saida = 'Planilha_CNPJs_Enriquecida.xlsx'
    try:
        df_enriquecido.to_excel(arquivo_saida, index=False, engine='openpyxl')
        caminho_saida = os.path.join(os.getcwd(), arquivo_saida)
        print(f"\n[PASSO 3] Processamento concluído.")
        print("\n--- SUCESSO FINAL! ---")
        print(f"O arquivo final foi criado em: '{caminho_saida}'")
    except Exception as e:
        print(f"\n[ERRO FATAL] Falha ao salvar o arquivo de saída. Erro: {e}")

if __name__ == "__main__":
    # IMPORTANTE: Altere o caminho abaixo para o local exato do seu arquivo.
    # Pode ser um arquivo .xlsx ou .csv
    caminho_do_arquivo = r"C:\Users\admin\Downloads\Empresas Dados - 72 empresas.xlsx"
    enriquecer_planilha(caminho_do_arquivo)