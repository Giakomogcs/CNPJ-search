import pandas as pd
import requests
import re
import time
import os
import urllib3

# --- CONFIGURAÇÕES ---
# Desabilitar avisos de requisição insegura (útil para alguns ambientes)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- PARÂMETROS DE TEMPO AJUSTÁVEIS ---
PAUSA_ENTRE_CNPJS = 1    # Pausa em segundos entre a consulta de cada CNPJ.
PAUSA_POR_TENTATIVA = 3  # Pausa base em segundos para as tentativas em caso de falha.
MAX_TENTATIVAS = 3       # Número máximo de tentativas de consulta à API.

# --- CAMINHOS DE ARQUIVOS ---
ARQUIVO_ENTRADA = 'Empresas Dados - 72 empresas.xlsx'
ARQUIVO_SAIDA = 'Planilha_CNPJs_Enriquecida.xlsx'
ARQUIVO_LOOKUP_CNAE = 'cnae_lookup.csv'

# --- CLASSES E FUNÇÕES AUXILIARES ---

def limpar_codigo(codigo):
    """Remove caracteres não numéricos de uma string."""
    if isinstance(codigo, str):
        return re.sub(r'\D', '', codigo)
    if isinstance(codigo, (int, float)):
        return str(int(codigo))
    return ''

class CNAEProcessor:
    """
    Classe para carregar e processar informações de CNAE a partir de um arquivo de lookup.
    """
    def __init__(self, caminho_lookup):
        try:
            self.df_lookup = pd.read_csv(caminho_lookup, dtype=str, keep_default_na=False)
            self.df_lookup['subclasse_codigo'] = self.df_lookup['subclasse_codigo'].apply(limpar_codigo)
            self.df_lookup.set_index('subclasse_codigo', inplace=True)
            print("Tabela de lookup CNAE carregada com sucesso.")
        except FileNotFoundError:
            print(f"[ERRO FATAL] Arquivo de lookup CNAE não encontrado em '{caminho_lookup}'.")
            self.df_lookup = None
        except Exception as e:
            print(f"[ERRO FATAL] Falha ao carregar o arquivo de lookup CNAE: {e}")
            self.df_lookup = None

    def get_cnae_details(self, cnae_code):
        """
        Busca os detalhes de um código CNAE na tabela de lookup.
        Retorna um dicionário com toda a hierarquia CNAE.
        """
        if self.df_lookup is None or self.df_lookup.empty:
            return {}

        cnae_limpo = limpar_codigo(str(cnae_code))

        if cnae_limpo in self.df_lookup.index:
            details = self.df_lookup.loc[cnae_limpo].to_dict()
            return {
                'cnae_secao_codigo': details.get('secao'),
                'cnae_secao_denominacao': details.get('secao_denominacao'),
                'cnae_divisao_codigo': details.get('divisao'),
                'cnae_divisao_denominacao': details.get('divisao_denominacao'),
                'cnae_grupo_codigo': details.get('grupo'),
                'cnae_grupo_denominacao': details.get('grupo_denominacao'),
                'cnae_classe_codigo': details.get('classe'),
                'cnae_classe_denominacao': details.get('classe_denominacao'),
                'cnae_subclasse_codigo': details.get('subclasse'),
                'cnae_subclasse_denominacao': details.get('subclasse_denominacao'),
            }
        return {}

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

def consultar_cnpj_api(cnpj):
    """Consulta um CNPJ na BrasilAPI com lógica de tentativas."""
    cnpj_limpo = limpar_codigo(cnpj)
    if not cnpj_limpo or len(cnpj_limpo) != 14:
        print(f"  -> CNPJ '{cnpj}' inválido. Pulando.")
        return None, 'CNPJ Inválido'
        
    url = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj_limpo}"
    
    for tentativa in range(MAX_TENTATIVAS):
        try:
            response = requests.get(url, timeout=20, verify=False)
            if response.status_code == 200:
                return response.json(), 'Sucesso'
            elif response.status_code == 404:
                print(f"  -> CNPJ {cnpj_limpo} não encontrado na base de dados.")
                return None, 'Não Encontrado'
            else:
                print(f"  -> Tentativa {tentativa + 1}/{MAX_TENTATIVAS} falhou com status {response.status_code}.")
        except requests.exceptions.RequestException as e:
            print(f"  -> Tentativa {tentativa + 1}/{MAX_TENTATIVAS} falhou com erro de conexão: {e}")
        
        if tentativa < MAX_TENTATIVAS - 1:
            pausa = (tentativa + 1) * PAUSA_POR_TENTATIVA
            print(f"     Aguardando {pausa} segundos antes de tentar novamente...")
            time.sleep(pausa)
            
    return None, 'Falha na Consulta'

def processar_linha(linha, cnae_processor):
    """Processa uma única linha do DataFrame, consultando a API e enriquecendo os dados."""
    cnpj_original = str(linha.get('cnpj', ''))
    print(f"\nProcessando CNPJ '{cnpj_original}'...")

    dados_api, status = consultar_cnpj_api(cnpj_original)

    linha_atualizada = linha.to_dict()
    linha_atualizada['status_consulta'] = status

    if dados_api:
        print("  -> SUCESSO: Dados obtidos da API. Preenchendo colunas...")

        # Mapeamento e preenchimento dos campos da API
        linha_atualizada.update({
            'razao_social': dados_api.get('razao_social'),
            'nome_fantasia': dados_api.get('nome_fantasia'),
            'data_inicio_atividade': dados_api.get('data_inicio_atividade'),
            'situacao_cadastral': dados_api.get('descricao_situacao_cadastral'),
            'data_situacao_cadastral': dados_api.get('data_situacao_cadastral'),
            'motivo_situacao_cadastral': dados_api.get('descricao_motivo_situacao_cadastral'),
            'natureza_juridica': dados_api.get('natureza_juridica'),
            'capital_social': dados_api.get('capital_social'),
            'logradouro': f"{dados_api.get('descricao_tipo_de_logradouro', '')} {dados_api.get('logradouro', '')}".strip(),
            'numero': dados_api.get('numero'),
            'complemento': dados_api.get('complemento'),
            'cep': dados_api.get('cep'),
            'bairro': dados_api.get('bairro'),
            'municipio': dados_api.get('municipio'),
            'uf': dados_api.get('uf'),
            'email': dados_api.get('email'),
            'telefone_1': dados_api.get('ddd_telefone_1'),
            'porte': dados_api.get('porte'),
            'matriz_filial': dados_api.get('descricao_identificador_matriz_filial'),
            'situacao_especial': dados_api.get('descricao_situacao_especial'),
            'data_situacao_especial': dados_api.get('data_situacao_especial'),
            'origem_dados': 'BrasilAPI'
        })

        # Campos derivados
        linha_atualizada['regiao'] = obter_regiao_por_uf(dados_api.get('uf'))
        linha_atualizada['regime_tributario'] = 'MEI' if dados_api.get('opcao_pelo_mei') else ('Simples Nacional' if dados_api.get('opcao_pelo_simples') else 'Outros/Normal')

        # Processamento do CNAE Principal
        cnae_principal = dados_api.get('cnae_fiscal')
        if cnae_principal:
            linha_atualizada['cnae_fiscal'] = cnae_principal
            cnae_details = cnae_processor.get_cnae_details(cnae_principal)
            linha_atualizada.update(cnae_details)

        # Processamento dos CNAEs Secundários
        cnaes_secundarios = dados_api.get('cnaes_secundarios', [])
        if cnaes_secundarios:
            lista_codigos = [item.get('codigo', '') for item in cnaes_secundarios]
            linha_atualizada['cnaes_secundarios'] = ", ".join([str(c) for c in lista_codigos if c])
        else:
            linha_atualizada['cnaes_secundarios'] = ""

    else:
        print(f"  -> FALHA: {status} para o CNPJ '{cnpj_original}'.")

    return linha_atualizada

def main():
    """Função principal para orquestrar o enriquecimento da planilha."""
    print("--- INICIANDO SCRIPT DE ENRIQUECIMENTO DE CNPJ ---")
    
    # Carrega a tabela de lookup CNAE
    cnae_processor = CNAEProcessor(ARQUIVO_LOOKUP_CNAE)
    if cnae_processor.df_lookup is None:
        return # Aborta se o lookup não puder ser carregado

    # Lê a planilha de entrada
    try:
        if ARQUIVO_ENTRADA.endswith('.csv'):
            df = pd.read_csv(ARQUIVO_ENTRADA)
        else:
            df = pd.read_excel(ARQUIVO_ENTRADA)
        print(f"\n[PASSO 1] Arquivo '{ARQUIVO_ENTRADA}' lido com sucesso.")
    except FileNotFoundError:
        print(f"\n[ERRO FATAL] Arquivo de entrada '{ARQUIVO_ENTRADA}' não encontrado.")
        return
    except Exception as e:
        print(f"\n[ERRO FATAL] Não foi possível ler o arquivo de entrada. Erro: {e}")
        return

    if 'cnpj' not in df.columns:
        print("\n[ERRO FATAL] A coluna 'cnpj' não foi encontrada na planilha.")
        return

    # Processa cada linha
    print(f"\n[PASSO 2] Iniciando consulta à API para {len(df)} CNPJs...")
    lista_linhas_atualizadas = [processar_linha(linha, cnae_processor) for _, linha in df.iterrows()]
    
    # Cria o DataFrame final
    df_enriquecido = pd.DataFrame(lista_linhas_atualizadas)
    
    # Reordena colunas para melhor visualização
    colunas_iniciais = ['cnpj', 'status_consulta', 'razao_social', 'nome_fantasia']
    cols_ordenadas = colunas_iniciais + [c for c in df_enriquecido.columns if c not in colunas_iniciais]
    df_enriquecido = df_enriquecido[cols_ordenadas]

    # Salva o arquivo de saída
    try:
        df_enriquecido.to_excel(ARQUIVO_SAIDA, index=False, engine='openpyxl')
        caminho_saida = os.path.join(os.getcwd(), ARQUIVO_SAIDA)
        print(f"\n[PASSO 3] Processamento concluído.")
        print("\n--- SUCESSO FINAL! ---")
        print(f"O arquivo final foi criado em: '{caminho_saida}'")
    except Exception as e:
        print(f"\n[ERRO FATAL] Falha ao salvar o arquivo de saída. Erro: {e}")

if __name__ == "__main__":
    main()