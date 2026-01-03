import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
from datetime import datetime
from matplotlib.ticker import FuncFormatter

# 1.Configurações iniciais
PASTA_ENTRADA = Path('entrada')
PASTA_SAIDA = Path('saida')
PASTA_LOGS = Path('logs')

def log_mensagem(mensagem):
    PASTA_LOGS.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(PASTA_LOGS / 'execucao.log', 'a', encoding='utf-8') as f:
        f.write(f"[{timestamp}] {mensagem}\n")
    print(mensagem)

# Formatar valores
def formatar_moeda(x, pos):
    """Converte valores brutos em formato legível ($1.2M, $500K)"""
    if x >= 1_000_000:
        return f'${x*1e-6:1.1f}M'
    elif x >= 1_000:
        return f'${x*1e-3:1.0f}K'
    return f'${x:1.0f}'

def processar_adventure_works():
    try:
        PASTA_SAIDA.mkdir(exist_ok=True)
        log_mensagem("🚀 Iniciando processamento AdventureWorks...")

        caminho_arquivo = PASTA_ENTRADA / 'AdventureWorks Sales.xlsx'

        # 2. Ler dados
        log_mensagem("📖 Lendo abas do Excel...")
        vendas = pd.read_excel(caminho_arquivo, sheet_name='Sales_data')
        produtos = pd.read_excel(caminho_arquivo, sheet_name='Product_data')
        territorios = pd.read_excel(caminho_arquivo, sheet_name='Sales Territory_data')

        # 3. Tratamento e cruzamento das informações
        df = pd.merge(vendas, produtos, on='ProductKey')
        df = pd.merge(df, territorios, on='SalesTerritoryKey')
        df['Lucro'] = df['Sales Amount'] - df['Total Product Cost']

        # 4. Visual gráfico
        sns.set_theme(style="whitegrid")
        plt.rcParams['font.family'] = 'sans-serif'
        formatter = FuncFormatter(formatar_moeda)

        # Vendas por categoria
        plt.figure(figsize=(12, 7))
        faturamento = df.groupby('Category')['Sales Amount'].sum().sort_values(ascending=True)
        
        ax1 = faturamento.plot(kind='barh', color='#2c3e50', width=0.7)
        ax1.xaxis.set_major_formatter(formatter)
        
        plt.title('Faturamento por Categoria (USD)', fontsize=16, fontweight='bold', color='#34495e', pad=25)
        plt.xlabel('Volume Total de Vendas', fontsize=12)
        plt.ylabel('Categoria de Produto', fontsize=12)

        # Rótulos de dados (Data Labels)
        for i, v in enumerate(faturamento):
            ax1.text(v + (v * 0.01), i, f'{formatar_moeda(v, None)}', va='center', fontsize=10, fontweight='bold')

        plt.tight_layout()
        plt.savefig(PASTA_SAIDA / 'dashboard_faturamento.png', dpi=300)
        plt.close()

        # Lucro por país
        plt.figure(figsize=(12, 7))
        lucro_pais = df.groupby('Country')['Lucro'].sum().sort_values(ascending=False)
        
        ax2 = sns.barplot(x=lucro_pais.index, y=lucro_pais.values, palette='Blues_d')
        ax2.yaxis.set_major_formatter(formatter)
        
        plt.title('Lucro Líquido por País (USD)', fontsize=16, fontweight='bold', color='#34495e', pad=25)
        plt.ylabel('Lucro Acumulado', fontsize=12)
        plt.xlabel('Mercado Geográfico', fontsize=12)

        # Rótulos de dados no topo das barras
        for i, v in enumerate(lucro_pais):
            ax2.text(i, v + (v * 0.02), f'{formatar_moeda(v, None)}', ha='center', fontsize=10, fontweight='bold')

        plt.tight_layout()
        plt.savefig(PASTA_SAIDA / 'dashboard_lucro.png', dpi=300)
        plt.close()

        # 5. Exportar para saída
        df.to_excel(PASTA_SAIDA / 'Relatorio_AdventureWorks_Final.xlsx', index=False)
        log_mensagem(f" Processo concluído com sucesso!")

    except Exception as e:
        log_mensagem(f"❌ ERRO CRÍTICO: {str(e)}")

if __name__ == "__main__":
    processar_adventure_works()