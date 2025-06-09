import os
import xml.etree.ElementTree as ET
import pandas as pd
import re
from pathlib import Path
import webbrowser

# Função para extrair informações de um arquivo XML de NFe
def extrair_dados_nfe(arquivo_xml):
    try:
        # Parsing do XML
        tree = ET.parse(arquivo_xml)
        root = tree.getroot()
        
        # Definir os namespaces
        ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
        
        # Extrair a chave de acesso
        inf_nfe = root.find('.//nfe:infNFe', ns)
        chave_acesso = inf_nfe.attrib['Id'].replace('NFe', '') if inf_nfe is not None and 'Id' in inf_nfe.attrib else 'N/A'
        
        # Extrair o número da NF
        numero_nf = root.find('.//nfe:nNF', ns).text if root.find('.//nfe:nNF', ns) is not None else 'N/A'
        
        # Extrair produtos, quantidades e valores
        produtos = []
        for det in root.findall('.//nfe:det', ns):
            prod = det.find('.//nfe:prod', ns)
            if prod is not None:
                nome_prod = prod.find('.//nfe:xProd', ns).text if prod.find('.//nfe:xProd', ns) is not None else 'N/A'
                qtd = prod.find('.//nfe:qCom', ns).text if prod.find('.//nfe:qCom', ns) is not None else '0'
                valor_unit = prod.find('.//nfe:vUnCom', ns).text if prod.find('.//nfe:vUnCom', ns) is not None else '0'
                valor_total = prod.find('.//nfe:vProd', ns).text if prod.find('.//nfe:vProd', ns) is not None else '0'
                
                produtos.append({
                    'nome': nome_prod,
                    'quantidade': float(qtd),
                    'valor_unitario': float(valor_unit),
                    'valor_total': float(valor_total)
                })
        
        # Extrair valor total da nota
        valor_liquido = root.find('.//nfe:vNF', ns).text if root.find('.//nfe:vNF', ns) is not None else '0'
        valor_liquido = float(valor_liquido)
        
        # Extrair informações adicionais
        info_adicionais = root.find('.//nfe:infAdFisco', ns)
        if info_adicionais is None:
            info_adicionais = root.find('.//nfe:infCpl', ns)
        
        info_texto = info_adicionais.text if info_adicionais is not None else ''
        
        # Extrair CENSO da escola
        censo_match = re.search(r'Censo\s+(\d+)', info_texto) if info_texto else None
        censo = censo_match.group(1) if censo_match else 'N/A'
        
        # Extrair nome da escola
        escola_match = re.search(r'Censo\s+\d+\s+-\s+([^-]+)', info_texto) if info_texto else None
        nome_escola = escola_match.group(1).strip() if escola_match else 'N/A'
        
        # Verificar se o somatório dos valores singulares está de acordo com o valor líquido total
        soma_produtos = sum(p['valor_total'] for p in produtos)
        valores_consistentes = abs(soma_produtos - valor_liquido) < 0.01  # Tolerância para diferenças de arredondamento
        
        return {
            'chave_acesso': chave_acesso,
            'numero_nf': numero_nf,
            'produtos': produtos,
            'valor_liquido': valor_liquido,
            'censo': censo,
            'nome_escola': nome_escola,
            'valores_consistentes': valores_consistentes
        }
    
    except Exception as e:
        print(f"Erro ao processar o arquivo {arquivo_xml}: {str(e)}")
        return None

# Função para gerar link para a nota na Receita Federal
def gerar_link_nfe(chave_acesso):
    return f"https://www.nfe.fazenda.gov.br/portal/consultaRecaptcha.aspx?tipoConsulta=completa&tipoConteudo=XbSeqxE8pl8=&nfe={chave_acesso}"

# Função para encontrar todos os arquivos XML de NFe nos diretórios
def encontrar_arquivos_nfe(diretorio_base):
    arquivos_nfe = []
    for root, dirs, files in os.walk(diretorio_base):
        for file in files:
            if file.endswith('-nfe.xml'):
                arquivos_nfe.append(os.path.join(root, file))
    return arquivos_nfe

# Função para gerar a tabela HTML com os dados extraídos
def gerar_tabela_html(dados_nfes):
    # Criar DataFrame para a tabela principal
    tabela_principal = []
    for nfe in dados_nfes:
        if nfe is not None:
            link_nfe = gerar_link_nfe(nfe['chave_acesso'])
            tabela_principal.append({
                'Chave de Acesso': f'<a href="{link_nfe}" target="_blank">{nfe["chave_acesso"]}</a>',
                'Número da NF': nfe['numero_nf'],
                'Valor Líquido': f'R$ {nfe["valor_liquido"]:.2f}',
                'CENSO': nfe['censo'],
                'Nome da Escola': nfe['nome_escola'],
                'Valores Consistentes': 'Sim' if nfe['valores_consistentes'] else 'Não',
                'Detalhes': f'<button onclick="toggleProdutos(\'produtos_{nfe["numero_nf"]}\')">Ver Produtos</button>'
            })
    
    df_principal = pd.DataFrame(tabela_principal)
    
    # Criar HTML para a tabela de produtos de cada NFe
    tabelas_produtos = []
    for nfe in dados_nfes:
        if nfe is not None:
            produtos_html = f'<div id="produtos_{nfe["numero_nf"]}" style="display:none;">\''
            produtos_html += f'<h3>Produtos da NF {nfe["numero_nf"]}</h3>'
            produtos_html += '<table class="table table-striped">'
            produtos_html += '<thead><tr><th>Produto</th><th>Quantidade</th><th>Valor Unitário</th><th>Valor Total</th></tr></thead>'
            produtos_html += '<tbody>'
            
            for produto in nfe['produtos']:
                produtos_html += f'<tr>'
                produtos_html += f'<td>{produto["nome"]}</td>'
                produtos_html += f'<td>{produto["quantidade"]:.2f}</td>'
                produtos_html += f'<td>R$ {produto["valor_unitario"]:.2f}</td>'
                produtos_html += f'<td>R$ {produto["valor_total"]:.2f}</td>'
                produtos_html += f'</tr>'
            
            produtos_html += '</tbody></table></div>'
            tabelas_produtos.append(produtos_html)
    
    # Gerar HTML completo
    html = '''
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Leitor de Notas Fiscais</title>
        <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
        <style>
            body { padding: 20px; }
            .table { margin-top: 20px; }
            h1 { margin-bottom: 20px; }
        </style>
        <script>
            function toggleProdutos(id) {
                var element = document.getElementById(id);
                if (element.style.display === "none") {
                    element.style.display = "block";
                } else {
                    element.style.display = "none";
                }
            }
        </script>
    </head>
    <body>
        <div class="container">
            <h1>Leitor de Notas Fiscais</h1>
            <div class="table-responsive">
    '''
    
    # Adicionar tabela principal
    html += df_principal.to_html(classes='table table-striped', escape=False, index=False)
    
    # Adicionar tabelas de produtos
    for tabela in tabelas_produtos:
        html += tabela
    
    html += '''
            </div>
        </div>
    </body>
    </html>
    '''
    
    return html

# Função principal
def main():
    diretorio_base = os.path.dirname(os.path.abspath(__file__))
    arquivos_nfe = encontrar_arquivos_nfe(diretorio_base)
    
    print(f"Encontrados {len(arquivos_nfe)} arquivos de NFe.")
    
    # Extrair dados de todas as NFes
    dados_nfes = []
    for arquivo in arquivos_nfe:
        print(f"Processando {arquivo}...")
        dados = extrair_dados_nfe(arquivo)
        if dados is not None:
            dados_nfes.append(dados)
    
    # Gerar HTML com os dados
    html = gerar_tabela_html(dados_nfes)
    
    # Salvar HTML em um arquivo
    arquivo_html = os.path.join(diretorio_base, 'relatorio_nfes.html')
    with open(arquivo_html, 'w', encoding='utf-8') as f:
        f.write(html)
    
    print(f"Relatório gerado com sucesso em {arquivo_html}")
    
    # Abrir o arquivo HTML no navegador padrão
    webbrowser.open('file://' + os.path.abspath(arquivo_html))

if __name__ == "__main__":
    main()