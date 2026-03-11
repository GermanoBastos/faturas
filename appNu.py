#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script completo para diagnóstico e correção de arquivos CSV
que exibem menos linhas que o esperado (ex: 2 linhas mostrando apenas 1)
Autor: Baseado na análise dos arquivos do histograma regulatório
"""

import os
import sys
import pandas as pd
import chardet
import csv
from pathlib import Path

def diagnosticar_csv(caminho_arquivo):
    """
    Função completa para diagnosticar problemas em arquivos CSV
    
    Args:
        caminho_arquivo (str): Caminho para o arquivo CSV
    """
    print("=" * 60)
    print(f"DIAGNÓSTICO COMPLETO DO ARQUIVO: {caminho_arquivo}")
    print("=" * 60)
    
    # 1. Verificar se o arquivo existe
    if not os.path.exists(caminho_arquivo):
        print(f"❌ ERRO: Arquivo não encontrado: {caminho_arquivo}")
        return
    
    # 2. Informações básicas do arquivo
    tamanho = os.path.getsize(caminho_arquivo)
    print(f"\n📁 INFORMAÇÕES BÁSICAS:")
    print(f"   - Tamanho: {tamanho} bytes ({tamanho/1024:.2f} KB)")
    print(f"   - Última modificação: {os.path.getmtime(caminho_arquivo)}")
    
    # 3. Detectar encoding automaticamente
    with open(caminho_arquivo, 'rb') as f:
        raw_data = f.read()
        resultado_encoding = chardet.detect(raw_data)
        encoding_detectado = resultado_encoding['encoding']
        confianca = resultado_encoding['confidence']
        print(f"\n🔤 ENCODING DETECTADO:")
        print(f"   - Encoding: {encoding_detectado}")
        print(f"   - Confiança: {confianca:.2%}")
    
    # 4. Ler arquivo como texto bruto
    print(f"\n📄 ANÁLISE DE LINHAS BRUTAS:")
    try:
        with open(caminho_arquivo, 'r', encoding=encoding_detectado) as f:
            linhas_brutas = f.readlines()
        
        print(f"   - Total de linhas no arquivo: {len(linhas_brutas)}")
        
        # Mostrar primeiras linhas com representação literal
        print(f"\n   Primeiras 5 linhas (representação literal):")
        for i, linha in enumerate(linhas_brutas[:5]):
            # repr() mostra caracteres especiais como \n, \r, etc.
            print(f"   Linha {i}: {repr(linha)}")
        
        # Verificar linhas vazias
        linhas_nao_vazias = [l for l in linhas_brutas if l.strip()]
        print(f"   - Linhas não vazias: {len(linhas_nao_vazias)}")
        
    except Exception as e:
        print(f"   ❌ Erro ao ler arquivo: {e}")
    
    # 5. Testar diferentes delimitadores e encodings
    print(f"\n🔍 TESTE DE LEITURA COM DIFERENTES CONFIGURAÇÕES:")
    
    delimitadores = [',', ';', '\t', '|', ' ']
    encodings = ['utf-8', 'utf-8-sig', 'latin1', 'cp1252', 'iso-8859-1']
    
    resultados = []
    
    for delim in delimitadores:
        for enc in encodings:
            try:
                df = pd.read_csv(caminho_arquivo, sep=delim, encoding=enc, 
                                  nrows=5, on_bad_lines='skip')
                resultados.append({
                    'delimitador': repr(delim) if delim != '\t' else 'TAB',
                    'encoding': enc,
                    'linhas': len(df),
                    'colunas': len(df.columns) if not df.empty else 0,
                    'status': 'OK'
                })
            except Exception as e:
                resultados.append({
                    'delimitador': repr(delim) if delim != '\t' else 'TAB',
                    'encoding': enc,
                    'linhas': 0,
                    'colunas': 0,
                    'status': f'Erro: {str(e)[:50]}'
                })
    
    # Mostrar resultados
    print(f"\n   {'Delimitador':<12} {'Encoding':<12} {'Linhas':<8} {'Colunas':<8} Status")
    print(f"   {'-'*12} {'-'*12} {'-'*8} {'-'*8} {'-'*30}")
    
    for r in resultados[:10]:  # Mostrar apenas os primeiros para não poluir
        print(f"   {r['delimitador']:<12} {r['encoding']:<12} {r['linhas']:<8} {r['colunas']:<8} {r['status']}")
    
    # 6. Tentar encontrar a melhor configuração
    print(f"\n✅ MELHOR CONFIGURAÇÃO ENCONTRADA:")
    melhores = [r for r in resultados if r['linhas'] > 1 and r['status'] == 'OK']
    
    if melhores:
        # Ordenar por número de linhas (decrescente)
        melhores.sort(key=lambda x: x['linhas'], reverse=True)
        melhor = melhores[0]
        
        print(f"   - Delimitador: {melhor['delimitador']}")
        print(f"   - Encoding: {melhor['encoding']}")
        print(f"   - Linhas lidas: {melhor['linhas']}")
        print(f"   - Colunas: {melhor['colunas']}")
    else:
        print(f"   ❌ Nenhuma configuração conseguiu ler mais de 1 linha")
    
    # 7. Verificar possíveis problemas específicos
    print(f"\n⚠️ VERIFICAÇÃO DE PROBLEMAS COMUNS:")
    
    problemas = []
    
    # Problema 1: Quebras de linha dentro de campos
    with open(caminho_arquivo, 'rb') as f:
        conteudo = f.read().decode(encoding_detectado, errors='ignore')
        if '\r\n\r\n' in conteudo or '\n\n' in conteudo:
            problemas.append("Possíveis linhas em branco entre registros")
        if conteudo.count('"') % 2 != 0:
            problemas.append("Número ímpar de aspas - possível campo não fechado")
    
    # Problema 2: Delimitador inconsistente
    with open(caminho_arquivo, 'r', encoding=encoding_detectado, errors='ignore') as f:
        primeira_linha = f.readline()
        if primeira_linha:
            for delim in [',', ';', '\t']:
                if primeira_linha.count(delim) > 0:
                    pass
            # Verificar se todas as linhas têm o mesmo número de delimitadores
            f.seek(0)
            linhas = f.readlines()[:10]
            contagens = [linha.count(',') for linha in linhas if linha.strip()]
            if len(set(contagens)) > 1:
                problemas.append(f"Número inconsistente de colunas entre linhas: {contagens}")
    
    if problemas:
        for prob in problemas:
            print(f"   ⚠️ {prob}")
    else:
        print(f"   ✅ Nenhum problema comum detectado")
    
    return melhores[0] if melhores else None


def corrigir_csv(caminho_arquivo, caminho_saida=None, delimitador_destino=','):
    """
    Corrige o arquivo CSV e salva uma versão limpa
    
    Args:
        caminho_arquivo (str): Caminho para o arquivo original
        caminho_saida (str): Caminho para o arquivo corrigido (opcional)
        delimitador_destino (str): Delimitador a ser usado no arquivo de saída
    
    Returns:
        bool: True se a correção foi bem-sucedida
    """
    print("\n" + "=" * 60)
    print("CORREÇÃO DO ARQUIVO CSV")
    print("=" * 60)
    
    if not caminho_saida:
        nome_original = Path(caminho_arquivo).stem
        extensao = Path(caminho_arquivo).suffix
        caminho_saida = f"{nome_original}_CORRIGIDO{extensao}"
    
    # Primeiro, diagnosticar para encontrar a melhor configuração
    melhor_config = diagnosticar_csv(caminho_arquivo)
    
    if not melhor_config:
        print("\n❌ Não foi possível determinar a configuração correta.")
        return False
    
    print(f"\n🔄 CORRIGINDO ARQUIVO...")
    print(f"   - Usando delimitador: {melhor_config['delimitador']}")
    print(f"   - Usando encoding: {melhor_config['encoding']}")
    print(f"   - Arquivo de saída: {caminho_saida}")
    
    try:
        # Ler o arquivo com a configuração detectada
        delim = melhor_config['delimitador']
        if delim == 'TAB':
            delim = '\t'
        else:
            delim = eval(delim) if delim.startswith("'") else delim
        
        df = pd.read_csv(
            caminho_arquivo, 
            sep=delim, 
            encoding=melhor_config['encoding'],
            engine='python',  # Mais flexível com arquivos problemáticos
            on_bad_lines='warn',  # Avisar sobre linhas problemáticas
            skipinitialspace=True,  # Ignorar espaços após delimitador
            quoting=csv.QUOTE_MINIMAL,
            keep_default_na=False,  # Não converter strings vazias para NaN
            na_filter=False  # Não filtrar valores NA
        )
        
        print(f"   - Linhas lidas: {len(df)}")
        print(f"   - Colunas: {len(df.columns)}")
        
        # Remover linhas totalmente vazias
        df_limpo = df.dropna(how='all')
        if len(df_limpo) < len(df):
            print(f"   - Linhas vazias removidas: {len(df) - len(df_limpo)}")
        
        # Salvar arquivo corrigido
        df_limpo.to_csv(
            caminho_saida, 
            sep=delimitador_destino, 
            encoding='utf-8-sig',  # UTF-8 com BOM para melhor compatibilidade com Excel
            index=False,
            quoting=csv.QUOTE_MINIMAL
        )
        
        print(f"\n✅ ARQUIVO CORRIGIDO COM SUCESSO!")
        print(f"   - Localização: {caminho_saida}")
        print(f"   - Tamanho: {os.path.getsize(caminho_saida)} bytes")
        
        # Verificar o arquivo corrigido
        df_verificacao = pd.read_csv(caminho_saida, encoding='utf-8-sig')
        print(f"   - Verificação: {len(df_verificacao)} linhas, {len(df_verificacao.columns)} colunas")
        
        return True
        
    except Exception as e:
        print(f"\n❌ Erro durante a correção: {e}")
        
        # Tentar abordagem alternativa: ler como texto e processar manualmente
        print("\n🔄 Tentando abordagem alternativa...")
        
        try:
            with open(caminho_arquivo, 'r', encoding=melhor_config['encoding'], errors='ignore') as f:
                linhas = f.readlines()
            
            # Limpar linhas
            linhas_limpas = []
            for linha in linhas:
                linha_limpa = linha.strip()
                if linha_limpa:  # Ignorar linhas vazias
                    # Remover caracteres problemáticos
                    linha_limpa = linha_limpa.replace('\r', '').replace('\n', '')
                    linhas_limpas.append(linha_limpa)
            
            print(f"   - Linhas após limpeza manual: {len(linhas_limpas)}")
            
            # Salvar versão limpa
            with open(caminho_saida, 'w', encoding='utf-8-sig') as f:
                for linha in linhas_limpas:
                    f.write(linha + '\n')
            
            print(f"✅ ARQUIVO CORRIGIDO (MODO ALTERNATIVO)!")
            return True
            
        except Exception as e2:
            print(f"❌ Erro na abordagem alternativa: {e2}")
            return False


def visualizar_arquivo(caminho_arquivo, num_linhas=10):
    """
    Visualiza o conteúdo do arquivo de forma organizada
    
    Args:
        caminho_arquivo (str): Caminho para o arquivo
        num_linhas (int): Número de linhas a visualizar
    """
    print("\n" + "=" * 60)
    print(f"VISUALIZAÇÃO DO ARQUIVO: {caminho_arquivo}")
    print("=" * 60)
    
    try:
        df = pd.read_csv(caminho_arquivo, encoding='utf-8-sig')
        print(f"\n📊 TOTAL: {len(df)} linhas x {len(df.columns)} colunas\n")
        
        if len(df) > 0:
            print("Primeiras 5 linhas:")
            print(df.head(5).to_string())
            
            if len(df) > 5:
                print(f"\n... e mais {len(df)-5} linhas")
        else:
            print("⚠️ Arquivo vazio ou sem dados")
            
    except Exception as e:
        print(f"❌ Erro ao visualizar: {e}")


def menu_principal():
    """
    Menu interativo para escolher as opções
    """
    print("\n" + "=" * 60)
    print("FERRAMENTA DE DIAGNÓSTICO E CORREÇÃO DE CSV")
    print("=" * 60)
    
    while True:
        print("\nOPÇÕES:")
        print("1. Diagnosticar arquivo CSV")
        print("2. Corrigir arquivo CSV")
        print("3. Diagnosticar e corrigir")
        print("4. Visualizar arquivo corrigido")
        print("5. Sair")
        
        opcao = input("\nEscolha uma opção (1-5): ").strip()
        
        if opcao == '5':
            print("Encerrando...")
            break
        
        if opcao in ['1', '2', '3', '4']:
            caminho = input("Caminho do arquivo CSV: ").strip()
            
            if not os.path.exists(caminho):
                print(f"❌ Arquivo não encontrado: {caminho}")
                continue
            
            if opcao == '1':
                diagnosticar_csv(caminho)
            
            elif opcao == '2':
                saida = input("Caminho de saída (ENTER para automático): ").strip()
                corrigir_csv(caminho, saida if saida else None)
            
            elif opcao == '3':
                saida = input("Caminho de saída (ENTER para automático): ").strip()
                if corrigir_csv(caminho, saida if saida else None):
                    visualizar_arquivo(saida if saida else f"{Path(caminho).stem}_CORRIGIDO{Path(caminho).suffix}")
            
            elif opcao == '4':
                visualizar_arquivo(caminho)
        
        else:
            print("❌ Opção inválida. Escolha 1-5.")


if __name__ == "__main__":
    # Se passar o caminho como argumento, executa automaticamente
    if len(sys.argv) > 1:
        caminho = sys.argv[1]
        print(f"Processando arquivo: {caminho}")
        melhor = diagnosticar_csv(caminho)
        
        if melhor:
            resp = input("\nDeseja corrigir o arquivo? (s/n): ").strip().lower()
            if resp == 's':
                saida = sys.argv[2] if len(sys.argv) > 2 else None
                corrigir_csv(caminho, saida)
    else:
        # Menu interativo
        menu_principal()
