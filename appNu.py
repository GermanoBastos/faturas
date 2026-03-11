#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script completo para converter CSV para Excel (XLSX)
com diagnóstico automático do formato correto
Autor: Baseado na análise dos arquivos do histograma regulatório
"""

import os
import sys
import pandas as pd
import chardet
import csv
from pathlib import Path
from datetime import datetime

def diagnosticar_csv(caminho_arquivo):
    """
    Diagnostica o formato correto do CSV
    
    Args:
        caminho_arquivo (str): Caminho para o arquivo CSV
    
    Returns:
        dict: Configuração detectada (delimitador, encoding) ou None
    """
    print(f"\n🔍 Diagnosticando arquivo: {caminho_arquivo}")
    
    if not os.path.exists(caminho_arquivo):
        print(f"❌ ERRO: Arquivo não encontrado: {caminho_arquivo}")
        return None
    
    # Detectar encoding
    with open(caminho_arquivo, 'rb') as f:
        raw_data = f.read()
        resultado = chardet.detect(raw_data)
        encoding_detectado = resultado['encoding']
        print(f"   Encoding detectado: {encoding_detectado} (confiança: {resultado['confidence']:.2%})")
    
    # Testar diferentes delimitadores
    delimitadores = [',', ';', '\t', '|', ' ']
    encodings = [encoding_detectado, 'utf-8', 'utf-8-sig', 'latin1', 'cp1252']
    
    melhor_config = None
    max_linhas = 0
    
    print(f"   Testando configurações...")
    
    for delim in delimitadores:
        for enc in encodings:
            try:
                df = pd.read_csv(
                    caminho_arquivo, 
                    sep=delim, 
                    encoding=enc,
                    nrows=10,  # Ler apenas primeiras linhas para teste
                    on_bad_lines='skip'
                )
                
                if len(df) > max_linhas:
                    max_linhas = len(df)
                    melhor_config = {
                        'delimitador': delim,
                        'encoding': enc,
                        'linhas_teste': len(df),
                        'colunas_teste': len(df.columns) if not df.empty else 0
                    }
                    
                    # Mostrar amostra
                    if len(df) > 0 and len(df.columns) > 1:
                        print(f"   ✓ {delim} + {enc}: {len(df)} linhas, {len(df.columns)} colunas")
                        
            except Exception:
                continue
    
    if melhor_config:
        print(f"\n✅ Melhor configuração encontrada:")
        print(f"   - Delimitador: {repr(melhor_config['delimitador'])}")
        print(f"   - Encoding: {melhor_config['encoding']}")
        print(f"   - Colunas detectadas: {melhor_config['colunas_teste']}")
        return melhor_config
    else:
        print(f"❌ Não foi possível detectar o formato do arquivo")
        return None


def csv_para_excel(caminho_csv, caminho_excel=None, config=None):
    """
    Converte CSV para Excel (XLSX) com diagnóstico automático
    
    Args:
        caminho_csv (str): Caminho para o arquivo CSV
        caminho_excel (str): Caminho para o arquivo Excel de saída (opcional)
        config (dict): Configuração manual (delimitador, encoding) - opcional
    
    Returns:
        bool: True se a conversão foi bem-sucedida
    """
    print("\n" + "=" * 70)
    print("CONVERSÃO CSV → EXCEL".center(70))
    print("=" * 70)
    
    # Definir nome de saída se não fornecido
    if not caminho_excel:
        nome_base = Path(caminho_csv).stem
        caminho_excel = f"{nome_base}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    print(f"📂 Arquivo fonte: {caminho_csv}")
    print(f"📁 Arquivo destino: {caminho_excel}")
    
    # Diagnosticar formato se não fornecido
    if not config:
        config = diagnosticar_csv(caminho_csv)
        if not config:
            return False
    
    try:
        # Ler CSV com a configuração detectada
        print(f"\n📖 Lendo arquivo CSV...")
        print(f"   - Delimitador: {repr(config['delimitador'])}")
        print(f"   - Encoding: {config['encoding']}")
        
        df = pd.read_csv(
            caminho_csv,
            sep=config['delimitador'],
            encoding=config['encoding'],
            engine='python',
            on_bad_lines='warn',
            skipinitialspace=True,
            keep_default_na=False,
            na_filter=False,
            low_memory=False
        )
        
        print(f"   ✅ Linhas lidas: {len(df)}")
        print(f"   ✅ Colunas: {len(df.columns)}")
        
        if len(df) == 0:
            print("❌ ERRO: Nenhuma linha lida do arquivo")
            return False
        
        # Mostrar amostra dos dados
        print(f"\n📊 AMOSTRA DOS DADOS (primeiras 3 linhas):")
        print(df.head(3).to_string())
        
        # Verificar duplicidade suspeita (quando deveriam ter 2 linhas e só mostra 1)
        if len(df) == 1 and 'deveriam' in input("\nHá suspeita de que o arquivo deveria ter mais linhas? (s/n): ").lower():
            print("\n🔧 Aplicando correção para arquivos com quebra de linha...")
            
            # Tentar ler como texto e dividir manualmente
            with open(caminho_csv, 'r', encoding=config['encoding']) as f:
                conteudo = f.read()
            
            # Tentar diferentes padrões de quebra
            linhas_brutas = []
            for separador in ['\r\n', '\n', '\r']:
                if separador in conteudo:
                    linhas_brutas = conteudo.split(separador)
                    if len(linhas_brutas) > 1:
                        print(f"   - Separador encontrado: {repr(separador)}")
                        print(f"   - Total de linhas brutas: {len(linhas_brutas)}")
                        break
            
            # Filtrar linhas vazias
            linhas_brutas = [l for l in linhas_brutas if l.strip()]
            print(f"   - Linhas não vazias: {len(linhas_brutas)}")
            
            if len(linhas_brutas) > 1:
                # Tentar ler cada linha como CSV
                dados = []
                cabecalho = linhas_brutas[0].split(config['delimitador'])
                
                for linha in linhas_brutas[1:]:
                    if linha.strip():
                        valores = linha.split(config['delimitador'])
                        if len(valores) == len(cabecalho):
                            dados.append(valores)
                
                if dados:
                    df = pd.DataFrame(dados, columns=cabecalho)
                    print(f"   ✅ Após correção: {len(df)} linhas")
        
        # Limpar dados (remover espaços extras)
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str).str.strip()
        
        # Salvar como Excel
        print(f"\n💾 Salvando arquivo Excel...")
        
        with pd.ExcelWriter(caminho_excel, engine='openpyxl') as writer:
            # Aba principal
            df.to_excel(writer, sheet_name='Dados', index=False)
            
            # Aba de diagnóstico
            diagnosticos = pd.DataFrame([{
                'Arquivo_origem': caminho_csv,
                'Data_conversao': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'Encoding': config['encoding'],
                'Delimitador': repr(config['delimitador']),
                'Linhas': len(df),
                'Colunas': len(df.columns),
                'Observacao': 'Convertido via script automático'
            }])
            diagnosticos.to_excel(writer, sheet_name='Diagnostico', index=False)
        
        print(f"✅ ARQUIVO EXCEL GERADO COM SUCESSO!")
        print(f"   - Local: {caminho_excel}")
        print(f"   - Tamanho: {os.path.getsize(caminho_excel)} bytes")
        
        # Verificar arquivo gerado
        df_verificacao = pd.read_excel(caminho_excel, sheet_name='Dados')
        print(f"   - Verificação: {len(df_verificacao)} linhas, {len(df_verificacao.columns)} colunas")
        
        return True
        
    except Exception as e:
        print(f"\n❌ ERRO DURANTE CONVERSÃO: {e}")
        return False


def converter_em_lote(pasta_origem, pasta_destino=None):
    """
    Converte todos os CSVs de uma pasta para Excel
    
    Args:
        pasta_origem (str): Pasta contendo os arquivos CSV
        pasta_destino (str): Pasta para salvar os Excel (opcional)
    """
    print("\n" + "=" * 70)
    print("CONVERSÃO EM LOTE".center(70))
    print("=" * 70)
    
    if not os.path.exists(pasta_origem):
        print(f"❌ Pasta não encontrada: {pasta_origem}")
        return
    
    if not pasta_destino:
        pasta_destino = os.path.join(pasta_origem, "EXCEL_CONVERTIDOS")
    
    os.makedirs(pasta_destino, exist_ok=True)
    
    # Listar arquivos CSV
    arquivos = list(Path(pasta_origem).glob("*.csv")) + list(Path(pasta_origem).glob("*.CSV"))
    
    if not arquivos:
        print(f"❌ Nenhum arquivo CSV encontrado em: {pasta_origem}")
        return
    
    print(f"📁 Pasta origem: {pasta_origem}")
    print(f"📁 Pasta destino: {pasta_destino}")
    print(f"📄 Arquivos encontrados: {len(arquivos)}")
    
    resultados = []
    
    for i, arquivo in enumerate(arquivos, 1):
        print(f"\n[{i}/{len(arquivos)}] Processando: {arquivo.name}")
        
        nome_excel = arquivo.stem + ".xlsx"
        caminho_excel = os.path.join(pasta_destino, nome_excel)
        
        # Diagnosticar formato
        config = diagnosticar_csv(str(arquivo))
        
        if config:
            sucesso = csv_para_excel(str(arquivo), caminho_excel, config)
            resultados.append({
                'arquivo': arquivo.name,
                'sucesso': sucesso,
                'destino': nome_excel if sucesso else None
            })
        else:
            resultados.append({
                'arquivo': arquivo.name,
                'sucesso': False,
                'destino': None
            })
    
    # Relatório final
    print("\n" + "=" * 70)
    print("RELATÓRIO FINAL".center(70))
    print("=" * 70)
    
    sucessos = sum(1 for r in resultados if r['sucesso'])
    falhas = len(resultados) - sucessos
    
    print(f"✅ Conversões com sucesso: {sucessos}")
    print(f"❌ Falhas: {falhas}")
    
    if falhas > 0:
        print("\nArquivos com falha:")
        for r in resultados:
            if not r['sucesso']:
                print(f"   - {r['arquivo']}")


def menu_principal():
    """
    Menu interativo para escolher as opções
    """
    print("\n" + "=" * 70)
    print("CONVERSOR CSV → EXCEL".center(70))
    print("=" * 70)
    
    while True:
        print("\nOPÇÕES:")
        print("1. Converter um arquivo CSV para Excel")
        print("2. Converter todos CSVs de uma pasta")
        print("3. Diagnosticar formato do arquivo")
        print("4. Sair")
        
        opcao = input("\nEscolha uma opção (1-4): ").strip()
        
        if opcao == '4':
            print("Encerrando...")
            break
        
        if opcao == '1':
            csv_path = input("Caminho do arquivo CSV: ").strip()
            excel_path = input("Caminho do arquivo Excel (ENTER para automático): ").strip()
            csv_para_excel(csv_path, excel_path if excel_path else None)
        
        elif opcao == '2':
            pasta = input("Caminho da pasta com CSVs: ").strip()
            destino = input("Pasta de destino (ENTER para subpasta 'EXCEL_CONVERTIDOS'): ").strip()
            converter_em_lote(pasta, destino if destino else None)
        
        elif opcao == '3':
            csv_path = input("Caminho do arquivo CSV: ").strip()
            diagnosticar_csv(csv_path)
        
        else:
            print("❌ Opção inválida.")


if __name__ == "__main__":
    # Se passar argumentos, executa modo automático
    if len(sys.argv) > 1:
        caminho = sys.argv[1]
        
        if os.path.isdir(caminho):
            # É uma pasta
            destino = sys.argv[2] if len(sys.argv) > 2 else None
            converter_em_lote(caminho, destino)
        else:
            # É um arquivo
            destino = sys.argv[2] if len(sys.argv) > 2 else None
            csv_para_excel(caminho, destino)
    else:
        # Menu interativo
        menu_principal()
