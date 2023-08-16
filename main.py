import pandas as pd
from Automacao import BaseDados
from AutomacaoEmail import EnvioEmail
from datetime import date
import os

data_relatorio = ""
cwd = os.getcwd().replace('\\', '/')
# pd.set_option('display.max_columns', None)

db = BaseDados('Emails', 'Lojas', 'Vendas')
db.local_arquivos('Projeto AutomacaoIndicadores/Bases de Dados')
db_emails = db.salvar_db_emails()
db_lojas = db.salvar_db_lojas()
db_vendas = db.salvar_db_vendas()
# print(db_emails)
# print(db_lojas)
# print(db_vendas)
top_loja_ano = []
indicadores_ano = []
top_loja_dia = []
indicadores_dia = []
for i in range(db.qtde_lojas()):
    # Indicadores Anuais
    filtro_loja = db.filtro_loja(i + 1)
    qtde_vendas = db.qtde_vendas(filtro_loja)
    qtde_produtos_diversos_ano = db.qtde_produtos_diversos(filtro_loja)
    meta_diversidade, valor_meta_diversidade_ano = db.ver_meta_diversidade_ano(filtro_loja)
    faturamento_ano = db.faturamento(filtro_loja)
    meta_faturamento, valor_meta_faturamento_ano = db.ver_meta_faturamento_ano(filtro_loja)
    ticket_medio_ano = db.ticket_medio(filtro_loja)
    meta_ticket, valor_meta_ticket_ano = db.ver_meta_ticket_ano(filtro_loja)

    indicadores_ano.append({'ID Loja': i + 1, 'Faturamento': faturamento_ano, 'Meta faturamento': meta_faturamento,
                            'Diversidade Produtos': qtde_produtos_diversos_ano,
                            'Meta diversidade produtos': meta_diversidade,
                            'Ticket Médio': ticket_medio_ano, 'Meta ticket': meta_ticket})

    arquivo_ano = filtro_loja.to_excel(f'Projeto AutomacaoIndicadores/Backup Arquivos Lojas/'
                                       f'{db.nome_loja(indicadores_ano, i + 1)}'
                                       f'_VendasAnual{date.today().strftime("%Y%m%d")}'
                                       f'.xlsx', index=False)

    caminho_arquivo_ano = f'{cwd}/Projeto AutomacaoIndicadores/Backup Arquivos Lojas/' \
                          f'{db.nome_loja(indicadores_ano, i + 1)}_VendasAnual{date.today().strftime("%Y%m%d")}.xlsx'

    # Indicadores Última data
    filtro_loja_data, data = db.filtro_data(i + 1)
    data_relatorio = data.strftime("%Y-%m-%d")
    qtde_vendas2 = db.qtde_vendas(filtro_loja_data)
    qtde_produtos_diversos_dia = db.qtde_produtos_diversos(filtro_loja_data)
    meta_diversidade2, valor_meta_diversidade_dia = db.ver_meta_diversidade_dia(filtro_loja_data)
    faturamento2 = db.faturamento(filtro_loja_data)
    meta_faturamento2, valor_meta_faturamento_dia = db.ver_meta_faturamento_dia(filtro_loja_data)
    ticket_medio_dia = db.ticket_medio(filtro_loja_data)
    meta_ticket2, valor_meta_ticket_dia = db.ver_meta_ticket_dia(filtro_loja_data)

    indicadores_dia.append({'ID Loja': i + 1, 'Faturamento': faturamento2, 'Meta faturamento': meta_faturamento2,
                            'Diversidade Produtos': qtde_produtos_diversos_dia,
                            'Meta diversidade produtos': meta_diversidade2,
                            'Ticket Médio': ticket_medio_dia, 'Meta ticket': meta_ticket2,
                            'Data': data.strftime("%Y-%m-%d")})

    arquivo_dia = filtro_loja_data.to_excel(f'Projeto AutomacaoIndicadores/Backup Arquivos Lojas/'
                                            f'{db.nome_loja(indicadores_dia, i + 1)}'
                                            f'_VendasDia{data.strftime("%Y%m%d")}.xlsx',
                                            index=False)

    caminho_arquivo_dia = f'{cwd}/Projeto AutomacaoIndicadores/Backup Arquivos Lojas/' \
                          f'{db.nome_loja(indicadores_dia, i + 1)}_VendasDia{data.strftime("%Y%m%d")}.xlsx'

    nome_loja = db.nome_loja(indicadores_dia, i + 1)
    gerente, destinatario = db.gerente_desitinatario(nome_loja)
    email = EnvioEmail()
    email.enviar_email(destinatario, gerente, nome_loja, data.strftime("%Y%m%d"),
                       faturamento2, valor_meta_faturamento_dia, ticket_medio_dia, valor_meta_ticket_dia,
                       qtde_produtos_diversos_dia, valor_meta_diversidade_dia,
                       faturamento_ano, valor_meta_faturamento_ano, ticket_medio_ano, valor_meta_ticket_ano,
                       qtde_produtos_diversos_ano, valor_meta_diversidade_ano,
                       caminho_arquivo_ano, caminho_arquivo_dia)

# Indicadores TOP e Piores Lojas
top_loja_ano, nome_loja_ano = db.top_loja(indicadores_ano)
top_loja_dia, nome_loja_dia = db.top_loja(indicadores_dia)
pior_loja_ano, nome_pior_loja_ano = db.pior_loja(indicadores_ano)
pior_loja_dia, nome_pior_loja_dia = db.pior_loja(indicadores_dia)

# TOP LOJA ANO
arquivo_top_ano = db.filtro_loja(top_loja_ano['ID Loja']).to_excel(
    f"Projeto AutomacaoIndicadores/Backup Arquivos Lojas/{nome_loja_ano}"
    f"_TopLojaAno{date.today().strftime('%Y%m%d')}.xlsx", index=False)

caminho_top_ano = f'{cwd}/Projeto AutomacaoIndicadores/Backup Arquivos Lojas/' \
                  f'{nome_loja_ano}_TopLojaAno{date.today().strftime("%Y%m%d")}.xlsx'

# PIOR LOJA ANO
arquivo_pior_ano = db.filtro_loja(pior_loja_ano['ID Loja']).to_excel(
    f"Projeto AutomacaoIndicadores/Backup Arquivos Lojas/{nome_pior_loja_ano}"
    f"_PiorLojaAno{date.today().strftime('%Y%m%d')}.xlsx", index=False)

caminho_pior_ano = f'{cwd}/Projeto AutomacaoIndicadores/Backup Arquivos Lojas/' \
                   f'{nome_pior_loja_ano}_PiorLojaAno{date.today().strftime("%Y%m%d")}.xlsx'

# TOP LOJA DIA
filtro_top_dia, nome_loja_top_dia = db.filtro_data(top_loja_dia['ID Loja'])

arquivo_top_dia = filtro_top_dia.to_excel(f"Projeto AutomacaoIndicadores/Backup Arquivos Lojas/{nome_loja_dia}"
                                          f"_TopLojaDia{data_relatorio}.xlsx", index=False)

caminho_top_dia = f'{cwd}/Projeto AutomacaoIndicadores/Backup Arquivos Lojas/' \
                  f'{nome_loja_dia}_TopLojaDia{data_relatorio}.xlsx'

# PIOR LOJA DIA
filtro_pior_dia, nome_loja_pior_dia = db.filtro_data(pior_loja_dia['ID Loja'])

arquivo_pior_dia = filtro_pior_dia.to_excel(f"Projeto AutomacaoIndicadores/Backup Arquivos Lojas/{nome_pior_loja_dia}"
                                            f"_PiorLojaDia{data_relatorio}.xlsx", index=False)

caminho_pior_dia = f'{cwd}/Projeto AutomacaoIndicadores/Backup Arquivos Lojas/' \
                   f'{nome_pior_loja_dia}_PiorLojaDia{data_relatorio}.xlsx'

gerente, destinatario = db.gerente_desitinatario('Diretoria')
email = EnvioEmail()
email.enviar_email_diretoria(destinatario, gerente, data_relatorio,
                             nome_loja_dia, top_loja_dia['Faturamento'],
                             nome_loja_ano, top_loja_ano['Faturamento'],
                             nome_pior_loja_dia, pior_loja_dia['Faturamento'],
                             nome_pior_loja_ano, pior_loja_ano['Faturamento'],
                             caminho_top_ano, caminho_pior_ano, caminho_top_dia, caminho_pior_dia)

print('Fim da automação.')
