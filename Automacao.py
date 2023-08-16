import os
import pandas as pd


class BaseDados:

    def __init__(self, email, loja, venda):
        self.emails = email  # 'Emails.xlsx'
        self.lojas = loja  # 'Lojas.csv'
        self.vendas = venda  # 'Vendas.xlsx'
        self.db_emails = None
        self.db_lojas = None
        self.db_vendas = None
        self.local = ''
        self.meta_faturamento_ano = 1650000
        self.meta_diversidade_produto_ano = 120
        self.meta_ticket_medio_ano = 500
        self.meta_faturamento_dia = 1000
        self.meta_diversidade_produto_dia = 4
        self.meta_ticket_medio_dia = 500

    def local_arquivos(self, local):
        direct = os.path.abspath(local)
        self.local = direct
        for arquivo in os.scandir(self.local):
            if self.emails in arquivo.name:
                self.emails = os.path.abspath(arquivo)
            elif self.lojas in arquivo.name:
                self.lojas = os.path.abspath(arquivo)
            elif self.vendas in arquivo.name:
                self.vendas = os.path.abspath(arquivo)
            elif r'~$' in arquivo.name:
                pass
            else:
                print(f'{arquivo.name} não localizado.')

    def salvar_db_emails(self):
        if '.xlsx' in self.emails:
            self.db_emails = pd.read_excel(self.emails)
        elif '.csv' in self.emails:
            self.db_emails = pd.read_csv(self.emails, encoding="ISO-8859-1", sep=';')
        return self.db_emails

    def salvar_db_lojas(self):
        if '.xlsx' in self.lojas:
            self.db_lojas = pd.read_excel(self.lojas)
        elif '.csv' in self.lojas:
            self.db_lojas = pd.read_csv(self.lojas, encoding="ISO-8859-1", sep=';')
        return self.db_lojas

    def salvar_db_vendas(self):
        if '.xlsx' in self.vendas:
            self.db_vendas = pd.read_excel(self.vendas)
        elif '.csv' in self.vendas:
            self.db_vendas = pd.read_csv(self.vendas, encoding="ISO-8859-1", sep=';')
        return self.db_vendas

    @staticmethod
    def titulos(db):
        print(list(db.columns))
        return list(db.columns)

    def filtro_loja(self, id_loja):
        filtro = self.db_vendas.loc[self.db_vendas['ID Loja'] == id_loja]
        return filtro

    # Data no formato 'YYYY-MM-DD'
    def filtro_data(self, id_loja):
        filtro = self.db_vendas.loc[self.db_vendas['ID Loja'] == id_loja]
        filtro = filtro.loc[self.db_vendas['Data'] == self.db_vendas['Data'].iloc[-1]]
        return filtro, self.db_vendas['Data'].iloc[-1]

    def qtde_lojas(self):
        qtde = self.db_lojas.nunique()
        return qtde['ID Loja']

    @staticmethod
    def qtde_vendas(db):
        vendas = len(db)
        return vendas

    @staticmethod
    def qtde_produtos_diversos(db):
        produtos = db.nunique()
        return produtos['Produto']

    @staticmethod
    def faturamento(db):
        faturamento = db['Valor Final'].sum()
        return faturamento

    def ticket_medio(self, db):
        if self.qtde_vendas(db) > 0:
            ticket_medio = self.faturamento(db) / self.qtde_vendas(db)
        else:
            ticket_medio = 0
        return ticket_medio

    def ver_meta_faturamento_ano(self, db):
        if self.faturamento(db) > self.meta_faturamento_ano:
            meta = 'Bateu meta'
        else:
            meta = 'Não bateu meta'
        return meta, self.meta_faturamento_ano

    def ver_meta_faturamento_dia(self, db):
        if self.faturamento(db) >= self.meta_faturamento_dia:
            meta = 'Bateu meta'
        else:
            meta = 'Não bateu meta'
        return meta, self.meta_faturamento_dia

    def ver_meta_diversidade_ano(self, db):
        if self.qtde_produtos_diversos(db) >= self.meta_diversidade_produto_ano:
            meta = 'Bateu meta'
        else:
            meta = 'Não bateu meta'
        return meta, self.meta_diversidade_produto_ano

    def ver_meta_diversidade_dia(self, db):
        if self.qtde_produtos_diversos(db) >= self.meta_diversidade_produto_dia:
            meta = 'Bateu meta'
        else:
            meta = 'Não bateu meta'
        return meta, self.meta_diversidade_produto_dia

    def ver_meta_ticket_ano(self, db):
        if self.ticket_medio(db) > self.meta_ticket_medio_ano:
            meta = 'Bateu meta'
        else:
            meta = 'Não bateu meta'
        return meta, self.meta_ticket_medio_ano

    def ver_meta_ticket_dia(self, db):
        if self.ticket_medio(db) > self.meta_ticket_medio_dia:
            meta = 'Bateu meta'
        else:
            meta = 'Não bateu meta'
        return meta, self.meta_ticket_medio_dia

    def top_loja(self, lista_dicionario):
        dicionario_loja = max(lista_dicionario, key=lambda x: x['Faturamento'])
        nome = self.db_lojas.loc[self.db_lojas['ID Loja'] == dicionario_loja['ID Loja'], 'Loja'].iloc[0]
        return dicionario_loja, nome

    def pior_loja(self, lista_dicionario):
        dicionario_loja = min(lista_dicionario, key=lambda x: x['Faturamento'])
        nome = self.db_lojas.loc[self.db_lojas['ID Loja'] == dicionario_loja['ID Loja'], 'Loja'].iloc[0]
        return dicionario_loja, nome

    def nome_loja(self, lista_dicionario, id_loja):
        nome = self.db_lojas.loc[self.db_lojas['ID Loja'] == lista_dicionario[id_loja - 1]['ID Loja'], 'Loja'].iloc[0]
        return nome

    def id_loja(self, nome_loja):
        id_loja = self.db_lojas.loc[self.db_lojas['Loja'] == nome_loja, 'ID Loja'].iloc[0]
        return id_loja

    def gerente_desitinatario(self, nome_loja):
        gerente = self.db_emails.loc[self.db_emails['Loja'] == nome_loja, 'Gerente'].iloc[0]
        destinatario = self.db_emails.loc[self.db_emails['Loja'] == nome_loja, 'E-mail'].iloc[0]
        return gerente, destinatario
