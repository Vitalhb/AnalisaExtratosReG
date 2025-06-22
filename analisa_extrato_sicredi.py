from pandas import DataFrame, MultiIndex
import pandas as pd
import math
from analisa_extrato import analisa_extrato, columns_historico, columns_extrato, categ_tuple

historico_sicredi = [
	"CESTA DE RELACIONAMENTO",
	"DB.CONV.PGTO FORNEC-NACIO",
	"DEBITO ARRECADACAO",
	"DEBITO CONVENIOS",
	"INTEGR.CAPITAL SUBSCRITO",
	"LIQ.COBRANCA SIMPLES",
	"LIQUIDACAO BOLETO",
	"LIQUIDACAO BOLETO SICREDI",
	"MANUTENCAO DE TITULOS",
	"PAGAMENTO PIX",
	"PAGAMENTO PIX SICREDI",
	"RECEBIMENTO PIX",
	"TARIFA COM R LIQUIDACAO",
	"TARIFAS - PGTO FORNECEDOR TED",
	"TED",
	"TED PAGAMENTO FORNECEDOR",
	"TRANSF ENTRE CONTAS",
]

class analisa_extrato_sicredi(analisa_extrato):
    bank_name = "Sicredi"
    categ_saidas = {
        "Tarifas": ["CESTA", "MANUTENCAO", "TARIFA"],
        "Tributos": ["DEBITO"],
        "Pagamentos": ["LIQUIDACAO", "DB.CONV", "PIX", "TED"],
        "Transf Env": ["TRANSF"],
    }
    categ_entradas = {
        "Saldo Inicial": ["Saldo"],
        "Recebimentos": ["COBRANCA", "PIX", "TED"],
        "Transf Rec": ["TRANSF"],
    }
    
    categs_list = ("Pagamentos","Recebimentos","Tarifas","Transf Env","Transf Rec","Tributos","Outros","Saldo Inicial","Saldo Final")

    extrato_prefix = "extrato R& Sicredi_"
    extrato_file_extension = "xls"

    def __init__(self, directory:str = None):
        super().__init__(categ_tuple(__class__.categ_entradas, __class__.categ_saidas), directory)
		
    def extrato_from_file(self, filename: str):
        if not __class__.bank_name in filename:
            raise(NameError("Banco invalido. Somente analisa " + __class__.bank_name))
        fullpath = __class__.directory + '/' + filename
        df = pd.read_excel(fullpath, header=8, parse_dates=True)
        last_idx = df[df.isna().all(1)].index[0]
        df = df.iloc[0:last_idx]
        df.columns= columns_extrato
        df['Data'] = pd.to_datetime(df['Data'], format="%d/%m/%Y")#, unit=)
        self.df_extrato = df

    def parse_historico(self) -> DataFrame:
        '''
			Separa o historico da Sicredi em:
				1- Descricao
				2- Codigo ou CPF/CNPJ
				3- Remetente/Destinatario
        '''
        parsedDf = self.df_extrato['Historico'].str.split(r"\s*(\d{5,})\s*", regex=True)
        slist = parsedDf.tolist()
        df1 = pd.DataFrame(slist)
        df1.columns = columns_historico
        return df1


if __name__ == "__main__":
    print("\nAnalisa extrato Sicredi\n")
    aes = analisa_extrato_sicredi("./extratos")
    aes.extrato_from_file("extrato R& Sicredi_202411.xls")
    aes.analisa_extrato()
    df = aes.df_extrato
    print(df)
    ano = df['Data'].iloc[-1].year
    mes = df['Data'].iloc[-1].month
    index = MultiIndex.from_tuples([(ano, mes)], names=["Ano", "Mes"])
    df_total = pd.DataFrame(df.groupby('Categoria')[['Valor']].sum()).transpose()
    df_total.index=index
    print(df_total)