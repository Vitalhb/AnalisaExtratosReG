from pandas import DataFrame, MultiIndex
import pandas as pd
import math
from analisa_extrato import analisa_extrato, columns_historico, columns_extrato, categ_tuple

historico_santander = [
	"CR COB BLOQ COMP CONF RECEBIMENTO",
	"CR COB DINHEIRO CONF RECEBIMENTO",
	"PAGAMENTO A FORNECEDORES",
	"PAGAMENTO DE BOLETO OUTROS BANCOS",
	"PAGAMENTO DE TITULO",
	"PGTO FORNECEDORES - TRIB FEDERAL",
	"PGTO FORNECEDORES - TRIB MUNICIPAL",
	"PGTO FORNECEDORES -CONCESSIONARIAS",
	"PIX ENVIADO",
	"PIX RECEBIDO",
	"TAR EMISSAO TED CIP PGTO FORNEC",
	"TARIFA AVULSA ENVIO PIX",
	"TARIFA BAIXA OU DEVOL DE TITULO",
	"TARIFA EXTRATO INTELIGENTE",
	"TARIFA MENSALIDADE PACOTE SERVICOS",
	"TED PGTO FORNECEDORES CIP",
	"TED RECEBIDA",
	"TRANSFERENCIA ENTRE CONTAS",
]

class analisa_extrato_santander(analisa_extrato):
    bank_name = "Santander"

    categ_saidas = {
        "Tarifas": ["TARIFA", "TAR"],
        "Tributos": ["TRIB FEDERAL", "TRIB ESTADUAL", "TRIB MUNICIPAL"],
        "Pagamentos": ["PAGAMENTO", "PGTO", "PIX", "TED"],
        "Transf Env": ["TRANSFERENCIA"],
    }

    categ_entradas = {
        "Saldo Inicial": ["SALDO"],
        "Recebimentos": ["RECEBIMENTO", "PIX", "TED"],
        "Transf Rec": ["TRANSFERENCIA"],
    }

    categs_list = ("Pagamentos","Recebimentos","Tarifas","Transf Env","Transf Rec","Tributos","Outros","Saldo Inicial","Saldo Final")
    
    extrato_prefix = "extrato R& Santander_"
    extrato_file_extension = "xls"

    def __init__(self, directory:str = None):
        super().__init__(categ_tuple(__class__.categ_entradas, __class__.categ_saidas), directory)
	
    def extrato_from_file(self, filename: str):
        if not __class__.bank_name in filename:
            raise(NameError("Banco invalido. Somente analisa " + __class__.bank_name))
        fullpath = __class__.directory + '/' + filename
        df = pd.read_excel(fullpath, header=2, parse_dates=True).dropna(how='all', axis='columns')
        df.columns= columns_extrato
        lastvalue = 0
        for index, row in df.iterrows(): # fill the empty rows of saldo with calculated values (needed for Santander).
            saldo = row["Saldo"]
            if math.isnan(saldo):
                saldo = lastvalue + row["Valor"]
                df.loc[index,"Saldo"] = saldo
            lastvalue = saldo
        df['Data'] = pd.to_datetime(df['Data'], format="%d/%m/%Y")#, unit=)
        self.df_extrato = df

    def parse_historico(self) -> DataFrame:
        '''
            Separa o historico do Santander em:
                1- Descricao
                2- Codigo ou CPF/CNPJ ou destinatario de PIX
        '''
        df1 = self.df_extrato['Historico'].str.extract(r"^(.{,35})([./\d]*)(.*)")
        for col in df1.columns:
            df1[col] = df1[col].str.strip()
        df1.columns = columns_historico
        return df1


if __name__ == "__main__":
    print("\nAnalisa extrato Santander\n")
    aes = analisa_extrato_santander("./extratos")
    aes.extrato_from_file("extrato R& Santander_202411.xls")
    aes.analisa_extrato()
    df = aes.df_extrato
    print(df)
    ano = df['Data'].iloc[-1].year
    mes = df['Data'].iloc[-1].month
    index = MultiIndex.from_tuples([(ano, mes)], names=["Ano", "Mes"])
    df_total = pd.DataFrame(df.groupby('Categoria')[['Valor']].sum()).transpose()
    df_total.index=index
    print(df_total)