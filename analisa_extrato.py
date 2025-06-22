from pandas import DataFrame, MultiIndex
import pandas as pd
from collections import namedtuple
from abc import ABCMeta, abstractmethod

categ_tuple = namedtuple('categ_tuple', 'entradas saidas')
columns_extrato = ["Data", "Historico", "Documento", "Valor", "Saldo"]
columns_historico = ['Descricao', 'Codigo', 'RemDest']

class analisa_extrato(metaclass=ABCMeta):
	categs = None
	directory = None

	def __init__(self, categs: categ_tuple, directory:str = None):
		__class__.categs = categs
		if directory:
			__class__.directory = directory
		elif not __class__.directory:
			raise(ValueError("Directory must be initialized!"))

	def categoria(tcategs: categ_tuple, descricao: str, valor: float) -> str:
		if valor < 0:
			categorias = tcategs.saidas # categorias para saidas
		else: 
			categorias = tcategs.entradas # categorias para entradas
			
		udescricao = descricao.upper()
		for key, values in categorias.items():
			if any(item in udescricao for item in values):
				return key
		return "Outros"

	@abstractmethod
	def extrato_from_file(self, filename: str) -> DataFrame:
		pass

	@abstractmethod
	def parse_historico(self, bank: str, df_extrato: DataFrame) -> DataFrame:
		pass

	def analisa_extrato(self):
		df1 = self.parse_historico()
		self.df_extrato = pd.concat([self.df_extrato, df1], axis=1)
		self.df_extrato['Categoria'] = self.df_extrato[['Descricao', 'Valor']].apply(lambda x: analisa_extrato.categoria(__class__.categs, x['Descricao'], x['Valor']), axis=1)



if __name__ == "__main__":
    print("\nAbstract class analisa_extrato\n")
