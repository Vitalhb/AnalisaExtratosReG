import argparse
import os
from glob import glob
from os.path import isfile, join
import re
import pandas as pd
import xlsxwriter as xw
from xlsxwriter.format import Format
from analisa_extrato_santander import analisa_extrato_santander
from analisa_extrato_sicredi import analisa_extrato_sicredi
from analisa_extrato import columns_extrato, analisa_extrato

banks = {
    "Sicredi": analisa_extrato_sicredi,
    "Santander": analisa_extrato_santander,
}

categ_table_columns = ["Descricao", "RemDest", "Data", "Valor"]

def get_col_widths(dataframe):
    # First we find the maximum length of the index column   
    idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
    # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
    return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) for col in dataframe.columns]

    
def format_columns(worksheet, df, columns, format, first_column = 0):
        # Auto-adjust columns' width
    for column in columns:
        col_idx = df.columns.get_loc(column) + first_column
        if (df[column].dtype == "float64"):
            width = max(len(column)+1.8, 12)
            worksheet.set_column(col_idx, col_idx, width, format)
    # for i, width in enumerate(get_col_widths(df)):
    #     worksheet.set_column(i, i, width)

def write_table(writer: pd.ExcelWriter, sheet_name: str, df_table: pd.DataFrame, title: str, total_column: str, start_column: int, start_row: int, formats: dict[Format]) -> int:
    
    worksheet = writer.sheets[sheet_name]
    worksheet.merge_range(start_row, start_column, start_row, start_column+len(df_table.columns)-1, title, formats['centered_bold'])
    df_table.to_excel(writer, sheet_name=sheet_name, index=False, header=True, startcol=start_column, startrow=start_row+1)
    #write total on the bottom
    total = df_table[total_column].sum()
    row_totais = len(df_table) + start_row + 2
    bold = formats['bold']
    bold.set_border(1)
    currency_bold = formats['currency_bold']
    currency_bold.set_border(1)
    worksheet.write(row_totais, start_column, "Total", bold)
    worksheet.write(row_totais, start_column+1, "", bold)
    worksheet.write(row_totais, start_column+2, "", bold)
    worksheet.write(row_totais, start_column+3, total, currency_bold)

    return row_totais


def extratos_to_excel(aes: analisa_extrato, file_list: list[str]):
    results_list = list()
    totais_list = list()
    categs_list = aes.categs_list
    zero_series = pd.Series(0.0, index=categs_list)
    for fname in file_list:
        aes.extrato_from_file(fname)
        aes.analisa_extrato()
        df_results = aes.df_extrato
        results_list.append(df_results)
        
        ano = df_results['Data'].iloc[-1].year
        mes = df_results['Data'].iloc[-1].month
        num_extr = pd.MultiIndex.from_tuples([(ano, mes)], names=["Ano", "Mes"]) 
        # sum_series = df_results.groupby('Categoria')[['Valor']].sum()
        sum_series = df_results.groupby('Categoria')['Valor'].sum()
        sum_series = (sum_series + zero_series).fillna(0)
        sum_series['Saldo Inicial'] = df_results['Saldo'].iloc[0]
        sum_series['Saldo Final'] = df_results['Saldo'].iloc[-1]
        df_total = pd.DataFrame(sum_series).transpose()
        df_total.index=num_extr
        totais_list.append(df_total)


    df_totais = pd.concat(totais_list)
    ano = df_totais.index.get_level_values(0)[-1]
    inicio = str(df_totais.index.get_level_values(0)[0] * 100 + df_totais.index.get_level_values(1)[0])
    fim = str(ano * 100 + df_totais.index.get_level_values(1)[-1])
    outfilename = "extratos_" + bank + "_" + inicio + '_' + fim + '.xlsx'

    with pd.ExcelWriter(outfilename, engine='xlsxwriter', datetime_format='DD-MM-YYYY') as writer:
        # Define formats for cells
        # call_format_currency = writer.book.add_format().set_num_format('#,##0.00;[Red]-#,##0.00')
        formats = {
            'currency': writer.book.add_format({'num_format': '#,##0.00;[Red]-#,##0.00'}),
            # 'currency': writer.book.add_format().set_num_format('#,##0.00;[Red]-#,##0.00'),
            'centered_bold': writer.book.add_format({'align': 'center',
                                                        'valign': 'vcenter',
                                                        'border': 1,
                                                        'bold': True}),
            'bold': writer.book.add_format({'bold': True}),
            'currency_bold': writer.book.add_format({'num_format': '#,##0.00;[Red]-#,##0.00',
                                                        'bold': True}),
        }

        # Analise tab
        df_totais.to_excel(writer, sheet_name='Analise', index=True, header=True, columns=categs_list)
        df_totais_len = len(df_totais)
        #save the total of each column
        worksheet = writer.sheets['Analise']
        worksheet.write(df_totais_len+2, 0, "Total", formats['centered_bold'])
        df_totais_ano = df_totais.groupby(level=0).sum()
        # df_totais_ano['Saldo Inicial'].iloc[0] = df_totais['Saldo Inicial'].iloc[0]
        # df_totais_ano['Saldo Final'].iloc[0] = df_totais['Saldo Final'].iloc[-1]
        saldo_inicial_pos = df_totais_ano.columns.get_loc('Saldo Inicial')
        saldo_final_pos = df_totais_ano.columns.get_loc('Saldo Final')
        df_totais_ano.iloc[0, saldo_inicial_pos] = df_totais['Saldo Inicial'].iloc[0]
        df_totais_ano.iloc[0, saldo_final_pos] = df_totais['Saldo Final'].iloc[-1]
        df_totais_ano.to_excel(writer, sheet_name='Analise', index=True, header=True, startcol=1, startrow=df_totais_len+2, columns=categs_list)

        # Auto-adjust columns' width
        format_columns(worksheet, df_totais, categs_list, formats['currency'], 2)
        worksheet.autofit()

        for result in results_list:
            #write extrato on the left
            tab_name = str(result['Data'].iloc[-1].year * 100 + result['Data'].iloc[-1].month)
            result.to_excel(writer, sheet_name=tab_name, index=False, header=True, columns=columns_extrato)
            worksheet = writer.sheets[tab_name]
            format_columns(worksheet, result, columns_extrato, formats['currency'])

            start_column = len(columns_extrato) + 1
            worksheet.set_column(start_column-1, start_column-1, 3)
            start_row = 0

            for categ in categs_list:
                if not categ.startswith("Saldo"):
                    df_categ = result[categ_table_columns][result['Categoria'] == categ].sort_values(["Descricao", "RemDest"])
                    if (start_row + len(df_categ)) > max(len(result), 35):
                        format_columns(worksheet, df_categ, categ_table_columns, formats['currency'], start_column)
                        start_row = 0
                        start_column += len(categ_table_columns) + 1
                        worksheet.set_column(start_column-1, start_column-1, 3)

                    start_row = write_table(writer, tab_name, df_categ, categ, 'Valor', start_column, start_row, formats) + 2

            format_columns(worksheet, df_categ, categ_table_columns, formats['currency'], start_column)
            worksheet.autofit()

if __name__ == '__main__':
    try:
        parser = argparse.ArgumentParser(description='Processa extratos da ReG - Santander e Sicredi.')
        parser.add_argument('-fe', '--folder_extratos', help='diretorio com os extratos', type=str, default="extratos")
        parser.add_argument('-y', '--year', type=str, default='2024', metavar='YYYY', help='Ano para processar (AAAA)')

        args = parser.parse_args()
        year = args.year
        initial_date = year + '01'
        final_date = year + '12'

        folder_extratos = args.folder_extratos
        print("Pasta de extratos:", folder_extratos)
        print("Ano: ", year)
        for bank, _class in banks.items():
            aes = _class(folder_extratos)
            year_month_re = re.compile(r'(\d{6})\.' + aes.extrato_file_extension + '$')
            files = glob(aes.extrato_prefix + "*." + aes.extrato_file_extension, root_dir=folder_extratos + "\\")
            file_list_bank = [file for file in files if (mo:=re.search(aes.extrato_prefix + "(\d{6})", file)) and mo.group(1) >= initial_date and mo.group(1) <= final_date]
            print("Banco " + bank + ", Extratos:", ", ".join(file_list_bank))
            analysis_result = extratos_to_excel(aes, file_list_bank)

    except argparse.ArgumentError as e:
        print('Error:', e)