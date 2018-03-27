# coding: latin-1
from xlrd import open_workbook
from xlwt import Workbook
import locale
import sqlite3
import argparse
import csv

locale.setlocale(locale.LC_ALL,'portuguese')

def cria_tabelas(cursor):
	cursor.execute('''create table tg_contas (
		account real,
		country text,
		amount real,
		product text)''')

def loads(book,conta,sheet,linha_inicial,coluna_inicial,cursor,pais):
	s = book.sheet_by_name(sheet)
	print 'Leitura da folha:',s.name
	wconta_brancos = 0
	for row in range(linha_inicial,s.nrows):
		values = []
		if s.cell(row,coluna_inicial+1).value == '':
			wconta_brancos = 1
		if wconta_brancos == 0:
			#print s.cell(row,coluna_inicial).value, s.cell(row,coluna_inicial+1).value
			values.append(conta)
			values.append(pais)
			values.append(s.cell(row,coluna_inicial).value)
			values.append(s.cell(row,coluna_inicial+1).value)
			cursor.execute('insert into tg_contas values (?,?,?,?)',values)
	db.commit()

# ainda para desenvolver
def get_product_description(product_id):
	prod_id = {'FS_01-010':'Food Service-Coffee-Beans',
				'FS_01-020':'Food Service-Coffee-Ground',
				'FS_01-030':'Food Service-Coffee-Instant',
				'FS_01-040':'Food Service-Coffee-Single Serve',
				'FS_02-010':'Food Service-Other Food-Tea',
				'FS_02-020':'Food Service-Other Food-Spices & Seasoning',
				'FS_02-030':'Food Service-Other Food-Other',
				'MM_01-010':'Mass Market-Coffee-Beans',
				'MM_01-020':'Mass Market-Coffee-Ground',
				'MM_01-030':'Mass Market-Coffee-Instant',
				'MM_01-040':'Mass Market-Coffee-Single Serve',
				'MM_02-010':'Mass Market-Other Food-Tea',
				'MM_02-020':'Mass Market-Other Food-Spices & Seasoning',
				'MM_02-030':'Mass Market-Other Food-Other',
				'OTH_03-010':'Other-Machines & Equipment-Grinders',
				'OTH_03-020':'Other-Machines & Equipment-HCS',
				'OTH_03-030':'Other-Machines & Equipment-OCS',
				'OTH_03-040':'Other-Machines & Equipment-Maintenance & other services',
				'OTH_03-050':'Other-Machines & Equipment-Spare Parts',
				'OTH_03-060':'Other-Machines & Equipment-Professional machines',
				'OTH_03-070':'Other-Machines & Equipment-Other equipment',
				'OTH_04-010':'Other-Cafes-Coffee Shop',
				'OTH_05-010':'Other-Other-Green Coffee',
				'OTH_05-020':'Other-Other-Other',
				'OTH_05-030':'Other-Other-POP material',
				'OTH_05-040':'Other-Other-Roasted and Semi-finished Products',
				'OTH_05-050':'Other-Other-Royalties',
				'OTH_05-060':'Other-Other-Services',
				'PL_01-010':'Private Label-Coffee-Beans',
				'PL_01-020':'Private Label-Coffee-Ground',
				'PL_01-030':'Private Label-Coffee-Instant',
				'PL_01-040':'Private Label-Coffee-Single Serve',
				'PL_02-010':'Private Label-Other Food-Tea',
				'PL_02-020':'Private Label-Other Food-Spices & Seasoning',
				'PL_02-030':'Private Label-Other Food-Other'}
	return prod_id[product_id]

def folha1(conta):
	xsheet1 = xbook.add_sheet('C'+str(conta)+'_ptsp')
	xlin1 = 0

	cur.execute('''select country, amount, product from tg_contas where account = ? ''',(conta,))
	linhas= cur.fetchall()
	for row in linhas:
		# correr as linhas
		# print row[0], row[1], row[2]
		xsheet1.write(xlin1,0,row[0])
		xsheet1.write(xlin1,1,row[1])
		xsheet1.write(xlin1,2,row[2])
		xlin1 = xlin1 +1

def folha2(conta):

	cost = {32730:'DM_0030', 32740:'DM_0030', 32207:'DM_0030', 31210:'NC'}
	sinal = {32730:-1, 32740:1, 32207:1, 31210:-1}

	xsheet2 = xbook.add_sheet('C'+str(conta)+'_ptsp_pv')
	xlin2 = 0

	cur.execute('''select product, sum(amount) from tg_contas where account = ? group by product order by 1 ''',(conta,))
	linhas= cur.fetchall()
	for row in linhas:
		# correr as linhas
		# print row[0], row[1]
		xsheet2.write(xlin2,0,row[0])
		xsheet2.write(xlin2,1,row[1])

		#criar já o layout para o meu ficheiro de envio
		xsheet2.write(xlin2,3,wanomes)
		xsheet2.write(xlin2,4,'NC')
		xsheet2.write(xlin2,5,'NC')
		xsheet2.write(xlin2,6,row[0])
		xsheet2.write(xlin2,7,get_product_description(row[0]))
		xsheet2.write(xlin2,8,round(row[1],2))
		xsheet2.write(xlin2,12,conta)
		xsheet2.write(xlin2,13,cost[conta])
		xsheet2.write(xlin2,14,sinal[conta])

		xlin2 = xlin2 +1


# inicio
# gestão de parametros
parser = argparse.ArgumentParser(description='Gera ficheiro TAGETIK - Conta 32730 / 32740 / 32207')
parser.add_argument('integers', metavar='WANOMES', type=int,  help='O ano e mês de trabalho no formato YYYYMM')
args = parser.parse_args()
wanomes = int(args.integers)
wano = str(wanomes)[:-2]
wmes = str(wanomes)[4:]

print 'Ano: '+wano
print 'Mês: '+wmes

db = sqlite3.connect(':memory:')
cur = db.cursor()

cria_tabelas(cur)

# Ler o ficheiro de PORTUGAL
wb = open_workbook(wmes+' '+wano+' MOVIMENTOS DIM PORT.xlsx')
loads(wb,32730,'32730 PORT',3,2,cur,'PT')
loads(wb,32740,'EXIST INICIAL PORT',3,2,cur,'PT')
loads(wb,32207,'32207 PORT',3,2,cur,'PT')
loads(wb,31210,'31210 PORT',3,2,cur,'PT')


# Ler o ficheiro de ESPANHA
wb = open_workbook(wmes+' '+wano+' MOVIMENTOS DIM ESP.xlsx')
loads(wb,32730,'32730 ESP',3,2,cur,'SP')
loads(wb,32740,'EXIST INICIAL ESP',3,2,cur,'SP')
loads(wb,32207,'32207 ESP',3,2,cur,'SP')
loads(wb,31210,'31210 ESP',3,2,cur,'SP')


# Criar o ficheiro de Saída
xbook = Workbook()

# Conta 32730
wconta = 32730
folha1(wconta)
folha2(wconta)

wconta = 32740
folha1(wconta)
folha2(wconta)

wconta = 32207
folha1(wconta)
folha2(wconta)

wconta = 31210
folha1(wconta)
folha2(wconta)

xbook.save(str(wanomes)+'_contasDM_ptsp.xls')


