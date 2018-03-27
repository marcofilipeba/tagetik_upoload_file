# coding: latin-1
from xlrd import open_workbook
import locale
import sqlite3
import argparse
import csv

locale.setlocale(locale.LC_ALL,'portuguese')


def cria_tabelas(cursor):
	cursor.execute('''create table tg_percent (
		anomes integer,
		client text,
		client_dsc text,
		profit text,
		profit_dsc text,
		valor real,
		total real,
		stotal real,
		contratosgrp real,
		salescostpessoal real,
		conta integer,
		costcenter text,
		sinal integer)''')
		
	cursor.execute('''create table tg_costcenter (
		mapa text,
		costcenter text,
		client text)''')
		
	cursor.execute('''create table tg_imput (
		anomes integer,
		mapa text,
		conta integer,
		valor real,
		imput text)''')

	cursor.execute('''create table tg_upload (
		anomes integer,
		mapa text,
		ano integer,
		mes text,
		empresa text,
		conta integer,
		costcenter text,
		profit text,
		client text,
		amount real)''')

def loads(book,sheet,tabela,nrcampos,cursor):
	s = book.sheet_by_name(sheet)
	print 'Load:',s.name
	for row in range(1,s.nrows):
		values = []
		for col in range(s.ncols):
			values.append(s.cell(row,col).value)
		campos = '?,'*nrcampos
		cursor.execute('insert into '+tabela+' values ('+campos[:-1]+')',values)

	# Este update e para nao ter problemas com os espacos que porventura podem vir a mais no texto da imputacao total e stotal
	if sheet == 'TG_IMPUT':
		cursor.execute('''update tg_imput set imput = trim(imput) where anomes = ?''',(wanomes,))
	db.commit()

def gerafich(cursor):
	cursor.execute('''select ano,mes,empresa,conta,costcenter,profit,client,amount from tg_upload where anomes=? order by mapa,conta,costcenter,profit,client''',(wanomes,))
	linhas= cursor.fetchall()
	with open('tg_'+str(wanomes)+'.csv', 'wb') as csvfile:
		escritor = csv.writer(csvfile, delimiter=';', quotechar='|', quoting=csv.QUOTE_MINIMAL)
		for row in linhas:
			escritor.writerow((row[0],row[1],row[2],row[3],row[4],row[5],row[6],locale.format('%.2f',row[7])))
	print 'Escrita: tg_'+str(wanomes)+'.csv'



# inicio
# gestão de parametros
parser = argparse.ArgumentParser(description='Gera ficheiro TAGETIK')
parser.add_argument('integers', metavar='WANOMES', type=int,  help='O ano e mês de trabalho no formato YYYYMM')
args = parser.parse_args()
wanomes = int(args.integers)

# leitura ficheiro
db = sqlite3.connect(':memory:');
cur = db.cursor()
wb = open_workbook('tg_data.xlsx')

cria_tabelas(cur)

loads(wb,'TG_PERCENT','TG_PERCENT',13,cur)
loads(wb,'TG_COSTCENTER','TG_COSTCENTER',3,cur)
loads(wb,'TG_IMPUT','TG_IMPUT',5,cur)


# CONTAS
# NetSales (31210)
# Stock Inicial (32740)
# Stock Final (32730)
# Cust Merc Vend (32207)
# Marketing Costs (32307 32510 32520 32314 32325 32624)
# G&A costs (32624)
cur.execute(''' select conta, costcenter, profit, client_dsc, sum(valor*sinal) svalor 
	from TG_PERCENT where ANOMES = ?
	group by conta, costcenter, profit, client_dsc order by 1''',(wanomes,))
linhas= cur.fetchall()
campos = '?,'*10
for row in linhas:
	values=[]
	values.extend((wanomes,'ISOLADAS',str(wanomes)[:-2],str(wanomes)[4:],'SZP',row[0],row[1],row[2],row[3],round(row[4],2)))
	cur.execute('insert into tg_upload values ('+campos[:-1]+')',values)

db.commit()

# SALES COST \ TRANSPORT
# Distribuídos em função de uma percentagem dos valores da conta 31210 (netsales)
cur.execute('''select mapa, conta, valor, imput from tg_imput where anomes = ?''',(wanomes,))
cur2 = db.cursor()
linhas= cur.fetchall()
campos = '?,'*10
for row in linhas:
	# correr as contas
	wmapa = row[0]
	wconta = row[1]
	wvalor = row[2]
	wimput = row[3]
	
	if wvalor != 0:
		# só executa quando há valor para distribuir
		
		cur2.execute('''select b.costcenter, a.profit, a.client, a.total+0.0, a.stotal+0.0, a.contratosgrp+0.0, a.salescostpessoal+0.0
				from tg_percent a, tg_costcenter b
				where a.anomes = ?
				   and a.conta = 31210
				   and b.client = a.client
				   and b.mapa = ? ''',(wanomes,wmapa))
		linhas2 = cur2.fetchall()
		for row2 in linhas2:
			wamount = 0
			if wimput == 'total':
				wamount = round(wvalor* row2[3],2)
			elif wimput == 'stotal':
				wamount = round(wvalor* row2[4],2)
			elif wimput == 'contratosgrp':
				wamount = round(wvalor* row2[5],2)
			else:
				wamount = round(wvalor* row2[6],2)
			values=[]
			values.extend((wanomes,wmapa,str(wanomes)[:-2],str(wanomes)[4:],'SZP',wconta,row2[0],row2[1],row2[2],wamount))
			cur.execute('insert into tg_upload values ('+campos[:-1]+')',values)
			
		db.commit()

		# verificar acerto do valor
		wamount = 0
		cur2.execute('''select sum(amount) from tg_upload
				where anomes = ? and mapa = ? and conta = ? ''',(wanomes,wmapa,wconta))
		linhas2 = cur2.fetchall()
		wamount = linhas2[0][0]
		
		# se os valores forem distintos
		if round(wamount,2) != round(wvalor,2):


			# e apenas vamos acertar valores ate um euro
			if abs(round(wamount,2) - round(wvalor,2)) < 1:

				# há diferença entre o valor dividido e o valor total, vamos fazer update no mais alto
				cur2.execute('''select mapa,conta,profit,costcenter,client,amount
						from tg_upload where anomes = ? and mapa = ? and conta = ? order by amount desc'''
						,(wanomes,wmapa,wconta))
				linhas2 = cur2.fetchall()
				#não correr o cursor, ficamos pela primeira linha que é a de valor mais elevado
				
				cur2.execute(''' update tg_upload set amount = round(amount+?,2)
							where anomes = ? and mapa = ? and conta = ?
							   and profit = ? and costcenter = ? and client = ? ''',
						(round(wvalor-wamount,2),wanomes,linhas2[0][0],linhas2[0][1],linhas2[0][2],linhas2[0][3],linhas2[0][4]))
				db.commit()

			else:
				# quando o acerto e superior a um euro nao fazemos e alertamos
				print "Acerto de VALOR >= 1 (ALERTA ERRO) ", wmapa, wconta, "{:.7f}".format(wamount), "{:.7f}".format(wvalor)


		# verificar acerto do valor (TESTE)
		wamount = 0
		cur2.execute('''select sum(amount) from tg_upload
				where anomes = ? and mapa = ? and conta = ? ''',(wanomes,wmapa,wconta))
		linhas2 = cur2.fetchall()
		wamount = linhas2[0][0]
		if round(wamount,2) != round(wvalor,2):
			print wmapa, wconta, "{:.7f}".format(wamount), "{:.7f}".format(wvalor)

# produzir o output
gerafich(cur)
