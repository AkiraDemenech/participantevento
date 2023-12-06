perfil_0 = 'Participante Simples'
perfil_1 = 'Participante + 1 Minicurso'
perfil_2 = 'Participante + 2 Minicursos'

arquivo = 'Inscritos CPCB 2024.xlsx - Geral.csv'
fila_simultanea = 10

latex_declara = '''\\documentclass[12pt]{article}
	\\usepackage[a4paper,landscape]{geometry}
	\\usepackage[T1]{fontenc}
	\\usepackage{ragged2e}'''

latex_abre = '''
\\begin{document}
	\\thispagestyle{empty} %% sem número de página
	\\sffamily %% sem serifa
	\\justifying %% justificado
	\\hyphenpenalty=0 %% evitar hifenização	
	\\hfill \\vfill %% centralizar verticalmente
	
	
	\\centerline{\\Huge Declaração}
	\\large
	
	\\vfill
	%\\bigskip
''' 	

latex_fecha = '''
	\\vfill
	%\\bigskip
	
	\\normalsize
	Assinaturas
	
	\\vfill \\hfill %% centralizar verticalmente

\\end{document}'''

modelo_participante = '''
Declaramos para os devidos fins que
	%s
	participou do 
	\\textbf{12° Congresso Paranaense de Ciências Biomédicas},
	%%promovido pelo curso de biomedicina
	entre os dias 
	13 e 15 de março de 2024,
	como OUVINTE'''
modelo_apresentador = '''
	e APRESENTADOR(A) do trabalho intitulado
	%s'''	
modelo_horas = ''',
	totalizando 
	%s 
	horas de duração'''	
modelo_minicurso = ''',
incluindo o minicurso
	"%s"'''
modelo_minicursos = ''',
incluindo os minicursos
	"%s"
	e 
	"%s"'''	


import csv


planilha = [[col.strip() for col in ln] for ln in csv.reader(open(arquivo, 'r', encoding='utf-8'))]
colunas = planilha.pop(0)    

col_cidade = colunas.index('Cidade')
col_cpf = colunas.index('CPF')
col_rg = colunas.index('RG')
col_telefone = colunas.index('Telefone')
col_email = colunas.index('E-mail')
col_nome = colunas.index('Nome')
col_resumo = colunas.index('Resumo')
col_perfil = colunas.index('Perfil')
col_minicurso_1 = colunas.index('Minicurso 1')
col_minicurso_2 = colunas.index('Minicurso 2')
col_minicurso_conferido = colunas.index('Minicurso conferido')

def algarismos (t):
	a = ''
	for d in t:
		a += d * d.isdigit()
	return a	
import threading
import time
import os

fila = []
fila_sem = threading.Semaphore()
fila_livre = threading.Semaphore(fila_simultanea)

def produtor (arq):
	fila_sem.acquire()
	fila.append(arq)
	fila_sem.release()
def consumidor ():
	while True:		
		fila_sem.acquire()
		while len(fila):
			time.sleep(1)
			threading.Thread(target=consumir, args=[fila.pop()]).start()
		fila_sem.release()	
def consumir (arq):
	fila_livre.acquire()
	print(os.system(f'pdflatex.exe -synctex=1 -interaction=nonstopmode "{arq}"'), '\t', arq)
	fila_livre.release()
threading.Thread(target=consumidor, daemon=True).start()	

for ln in planilha:
	cidade = ln[col_cidade].upper()
	nome = ln[col_nome].upper()
	email = ln[col_email].lower()
	
	telefone = algarismos(ln[col_telefone])
	
	cpf = algarismos(ln[col_cpf])
	rg = algarismos(ln[col_rg])

	perfil = ln[col_perfil]
	resumo = ln[col_resumo]
	
	print(cpf,rg,'\t',nome,email,'\t',perfil)

	horas = 0
	tipo = -1
	if perfil_0 in perfil:
		tipo = 0
	elif perfil_1 in perfil:
		tipo = 1
	elif perfil_2 in perfil:
		tipo = 2	
	else:	
		print('Perfil não identificado\n')

	with open(f'{tipo} {nome} {cpf} {rg}.tex', 'w', encoding='utf-8') as tex:
		print(latex_declara, file=tex)
		print(latex_abre, file=tex)
		print(end=(modelo_participante % nome), file=tex)
		if len(resumo) > 1:
			print(end=(modelo_apresentador % resumo), file=tex)
		print(end=(modelo_horas % str(horas)), file=tex)

		if tipo == 1:
			print(end=(modelo_minicurso % ln[col_minicurso_1]), file=tex)
		elif tipo == 2:
			print(end=(modelo_minicursos % (ln[col_minicurso_1], ln[col_minicurso_2])), file=tex)	

		print('.\n\n', latex_fecha, '\n%%', cidade, '\n%%', email, '\n%%', telefone, file=tex)

	
		a = tex.name

	produtor(a)
input('Enter para encerrar')	