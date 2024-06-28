import openpyxl
from PIL import Image, ImageDraw, ImageFont

# essa bosta carrega a poha do arquivo do excel
try:
    planilha_arquivo = openpyxl.load_workbook('Teste_emissao_certificado.xlsx')
except PermissionError as e:
    print(f"Erro: Permissão negada ao tentar abrir o arquivo. Detalhes: {e}")
    exit(1)

# aqui mescolhe a planilha desejada, mesmo tendo só uma coloca, deu uma puta dor de cabeça sem hahahaha
try:
    planilha_alunos = planilha_arquivo['Respostas']
except KeyError as e:
    print(f"Erro: A planilha 'Respostas' não foi encontrada no arquivo. Detalhes: {e}")
    exit(1)

# Aqui começa a ler da segunda linha por qaue a primeira é o tidulo da coluna
for indice, linha in enumerate(planilha_alunos.iter_rows(min_row=2)):

    emissao_certificado = linha[0].value
    nome_aluno = linha[1].value
    carga_horaria = linha[5].value
    professor = linha[6].value
    nome_curso = linha[7].value

    # busca a font, lembre-se de ficar no mesmo diretorio se não vai ter que mostrar
    fonte_nome = ImageFont.truetype('./tahomabd.ttf', 90)
    fonte_geral = ImageFont.truetype('./tahoma.ttf', 80)
    fonte_data = ImageFont.truetype('./tahoma.ttf', 55)

    # busca certificado
    image = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)

    # posiçoes do texto esse foi complicado kkkk
    desenhar.text((1020, 827), nome_aluno, fill='black', font=fonte_nome)
    desenhar.text((1060, 950), nome_curso, fill='black', font=fonte_geral)
    desenhar.text((1435, 1065), professor, fill='black', font=fonte_geral)
    desenhar.text((1480, 1182), str(carga_horaria), fill='black', font=fonte_geral)

    # convete a coluna (emissao_certificado) para str 
    desenhar.text((2220, 1930), emissao_certificado.strftime('%d/%m/%Y'), fill='blue', font=fonte_data)

    # renomear com nome do aluno e indice
    image.save(f'./{indice} {nome_aluno} certificado.png')
