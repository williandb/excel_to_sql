import pandas as pd
import os
import sqlalchemy
import sys
import termcolor

nomeConexaoODBC = sys.argv[1]
usuario = sys.argv[2]
senha = sys.argv[3]
'''
print(usuario)
print(senha)
print(nomeConexaoODBC)
'''
# Faz a conexão com o banco de dados
engine = sqlalchemy.create_engine("sybase+pyodbc://{}:{}@{}".format(usuario,senha,nomeConexaoODBC))
engine.connect()

# Percorre os arquivos
for arquivo in os.listdir('.'):  
  nomeArquivo, extensao = arquivo.split('.')  
  # Filtra apenas os arquivos do excel
  if extensao == 'xlsx':
    texto = 'Processando o arquivo {}.{}'.format(nomeArquivo,extensao)    
    print(texto+(" "* (40-len(texto))),end='\t')
    # Lê o arquivo para o dataframe
    df = pd.read_excel(os.path.abspath(arquivo),index_col=None)
    # Retira as colunas sem informação
    df = df[df.filter(regex='^(?!Unnamed)').columns]
    # Cria a tabela e insere os registros no banco
    engine.execute('drop table "{}"'.format(nomeArquivo))
    df.to_sql(nomeArquivo, con=engine, if_exists='replace')
    print(termcolor.colored('[OK]','green'))