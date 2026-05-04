import xlsxwriter
import pyodbc

def pedeSenha():
    usuario = input("Digite seu usuário: ")
    senha = input("Digite sua senha: ")

    try:
        conexao = conectaServidor(usuario, senha)
        return True, conexao

    except Exception as erro:
        print("Falha na conexão. Verifique usuário e senha.")
        print("Erro:", erro)
        return False, None


def conectaServidor(usuario, senha):
    conexao = pyodbc.connect(
    f"DRIVER={{ODBC Driver 17 for SQL Server}};"
    f"SERVER=10.10.4.14,61433;"
    f"DATABASE=analise_recursos;"
    f"UID={usuario};"
    f"PWD={senha}"
    )
    print("Conectado com sucesso!")
    return conexao

def criaExcel(nome_arquivo):
    workbook = xlsxwriter.Workbook(nome_arquivo)
    return workbook

def main():
    print("Programa iniciado")

    conectado = False
    conexao = None

    while conectado is not True:
        conectado, conexao = pedeSenha()

    


    workbook = criaExcel("RelCust.xlsx")

    
    conexao.close()
    workbook.close()

if __name__ == "__main__":
    main()