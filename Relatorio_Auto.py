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

def criarRelatorio():
    ##aqui a ideia é chamar as funções para criar o relatório, como se fosse uma main

    conectado = False
    conexao = None

    while conectado is not True:
        conectado, conexao = pedeSenha()
    


    workbook = criaExcel("RelCust.xlsx")

    
    conexao.close()
    workbook.close()

    return


def main():
    
    menu()



def menu():
    while True:
        print("===== MENU PRINCIPAL =====")
        print("1 - Criar relatório")
        print("2 - Sair")

        opcao = input("Escolha uma opção: ")

        if opcao == "1":

            while True:
                print("Selecione a hierarquia do projeto:")
                print("1 - Estado, Regional, Municipio, Escola, Turma")
                print("2 - Municipio, Escola, Turma")
                print("3 - Municipio, Regional, Escola, Turma")
                print("4 - Inserir outra hierarquia")
                print("5 - Sair")
                opcaoHierarquia = input("Escolha uma opção: ")

                if opcaoHierarquia == "1":
                    criarRelatorio("Estado,Regional,Municipio,Escola,Turma")
                    break

                elif opcaoHierarquia == "2":
                    criarRelatorio("Municipio,Escola,Turma")
                    break

                elif opcaoHierarquia == "3":
                    criarRelatorio("Municipio,Regional,Escola,Turma")
                    break

                elif opcaoHierarquia == "4":
                    hierarquia = input("Escreva a hierarquia, separando por virgula e sem espaços. (Ex: Municipio,Escola,Turma)")
                    criarRelatorio(hierarquia)
                    break

                elif opcaoHierarquia == "5":
                    print("Saindo")
                    break

                else:
                    print("Opção inválida. Tente novamente.\n")

            

        elif opcao == "2":
            print("Encerrando programa...")
            break

        else:
            print("Opção inválida. Tente novamente.\n")

if __name__ == "__main__":
    main()