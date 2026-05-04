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

def adicionaAbas(hierarquia, workbook):
    vetor_hierarquia = hierarquia.split(",")

    abaMenu = workbook.add_worksheet("MENU")
    editarMenu(abaMenu, workbook, hierarquia)

    for item in vetor_hierarquia:
        print("Criando aba " + item)
    
        if item.upper() == "PAIS":
            abaPais = workbook.add_worksheet("PAIS")
    
        elif item.upper() == "ESTADO":
            abaEstado = workbook.add_worksheet("ESTADO")
    
        elif item.upper() == "REGIONAL":
            abaRegional = workbook.add_worksheet("REGIONAL")
    
        elif item.upper() == "MUNICIPIO":
            abaMunicipio = workbook.add_worksheet("MUNICIPIO")
    
        elif item.upper() == "ESCOLA":
            abaEscola = workbook.add_worksheet("ESCOLA")
    
        elif item.upper() == "TURMA":
            abaTurma = workbook.add_worksheet("TURMA")


def editarMenu(abaMenu, workbook, hierarquia):
    vetor_hierarquia = [item.strip() for item in hierarquia.split(",")]
    abaMenu.hide_gridlines(2)
    abaMenu.set_column('A:Z', 12)

    # =========================
    # FORMATOS
    # =========================
    formato_titulo = workbook.add_format({
        'bold': True,
        'font_size': 11,
        'align': 'center',
        'bg_color': '#FFE6B3',
        'valign': 'vcenter'
    })

    formato_bloco_claro = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#F3F3F3'
        
    })

    formato_bloco_muito_claro = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#FFFFFF'
        
    })

    formato_rodape = workbook.add_format({
        'bg_color': '#FCD904'
        
    })

    formato_link_claro = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#F3F3F3',
        'font_color': 'blue',
        'underline': 1
    })

    formato_link_escuro = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#E6E6E6',
        'font_color': 'blue',
        'underline': 1
    })

    formato_branco = workbook.add_format({
        'bg_color': '#FFFFFF',
        'border': 0
    })

    for linha_branca in range(0, 200):
        abaMenu.set_row(linha_branca, 15)  
        for col in range(0, 26):  
            abaMenu.write_blank(linha_branca, col, '', formato_branco)

    # =========================
    # TAMANHO DAS COLUNAS
    # =========================
    abaMenu.set_column('A:A', 2)
    abaMenu.set_column('B:H', 9)
    abaMenu.set_column('I:I', 25)
    

    # =========================
    # ESPAÇOS SUPERIORES
    # =========================
    abaMenu.merge_range('B2:I4', '', formato_bloco_claro)
    abaMenu.insert_image('B2', 'logoCaed.png')
    abaMenu.merge_range('B5:I5', '', formato_rodape)
    abaMenu.insert_image('I2', 'img_relatorio.png')
    # =========================
    # TÍTULO PRINCIPAL
    # =========================
    titulo = f"{vetor_hierarquia[0].upper()}S - 1ª AVALIAÇÃO DE FLUÊNCIA 2026"
    abaMenu.merge_range('B6:I7', titulo, formato_titulo)

    # =========================
    # BLOCOS DE NAVEGAÇÃO
    # =========================
    linha = 8

    for i, item in enumerate(vetor_hierarquia):

        nome = item.capitalize()

        # Alterna cor como no modelo
        if i % 2 == 0:
            formato_esquerda = formato_bloco_claro
            formato_direita = formato_link_escuro
        else:
            formato_esquerda = formato_bloco_muito_claro
            formato_direita = formato_link_claro

        # Texto da esquerda
        abaMenu.merge_range(f'B{linha}:G{linha}', f'Nível {nome}', formato_esquerda)

        # Link da direita
        abaMenu.merge_range(
            f'H{linha}:I{linha}',
            f'Monitoramento por {nome}',
            formato_direita
        )

        # Link interno para a aba correspondente
        abaMenu.write_url(
            f'H{linha}',
            f"internal:'{nome.upper()}'!A1",
            formato_direita,
            string=f'Monitoramento por {nome}'
        )

        linha += 1

    # =========================
    # RODAPÉ AMARELO
    # =========================
    abaMenu.merge_range(f'B{linha}:I{linha}', '', formato_rodape)  

def criarRelatorio(hierarquia):
    ##aqui a ideia é chamar as funções para criar o relatório, como se fosse uma main

    #conectado = False
    #conexao = None

    #while conectado is not True:
    #    conectado, conexao = pedeSenha()

    workbook = criaExcel("RelCust.xlsx")
    adicionaAbas (hierarquia, workbook)
    
    #conexao.close()
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
                print("1 - Sair")
                print("2 - Inserir outra hierarquia")
                print("3 - Estado, Regional, Municipio, Escola, Turma")
                print("4 - Municipio, Escola, Turma")
                print("5 - Municipio, Regional, Escola, Turma")
                print("6 - Pais, Estado, Regional, Municipio, Escola, Turma")
                
                opcaoHierarquia = input("Escolha uma opção: ")

                if opcaoHierarquia == "3":
                    criarRelatorio("Estado,Regional,Municipio,Escola,Turma")
                    break

                elif opcaoHierarquia == "4":
                    criarRelatorio("Municipio,Escola,Turma")
                    break

                elif opcaoHierarquia == "5":
                    criarRelatorio("Municipio,Regional,Escola,Turma")
                    break
                    
                elif opcaoHierarquia == "6":
                    criarRelatorio("Pais,Estado,Regional,Municipio,Escola,Turma")
                    break    

                elif opcaoHierarquia == "2":
                    hierarquia = input("Escreva a hierarquia, separando por virgula e sem espaços. (Ex: Municipio,Escola,Turma)")
                    criarRelatorio(hierarquia)
                    break

                elif opcaoHierarquia == "1":
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