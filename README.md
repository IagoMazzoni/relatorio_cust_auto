# relatorio_cust_auto
A ideia é criar um código bem fácil de manutenir, atualizar e bem genérico para gerar o relatório de monitoramento

Até o breve momento o código é capaz de
1- fazer a conexão com o servidor 14
2- cria um relatório .xlsx vazio
3- obviamente mas não menos importante, fecha a conexão com o banco e com o relatório pre aberto

Bibliotecas usadas:
1- pyodbc (pip install pyodbc)-> usada para fazer conexão com o banco 
2- xlsxwriter (pip install xlsxwriter)-> usada para criar e editar o arquivo xlsx. um ponto importante sobre essa biblioteca é que ela não é capaz de editar um excel ja pronto, ela sempre cria um novo, por isso o código cria e abre o excel, edita e depois fecha tudo de uma vez. Optei por ela por causa das funcionalidades de edição e formatação, talvez eu me arrependa num futuro, não tive paciencia para ler a documentação toda.