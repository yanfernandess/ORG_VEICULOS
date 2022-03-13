## PROJETO ORGANIZAÇÃO DE VEICULOS

Proposta

Estruturar o desenvolvimento de um sistema,
passo a passo, detalhando as etapas e decisões que foram tomadas.
Explicar a razão das decisões, tecnologias escolhidas, etc. Apresentar tudo o que fez, um levantamento das dificuldades, resultados, diagramas, códigos, etc.

Desafio

Uma empresa possui diversos carros alugados disponíveis para seus funcionários,
porém a gerência destes carros não é bem administrada. Isto leva a problemas como pessoas pagarem carros diferentes para o mesmo lugar,
ou pessoas indo a pé com carro disponível. Este problema foi apresentado à área de TI, o que fazer para solucioná-lo?


--------------------------Sobre a proposta-----------------------------------------------------------

Quando o Tiago me passou a proposta e o desafio para realizar este projeto eu sabia que não seria fácil por não ter experiência como desenvolvedor, então
comecei a estudar sobre e consultar pessoas que já vivenciaram esse desafio e saber por onde eu começaria e cheguei a conclusão que iria começar com um 
rascunho no caderno fazendo um diagrama e escrevendo todas as minhas ideias, como eu acho que resolveria o problema de Aluguel vs Funcionários então saiu as primeiras ideias 
e comecei a desenvolver em java mas como eu vi que o desafio também pedia poucos dias e eficiência decidi escolher o Access e VBA para desenvolver as tabelas,
formulários, banco de dados e etc. Pois no Access o desenvolvimento de Tabelas e formulários(telas) não precisam de código para cria-las só no intermediário e no final
do projeto eu iria fazer as Querys e desenvolver algo para os campos e botões, algo que eu demoraria em torno de 2 semanas em JAVA ou qualquer outra linguagem que eu sei
consegui resolver em 4 dias (claro que tive que estudar e questionar muito sobre como realizar mas consegui terminar sozinho. Acho que minha maior dificuldade foi aprender 
a questão de códificação do VBA pois é muito diferente do que estava acostumado a realizar e foi um desafio muito bom e de muito aprendizado, tentei dar meu melhor.


--------------------------Sobre o desafio------------------------------------------------------------


Criei um sistema que tem 7 telas:

Login
Menu principal
Aluguel
Pagamentos
Cadastro de Locadoras
Cadastro de veículos
Cadastro de Funcionários

As telas de cadastro são as principais que fazem o sistema ser alimentado e é por ai que vou organizar em quantas locadoras nossa empresa costuma ter contato para o aluguel,
quantos veículos temos disponíveis para uso dos nossos funcionários e por ultimo os funcionários, e também temos a tela de pagamentos onde você vai realizar o cadastro quando 
você alugou o veículo. Todos esses cadastro levam a tela principal onde você escolhe o nome do veículo/placa, funcionário/área e a data que foi alugado pelo funcionário que seria
o check-in ai ele registra na tabela de aluguel e nessa mesma tela foi criado um campo de pesquisa/informação, neste campo informa quais os veículos que estão alugados e em poder
e nos combos criados nos campos dessa tela vão ficar os disponíveis para ser alugados. Quando for feita a devolução do veículo você tem que dar double clique em um veículo da parte
de pesquisa da tela_aluguel e dar um clique no botão salvar ai ele registra que foi devolvido e fica disponivel para uso. Então a solução foi que devemos usar esse sistema
constante para que tenha a informação de quantos veículos temos para uso e quantos estão sendo usados isso organiza tudo novamente.


----------------------------- Sobre a codificação usada---------------------------------------------

Vou citar abaixo sobre as codificações usadas nas telas.


FRM_ALUGUEL

Criei dois Select para buscar todos os veículos disponíveis e os funcionários que não alugaram veículos ainda, então os dois campos são principalmente os dois selects e mais a Query que vou citar abaixo
o campo data check in eu criei um codigo para buscar o dia que está sendo feito o aluguel do veículo, travei os últimos dois campos para que nada fosse preenchido manualmente e sim automático. Também criei
um código para que limpasse os campos assim que utilizados e no final de tudo fiz um MSGBOX(uma tela que aparece uma mensagem no final que eu salvasse o check_out).

FRM_FUNCIONARIO
FRM_LOCADORA
FRM_PAGAMENTOS
FRM_VEICULOS

Criei um código para ele adicionar o cadastro que foi feito e também fiz um código para limpar os campos assim que fosse clicado no botão de cadastrar, criei um select para conversar com as tabelas e 
registrar, criei também um MSGBOX cadastrado com sucesso.

FRM_LOGIN

Criei um select para buscar na tabela de Login se tem o funcionário cadastrado, mesma coisa de todos os outros fiz o empty que seria para limpar todos os campos após cadastrados
e no final o MSGBOX de Seja Bem-vindo !!!

FRM_MENU_PRINCIPAL

Criei para todos os botoes um OpenForm que seria para abrir a tela selecionada.




