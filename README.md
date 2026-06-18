Gerenciador de Relatórios: 
Sistema em Python para controle, priorização e arquivamento de relatórios de faturamento. 

Funcionalidades 
Ordenação Automática: Organiza a fila com os faturamentos mais próximos no topo. 
Alertas por Cores: 
🔴 Crítico: Faturamento em até 3 dias. 
🟡 Atenção: Faturamento em até 7 dias. 
Controle de Status: Botões para marcar como Feito (Pronto) ou enviado (arquivar). 
Histórico com Busca: Filtro rápido por Empresa, Contrato ou Código. 
Exportação: Gera planilhas Excel (.xlsx) do histórico de entregas. 

Requisitos e Instalação: 
Ter instalado o VSCode com as extensões de python: LiveServer, Pylance, Python, Debugger, Environments, Indent, Path, Ruff; 
Executar no terminal VSCode: bash pip install customtkinter pandas openpyxl 
A Linguagem python no computador; 

Bibliotecas: 
Visual:
customtkinter   5.2.2 
darkdetect      0.8.0 
Exportação:
et_xmlfile      2.0.0 
openpyxl        3.1.5 
Banco de dados:
numpy           2.4.6 
pandas          3.0.3 
python-dateutil 2.9.0.post0 
tzdata          2026.2 
Gerenciamento:
packaging       26.2 
pip             26.1.2 
six             1.17.0

ESCOPO

Etapa 1: Base de Dados e Estrutura: 
Criar a memória do sistema: 
Configuração do banco de dados invisível (SQLite3) para que as informações fiquem salvas. 
Criação da Fila: 
Um cálculo automático que analisa a data de hoje e  quais empresas estão mais próximas do dia de faturamento, montando assim a lista de prioridade. 

Etapa 2: Front-End: 
Design de tela com o CustomTkinter 
Alertas por Cores:
Linhas que ficam vermelhas automaticamente se o faturamento for em até 3 dias, ou amarelas se for em até 7 dias. 
Botão Duplo (Double Check):
Separação do trabalho em dois cliques para evitar erros: um botão para registrar que o relatório está feito e outro para registrar que ele foi enviado para o cliente. 

Etapa 3: Busca e Prestação de Contas: 
Barra de pesquisa que filtra o histórico. 
Exportação para Excel: Um botão que pega todo o seu histórico de entregas concluídas e transforma em uma planilha. 
