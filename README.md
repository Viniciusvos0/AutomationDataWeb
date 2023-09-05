# AutomationDataWeb
Solução de automatização web com scraping e load no networkDirectory.
Esta solução em python tem como objetivo extrair dados de um site intranet 
de uma empresa.
Inicialmente ele abre o navegador web e faz todas as autenticações necessárias
inclusive no autenticador de segurança do windows(mediante fornecimento do 
usuário e senha no código). Devido a solicitação de autenticação de segurança 
do windows não foi possível utilizar o modo headles, portanto foram utilizados
outros métodos para ocultar a automatização no navegador.
Por fim este script busca, cria e personaliza o nome das pastas em que deve salvar
os arquivos extraídos.
Para esta solução foram utilizadas as seguintes ferramentas:


os: Módulo Python.

config:módulo de configurações Python

datetime: Módulo Python.

pandas: Biblioteca Python

sys: Módulo Python.

selenium: Framework

time: Módulo Python.

autoit: Biblioteca externa

openpyxl: Biblioteca externa

locale: Módulo Python.

pyautogui: Biblioteca externa
