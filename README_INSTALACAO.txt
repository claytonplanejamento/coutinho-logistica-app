COUTINHO LOGISTICA - V6 PWA
===========================

O QUE E ESTA VERSAO
-------------------
Esta pasta e uma PWA (Progressive Web App) real. O icone da tela inicial vem dos arquivos PNG reais do manifest, por isso o Android/Chrome consegue usar a logo da Coutinho Logistica corretamente.

IMPORTANTE
----------
A PWA precisa ser publicada em um endereco HTTPS. Abrir o index.html diretamente do celular NAO instala como PWA.
O backend continua sendo o Google Apps Script atual.

ARQUIVOS
--------
index.html             -> casca mobile instalavel
manifest.json          -> nome, cores e icones do app
service-worker.js      -> requisito de instalacao/offline shell
app-config.js          -> URL do Web App do Apps Script
icons/                 -> icones reais para Android/iOS
logo-coutinho.png      -> arte original

BACKEND CONFIGURADO
-------------------
https://script.google.com/macros/s/AKfycbw9_GypoVyYQ6qArmPOTsVXTBR1h4AoRwwwwjR4JCu4Gu3cj9TRDFtgEQX7UTaDPM3C/exec

Se voce criar uma NOVA implantacao do Apps Script e a URL mudar, edite apenas app-config.js.

PUBLICACAO GRATUITA - GITHUB PAGES
----------------------------------
1. Crie um repositorio no GitHub, por exemplo: coutinho-logistica-app
2. Envie TODO o conteudo desta pasta mantendo a pasta icons.
3. No repositorio: Settings > Pages.
4. Em Build and deployment, selecione Deploy from a branch.
5. Branch: main / root e Save.
6. O GitHub fornecera uma URL HTTPS parecida com:
   https://SEU-USUARIO.github.io/coutinho-logistica-app/
7. Abra essa URL no Chrome do Android.
8. Aguarde o botao INSTALAR APP ou use o menu do Chrome > Instalar app / Adicionar a tela inicial.

PARA TROCAR UMA VERSAO DO APPS SCRIPT
-------------------------------------
Se voce apenas cria uma NOVA VERSAO dentro da MESMA implantacao, a URL /exec permanece igual. Nao precisa alterar a PWA.

ICONE ANTIGO EM CACHE
---------------------
Se o Android continuar mostrando um icone antigo:
1. Desinstale/remova o atalho anterior.
2. Chrome > Configuracoes > Configuracoes do site > Dados armazenados (ou limpe os dados do site da URL da PWA).
3. Abra novamente a URL HTTPS da PWA.
4. Instale novamente.

OBSERVACAO
----------
O Apps Script ja usa setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL), necessario para abrir dentro desta PWA.
