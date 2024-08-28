Automação para Criação de Assinatura de Email:

Visão Geral

Este projeto automatiza a criação de assinaturas de email personalizadas utilizando Python, com as bibliotecas openpyxl para manipulação de planilhas Excel e Pillow para processamento de imagens. O script gera imagens de assinaturas contendo nome, cargo e ramal de cada colaborador, prontos para uso em emails.

Requisitos

Planilha Excel (dados.xlsx):

Deve conter os dados como nome, cargo e ramal nas colunas apropriadas.
A planilha deve estar localizada no mesmo diretório do script (app.py).

Fonte Personalizada:

As fontes utilizadas no design da assinatura precisam estar disponíveis no diretório do projeto.
Exemplo: Scansky Condensed Semi Bold.ttf, Scansky Condensed Light.ttf.

Modelo de Assinatura de Email:

Imagem base (assinatura_email.png) em formato PNG ou JPG, que servirá de fundo para as assinaturas.
A imagem deve estar localizada no mesmo diretório do script.

Como Funciona?

Carregamento da Planilha:

O script carrega os dados de dados.xlsx e percorre cada linha, extraindo nome, cargo e ramal.

Carregamento das Fontes:

As fontes são carregadas com tamanhos específicos, que podem ser ajustados conforme necessário para se adequar ao design da assinatura.

Processamento de Imagem:

A biblioteca Pillow é usada para desenhar o texto sobre a imagem base, posicionando nome, cargo e ramal de acordo com as coordenadas fornecidas.

Geração e Salvamento:

Cada assinatura é gerada e salva na pasta assinaturas/ com o nome do funcionário seguido de _cartao.png.

Personalização e Ajustes

Coordenadas: As coordenadas dos textos na imagem podem ser ajustadas diretamente no código para alinhar melhor com o modelo de assinatura.

Tamanho da Fonte: Pode ser ajustado conforme necessário para combinar com o design da assinatura.

Execução

Certifique-se de que a planilha dados.xlsx, a imagem assinatura_email.png e as fontes estejam no diretório do projeto.

Execute o script app.py.

As assinaturas serão geradas automaticamente e salvas na pasta assinaturas/.

Considerações Finais

Esta automação torna o processo de criação de assinaturas de email personalizado rápido e consistente, ideal para empresas que desejam padronizar as comunicações visuais internas.
