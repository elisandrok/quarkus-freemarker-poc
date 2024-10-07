API de Geração de Relatórios - Quarkus, Apache POI, FreeMarker e LibreOffice

Esta POC demonstra a criação de uma API para geração de relatórios utilizando Quarkus, Apache POI, FreeMarker e LibreOffice em modo headless. A API permite o upload de arquivos .odt ou .docx, faz a substituição de variáveis no documento usando FreeMarker, e converte o resultado final para PDF usando o LibreOffice sem interface gráfica.
Tecnologias Utilizadas:

    Quarkus: Framework Java otimizado para microsserviços.
    Apache POI: Manipulação de documentos .docx.
    FreeMarker: Motor de templates para substituir variáveis nos documentos.
    LibreOffice (Headless): Conversão de arquivos .odt e .docx para PDF.

Funcionalidades:

    Recebe templates .docx.
    Substitui variáveis dentro dos documentos usando dados dinâmicos.
    Retorna o documento gerado em formato PDF.

Execução

Para executar a aplicação em modo de desenvolvimento, use o seguinte comando:

./mvnw quarkus:dev

Documentação

A API está documentada e pode ser acessada via Swagger UI em: <http://localhost:8080/swagger-ui/>
