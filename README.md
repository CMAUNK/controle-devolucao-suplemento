# Suplemento do Excel — Controle Devolução

Este suplemento importa no Excel as linhas TSV exportadas pelo aplicativo Controle Devolução.

## Uso

1. No aplicativo, baixe o arquivo TSV da base RECEBENDO ou EXPEDINDO.
2. No painel do Excel, selecione a mesma base e o TSV correspondente.
3. Confira a prévia e clique em **Adicionar linhas na tabela**.
4. Se algum manifesto já existir ou se repetir no arquivo, o suplemento mostra um aviso e permite cancelar ou confirmar a inclusão das duplicatas.
5. Use **Limpar arquivo** para remover a seleção e a prévia do painel; isso não apaga linhas da planilha.

O suplemento associa dados pelos cabeçalhos e escreve na pasta de trabalho ativa pelo Office JavaScript API. Ele não envia o conteúdo do TSV ou da planilha para um servidor. O código hospedado aqui é público; não inclua arquivos de dados, PDFs, credenciais ou tokens neste repositório.

O suplemento reutiliza linhas vazias já existentes na tabela antes de criar linhas novas. Detecta colunas com fórmulas e preserva as fórmulas já existentes; quando adiciona linhas novas, copia as fórmulas da tabela para manter os cálculos com referências ajustadas pelo Excel. Confira o destino e as duplicatas antes de confirmar a importação.
