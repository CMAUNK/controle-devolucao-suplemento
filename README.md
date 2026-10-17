# Suplemento do Excel — Controle Devolução

Este suplemento permite importar para uma tabela do Excel as linhas TSV exportadas pelo aplicativo Controle Devolução.

## Uso

1. No aplicativo, baixe o arquivo TSV da base RECEBENDO ou EXPEDINDO.
2. Abra a planilha de trabalho no Excel Online e carregue `manifest.xml` em **Suplementos → Carregar Meu Suplemento**.
3. No painel do suplemento, escolha a base correspondente, selecione o TSV e confira a prévia.
4. Clique em **Adicionar linhas na tabela** para inserir os registros na tabela da aba correspondente.

Os dados são lidos no navegador e escritos na pasta de trabalho ativa pelo Office JavaScript API. O suplemento não envia o conteúdo do TSV ou da planilha a um servidor. O código hospedado aqui é público; não inclua arquivos de dados, PDFs, credenciais ou tokens neste repositório.

O suplemento associa valores pelos nomes dos cabeçalhos, conserva os campos que não vierem no arquivo e aplica a fórmula de Aging nas novas linhas de BASE-RECEBENDO. Antes do uso na planilha oficial, valide o fluxo em uma cópia e confira os nomes das tabelas e colunas.
