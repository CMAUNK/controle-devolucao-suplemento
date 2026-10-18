# Suplemento do Excel — Controle Devolução

Este suplemento importa no Excel as linhas TSV exportadas pelo aplicativo Controle Devolução. Ele tem destinos para BASE-RECEBENDO, BASE-EXPEDINDO e BASE 100km.

## Uso

1. No aplicativo, baixe o arquivo TSV da base RECEBENDO, EXPEDINDO ou 100km.
2. No painel do Excel, selecione o mesmo destino e o TSV correspondente.
3. Confira a prévia dos dados e clique em **Conferir antes de inserir**. O painel confere a base selecionada, avisa sobre linhas repetidas, estima o uso de linhas vazias e identifica colunas com fórmulas.
4. Revise o resumo. Se houver duplicados, escolha **Inserir somente novos**, **Inserir todos** ou **Cancelar**. Sem duplicados, confirme em **Confirmar e inserir**.
5. Confira o resumo final: linhas inseridas, linhas vazias reutilizadas, novas linhas criadas e fórmulas detectadas.
6. Use **Limpar arquivo** para remover a seleção e a prévia do painel; isso não apaga linhas da planilha.

O botão **Confirmar e inserir** pula a tela de conferência. Ele insere somente registros novos e ignora os repetidos automaticamente; o resumo final informa quantas linhas foram ignoradas.

O destino 100km é a aba `BASE`, identificada pelo cabeçalho `SM`. Motorista e ajudante podem ser importados em linhas separadas com a mesma SM; o suplemento verifica duplicidade pelo par SM + CPF, de modo que pessoas diferentes da mesma viagem não sejam tratadas como duplicadas. O suplemento associa dados pelos cabeçalhos e escreve na pasta de trabalho ativa pelo Office JavaScript API. Ele não envia o conteúdo do TSV ou da planilha para um servidor. O código hospedado aqui é público; não inclua arquivos de dados, PDFs, credenciais ou tokens neste repositório.

O suplemento reutiliza linhas vazias já existentes na tabela antes de criar linhas novas. As linhas preenchidas herdam o formato visual e o alinhamento da linha preenchida mais próxima (fonte, tamanho, alinhamento, bordas, cores e formatos numéricos). Na planilha 100km, o nome do motorista é apresentado em formato de nome e o CPF é mantido como texto, com 11 dígitos, pontuação e alinhamento à esquerda. Ele detecta colunas com fórmulas e as preserva nas linhas reutilizadas; ao criar novas linhas, copia as fórmulas existentes da tabela para manter os cálculos com referências ajustadas pelo Excel. Confira o destino e as duplicatas antes de confirmar a importação.

No destino `BASE 100km`, o TSV inclui `Janta`, mas não inclui `Matrícula` nem `Total`. O suplemento pode preencher Janta quando ela tiver um valor; Matrícula não é importada e Total é preservado ou copiado como fórmula da tabela.

Ao inserir na `BASE 100km`, o suplemento mantém KM numérico e aplica o formato visual `Km` (por exemplo, `485 Km`) para que as fórmulas da planilha continuem funcionando.
