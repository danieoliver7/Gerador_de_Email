# Ateliê Email — gerador local

Aplicação React para importar o copy de um DOCX, revisar o conteúdo e exportar HTML de email. O processamento acontece no navegador, sem API paga, conta ou hospedagem.

## Executar no Windows

No PowerShell, nesta pasta:

```powershell
npm install
npm run dev
```

Abra o endereço exibido pelo Vite, normalmente `http://127.0.0.1:5173/`.

## Cabeçalho automático

Por padrão, todos os emails incluem um cabeçalho visual: uma faixa preta (`#000000`) de 100 px de altura, sem imagem, link ou texto. Ele ocupa a largura interna do email e acompanha a configuração **Largura do email**, sem uma largura independente. Não é necessário incluir marcador no Word.

No painel direito, abra **Configurações > Cabeçalho** para ativar ou desativar a faixa, escolher uma cor, editar seu valor hexadecimal e ajustar a altura e os espaçamentos acima e abaixo. Os espaçamentos começam em 0 px. As mudanças aparecem na prévia **Email**, nos modos **Desktop** e **Mobile**, e no HTML exportado.

O gerador coloca o cabeçalho antes do copy. Se os primeiros parágrafos do Word começarem com os rótulos `SUB:`, `Assunto:` ou `Linha fina:`, eles permanecem antes da faixa, na ordem importada. Na estrutura atual, esses rótulos são parágrafos de texto do Word; a aplicação ainda não tem um componente SUB nem um separador visual independente. O cabeçalho fica fora das condições de audiência e não altera textos, botões ou links.

Para conferir no navegador:

1. Importe um DOCX com `SUB: Exemplo` e `Linha fina: Exemplo` nos primeiros parágrafos, seguidos pelo copy; na aba **Email**, confira a faixa depois desses dois parágrafos. Com um DOCX sem esses rótulos iniciais, confira a faixa antes do copy.
2. Em **Configurações > Cabeçalho**, mude a cor, a altura e os espaçamentos. Alterne **Desktop** e **Mobile**, desative e reative **Exibir cabeçalho** e confira a prévia.
3. Abra a aba **HTML** e exporte o arquivo. Confirme a cor, a altura e uma única ocorrência da faixa. Recarregue a página ou use **Salvar projeto JSON** e **Abrir projeto** para confirmar que os ajustes são restaurados. Projetos criados antes desta função recebem automaticamente os valores padrão do cabeçalho ao abrir.

## Testar o botão com um Word real

1. Crie um documento no Word com estes parágrafos, cada linha separada por **Enter**:

   ```text
   Olá, cliente!
   Preparamos uma oportunidade especial para você.
   %acessar%
   Continue acompanhando nossas novidades.
   %BAIXAR AGORA%
   ```

2. Selecione apenas `BAIXAR AGORA` (sem os `%`) e adicione um hyperlink no Word. Para testar fragmentação interna, aplique negrito em parte do marcador. Salve como `.docx`.
3. Importe o arquivo. Na aba **Conteúdo**, os dois botões devem aparecer nas posições dos marcadores. O segundo usa `BAIXAR AGORA` como texto e o hyperlink como destino inicial. Clique em **Editar botão** ou escolha uma ocorrência no painel direito.
4. Altere o texto e a URL de um botão. Confira que o outro continua com os valores originais. Clique em **Restaurar texto e URL iniciais** para repor os valores importados.
5. Abra **Email** para ver a prévia e **HTML** para inspecionar as tabelas e os estilos inline. **Exportar HTML** baixa o arquivo final para revisão no Adobe Journey Optimizer.
6. Recarregue a página: o projeto deve reabrir automaticamente. Você também pode usar **Salvar projeto JSON** e **Abrir projeto** para fazer backup e restaurar manualmente.

Um DOCX de exemplo com `%acessar%` e `%BAIXAR AGORA%` vinculado a uma URL de teste está em [examples/botoes-acessar.docx](examples/botoes-acessar.docx). Ele pode ser recriado com `node scripts/create-button-example.mjs`.

Qualquer marcador `%texto%` vira botão quando ocupa sozinho um parágrafo ou uma linha dentro dele, mesmo que o Word tenha separado o marcador em diferentes trechos de formatação. `%acessar%` continua usando o texto e a URL fornecidos no template original. Para outros marcadores, o texto entre `%` vira o rótulo e o hyperlink do Word vira a URL inicial. Se não houver link, ou ele estiver incompleto como `https://www/`, o botão recebe `href="#"` e pode ser exportado sem preencher um destino. Um marcador no meio de uma frase gera aviso e mantém o texto original. Uma URL personalizada não vazia precisa ser HTTP ou HTTPS com domínio completo; valores inseguros bloqueiam a exportação.

## Testar condições de audiência com um Word real

Crie um `.docx` no Word com cada marcador em um parágrafo próprio (use **Enter**):

```text
Olá, cliente!

##BASE_ATIVO##
Texto exclusivo para a base de ativos.
%acessar%
##FIM_BASE_ATIVO##

Este texto aparece para todos.

##CLIENTES_VIP##
Conteúdo VIP.
##FIM_CLIENTES_VIP##

Até breve!
```

Se preferir, importe o arquivo pronto [examples/condicoes-audiencia.docx](examples/condicoes-audiencia.docx). Ele inclui SUB, linha fina, duas audiências, um botão em cada uma, negrito, link e marcadores fragmentados. O script `node scripts/create-condition-example.mjs` recria esse arquivo.

Pode aplicar negrito ou itálico em parte dos marcadores e do texto; o Word frequentemente divide o conteúdo em vários *runs*. Ao importar, a aba **Conteúdo** mostrará os limites `[Condição: BASE_ATIVO]` e `[Fim da condição]`, com o botão dentro da primeira região. Use **Perfil pertence** e **Perfil não pertence** para simular as duas prévias. A aba **HTML** e o arquivo de **Exportar HTML** sempre conservam as expressões literais `{%#if profile.segmentMembership.get("ups").get("BASE_ATIVO").status = "existing"%}` e `{%/if%}`; a simulação não altera o arquivo exportado.

É possível usar qualquer nome de audiência com letras ASCII, números e underscores, preservando a capitalização. Os pares devem corresponder exatamente. Marcadores incompletos, no meio de texto, incompatíveis, vazios ou aninhados são exibidos como erros e bloqueiam a exportação; corrija o Word e importe novamente. **Salvar projeto JSON** e o salvamento local preservam os limites condicionais e as edições dos botões. Imagens embutidas continuam indisponíveis no importador atual e são sinalizadas com aviso, inclusive quando estiverem entre marcadores.

## Como funciona

- `src/lib/document.ts`: converte DOCX com Mammoth, sanitiza o HTML e preserva texto, ordem, negrito, itálico e links na estrutura de blocos.
- `src/lib/custom-components.ts`: registra marcadores e seus templates HTML confiáveis. `%acessar%` usa exatamente o código de tabelas fornecido, inclusive a URL e o texto originais. Outros `%texto%` usam o mesmo visual, com texto e URL vindos do Word. Novos componentes específicos podem ser registrados sem mudar o importador.
- `src/lib/conditions.ts`: reconhece e valida pares `##AUDIENCIA##` / `##FIM_AUDIENCIA##` no modelo de blocos, incluindo marcadores fragmentados em runs do Word.
- `src/lib/email.ts`: gera MJML com o cabeçalho, slots para componentes e limites de audiência. Depois de compilar, insere o HTML confiável dos botões e as expressões AJO entre linhas completas de tabelas, preservando o código original dos componentes.
- `src/lib/header.ts`: define os valores padrão e valida a configuração do cabeçalho.
- `src/lib/project.ts`: valida, salva e reabre o documento estruturado, as edições individuais dos botões e as configurações visuais, incluindo o cabeçalho. O navegador guarda o projeto localmente; o JSON permite backup manual.
- `src/App.tsx`: importação, revisão, configuração do cabeçalho, edição dos botões, simulação de audiência na prévia e exportação.

O código do botão não inclui VML para Outlook clássico. Esse cliente pode ignorar os cantos arredondados de `border-radius`; o template foi preservado sem acrescentar fallback visual. Imagens embutidas no Word ainda são ignoradas com aviso. O editor drag and drop, a leitura de prints e o Ollama não fazem parte desta funcionalidade.

Para verificar os módulos e a compilação:

```powershell
npm test
npm run build
```

Os testes cobrem DOCX com marcadores fragmentados, formatação vizinha, condições simples e múltiplas, botões dentro de audiências, erros de validação, expressões AJO literais, posições no HTML e persistência do projeto.

A proposta para a próxima entrega está em [docs/ETAPA_2.md](docs/ETAPA_2.md). Ela será ajustada com exemplos reais antes da implementação.
