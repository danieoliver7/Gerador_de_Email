# Etapa 2 — proposta para revisão

## Resultado esperado

Enviar o DOCX e um print do layout de referência. A aplicação usa o Qwen3-VL instalado no Ollama local para sugerir estrutura e estilo, monta uma nova prévia MJML e permite aceitar ou descartar a sugestão. O **DOCX continua sendo a única fonte do copy e dos links**.

## Escopo proposto

1. **Ollama local:** detectar o serviço em `127.0.0.1:11434`, verificar o modelo `qwen3-vl:4b-instruct-q4_K_M`, mostrar estados claros (desligado, pronto, analisando, erro) e oferecer um comando Windows que inicia aplicação e Ollama quando necessário. Restringir ambos ao loopback.
2. **Print:** aceitar PNG/JPEG; validar tipo, tamanho e dimensões; reduzir a resolução em memória quando necessário. O print não será enviado a serviços externos nem salvo automaticamente.
3. **Análise visual:** pedir ao modelo somente um plano estruturado de layout: largura, cores, espaçamentos, alinhamento, hierarquia, seções e até duas colunas. Usar saída JSON Schema e validar localmente antes de aplicar. O plano não terá campos para texto, URLs ou HTML arbitrário.
4. **Composição determinística:** aplicar o plano ao conteúdo estruturado do Word. Cada bloco do DOCX deve aparecer exatamente uma vez e na mesma ordem; negritos, itálicos e hyperlinks válidos devem permanecer. Se o print mostrar algo que não tem arquivo de origem (imagem, logo ou link), mostrar aviso e não inventar esse recurso.
5. **Revisão:** comparar a versão atual com a sugestão e oferecer **Aplicar** e **Descartar**. Falhas do modelo deixam o template simples da Etapa 1 utilizável.
6. **Exportação:** gerar MJML e HTML responsivo com tabelas e estilos adequados a email. A prévia fica isolada em iframe sem scripts.

## Fora desta etapa

- Editor drag-and-drop (GrapesJS), condicionais IF/ELSE e atributos Adobe Journey Optimizer.
- Reconstrução de imagens ou logos a partir do print.
- Reescrita do copy pelo modelo.
- Contas, hospedagem, APIs pagas ou IA externa.

## Desenho técnico

`DOCX → Mammoth → conteúdo estruturado` permanece como está. `Print → Ollama /api/chat local → JSON de layout validado` é uma trilha separada. O renderizador MJML combina os dois, mas só o conteúdo do Word fornece palavras e URLs. A conexão com o Ollama será isolada em um módulo TypeScript e usará `stream:false`, um JSON Schema restrito e tempo de carregamento limitado; o navegador usa a origem local já autorizada pelo Ollama. Um script local de inicialização poderá abrir o Ollama se ele não estiver ativo, sem criar um segundo servidor na mesma porta.

## Critérios de aceite

- Um print de email de uma ou duas colunas gera uma sugestão que pode ser aplicada e desfeita.
- Uma comparação automática confirma que todos os trechos de texto do Word aparecem uma única vez, na ordem original, no modelo intermediário antes da exportação.
- Nenhum URL ou imagem aparece no HTML sem vir do DOCX ou de um arquivo fornecido pelo usuário.
- Um JSON inválido do modelo é rejeitado com mensagem compreensível, mantendo a geração da Etapa 1.
- Testes cobrem validação do plano, preservação do copy, falha do Ollama e exportação; um teste manual usa o Qwen local com um DOCX e print de referência.
- A aplicação e a análise funcionam sem conexão com a internet.

## Decisões pendentes

1. Receber um par DOCX + print representativo para calibrar e validar a análise.
2. Definir se upload de imagens/logos originais entra já nesta etapa. Sem o arquivo original, imagens do print serão omitidas com aviso.
3. Definir como o Word indica botões e como obter a URL de cada CTA. Sem URL explícita, nenhum link clicável será criado.

**Status:** proposta; nenhuma implementação da Etapa 2 foi iniciada.

## Referências técnicas

- [Visão no Ollama](https://docs.ollama.com/capabilities/vision)
- [Saída estruturada](https://docs.ollama.com/capabilities/structured-outputs)
- [API de chat](https://docs.ollama.com/api/chat)
- [FAQ: keep_alive e origens locais](https://docs.ollama.com/faq)
