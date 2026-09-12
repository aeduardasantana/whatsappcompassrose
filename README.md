# Compass WhatsApp — V1

Aplicação PWA estática para gestão local de campanhas de WhatsApp, com Excel como entrada/saída e sem banco de dados.

## V1 incluída
- Importação XLSX/XLS/CSV
- Validação básica de telefone
- Seleção e busca de contatos
- Campanha com remetente e intervalo mínimo/máximo
- Até 5 variações de mensagem
- Variáveis {{nome}}, {{empresa}}, {{telefone}}
- Rodízio equilibrado, sequencial ou aleatório
- Fila local com iniciar/pausar/continuar
- Persistência via localStorage
- Relatório e exportação XLSX
- PWA instalável
- Deploy automático em GitHub Pages

## Importante
A V1 ainda não envia mensagens à API do WhatsApp. O módulo de transporte foi deliberadamente desacoplado para que a base do sistema continue 100% local e segura. A próxima etapa é escolher o mecanismo de envio compatível com a regra de não expor credenciais no código público.

## Publicar no GitHub Pages
1. Crie um repositório e envie estes arquivos para a branch `main`.
2. Abra **Settings → Pages**.
3. Em **Build and deployment**, selecione **GitHub Actions**.
4. Faça um novo push ou execute manualmente o workflow `Deploy GitHub Pages`.
