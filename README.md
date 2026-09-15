# Dashboard Reclame Aqui — J&T Express

Dashboard web preparado para a planilha `RA_seed.xlsx`, com separação de **Base Própria x Franquia**, filtros por **Regional (MG / SPN / Todos)** e **RM**, SLA automático de 24h, ranking por RM e motivo, edição da base dentro do próprio site, importação/exportação XLSX e tradução PT-BR / Chinês Simplificado.

## Regras implementadas

- **Franquia**: qualquer base cujo nome começa com `F ` (ex.: `F SOD 02-SP`). As demais são **BASE PRÓPRIA**.
- **Pedido em aberto**: `SOLUÇÃO` vazia ou igual a `EM ABERTO`.
- **SLA automático**: diferença entre o horário atual de São Paulo e `DATA RECLAMAÇÃO`.
  - `<= 24h` = Dentro do prazo
  - `> 24h` = Fora do prazo
- A coluna original **DENTRO DO PRAZO** continua disponível para preenchimento manual, mas o ranking usa o SLA automático de 24h solicitado.
- **Tratativa**: recebe uma sugestão automática conforme o motivo e pode ser editada manualmente.
- **Sincronização multiusuário**: toda gravação incrementa a versão dos dados no banco. Os navegadores verificam mudanças a cada ~2,5 s e atualizam a tela. No Railway, use PostgreSQL para que todos os usuários compartilhem a mesma base.

## Rodar no Windows/local

1. Instale Python 3.12+.
2. Abra o terminal na pasta do projeto.
3. Instale as dependências:

```bash
pip install -r requirements.txt
```

4. Execute:

```bash
python app.py
```

5. Abra `http://127.0.0.1:5000`.

Localmente, se `DATABASE_URL` não existir, o sistema usa SQLite em `data/dashboard.db`.

## Publicar no Railway

1. Crie um projeto no Railway e envie este projeto via GitHub ou upload.
2. Adicione um serviço **PostgreSQL** ao mesmo projeto.
3. O Railway normalmente disponibiliza `DATABASE_URL` automaticamente para o serviço conectado. Se necessário, adicione essa variável manualmente usando a URL do PostgreSQL.
4. Adicione também `SECRET_KEY` com um valor aleatório.
5. Deploy. O projeto já inclui `railway.json`, `Procfile` e `Dockerfile`.
6. O healthcheck é `/health`.

> Importante: não use apenas o SQLite do container no Railway para produção. O filesystem do deploy pode ser recriado e não é a melhor opção para sincronização persistente entre usuários. PostgreSQL é o modo recomendado.

## Importar uma nova planilha

Na aba **Editar Base**, clique em **Importar XLSX**. A importação substitui os dados atuais, mas tenta preservar `Tratativa` e `DENTRO DO PRAZO` por `ID RA` quando o mesmo ID existe na nova planilha.

A planilha pode usar os cabeçalhos originais:

`ID RA`, `DATA RECLAMAÇÃO`, `ESTADO`, `RASTREIO`, `Motorista`, `BASE`, `BASE/FRANQUIA`, `RM`, `DATA DIRECIONAMENTO`, `Motivo`, `SOLUÇÃO`, `DENTRO DO PRAZO`, `Observação`.

O sistema também reconhece a coluna adicional `TRATATIVA`.

## Tradução

O botão **PT-BR / 简体中文** traduz toda a interface e os valores categóricos conhecidos (motivos, soluções, tipo de base, SLA e tratativas sugeridas). Campos livres digitados pelo usuário, como observações personalizadas, permanecem no idioma em que foram escritos.
