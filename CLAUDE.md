# Fluxo de caixa das obras (EPO): ferramenta, painel publicado e planilhas

## Regra: toda modificação tem de chegar à página publicada, conferida

Página: https://claude.ai/artifact/9TpDGLVKmKfKstM6MiEAvy (montada a partir de `site/`).
Sempre que entrar qualquer modificação (incorrido, cronograma, premissa, estética, painel novo), depois de regenerar o site
(`viewer.py` → `make_site.py` → `pdfs.js` → `make_site.py`, xlsx + `.b64.txt`) faça a rotina de publicação, sem exceção:

1. `python3 gerador/conferir_publicacao.py preparar`
   Tem que sair sem ERRO. Ele confere localmente:
   - o corte do cartão é igual ao corte dos dados;
   - nenhum painel sumiu em relação ao que está publicado;
   - o `.b64.txt` corresponde ao `.xlsx`;
   - o xlsx do site é igual ao entregável da raiz;
   - o PDF não é mais antigo que os dados.

   Depois grava `site/.pendente.json` com os lotes a publicar e os caminhos a reler.
2. Publicar cada lote. Artifact publish com:
   - `file_path` = `site/index_artifact.html`;
   - `url` = a página acima;
   - `files` = o lote.

   Antes, se a publicação for recusada por arquivos "não lidos", liste os arquivos da página (`scope: files`).
3. Reler da página publicada todos os caminhos de `reler`. Artifact read com `paths` e `out_dir` numa pasta de conferência.
4. `python3 gerador/conferir_publicacao.py conferir <pasta de conferência>`
   Tem que dar "OK: página publicada = arquivos locais". Só então atualiza `site/.publicado.json`.
5. Commit inclui `site/.publicado.json`. Ele é o registro do que está confirmado no ar.

Nunca publique só os cartões ou só os `.b64.txt`: os painéis vêm de `obras/<id>.json` e os PDFs de `pdf/<id>.pdf`.
O gancho `Stop` em `.claude/settings.json` (`conferir_publicacao.py gancho`) impede encerrar o turno enquanto houver
arquivo do site diferente do último estado confirmado.

## Outras convenções

- Controle em INCC. Detalhes do gerador em `gerador/LEIA-ME.md`.
- Aba/painel ESQUEMÁTICO: skill `.claude/skills/esquematico-avanco`.
- Mudanças só estéticas não podem alterar números. Confira o 2027 de cada cenário e a auditoria (19 OK) a cada build.
