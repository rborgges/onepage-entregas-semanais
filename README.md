# Onepage de Entregas Semanais

Gera automaticamente um slide (`.pptx`) com o resumo semanal de entregas da equipe, a partir de um arquivo de texto simples editado no Notepad++, VS Code ou qualquer editor.

## Arquivos do projeto

| Arquivo | Descrição |
|---|---|
| `Onepage_Entregas_Semanais.pptx` | Template base (não editar diretamente) |
| `preencher_onepage.py` | Script que lê o `.txt` e gera o slide preenchido |
| `dados_semana_exemplo.txt` | Exemplo do formato de dados esperado |
| `dados_semana.txt` | Seu arquivo de dados da semana (você cria/edita) |

## Requisitos

- Python 3.8+
- Biblioteca `python-pptx`

Instalar:

```bash
python -m pip install python-pptx
```

## Como usar

1. Copie `dados_semana_exemplo.txt` para `dados_semana.txt` (ou outro nome) e edite com os dados da semana.
2. Rode o script:

```bash
python preencher_onepage.py dados_semana.txt
```

3. Isso gera `Onepage_Entregas_Semanais_preenchido.pptx` na mesma pasta.
4. Abra o arquivo, confira, e exporte como PDF (**Arquivo → Exportar → PDF**) para anexar no e-mail semanal.

### Opções

```bash
python preencher_onepage.py dados_semana.txt --template Onepage_Entregas_Semanais.pptx --output relatorio_semana32.pptx
```

- `--template`: caminho do `.pptx` modelo (padrão: `Onepage_Entregas_Semanais.pptx`)
- `--output`: nome do arquivo gerado (padrão: `Onepage_Entregas_Semanais_preenchido.pptx`)

## Formato do arquivo de dados (`.txt`)

```
EQUIPE: Time de Growth
PERIODO: 04/08 a 08/08/2026
RESPONSAVEL: Fulano de Tal

ENTREGAS:
Landing page nova | Ana | Concluído | No ar desde quarta
Integração com API de pagamento | Bruno | Em andamento | 80% concluído
Ajuste no checkout | Carla | Atrasado | Aguardando fornecedor

DESTAQUES:
Lançamento da campanha X superou a meta em 20%
Novo processo de onboarding reduziu tickets de suporte

RISCOS:
Atraso do fornecedor Z pode impactar o checkout
Precisamos de mais 1 dev para a sprint que vem
```

**Regras:**

- Cada linha em `ENTREGAS:` tem 4 campos separados por `|`:
  `Nome da entrega | Responsável | Status | Observação` (observação é opcional)
- Status reconhecidos (cor automática, sem diferenciar maiúsculas/acentos):
  - **Concluído** → verde
  - **Em andamento** → amarelo
  - **Atrasado** / **Bloqueado** → vermelho
  - Qualquer outro texto entra sem cor especial
- `DESTAQUES:` e `RISCOS:` aceitam quantas linhas quiser — uma por bullet
- A tabela se ajusta automaticamente ao número de entregas (linhas são adicionadas ou removidas, e a altura é redimensionada para não sobrepor os cards de baixo)
- Se houver **muitas** entregas (8+), o script avisa no terminal que o espaço pode ficar apertado — nesse caso, resuma as observações ou divida em mais de um envio

## Rotina sugerida

Toda sexta-feira:

1. Atualize `dados_semana.txt` com as entregas da semana
2. Rode `python preencher_onepage.py dados_semana.txt`
3. Exporte o `.pptx` gerado como PDF
4. Anexe ao e-mail semanal da equipe

## Problemas comuns

- **`ModuleNotFoundError: No module named 'pptx'`** → rode `python -m pip install python-pptx`
- **Erro de launcher do pip no Windows** → use `python -m pip install python-pptx` em vez de `pip install ...`
- **Acentos estranhos no `.txt`** → salve o arquivo como UTF-8 no seu editor (no Notepad++: Codificação → UTF-8)
