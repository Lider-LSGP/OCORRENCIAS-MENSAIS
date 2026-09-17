# Ocorrências 2.0 — Conferência & Importação (Domínio → EasyApp)

App Streamlit para **conferir o que já está lançado** e **gerar as planilhas de
importação** de ocorrências, por tipo e por empresa.

## Módulos

| Módulo | Status | Observações |
|---|---|---|
| Faltas Mês Todo | ✅ Completo | filtro `valor_inf` ≥ limiar (padrão 20) |
| Férias | 🚧 Esqueleto | mesma estrutura de saída; aguardando regras |
| Afastamentos | 🚧 Esqueleto | idem |
| Rescisões / Avisos | 🚧 Esqueleto | idem |

## Entradas

1. **Base de colaboradores do sistema** (exportação completa, com `Id:`, `Nome`,
   `Matricula`, `Empresa:`, `Tipo Escala`, `Posto Trabalho:`, `Ativo`,
   `DT/Demissão`, `Tipo/PCD:`).
2. **Relatório Sintético de Ocorrências do EasyApp** (opcional, mas recomendado —
   é ele que alimenta a conferência de duplicidade). Cabeçalho `PARCEIROID` é
   localizado automaticamente.
3. **Relatórios da Domínio** (um ou vários; o tipo é detectado pelas colunas).

## Regras implementadas — Faltas Mês Todo

- `valor_inf` ≥ 20 ⇒ falta de mês todo.
- Referência = coluna `data` do relatório; a ocorrência é lançada no **mês
  seguinte** (08/2026 → ocorrência 01/09→30/09, etc.). `qtddiasafastamento` =
  dias do mês (28/29/30/31, bissexto ok); retorno = dia seguinte ao fim.
- `tipoocorrencia_id` = 13; demais campos fixos conforme registros-modelo do
  layout; colunas `id`, `codigoresponsavelcobertura`, `nomeresponsavelcobertura`,
  `usoregcreated_at`, `system_unit_id`, `totaldiasdescontado` sempre vazias.
- Datas em `aaaa-mm-dd`; `mes_ano` em `mm/aaaa` (mês/ano do cadastro = hoje).
- **Casamento** do colaborador (porque `i_empregados` Domínio ≠ `Id` EasyApp):
  matrícula eSocial → nome+empresa → nome. Códigos podem se repetir entre
  empresas, por isso a empresa (`codi_emp`) faz parte da chave.
- **Conferência**: já existe ocorrência com a mesma `datainicio` para o
  colaborador ⇒ STATUS **JÁ LANÇADO** e a linha NÃO sai no CSV de importação.

## Saídas (sempre em par)

- `importacao_<tipo>[_empresa].csv` — pronto para importar no EasyApp.
- `resultado_<tipo>[_empresa].xlsx` — conferência com **STATUS** colorido
  (VÁLIDO / EM ABERTO / JÁ LANÇADO), dados de cruzamento e **OBS**
  (não encontrado, nome divergente, demitido/não ativo, PCD...).

Escolha na barra lateral: **juntar empresas** (1 arquivo por tipo) ou
**dividir por tipo + empresa**.

## Empresas (codi_emp)

1. VSP VIGILANCIA E SEGURANCA PATRIMONIAL L
2. ATIVA TERCEIRIZACAO DE MAO DE OBRA LTDA
3. LIDER MULTISSERVICOS LTDA
4. LIDER LIMPE LIMPEZA COMERCIAL LTDA

## Rodar

```bash
pip install -r requirements.txt
streamlit run app.py
```

Deploy: suba a pasta no GitHub e publique no Streamlit Community Cloud
(apontando para `app.py`).

## Estrutura

```
app.py                # interface Streamlit
core/
  leitura.py          # leitura tolerante (xlsx / xls BIFF / SYLK Domínio)
  empresas.py         # cadastro e aliases das 4 empresas
  colaboradores.py    # base + casamento em cascata (matrícula → nome+emp → nome)
  ocorrencias.py      # índice do que já está lançado (conferência)
  layout.py           # colunas/constantes do layout EasyApp + datas do mês
  faltas.py           # módulo Faltas + esqueleto dos demais tipos
  saida.py            # CSV importação, XLSX resultado colorido, ZIP
```
