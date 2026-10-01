# Revisitas de campo

Envie um Excel (.xlsx/.xlsm) e baixe o PDF. Não há filtros ou configuração de colunas.

O aplicativo usa a aba CAMPO; se ela não existir, usa a primeira aba. Reconhece os títulos nas primeiras 30 linhas. As colunas obrigatórias são CHAVE_VISITA, LOGRADOURO, STATUS e pelo menos uma coluna de turno.

Entram os status Ausente e Ausente seguido de número, como AUSENTE 1 e AUSENTE 2 (ignora maiúsculas, acentos e espaços extras). Pendente fica fora porque ainda não houve visita. A primeira coluna mostra apenas o endereço: logradouro, N°, complemento e município. A segunda coluna, Última visita, mostra o último período visitado. A terceira indica a próxima tentativa: AUSENTE 1 sugere a 2ª visita, AUSENTE 2 sugere a 3ª, e assim por diante. Para Ausente sem número, usa “Realizar próxima visita”. Traços nas colunas de turno são tratados como campos vazios.

O último valor preenchido na sequência Turno → Turno2 → Turno4 → Turno6 determina a recomendação: manhã sugere tarde, e tarde sugere manhã, em outro dia. As colunas Dia, Dia3, Dia5 e Dia7 não são usadas. O relatório indica períodos, sem inventar horários exatos. Turnos vazios, desconhecidos ou noturnos exigem confirmação. Se o último valor for inválido, não usamos silenciosamente um turno antigo. Cada linha de ausência é preservada.

## Executar

O trecho da coluna TRECHO aparece abaixo do título do PDF. Se houver vários trechos entre os ausentes, todos são listados.

## Publicar no Streamlit Community Cloud

Conecte este repositório no Streamlit Community Cloud, selecione a branch publicada e indique `app.py` como arquivo principal. As dependências são instaladas a partir de `requirements.txt`. Use Python 3.11 ou 3.12. Os arquivos Excel e relatórios locais não são enviados ao GitHub.

```powershell
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Após instalar as dependências, também é possível abrir iniciar.bat. Os uploads são processados em memória.

## Testar

```powershell
python -m unittest discover -s tests -v
```
