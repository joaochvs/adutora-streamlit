# Revisitas de campo

Envie um Excel (.xlsx/.xlsm) e baixe o PDF. Não há filtros ou configuração de colunas.

O aplicativo usa a aba CAMPO; se ela não existir, usa a primeira aba. Reconhece os títulos nas primeiras 30 linhas. As colunas obrigatórias são CHAVE_VISITA, LOGRADOURO, STATUS e pelo menos uma coluna de turno.

Entram os status Ausente e Ausente seguido de número, como AUSENTE 1 e AUSENTE 2 (ignora maiúsculas, acentos e espaços extras). Pendente fica fora porque ainda não houve visita. A primeira coluna mostra o endereço: logradouro, N°, complemento e município. A segunda, Histórico, mostra V1, V2 etc., com horário, período, data e dia da semana disponíveis. A terceira usa textos curtos: “2ª visita = Tarde” e “3ª visita = Folder (sugestão)”, seguidos do próximo dia de atendimento. AUSENTE 1 sugere a 2ª visita, AUSENTE 2 sugere Folder na 3ª. Para Ausente sem número, a tentativa é inferida da última visita preenchida. Traços são tratados como campos vazios.

Os pares Turno/Dia, Turno2/Dia3, Turno4/Dia5 e Turno6/Dia7 representam V1 a V4. O último par preenchido determina a recomendação: manhã sugere tarde, e tarde sugere manhã. A próxima visita é no dia seguinte, de segunda a sábado; sábado sugere segunda-feira. Quando a próxima visita cai no sábado, pode ser de manhã ou à tarde. Datas brasileiras, datas do Excel e nomes dos dias da semana são aceitos. Se houver somente o nome do dia, não se inventa uma data.

Nos trechos 8 a 11, as sugestões exibem as faixas Manhã: 08:00 - 11:59 e Tarde: 12:00 - 15:00, inclusive no sábado. Nos trechos 15 a 19, uma última visita noturna sugere manhã; não há sugestão noturna. A terceira tentativa sugere Folder. Dias ou turnos vazios/desconhecidos exigem confirmação; se o último par for incompleto ou inválido, não usamos silenciosamente uma visita antiga. Cada linha de ausência é preservada.

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
