# Excel export
Prosty program do wyciągania danych z wielu plików Excela i zapisywania ich jako tabela w `docx`.



### Sciągnij paczkę z GH:

```bash
$ git clone https://github.com/mdyzma/excel-export.git

```

### Zainstaluj zależności

Przejdź do głównego folderu:

```bash
$ cd excel-export
```
Zainstaluj zależności:

```bash
pip install -r  requirements.txt
```

### Uruchom

Przejdź do głównego folderu:

```bash
$ cd excel-export
```
Odpal program i podaj parametry

`--input <pliki-excela-na-dysku>` oraz `--nsamples <liczba-wylosowanych-wynikow>`. Domyślna wartość dla liczby próbek wynosi `100`. Możliwe jest dodanie trzeciego parametru - folderu w którym zapisane będą wyniki `--destination <nazw-folderu-do-zapisania-wyników>`. Domyślna wartość to folder `~/exls_exports/`. Parametry oczywiście bez `<` i `>`

Przykład dla Windows:

```bash
python main.py --input %UerProfile%\Dokumenty\excel-files --nsamples 1000
```

Przykład dla Linux:

```bash
python main.py --input ~/excel-files --nsamples 1000
```



### Help


```bash
python main.py --help
```