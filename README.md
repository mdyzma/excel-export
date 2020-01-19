# Excel export
Prosty program do wyciągania danych z wielu plików Excela i zapisywania ich jako tabela w `docx`.

![](short.gif)

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


Aby moc wygenerowac PDFa potrzebna jest tez biblioteka whtmltopdf. Instrukcja instalacji:

https://github.com/JazzCore/python-pdfkit/wiki/Installing-wkhtmltopdf


Po instalacji nalezy dodac folder `bin` do zmiennych środowiskowych.


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

usage: main.py [-h] [--input INPUT] [--destination DESTINATION]
               [--nsamples NSAMPLES]

optional arguments:
  -h, --help            show this help message and exit
  --input INPUT, -i INPUT
                        Ścieżka do folderu z plikami xls
  --destination DESTINATION, -d DESTINATION
                        Folder do zapiasania wynikow
  --nsamples NSAMPLES, -s NSAMPLES
                        Liczba wierszy wylosowana losowo z calego datasetu.

```