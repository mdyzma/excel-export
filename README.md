# Excel export
Prosty program do wyciagania danych z Excela.



### Sciągnij paczke z GH:

```bash
$ git clone https://github.com/mdyzma/excel-export.git

```

### Zainstaluj zaleznosci

Przejdz do glownego folderu:

```bash
$ cd excel-export
```
Zainstaluj zaleznosci:

```bash
pip install -r  requirements.txt
```

### Uruchom

Przejdz do glownego folderu:

```bash
$ cd excel-export
```
Odpal program i podaj paramentry

`--input <pliki-excela-na-dysku>` oraz `--nsamples <liczba-wylosowanych-wynikow>`. Możliwe jest dodanie trzeciego parametru - folderu w którym zapisane beda wyniki `--destination <nazw-folderu-do-zapisania-wyników>`. Parametry oczywiście bez `<` i `>`

Na przykład (Windows):

```bash
python main.py --input Dokumenty\excelfiles --nsamples 1000
```

Przykład Linux:

```bash
python main.py --input ~/excel-files --nsamples 1000
```



### Help


```bash
python main.py --help
```