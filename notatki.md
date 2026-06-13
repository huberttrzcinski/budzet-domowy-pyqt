# Notatki do zaliczenia — Python, PyQt5, Excel, OOP, wykresy

Ten plik służy jako szybka ściąga ze składni potrzebnej do stworzenia aplikacji desktopowej w Pythonie.

---

# 1. Podstawowe importy

```python
import sys
from dataclasses import dataclass

import pandas as pd

from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QPushButton,
    QFileDialog,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QComboBox,
    QDateEdit,
    QMessageBox
)

from PyQt5.QtCore import QDate

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
```

---

# 2. Klasa danych — pojedynczy wydatek

```python
@dataclass
class Wydatek:
    data: str
    kategoria: str
    kwota: float
    opis: str
```

Przykład utworzenia obiektu:

```python
wydatek = Wydatek("2026-06-01", "Jedzenie", 45.50, "Zakupy")
```

---

# 3. Klasa zarządzająca danymi

```python
class BudgetManager:
    def __init__(self):
        self.wydatki = []

    def dodaj_wydatek(self, wydatek):
        self.wydatki.append(wydatek)

    def czysc_dane(self):
        self.wydatki.clear()

    def oblicz_sume_kategorii(self):
        wynik = {}

        for wydatek in self.wydatki:
            if wydatek.kategoria not in wynik:
                wynik[wydatek.kategoria] = 0

            wynik[wydatek.kategoria] += wydatek.kwota

        return wynik
```

Użycie:

```python
manager = BudgetManager()
manager.dodaj_wydatek(wydatek)
print(manager.oblicz_sume_kategorii())
```

---

# 4. Wczytywanie danych z Excela

Plik Excel powinien mieć kolumny:

```text
Data | Kategoria | Kwota | Opis
```

Kod:

```python
def wczytaj_z_excela(self, sciezka_pliku):
    df = pd.read_excel(sciezka_pliku)

    wymagane_kolumny = ["Data", "Kategoria", "Kwota", "Opis"]

    for kolumna in wymagane_kolumny:
        if kolumna not in df.columns:
            raise ValueError(f"Brakuje kolumny: {kolumna}")

    self.czysc_dane()

    for _, row in df.iterrows():
        data = str(row["Data"])
        kategoria = str(row["Kategoria"])
        kwota = float(row["Kwota"])
        opis = str(row["Opis"])

        wydatek = Wydatek(data, kategoria, kwota, opis)
        self.dodaj_wydatek(wydatek)
```

Najważniejsza linia:

```python
df = pd.read_excel(sciezka_pliku)
```

---

# 5. Otwieranie pliku przez QFileDialog

```python
sciezka_pliku, _ = QFileDialog.getOpenFileName(
    self,
    "Wybierz plik Excel",
    "",
    "Pliki Excel (*.xlsx)"
)
```

Sprawdzenie, czy użytkownik wybrał plik:

```python
if sciezka_pliku:
    self.manager.wczytaj_z_excela(sciezka_pliku)
```

---

# 6. Szkielet aplikacji PyQt5

```python
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Budżet domowy")
        self.setGeometry(200, 100, 1000, 600)

        self.manager = BudgetManager()

        self.utworz_interfejs()
```

Uruchomienie programu:

```python
if __name__ == "__main__":
    app = QApplication(sys.argv)

    okno = MainWindow()
    okno.show()

    sys.exit(app.exec_())
```

---

# 7. Tworzenie głównego layoutu

```python
centralny_widget = QWidget()
self.setCentralWidget(centralny_widget)

layout_glowny = QVBoxLayout()

centralny_widget.setLayout(layout_glowny)
```

Layout poziomy:

```python
layout_przyciski = QHBoxLayout()
```

Dodawanie elementów:

```python
layout_glowny.addLayout(layout_przyciski)
layout_glowny.addWidget(self.tabela)
```

---

# 8. Przycisk i podłączenie funkcji

```python
self.btn_wczytaj = QPushButton("Wczytaj plik Excel")
self.btn_wczytaj.clicked.connect(self.wczytaj_excel)
```

Funkcja wykonywana po kliknięciu:

```python
def wczytaj_excel(self):
    print("Kliknięto przycisk")
```

Schemat:

```python
przycisk.clicked.connect(nazwa_funkcji)
```

Uwaga: przy `connect` nie dajemy nawiasów.

Poprawnie:

```python
self.btn.clicked.connect(self.funkcja)
```

Źle:

```python
self.btn.clicked.connect(self.funkcja())
```

---

# 9. Tabela QTableWidget

Tworzenie tabeli:

```python
self.tabela = QTableWidget()
self.tabela.setColumnCount(4)
self.tabela.setHorizontalHeaderLabels(["Data", "Kategoria", "Kwota", "Opis"])
```

Ustawienie liczby wierszy:

```python
self.tabela.setRowCount(len(self.manager.wydatki))
```

Wpisanie danych do tabeli:

```python
self.tabela.setItem(nr_wiersza, 0, QTableWidgetItem(wydatek.data))
self.tabela.setItem(nr_wiersza, 1, QTableWidgetItem(wydatek.kategoria))
self.tabela.setItem(nr_wiersza, 2, QTableWidgetItem(str(wydatek.kwota)))
self.tabela.setItem(nr_wiersza, 3, QTableWidgetItem(wydatek.opis))
```

Cała funkcja odświeżania tabeli:

```python
def odswiez_tabele(self):
    self.tabela.setRowCount(len(self.manager.wydatki))

    for nr_wiersza, wydatek in enumerate(self.manager.wydatki):
        self.tabela.setItem(nr_wiersza, 0, QTableWidgetItem(wydatek.data))
        self.tabela.setItem(nr_wiersza, 1, QTableWidgetItem(wydatek.kategoria))
        self.tabela.setItem(nr_wiersza, 2, QTableWidgetItem(str(wydatek.kwota)))
        self.tabela.setItem(nr_wiersza, 3, QTableWidgetItem(wydatek.opis))
```

---

# 10. Pola tekstowe QLineEdit

```python
self.kwota_input = QLineEdit()
self.kwota_input.setPlaceholderText("Kwota")
```

Pobieranie tekstu:

```python
kwota = self.kwota_input.text()
```

Czyszczenie pola:

```python
self.kwota_input.clear()
```

Zamiana tekstu na liczbę:

```python
kwota = float(self.kwota_input.text().replace(",", "."))
```

---

# 11. ComboBox — lista kategorii

```python
self.kategoria_input = QComboBox()
self.kategoria_input.addItems([
    "Jedzenie",
    "Transport",
    "Rozrywka",
    "Rachunki",
    "Inne"
])
```

Pozwolenie na wpisanie własnej kategorii:

```python
self.kategoria_input.setEditable(True)
```

Pobieranie wybranej kategorii:

```python
kategoria = self.kategoria_input.currentText()
```

---

# 12. Data — QDateEdit

```python
self.data_input = QDateEdit()
self.data_input.setDate(QDate.currentDate())
self.data_input.setCalendarPopup(True)
```

Pobieranie daty jako tekst:

```python
data = self.data_input.date().toString("yyyy-MM-dd")
```

---

# 13. Dodawanie wydatku ręcznie

```python
def dodaj_wydatek_recznie(self):
    try:
        data = self.data_input.date().toString("yyyy-MM-dd")
        kategoria = self.kategoria_input.currentText()
        kwota = float(self.kwota_input.text().replace(",", "."))
        opis = self.opis_input.text()

        wydatek = Wydatek(data, kategoria, kwota, opis)

        self.manager.dodaj_wydatek(wydatek)

        self.kwota_input.clear()
        self.opis_input.clear()

        self.odswiez_tabele()
        self.odswiez_wykres()

    except ValueError:
        QMessageBox.warning(
            self,
            "Błąd danych",
            "Kwota musi być liczbą, np. 25.50"
        )
```

---

# 14. Komunikaty QMessageBox

Błąd krytyczny:

```python
QMessageBox.critical(
    self,
    "Błąd",
    "Nie udało się wczytać pliku"
)
```

Ostrzeżenie:

```python
QMessageBox.warning(
    self,
    "Błąd danych",
    "Kwota musi być liczbą"
)
```

Informacja:

```python
QMessageBox.information(
    self,
    "Informacja",
    "Dane zostały zapisane"
)
```

---

# 15. Obsługa błędów try/except

```python
try:
    kwota = float(self.kwota_input.text())
except ValueError:
    QMessageBox.warning(self, "Błąd", "Podaj poprawną liczbę")
```

Przy wczytywaniu pliku:

```python
try:
    self.manager.wczytaj_z_excela(sciezka_pliku)
except Exception as blad:
    QMessageBox.critical(
        self,
        "Błąd",
        f"Nie udało się wczytać pliku:\n{blad}"
    )
```

---

# 16. Matplotlib w PyQt5

Klasa do wykresu:

```python
class WykresWidget(FigureCanvas):
    def __init__(self):
        self.figure = Figure(figsize=(5, 4))
        self.ax = self.figure.add_subplot(111)

        super().__init__(self.figure)

    def rysuj_wykres(self, dane):
        self.ax.clear()

        if not dane:
            self.ax.set_title("Brak danych do wyświetlenia")
            self.draw()
            return

        kategorie = list(dane.keys())
        kwoty = list(dane.values())

        self.ax.bar(kategorie, kwoty)
        self.ax.set_title("Suma wydatków według kategorii")
        self.ax.set_xlabel("Kategoria")
        self.ax.set_ylabel("Kwota [zł]")
        self.ax.tick_params(axis="x", rotation=30)

        self.figure.tight_layout()
        self.draw()
```

Utworzenie wykresu w oknie:

```python
self.wykres = WykresWidget()
layout_glowny.addWidget(self.wykres)
```

Odświeżenie wykresu:

```python
def odswiez_wykres(self):
    dane = self.manager.oblicz_sume_kategorii()
    self.wykres.rysuj_wykres(dane)
```

---

# 17. Wykres słupkowy

```python
self.ax.bar(kategorie, kwoty)
```

---

# 18. Wykres kołowy

Zamiast:

```python
self.ax.bar(kategorie, kwoty)
```

można dać:

```python
self.ax.pie(kwoty, labels=kategorie, autopct="%1.1f%%")
self.ax.set_title("Udział kategorii w wydatkach")
```

Cała wersja dla wykresu kołowego:

```python
self.ax.clear()

kategorie = list(dane.keys())
kwoty = list(dane.values())

self.ax.pie(kwoty, labels=kategorie, autopct="%1.1f%%")
self.ax.set_title("Udział kategorii w wydatkach")

self.draw()
```

---

# 19. Czyszczenie danych

W klasie managera:

```python
def czysc_dane(self):
    self.wydatki.clear()
```

W oknie:

```python
def wyczysc_dane(self):
    self.manager.czysc_dane()
    self.odswiez_tabele()
    self.odswiez_wykres()
```

Przycisk:

```python
self.btn_wyczysc = QPushButton("Wyczyść dane")
self.btn_wyczysc.clicked.connect(self.wyczysc_dane)
```

---

# 20. Eksport danych do Excela

Opcjonalnie, jeśli będzie wymagane:

```python
def eksportuj_do_excela(self, sciezka_pliku):
    dane = []

    for wydatek in self.wydatki:
        dane.append({
            "Data": wydatek.data,
            "Kategoria": wydatek.kategoria,
            "Kwota": wydatek.kwota,
            "Opis": wydatek.opis
        })

    df = pd.DataFrame(dane)
    df.to_excel(sciezka_pliku, index=False)
```

Dialog zapisu pliku:

```python
sciezka_pliku, _ = QFileDialog.getSaveFileName(
    self,
    "Zapisz plik Excel",
    "",
    "Pliki Excel (*.xlsx)"
)
```

---

# 21. Filtrowanie po kategorii

```python
def filtruj_po_kategorii(self, kategoria):
    wynik = []

    for wydatek in self.wydatki:
        if wydatek.kategoria == kategoria:
            wynik.append(wydatek)

    return wynik
```

Krótsza wersja:

```python
def filtruj_po_kategorii(self, kategoria):
    return [w for w in self.wydatki if w.kategoria == kategoria]
```

---

# 22. Obliczanie sumy wszystkich wydatków

```python
def oblicz_sume_wszystkich_wydatkow(self):
    suma = 0

    for wydatek in self.wydatki:
        suma += wydatek.kwota

    return suma
```

Krótsza wersja:

```python
def oblicz_sume_wszystkich_wydatkow(self):
    return sum(w.kwota for w in self.wydatki)
```

---

# 23. Liczba wydatków

```python
liczba = len(self.manager.wydatki)
```

Wyświetlenie w QLabel:

```python
self.label_suma.setText(f"Liczba wydatków: {liczba}")
```

---

# 24. QLabel — tekst w oknie

```python
self.label = QLabel("Tekst początkowy")
```

Zmiana tekstu:

```python
self.label.setText("Nowy tekst")
```

Przykład:

```python
suma = self.manager.oblicz_sume_wszystkich_wydatkow()
self.label_suma.setText(f"Suma wydatków: {suma:.2f} zł")
```

---

# 25. Formatowanie liczby

```python
kwota = 45.5
tekst = f"{kwota:.2f} zł"
```

Wynik:

```text
45.50 zł
```

---

# 26. Najczęstszy schemat programu

```text
Excel
↓
pandas read_excel()
↓
tworzenie obiektów Wydatek
↓
zapisanie obiektów w BudgetManager
↓
odświeżenie tabeli
↓
odświeżenie wykresu
```

---

# 27. Minimalny szkielet okna PyQt5

```python
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Moja aplikacja")
        self.setGeometry(200, 100, 800, 600)

        centralny_widget = QWidget()
        self.setCentralWidget(centralny_widget)

        layout = QVBoxLayout()

        przycisk = QPushButton("Kliknij")
        layout.addWidget(przycisk)

        centralny_widget.setLayout(layout)


if __name__ == "__main__":
    app = QApplication(sys.argv)

    okno = MainWindow()
    okno.show()

    sys.exit(app.exec_())
```

---

# 28. Minimalny przykład pandas + Excel

```python
import pandas as pd

df = pd.read_excel("dane.xlsx")

print(df.head())
print(df.columns)
```

Przejście po wierszach:

```python
for _, row in df.iterrows():
    print(row["Data"], row["Kategoria"], row["Kwota"], row["Opis"])
```

---

# 29. Minimalny przykład wykresu matplotlib

```python
import matplotlib.pyplot as plt

kategorie = ["Jedzenie", "Transport", "Rozrywka"]
kwoty = [100, 50, 80]

plt.bar(kategorie, kwoty)
plt.title("Wydatki według kategorii")
plt.xlabel("Kategoria")
plt.ylabel("Kwota [zł]")
plt.show()
```

---

# 30. Instalacja bibliotek

W terminalu:

```bash
pip install PyQt5 pandas matplotlib openpyxl
```

Jeśli jest plik requirements.txt:

```bash
pip install -r requirements.txt
```

---

# 31. Uruchomienie programu

```bash
python main.py
```

Na Macu czasem:

```bash
python3 main.py
```

---

# 32. Najczęstsze błędy

## Brak biblioteki

Błąd:

```text
ModuleNotFoundError: No module named 'PyQt5'
```

Rozwiązanie:

```bash
pip install PyQt5
```

---

## Excel się nie wczytuje

Błąd może wynikać z braku biblioteki:

```bash
pip install openpyxl
```

---

## Zła nazwa kolumny w Excelu

Program wymaga:

```text
Data
Kategoria
Kwota
Opis
```

Nazwy muszą być takie same jak w kodzie.

---

## Błąd przy kwocie

Jeżeli użytkownik wpisze tekst zamiast liczby, trzeba użyć:

```python
try:
    kwota = float(self.kwota_input.text().replace(",", "."))
except ValueError:
    QMessageBox.warning(self, "Błąd", "Kwota musi być liczbą")
```

---

## Przycisk nie działa

Sprawdź, czy jest:

```python
self.btn.clicked.connect(self.funkcja)
```

A nie:

```python
self.btn.clicked.connect(self.funkcja())
```

---

# 33. Co warto pamiętać na zaliczeniu

Najważniejsze fragmenty:

```python
df = pd.read_excel(sciezka_pliku)
```

```python
plik, _ = QFileDialog.getOpenFileName(self, "Wybierz plik", "", "Excel Files (*.xlsx)")
```

```python
self.btn.clicked.connect(self.funkcja)
```

```python
self.tabela.setItem(wiersz, kolumna, QTableWidgetItem("tekst"))
```

```python
self.ax.bar(kategorie, kwoty)
```

```python
self.draw()
```

```python
sys.exit(app.exec_())
```

---

# 34. Kolejność pisania programu na zaliczeniu

1. Zrobić klasy danych: `Wydatek`, `BudgetManager`.
2. Zrobić główne okno `MainWindow`.
3. Dodać przycisk wczytywania pliku Excel.
4. Dodać tabelę `QTableWidget`.
5. Dodać funkcję `wczytaj_excel()`.
6. Dodać funkcję `odswiez_tabele()`.
7. Dodać wykres matplotlib.
8. Dodać funkcję `odswiez_wykres()`.
9. Dodać czyszczenie danych.
10. Dopiero na końcu poprawiać wygląd aplikacji.

---

# 35. Prosty wzór do zapamiętania

```text
Dane → Logika → GUI → Wykres
```

Czyli:

```text
Wydatek → BudgetManager → MainWindow → WykresWidget
```
