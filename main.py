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


# ==========================================
# MODEL DANYCH
# ==========================================

@dataclass
class Wydatek:
    data: str
    kategoria: str
    kwota: float
    opis: str


# ==========================================
# LOGIKA PROGRAMU
# ==========================================

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


# ==========================================
# WYKRES MATPLOTLIB
# ==========================================

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


# ==========================================
# GŁÓWNE OKNO APLIKACJI
# ==========================================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Budżet domowy")
        self.setGeometry(200, 100, 1000, 600)

        self.manager = BudgetManager()

        self.utworz_interfejs()

    def utworz_interfejs(self):
        centralny_widget = QWidget()
        self.setCentralWidget(centralny_widget)

        layout_glowny = QVBoxLayout()

        # Przyciski
        layout_przyciski = QHBoxLayout()

        self.btn_wczytaj = QPushButton("Wczytaj plik Excel")
        self.btn_wczytaj.clicked.connect(self.wczytaj_excel)

        self.btn_wyczysc = QPushButton("Wyczyść dane")
        self.btn_wyczysc.clicked.connect(self.wyczysc_dane)

        layout_przyciski.addWidget(self.btn_wczytaj)
        layout_przyciski.addWidget(self.btn_wyczysc)

        # Tabela
        self.tabela = QTableWidget()
        self.tabela.setColumnCount(4)
        self.tabela.setHorizontalHeaderLabels(["Data", "Kategoria", "Kwota", "Opis"])

        # Formularz dodawania wydatku
        layout_formularz = QHBoxLayout()

        self.data_input = QDateEdit()
        self.data_input.setDate(QDate.currentDate())
        self.data_input.setCalendarPopup(True)

        self.kategoria_input = QComboBox()
        self.kategoria_input.addItems([
            "Jedzenie",
            "Transport",
            "Rozrywka",
            "Rachunki",
            "Inne"
        ])
        self.kategoria_input.setEditable(True)

        self.kwota_input = QLineEdit()
        self.kwota_input.setPlaceholderText("Kwota")

        self.opis_input = QLineEdit()
        self.opis_input.setPlaceholderText("Opis")

        self.btn_dodaj = QPushButton("Dodaj wydatek")
        self.btn_dodaj.clicked.connect(self.dodaj_wydatek_recznie)

        layout_formularz.addWidget(QLabel("Data:"))
        layout_formularz.addWidget(self.data_input)
        layout_formularz.addWidget(QLabel("Kategoria:"))
        layout_formularz.addWidget(self.kategoria_input)
        layout_formularz.addWidget(QLabel("Kwota:"))
        layout_formularz.addWidget(self.kwota_input)
        layout_formularz.addWidget(QLabel("Opis:"))
        layout_formularz.addWidget(self.opis_input)
        layout_formularz.addWidget(self.btn_dodaj)

        # Wykres
        self.wykres = WykresWidget()

        # Dodanie elementów do głównego layoutu
        layout_glowny.addLayout(layout_przyciski)
        layout_glowny.addWidget(self.tabela)
        layout_glowny.addLayout(layout_formularz)
        layout_glowny.addWidget(self.wykres)

        centralny_widget.setLayout(layout_glowny)

    def wczytaj_excel(self):
        sciezka_pliku, _ = QFileDialog.getOpenFileName(
            self,
            "Wybierz plik Excel",
            "",
            "Pliki Excel (*.xlsx)"
        )

        if sciezka_pliku:
            try:
                self.manager.wczytaj_z_excela(sciezka_pliku)
                self.odswiez_tabele()
                self.odswiez_wykres()

            except Exception as blad:
                QMessageBox.critical(
                    self,
                    "Błąd",
                    f"Nie udało się wczytać pliku:\n{blad}"
                )

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

    def wyczysc_dane(self):
        self.manager.czysc_dane()
        self.odswiez_tabele()
        self.odswiez_wykres()

    def odswiez_tabele(self):
        self.tabela.setRowCount(len(self.manager.wydatki))

        for nr_wiersza, wydatek in enumerate(self.manager.wydatki):
            self.tabela.setItem(nr_wiersza, 0, QTableWidgetItem(wydatek.data))
            self.tabela.setItem(nr_wiersza, 1, QTableWidgetItem(wydatek.kategoria))
            self.tabela.setItem(nr_wiersza, 2, QTableWidgetItem(str(wydatek.kwota)))
            self.tabela.setItem(nr_wiersza, 3, QTableWidgetItem(wydatek.opis))

    def odswiez_wykres(self):
        dane = self.manager.oblicz_sume_kategorii()
        self.wykres.rysuj_wykres(dane)


# ==========================================
# URUCHOMIENIE PROGRAMU
# ==========================================

if __name__ == "__main__":
    app = QApplication(sys.argv)

    okno = MainWindow()
    okno.show()

    sys.exit(app.exec_())
