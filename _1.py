from openpyxl import load_workbook
excel = load_workbook('1.xlsx')

class klient:
    def __init__(self, imie_i_nazwisko, pesel, adres_zamieszkania, nr_tel, email, adres_inwestycji, zuzycie_energi):
        self.imie_i_nazwisko = imie_i_nazwisko
        self.pesel=pesel
        self.adres_zamieszkania = adres_zamieszkania
        self.nr_tel = nr_tel
        self.email = email
        self.adres_inwestycji = adres_inwestycji
        self.zuzycie_energi = zuzycie_energi
    def dane_klienta(self):
        print ("\nPodaj imię i nazwisko klienta")
        imie_i_nazwisko = input()
        print ("\nPodaj podaj PESEL klienta")
        pesel = input()
        print ("\nPodaj adres zamieszkania klienta")
        adres_zamieszkania = input()
        print ("\nPodaj podaj numer telefonu klienta")
        nr_tel = input()
        print ("\nPodaj podaj adres email klienta")
        email = input()
    def adres(self):
        print ("\nCzy adres zamieszkania jest adresem pod którym będzie wykonywana inwestycja?")
        taknie = ["Wybierz","1. Tak", "2. Nie"]
        for tn in taknie:
            print (tn)
            tn = int(input())
        if (tn>1): 
            print ("\nPodaj adres inwestycji")
            adres_inwestycji = input()
        else:
            adres_inwestycji = "j/w"
    def zuzycie(self):
        print ("\nPodaj roczne zużycie energii w kWh")
        zuzycie_energi = int(input())

# powitanie
print ("Dzień Dobry!")
print ("Konfigurator pomoże Ci przygotować oferte dla klienta")

# dane klienta

def main():
    # tworzymy dwa obiekty klasy Osoba
    osoba = klient()

osoba.dane_klienta(self)
osoba.adres(self)
osoba.zuzycie(self)



# zużycie energi

#print ("\nPodaj roczne zużycie energii w kWh")
#zuzycie = int(input())



# kierunek domu
excel.active = 0
kierunki = excel.active

print ("\nJakie jest ułożenie budynku względem południa")
ulozenie = ["Wybierz:",kierunki['A1'].value, kierunki['A2'].value ]
for u in ulozenie:
    print (u)
ulozenie = int(input())

# współczynnik produkcji

if (ulozenie>1):
    wspolczynnik = float(kierunki['B2'].value)
else:
    wspolczynnik = float(kierunki['B1'].value)

#wybór pokrycia dachowego
excel.active = 5
dach = excel.active

typ_dachu = ["\nPodaj rodzaj pokrycia dachu:", dach['A2'].value, dach['A3'].value]
for d in typ_dachu:
    print (d)
wybor_pokrycia = int(input())
if (wybor_pokrycia>1):
    wybrany_dach = dach['A3'].value
    cena_dachu = int(dach['C3'].value)
    cena_montazu = int(dach['D3'].value)
else:
    wybrany_dach = dach['A2'].value
    cena_dachu = int(dach['C2'].value)
    cena_montazu = int(dach['D2'].value)

# sugerowana moc instalacji
sugerowana = int(osoba.zuzycie_energi/wspolczynnik)
print ("\nSygerowana moc instalacji wynosi:", sugerowana)

# proponowana
print ("\nPodaj proponowaną moc instalacji w Wp")
proponowana = int(input())

#wybór paneli
excel.active = 1
panele = excel.active

print ("\nJakie chcesz panele?")
lista_paneli = ["Wybierz panel:", panele['A2'].value, panele['A3'].value]
for p in lista_paneli:
    print (p)
wybor_paneli = int(input())
if (wybor_paneli>1):
    model_panelu = panele['A3'].value
    moc_panelu = int(panele['B3'].value)
    cena_panelu = int(panele['C3'].value)
else:
    model_panelu = panele['A2'].value
    moc_panelu = int(panele['B2'].value)
    cena_panelu = int(panele['C2'].value)


# ilość paneli i moc falownika
ilosc_paneli = int(proponowana / moc_panelu)
moc_instalacji = (ilosc_paneli * moc_panelu)
moc_falownika = moc_instalacji

moc_instalacji = moc_instalacji/1000
cena_montazu =int(panele['D2'].value)

# dobranie falownika do mocy
excel.active = 2
falowniki = excel.active

if (moc_falownika <= int(falowniki['B2'].value)):
    model_falownika = falowniki['A2'].value
    cena_falownika = falowniki['C2'].value
elif (moc_falownika <= int(falowniki['B3'].value)):
    model_falownika = falowniki['A3'].value
    cena_falownika = falowniki['C3'].value
elif (moc_falownika <= int(falowniki['B4'].value)):
    model_falownika = falowniki['A4'].value
    cena_falownika = falowniki['C4'].value
elif (moc_falownika <= int(falowniki['B5'].value)):
    model_falownika = falowniki['A5'].value
    cena_falownika = falowniki['C5'].value
elif (moc_falownika <= int(falowniki['B6'].value)):
    model_falownika = falowniki['A6'].value
    cena_falownika = falowniki['C6'].value
else:
    print ("wycena indywidualna")
montaz_falownika = int(falowniki['D2'].value)
# Zabezpieczenia elektryczne
excel.active = 3
zabezpieczenia = excel.active

if (moc_falownika>int(zabezpieczenia['B2'].value)):
    cena_zabezpieczen = int(zabezpieczenia['C2'].value)
else:
    cena_zabezpieczen = int(zabezpieczenia['C3'].value)

# przewody elektryczne
excel.active = 4
przewody = excel.active
# wyliczenia


cena_konstrukcji = ilosc_paneli * cena_dachu
cena_paneli = ilosc_paneli * cena_panelu

cena_przewodow = ((2*ilosc_paneli)+40) * przewody['C2'].value
robocizna = (ilosc_paneli * cena_montazu)+montaz_falownika

# cena instalacji
def oblicz_cene_instalacji ():
    cena_instalacji = int(cena_przewodow + cena_zabezpieczen + cena_falownika+ cena_paneli+cena_konstrukcji+robocizna)

#wyświetlanie danych klienta  
print ("\n\nOferta instalacji fotowoltaicznej o mocy:", moc_instalacji, "kWp")

print ("\n\nprzygotowana dla:")
print (osoba.imie_i_nazwisko)
print ("PESEL:",osoba.pesel)
print ("adres",osoba.adres_zamieszkania)
print ("adres inwestycji:")
print (osoba.adres_inwestycji)
print ("tel:", osoba.nr_tel)
print ("email:", osoba.email)


#wyswietlenie wybranego panelu
print ("\nPanele:", ilosc_paneli," x ", model_panelu)
print ("Model falownika:", model_falownika,)
print("\n Wycena:")
print("\nCena falownika:", cena_falownika,"zł")
print("Cena paneli:", cena_paneli,"zł")
print("Cena konstrukcji:", cena_konstrukcji,"zł")
print("Cena przewodów elektrycznych", cena_przewodow,"zł")
print("Cena zabezpieczeń elektrycznych", cena_zabezpieczen,"zł")
print("Cena montazu: ", robocizna,"zł")
print("\n Łącznie cena netto:", cena_instalacji,"zł")
cena_brutto = float(cena_instalacji * 1.08)
print("\n Łącznie cena brutto:", round(cena_brutto,2), "zł (VAT8%)")