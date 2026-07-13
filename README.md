# Csapat naptár

Szabadság / betegszabadság / heti munkarend nyilvántartó a csapat számára (megtekintésre).

**FONTOS:** amíg csak minta adat van benne, publikus repóban tesztelheted. Mielőtt valódi
dolgozói adatot (főleg betegszabadságot) töltesz bele, vedd le a nyilvános oldalt, vagy
válts privát/belső hosztingra (GDPR — egészségügyi adat).

## Feltöltés GitHub-ra (böngészőből, git nélkül)

1. Menj a [github.com/new](https://github.com/new) oldalra, hozz létre egy publikus repót
   (pl. `csapat-naptar`), ne pipáld be a README-t.
2. A repó oldalán kattints az **"uploading an existing file"** linkre.
3. Húzd be ennek a mappának a tartalmát: az `index.html` fájlt (a `forras` mappa nem kell
   fel a webre, az csak neked kell a frissítéshez — azt ne töltsd fel, vagy külön commitban,
   nem számít, csak az `index.html` kell a gyökérbe).
4. Commit.
5. **Settings → Pages** → Source: *Deploy from a branch* → Branch: `main` / `/(root)` → Save.
6. Pár perc múlva elérhető: `https://<felhasználóneved>.github.io/csapat-naptar/`

## Frissítés (naponta ezt használd — nincs szükség Excelre vagy scriptre)

1. Nyisd meg az élő oldalt, kattints jobb fent a **"Szerkesztés"** gombra.
2. Add meg a jelszót: `vezeto2026` (ezt megváltoztathatod — szólj, vagy keresd meg a
   `forras/generate_html.py` fájlban az `EDIT_PASSWORD` sort, és futtasd újra a scriptet).
3. Kattints bármelyik cellára (napi jelenlét, heti beosztás), és válaszd ki az új kódot a
   felugró listából. Az Éves összesítőben a "Keret" mezőt is közvetlenül átírhatod.
4. Ha végeztél, kattints a **"Mentés"** gombra — ez letölt egy új `index.html` fájlt a
   frissített adatokkal.
5. Töltsd fel ezt a letöltött `index.html`-t a GitHub repóba (ugyanoda, felülírja a régit).

## Ritkán használt: teljes újraépítés Excelből

Ha új évet kell hozzáadni, vagy elölről akarod kezdeni az egészet, a `forras` mappában lévő
`csapat_nyilvantartas.xlsx`-et és a `build_xlsx.py` / `generate_html.py` scripteket
használhatod (Python + openpyxl szükséges):
```
cd forras
python3 build_xlsx.py      # xlsx alap felépítése (évek, ünnepnapok, rotáció)
python3 generate_html.py   # xlsx -> index.html
```
Az így kapott `csapat_naptar.html`-t nevezd át `index.html`-re és töltsd fel.

## Tartalom

- `index.html` — a kész weboldal; a "Szerkesztés" móddal itt módosítod a napi adatokat
- `forras/csapat_nyilvantartas.xlsx` — a kiinduló adatforrás (csak a teljes újraépítéshez kell)
- `forras/generate_html.py` — script, ami az xlsx-ből legyártja a kezdő index.html-t
- `forras/build_xlsx.py` — az xlsx sablon generátora (évek, ünnepnapok, műszakrotáció)
