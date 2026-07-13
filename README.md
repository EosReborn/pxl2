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

## Frissítés

1. Nyisd meg és szerkeszd a `forras/csapat_nyilvantartas.xlsx` fájlt (napi kódok, heti
   beosztás).
2. Futtasd (Python + openpyxl szükséges hozzá):
   ```
   cd forras
   python3 generate_html.py
   ```
3. Az így frissült `forras/csapat_naptar.html` fájlt nevezd át `index.html`-re, és töltsd
   fel újra a GitHub repóba (ugyanoda, felülírja a régit — vagy `git add/commit/push`, ha
   git-tel dolgozol).

## Tartalom

- `index.html` — a kész weboldal, ezt kell feltölteni
- `forras/csapat_nyilvantartas.xlsx` — a szerkeszthető adatforrás
- `forras/generate_html.py` — script, ami az xlsx-ből legyártja az index.html-t
- `forras/build_xlsx.py` — az eredeti xlsx sablon generátora (csak referenciának)
