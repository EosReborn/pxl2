# Csapat naptár

Szabadság / betegszabadság / heti munkarend nyilvántartó. Három szintű hozzáférés:

- **`index.html`** — publikus főoldal, **nem tartalmaz semmilyen személyes adatot**. Ha
  valaki csak rátalál a repóra vagy a domainre, semmit nem lát.
- **`admin.html`** — a teljes nézet (minden dolgozó, szerkesztés, GitHub-mentés). Ez
  **sehonnan nincs belinkelve** — csak azok tudják elérni, akiknek elküldöd a linket
  (`https://<felhasználóneved>.github.io/<repo>/admin.html`). Ezt te és a másik megbízott
  ember használja.
- **`people/<nev>.html`** — 16 külön fájl, mindegyik **csak az adott dolgozó saját adatát**
  tartalmazza (a saját neve sincs benne másokénak, a fájlban ténylegesen nincs más
  szerepeltetve). Ezt a linket egyenként küldöd ki emailben mindenkinek a saját nevére.

Ez valódi adatminimalizálás (GDPR szempontból is jobb, mint egy közös fájl): egy dolgozó
linkjében fizikailag nincs benne más kollégája szabadsága vagy betegsége — nem szűrésről van
szó, a fájl tartalma is más.

**Amíg placeholder nevek vannak** (`Munkatárs 01` stb.), a `people/` fájlnevek is ilyenek
(pl. `munkatars-01.html`). Ha valódi neveket írsz be az xlsx-be és újragenerálod, a
fájlnevek is a valós nevek alapján készülnek majd — ilyenkor a régi linkek érvénytelenné
válnak, újra ki kell küldened őket.

## Feltöltés GitHub-ra (böngészőből, git nélkül)

1. Menj a [github.com/new](https://github.com/new) oldalra, hozz létre egy **publikus**
   repót (pl. `csapat-naptar`), ne pipáld be a README-t. (A publikus repó nem baj, mert az
   `index.html` üres, az `admin.html` és a `people/*.html` linkjét pedig nem ismeri senki,
   akinek nem küldted el — GitHub Pages-en minden fájl technikailag elérhető a saját URL-jén,
   de senki nem fog rátalálni, ha nincs belinkelve és nem sejti a nevét.)
2. A repó oldalán kattints az **"uploading an existing file"** linkre.
3. Húzd be ide ennek a mappának a tartalmát: `index.html`, `admin.html`, a teljes `people/`
   mappát. (A `forras` mappa nem kell fel a webre, az csak neked kell a frissítéshez.)
4. Commit.
5. **Settings → Pages** → Source: *Deploy from a branch* → Branch: `main` / `/(root)` → Save.
6. Pár perc múlva élesben:
   - Főoldal (üres): `https://<felhasználóneved>.github.io/csapat-naptar/`
   - Admin: `https://<felhasználóneved>.github.io/csapat-naptar/admin.html`
   - Egy dolgozó: `https://<felhasználóneved>.github.io/csapat-naptar/people/munkatars-01.html`

## Frissítés (naponta ezt használd — nincs szükség Excelre vagy scriptre)

**Első alkalommal (egyszeri beállítás), az `admin.html`-en:**

1. GitHub-on hozz létre egy tokent: Settings → Developer settings → Fine-grained tokens →
   New token → csak erre a repóra korlátozva → Permissions: **Contents: Read and write**.
2. Kattints a fogaskerék ikonra (⚙) jobb fent, és töltsd ki: GitHub felhasználóneved, a
   repó neve, branch (`main`), fájl (**`admin.html`** — ezt írd át, alapértelmezésben
   `index.html` van beírva!), és illeszd be a tokent. Mentés. (Csak a te böngésződben marad
   meg.)

**Utána, mindig az `admin.html`-en:**

1. Kattints jobb fent a **"Szerkesztés"** gombra (ha már be van állítva a token, rögtön
   bekapcsol, nem kérdez semmit).
2. Kattints bármelyik cellára, válaszd ki az új kódot. A "Keret" mezőt is közvetlenül
   átírhatod.
3. Kattints a **"Mentés"** gombra — visszaírja az `admin.html`-t a GitHub repóba.

**Fontos:** a `admin.html` szerkesztése és mentése **nem** frissíti automatikusan a
`people/*.html` fájlokat — azok külön fájlok. Ha napi adatot módosítasz, a
"Ritkán használt: teljes újraépítés"-sel (lásd lent) tudod újragenerálni az összes
személyes oldalt is egyszerre a friss adatokkal, és felülírva feltölteni a `people/`
mappát.

## Ritkán használt: teljes újraépítés Excelből (ez frissíti a people/ mappát is)

Ha új évet kell hozzáadni, ünnepnapokat újraszámolni, vagy a napi szerkesztések után az
összes személyes oldalt is frissíteni akarod, a `forras` mappában lévő scripteket
használd (Python + openpyxl szükséges):
```
cd forras
python3 build_xlsx.py      # xlsx alap felépítése (évek, ünnepnapok, rotáció)
python3 generate_html.py   # xlsx -> admin.html + index.html + people/*.html
```
Az így kapott `admin.html`, `index.html` és `people/` mappát töltsd fel újra (felülírva a
régieket).

## Tartalom

- `index.html` — publikus, adatmentes főoldal
- `admin.html` — teljes nézet (csak a megbízott adminoknak küldött link)
- `people/*.html` — 16 személyes, adatminimalizált oldal (egyenként kiküldött linkek)
- `forras/csapat_nyilvantartas.xlsx` — a kiinduló adatforrás (csak a teljes újraépítéshez kell)
- `forras/generate_html.py` — script, ami az xlsx-ből legyártja mindhárom szintet
- `forras/build_xlsx.py` — az xlsx sablon generátora (évek, ünnepnapok, műszakrotáció)
