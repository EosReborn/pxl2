# Csapat naptár

Szabadság / betegszabadság / heti munkarend nyilvántartó. Három szintű hozzáférés:

- **`index.html`** — publikus főoldal, **nem tartalmaz semmilyen személyes adatot**. Ha
  valaki csak rátalál a repóra vagy a domainre, semmit nem lát.
- **`admin.html`** — a teljes nézet (minden dolgozó, szerkesztés, GitHub-mentés). Ez
  **sehonnan nincs belinkelve** — csak azok tudják elérni, akiknek elküldöd a linket
  (`https://<felhasználóneved>.github.io/<repo>/admin.html`). Ezt te és a másik megbízott
  ember használja.
- **`people/<azonosító>.html`** — 15 külön fájl (a 15 dolgozói azonosító alapján:
  24009, 18639, 25856, 33501, 20161, 21319, 35883, 27556, 33373, 21180, 35885, 90008327,
  35884, 29739, 22166), mindegyik **csak az adott dolgozó saját adatát** tartalmazza (más
  kollégája neve/azonosítója sincs benne, a fájlban ténylegesen nincs más szerepeltetve).
  Ezt a linket egyenként küldöd ki emailben/más csatornán mindenkinek.

Ez valódi adatminimalizálás (GDPR szempontból is jobb, mint egy közös fájl): egy dolgozó
linkjében fizikailag nincs benne más kollégája szabadsága vagy betegsége — nem szűrésről van
szó, a fájl tartalma is más. Azonosítószám használata a névhez képest plusz védelem: a link
maga sem árulja el a nevet, ha véletlenül rossz kézbe kerülne.

**A "Mentés" gomb az admin.html-lel együtt automatikusan frissíti az összes érintett
`people/<azonosító>.html` fájlt is** a GitHub-on — nem kell külön lépés, nem kell Excelt
vagy scriptet futtatni. Csak akkor kell a `forras` mappa scriptjeivel teljesen újraépíteni
(lásd lent), ha új azonosítót/dolgozót adsz hozzá, vagy évet kell hozzáadni, vagy az
ünnepnapokat újra kell számolni. Ha valakinek megváltozik az azonosítója/neve, a régi
`people/<régi>.html` linkje érvénytelenné válik és egy új jön létre — ilyenkor az újat kell
kiküldeni.

### Jelszóvédelem (opcionális, dolgozónként külön beállítható)

Az "Éves összesítő" fülön, szerkesztő módban minden kártyán van egy "Személyes oldal
jelszava" mező, ahová beírhatsz egy jelszót az adott dolgozónak (vagy a 🎲 gombbal
generáltathatsz egyet). Ha egy dolgozónak van jelszava, a Mentés gombra kattintva a
`people/<azonosító>.html` fájlja **titkosítva** kerül fel a GitHub-ra — jelszó nélkül az
oldal forráskódjában sem olvasható ki az adat (nem csak egy jelszót kérő felugró ablakról
van szó, maga az adat van összekeverve a jelszóval, böngészőben, PBKDF2 + AES-GCM
titkosítással). Jelszó nélküli dolgozóknál minden a régi módon működik (nincs extra
kattintás, nincs teljesítményveszteség).

Fontos:
- **A jelszót és a linket mindig KÜLÖN csatornán küldd ki** (pl. link emailben, jelszó
  SMS-ben vagy szóban) — ha együtt küldöd őket, a jelszó nem véd semmit.
- A jelszó csak az admin.html-ben (a te böngésződben, mentéskor) és a dolgozó saját
  fejében létezik — sehol máshol nincs eltárolva. Ha a dolgozó elfelejti, az admin.html-en
  új jelszót adhatsz neki (vagy törölheted, ha nem kell védelem).
- Ha elfelejted a jelszót és nincs is beírva sehova, a régi, már feltöltött titkosított
  oldal **nem visszafejthető** — ilyenkor adj meg egy új jelszót az admin.html-en, mentsd
  el, és küldd ki újra a dolgozónak.
- Egy teljes Excel-alapú újraépítés (lásd lent) plaintext (jelszó nélküli) személyes
  oldalakat ír ki — utána egyszer kattints Mentésre az admin.html-en, hogy a jelszavak és
  a titkosítás visszaálljon.

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
   - Egy dolgozó: `https://<felhasználóneved>.github.io/csapat-naptar/people/24009.html`

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
3. Kattints a **"Mentés"** gombra — visszaírja az `admin.html`-t a GitHub repóba, **majd
   automatikusan frissíti mind a 15 `people/<azonosító>.html` fájlt is** ugyanazzal az
   adattal. Egy pillanatig eltarthat (15 külön fájlt ír fel egymás után) — a folyamat
   végén egy visszajelzés mutatja, hány oldal frissült sikeresen.

A "Kivételek" fülön (mind az admin, mind a személyes oldalakon) alul egy összesítő tábla
mutatja dolgozónként és típusonként (Éjjel/Délután/Délelőtt-csere, Vállalati szabadság,
Saját szabadság, Pihenő, Egyéb), hányszor tért el valaki az automatikus rotációtól az adott
évben, plusz egy végösszeg sort is.

**Napi kódok:** `SZ` = vállalati szabadság, `SZMV` = saját szabadság, `B` = betegszabadság,
`UN` = ünnepnap, `P` = pihenőnap. A két szabadságtípus (SZ + SZMV) együtt terheli ugyanazt az
éves "Szabadságkeretet" — a "Hátralévő szabadság" mindkettőt levonja belőle.

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
- `people/*.html` — 15 személyes, adatminimalizált oldal (egyenként kiküldött linkek)
- `forras/csapat_nyilvantartas.xlsx` — a kiinduló adatforrás (csak a teljes újraépítéshez kell)
- `forras/generate_html.py` — script, ami az xlsx-ből legyártja mindhárom szintet
- `forras/build_xlsx.py` — az xlsx sablon generátora (évek, ünnepnapok, műszakrotáció)
