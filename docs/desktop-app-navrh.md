# Návrh: najjednoduchšia cesta na desktop appku

## Odporúčanie (minimum práce): zostať v Go a použiť **Fyne**

Pre tento projekt je najjednoduchšie riešenie **neprepisovať logiku do iného jazyka** a len vymeniť CLI vstupy za malé GUI.

Prečo práve Fyne:
- **Go ostáva rovnaké**: vieš znovu použiť väčšinu existujúcej logiky práce s Excelom.
- **Jednoduchý deploy**: buildneš jeden natívny binárny súbor pre Windows/macOS/Linux.
- **Minimalistický moderný vzhľad**: Fyne má čisté defaultné widgety a dark/light tému.
- **Lightweight workflow**: netreba Node.js runtime ani browser wrapper.

## Čo sa zmení v aplikácii

Súčasná CLI app robí tieto kroky:
1. otvorenie Excel súboru,
2. výber hárku,
3. zadanie formátovania stĺpcov,
4. interval riadkov,
5. uloženie nového súboru.

V GUI sa to dá spraviť ako jeden jednoduchý formulár:
- Tlačidlo „Otvoriť súbor" (file picker)
- Dropdown „Hárok"
- Textové pole „Stĺpce" (napr. `A-N,B-E,C-V`)
- Pole „Od riadku" a „Do riadku"
- Tlačidlo „Spracovať a uložiť"

## Návrh minimálnej architektúry

1. **Vyseparovať doménovú logiku** z `App` (validácia vstupu + `FormatManager`) do samostatnej služby.
2. Spraviť nový súbor `ui_fyne.go`, ktorý túto službu zavolá.
3. Pôvodné CLI môže ostať (napr. pre batch), ale default `main` môže štartovať GUI.

Týmto dostaneš desktop app bez veľkého refaktoru.

## Implementačný plán (1–2 dni)

1. Presunúť parsing vstupu (stĺpce, rozsahy) do samostatných funkcií.
2. Pridať Fyne dependency (`fyne.io/fyne/v2`).
3. Vytvoriť jedno okno s formulárom a progress/error hláškou.
4. Napojiť na existujúci `FormatManager`.
5. Build pre Windows:
   - `go build -ldflags "-H=windowsgui"`

## Prečo nie prepis do iného jazyka (zatiaľ)

- **Tauri/Electron**: UI síce moderné, ale pridáš JS toolchain a väčší build pipeline.
- **Python + Tkinter/PySide**: rýchle prototypovanie, ale distribúcia býva komplikovanejšia než jeden Go binary.

Ak je cieľ „čo najrýchlejšie, čisté, malé“, Go + Fyne má najlepší pomer jednoduchosť/výsledok.

## Bonus: budúce vylepšenia bez komplikácií

- drag & drop Excel súboru,
- zapamätanie poslednej cesty,
- posledné použité nastavenia,
- jednoduchý log panel s výsledkom spracovania.
