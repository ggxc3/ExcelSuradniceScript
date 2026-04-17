# Excel Súradnice Script (Electron + TypeScript)

Desktop aplikácia pre spracovanie Excel súradníc bola kompletne prerobená na **Electron** s **TypeScript** backendom aj UI.

## Funkcionalita

Aplikácia zachováva pôvodné správanie:
- výber vstupného a výstupného Excel súboru,
- výber hárku,
- definícia stĺpcov vo formáte `A-N,B-E,C-V`,
- spracovanie rozsahu riadkov,
- formátovanie súradníc `N/E`,
- formátovanie výšky `V`,
- odstraňovanie riadkov s neplatnými súradnicami.

## Spustenie lokálne

```bash
npm install
npm run dev
```

## Testy

```bash
npm test
```

## Build

```bash
npm run build
```

## Windows release (portable EXE)

```bash
npm run dist:win
```
