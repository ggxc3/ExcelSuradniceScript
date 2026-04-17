# Excel Súradnice Script (Tauri)

Projekt bol kompletne migrovaný z Go/WinForms na **Tauri**.

## Spustenie

```bash
cargo tauri dev
```

## Testy

```bash
cargo test --manifest-path src-tauri/Cargo.toml
```

## Windows chyba „Systém Windows nemôže získať prístup…“

Táto chyba sa typicky objaví pri **priamej exekúcii stiahnutého nesignovaného EXE** (SmartScreen / antivirus / práva v priečinku).

Release teraz publikuje späť **single-file EXE**:
- `ExcelSuradniceScript-vX.Y.Z-windows-amd64.exe`

Ak sa chyba objaví znova, potrebujeme od používateľa tieto info:
1. presnú cestu, z ktorej sa EXE spúšťa (Downloads/OneDrive/sieťový disk),
2. či EXE po stiahnutí vidno v Properties tlačidlo **Unblock**,
3. názov antivírusu a či nepresunul EXE do karantény,
4. či spustenie z lokálnej cesty `C:\\Tools\\...` funguje,
5. screenshot celej chyby + timestamp.
