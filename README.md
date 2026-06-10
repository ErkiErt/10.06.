# Plastok-Zenith toote- ja materjalibaas

See kaust on valmis GitHubi üles laadimiseks. Peamine väljund on Exceli töövihik:

`outputs/zenith_catalog_database_20260609/Zenith_Kataloogi_Tootebaas_2026-06-09.xlsx`

## Sisu

- `data_sources/` - algallikad: Zenith kataloog, Zenith 2025 hinnakiri Plastokile ja Plastok DATABASE.xlsm.
- `outputs/zenith_catalog_database_20260609/` - valmis Excel ja genereeritud `source_extract.json`.
- `scripts/` - lähteandmete väljavõtte ja Exceli ehitamise skriptid.
- `legacy_reference/` - varasemad tööpõhjad ja kontrollfailid.
- `docs/QUALITY_CHECK.md` - andmetäpsuse ja auditistaatuste kokkuvõte.

## Tööloogika

Plastoki oma materjalibaas on esmane kontroll. Zenithist tellitakse juurde siis, kui Plastokil pole sobivat materjali, mõõtu, sertifikaati, kogust või oma laoseis/materjal ei kata kliendi vajadust.

Failis ei käsitleta oletusi faktidena:

- otsesed allikaväljad tulevad kataloogist, hinnakirjast või Plastok DATABASE failist;
- valikuabi, kasutusfiltrid ja Plastoki normaliseeritud materjaligrupid on märgitud `TULETATUD_*`;
- read, mida PDF tekstikiht või hinnakiri automaatselt ei kinnitanud, on auditilehtedel ja `21_KÄSITSI_KONTROLL` lehel.

## Uuesti genereerimine

Python sõltuvused:

```bash
pip install -r requirements.txt
```

Lähteandmete ekstrakt:

```bash
python scripts/extract_zenith_sources.py
```

Exceli ehitus:

```bash
npm install
npm run build
```

Märkus: Exceli ehitusskript kasutab `@oai/artifact-tool` teeki. Kui see ei ole avalikust npm registrist kättesaadav, ava projekt Codexi tööruumis, kus see sõltuvus on olemas.

## Tundlikkus

Kaust sisaldab hinnakirja ja Plastoki materjalibaasi. Enne avalikku GitHubi üleslaadimist kontrolli, kas repo peab olema privaatne.
