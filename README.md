# Knygospigiau Newsletter Editor v2.0

Modernus naujienlaiškių redaktorius knygynui "Knygospigiau".

## Versijos

### macOS
- `Knygospigiau Newsletter.app` - pilnai sukompiliuota macOS versija
- Tinka macOS 11.0 (Big Sur) ir naujesnėms versijoms
- Apple Silicon (M1/M2) ir Intel procesoriai

### Windows
- Windows versija ruošiama

## Naujausi pakeitimai (v2.0)

- Visiškai atnaujintas vartotojo sąsajos dizainas
- Patobulintas teksto redaktorius su emoji palaikymu
- Automatinis šventinių akcijų blokų generavimas
- Patobulinta laiškų siuntimo sistema
- Optimizuotas programos veikimas

## Funkcijos

- Rich text redaktorius su HTML formatavimu
- Emoji įterpimo galimybė
- Automatinis šventinių akcijų blokų generavimas
- Laiškų siuntimas el. paštu
- Žurnalo registravimas
- Moderni, minimalistiška vartotojo sąsaja

## Reikalingi failai

Programai veikti reikalingi šie failai:

### Pagrindiniai failai:
- `email_sender_app.py` - pagrindinis programos failas
- `email_sender_ui.py` - vartotojo sąsajos kodas
- `send_emails.py` - laiškų siuntimo funkcijos
- `emoji_picker.py` - emoji pasirinkimo funkcionalumas

### Duomenų failai:
- `data/emoji_data.json` - emoji duomenų bazė
- `data/holiday_templates.json` - šventinių akcijų šablonai
- `data/email_list.txt` - gavėjų el. pašto adresų sąrašas
- `data/email_log.txt` - laiškų siuntimo žurnalas

### Šablonų failai:
- `templates/emoji_picker.html` - emoji pasirinkimo šablonas

### Konfigūracijos failai:
- `requirements.txt` - Python paketų sąrašas
- `README.md` - programos dokumentacija

## Diegimas

1. Atsisiųskite visus reikalingus failus iš GitHub repozitorijos
2. Įdiekite reikalingus paketus:
   ```bash
   pip install -r requirements.txt
   ```
3. Sukurkite `data` aplanką ir įdėkite į jį reikalingus failus:
   - `emoji_data.json`
   - `holiday_templates.json`
   - `email_list.txt` (su gavėjų el. pašto adresais)
4. Paleiskite programą:
   ```bash
   python email_sender_app.py
   ```

## Reikalavimai

- Python 3.11+
- PyQt5
- Pillow
- Jinja2

## Licencija

MIT

## Autorius

Robertas Šimkus 