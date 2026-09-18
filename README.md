# ChromeTabCreator

Jednoduchá desktopová aplikace pro Windows na vytváření zástupců Google Chrome s předdefinovanou sadou karet. Umožňuje jedním kliknutím otevřít libovolnou skupinu webů (např. pracovní nástroje, zpravodajství nebo vývojářské služby).

## Funkce

- Vytváření `.lnk` zástupců přímo na plochu Windows.
- Podpora libovolného počtu URL adres v rámci jednoho zástupce.
- Rychlé předvolby pro často používané weby (ChatGPT, Claude, GitHub apod.).
- Možnost vynutit tmavý režim (`--force-dark-mode`).
- Možnost otevírání v novém okně (`--new-window`).
- Možnost okamžitého otestování zadaných karet bez nutnosti ukládání.

## Požadavky

- Windows 10 / 11
- Python 3.10+
- Google Chrome

## Instalace a spuštění

1. Klonování repozitáře:
```bash
git clone https://github.com/MonsterMarian/ChromeTabCreator.git
cd ChromeTabCreator
```

2. Instalace závislostí:
```bash
pip install -r requirements.txt
```

3. Spuštění aplikace:
```bash
python app.py
```

Pro spuštění bez zobrazení okna příkazového řádku:
```bash
pythonw app.py
```

## Použité knihovny

- **CustomTkinter** – desktopové grafické rozhraní
- **Pillow** – práce s ikonami
- **pywin32** – Windows Shell API pro vytváření zástupců (.lnk)
