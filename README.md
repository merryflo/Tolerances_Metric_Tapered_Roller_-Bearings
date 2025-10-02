# program description
- program to calculate Tolerances for metric radial bearings (except TRB)

# How to generate executable file


VARIANTA BUNA!!!!!
```bash
pyinstaller --name Tolerances_Metric_Radial_Tapered_Roller_Bearings --onefile --windowed --icon "./docs/icon.ico" --add-data "./docs:." --add-data "./.env.prod:." Tolerances_Metric_Radial_Tapered_Roller_Bearings.py
```
var
```bash
pyinstaller --name Tolerances_Metric_Radial_Tapered_Roller_Bearings --onefile --icon "./docs/icon.ico" --add-data "./docs:." --add-data "./.env.prod:." Tolerances_Metric_Radial_Tapered_Roller_Bearings.py
```

# Libraries and Tools
- [pyinstaller](https://github.com/pyinstaller/pyinstaller) - run the packaged program (`.exe` file) without installing a Python interpreter
- 