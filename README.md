# Automatic Inspection with GPIB

A Python program for inspecting a power electronic module using GPIB-connected test equipment. It controls a multimeter, an AC power supply, and an electronic load, guides the operator through the inspection in English or Japanese, and writes the results directly to an Excel report.

The program automates instrument settings, measurements, and calculations while the operator handles wiring, breakers, and potentiometer adjustments.

## What it does

- Collects product information, serial numbers, room temperature, and tester details.
- Provides step-by-step inspection prompts in English and Japanese.
- Sets the input voltage and electronic load for each test.
- Records input voltage, input current, input power, output voltage, and output current.
- Calculates output power and efficiency, and records power factor.
- Measures output voltage at different input voltages and load currents.
- Guides current-limit and maximum/minimum output-voltage checks.
- Updates `BatteryCharger.xlsx` and creates a report copy named using the charger serial number.

## Test equipment

The included documentation describes the following setup:

| Instrument | Model | Address used in `main.py` |
| --- | --- | --- |
| Digital multimeter | Agilent 34401A | `GPIB0::8::INSTR` |
| AC power supply | Kikusui PCR2000LA | `GPIB0::1::INSTR` |
| Electronic load | Kikusui PLZ1004W | `GPIB0::2::INSTR` |

A Keysight USB/GPIB interface connects the instruments to the computer. The documentation includes connection diagrams for the SWSB24-10-200 and EHIS1-24-5L battery chargers.

The checked-in script opens one multimeter connection. Instrument commands, test settings, GUI prompts, and spreadsheet cell locations are specific to this inspection setup.

## Required software

The setup described in the documentation uses Windows with:

- **Python 3**, with Tcl/Tk support for the GUI.
- **Keysight IO Libraries Suite**, which provides the VISA backend used by PyVISA.
- **Keysight USB/GPIB interface drivers** for the adapter.
- **PyVISA, openpyxl, tkcalendar, and keyboard**, installed in the Python environment running the script.

The documentation recommends Keysight VISA and specifically advises against installing NI-VISA for this setup. See its troubleshooting section if NI-VISA is already installed.

Install the Python packages with:

```bash
python -m pip install pyvisa openpyxl tkcalendar keyboard
```

`tkinter` comes with a standard Windows Python installation when Tcl/Tk support is selected. The other imports, including `datetime`, `shutil`, `os`, `time`, and `string`, are part of Python's standard library.

Visual Studio Code is the editor suggested in the documentation; it is optional when running the script. Excel is useful for viewing the completed report, but the program edits the workbook through openpyxl.

## Running the inspection

1. Follow the wiring and equipment checks in **AutomaticInspection Documentation.pdf** before running the program.
2. Confirm that the instruments are visible through the VISA software and their GPIB addresses match the table above. If needed, update the `rm.open_resource(...)` calls at the top of `main.py`.
3. Keep `main.py` and `BatteryCharger.xlsx` in the same directory. Back up the workbook before an inspection; the program writes to it directly.
4. Open a terminal in that directory and run:

   ```bash
   python main.py
   ```

5. Select **English** or **日本語**, enter the inspection details, and press Enter to advance through the prompts. Wait for each processing window to close before continuing.
6. Complete the sequence and check the generated `BatteryCharger_<serial>.xlsx` report.

The script opens and resets all three instruments before displaying the language selection. It requires the connected equipment even to reach the GUI.

## Files

| File | Purpose |
| --- | --- |
| `main.py` | GUI, instrument control, inspection sequence, and Excel report generation |
| `BatteryCharger.xlsx` | Inspection workbook used and updated by the program |
| `AutomaticInspection Documentation.pdf` | English and Japanese instructions, wiring diagrams, software requirements, and troubleshooting |

The PDF also describes an executable distribution with an `_internal` folder. Those packaged files are not included here; this repository contains the Python source.

## Notes

Run the program from the directory containing the workbook and close the workbook in Excel before starting. If an instrument cannot be opened, check its address, GPIB connection, adapter driver, and VISA installation.

The date step currently records today's date. Calendar-selection methods are present in the code but are not called by the current prompt sequence.

This program operates the power supply and load during the inspection. Follow the equipment procedure in the documentation, and turn off power before changing connections.
