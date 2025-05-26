# pybaseline_file

**pybaseline_file** is a graphical Python application for batch processing and baseline correction of SWV (Square Wave Voltammetry) data files. It provides a user-friendly interface to select input folders, configure data import options, process multiple `.txt` files, and export results with baseline correction and peak analysis.

## Features

- **Batch Processing:** Analyze all `.txt` files in a selected folder.
- **Baseline Correction:** Uses asPLS (asymmetric penalized least squares) for robust baseline estimation.
- **Peak Detection:** Automatically finds the main peak in each signal.
- **Signal Smoothing:** Applies Savitzky-Golay filter for noise reduction.
- **Export Options:** Save cleaned data as CSV or Excel files.
- **Result Summary:** Generates a summary Excel file with peak voltages, currents, and calculated charges.
- **Visualization:** Saves annotated plots of each processed signal.
- **Progress Tracking:** Displays a progress bar and log during processing.
- **Cross-platform:** Works on Windows, macOS, and Linux.

## Requirements

- Python 3.8+
- [numpy](https://numpy.org/)
- [pandas](https://pandas.pydata.org/)
- [matplotlib](https://matplotlib.org/)
- [pybaselines](https://pybaselines.readthedocs.io/)
- [scipy](https://scipy.org/)
- [openpyxl](https://openpyxl.readthedocs.io/)
- [tkinter](https://docs.python.org/3/library/tkinter.html) (usually included with Python)

Install dependencies with:

```sh
pip install numpy pandas matplotlib pybaselines scipy openpyxl
```

## Usage

1. **Run the Application**

   ```sh
   python pybaseline_file.py
   ```

2. **Select Input Folder**

   - Click `Parcourir` to choose the folder containing your `.txt` files.

3. **Configure Import Options**

   - Choose the column separator (Tabulation, Virgule, Point-virgule, Espace).
   - Choose the decimal separator (Point or Virgule).
   - Select export format: no export, CSV, or Excel.

4. **Start Analysis**

   - Click `Lancer l'analyse` to process all files.
   - Progress and logs will be displayed in the interface.

5. **View Results**

   - Processed data and plots are saved in a new folder named `<input_folder> (results)` next to your input folder.
   - Click `Ouvrir le dossier de résultats` to open the results folder.

## Input File Format

- Each `.txt` file should contain at least two columns: Potential and Current.
- The first row is skipped (assumed to be a header).
- The script expects the data to be in columns 0 and 1.

## Output

- **Plots:** Annotated PNG plots for each file.
- **Cleaned Data:** CSV or Excel files for each input file (optional).
- **Summary Excel:** A summary file with peak voltages, currents, and calculated charges for each electrode.

## How It Works

- **Baseline Correction:** Uses [`aspls`](https://pybaselines.readthedocs.io/en/latest/api/pybaselines.whittaker.html#pybaselines.whittaker.aspls) from `pybaselines.whittaker` to estimate and subtract the baseline.
- **Peak Detection:** Finds the maximum (after smoothing) away from the edges.
- **Charge Calculation:** For each electrode, the charge is calculated as current divided by frequency (50 Hz by default).

## Troubleshooting

- If you get errors about missing modules, ensure all dependencies are installed.
- If files are not processed, check that they match the expected format and encoding (`latin1`).

## License

MIT License. See [LICENCE](LICENCE) for details.

**Note:**  
This application uses the asPLS algorithm from [pybaselines](https://github.com/derb12/pybaselines) for baseline correction.