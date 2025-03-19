# Excel to Grepolis BBCode Converter

## Description
This Python script provides a graphical user interface (GUI) to convert Excel tables into [Grepolis](https://en.grepolis.com/) BBCode format. It allows users to extract data from an Excel file and format it into a Grepolis-compatible table, preserving optional BBCode tags for players, cities, and alliances based on cell colors.

## Pictures
### GUI
![GUI_example.png](doc%2FGUI_example.png)
### Input
![excel_input_example.png](doc%2FExcel_input_example.png)
### Output
![txt_output_example.png](doc%2Ftxt_output_example.png)
## Features
- **GUI for easy file selection and conversion**
- **Automatic detection of table boundaries**
- **BBCode formatting based on cell colors:**
  - **Yellow (FFFF00):** `[player]...[/player]`
  - **Green (92D050):** `[town]...[/town]`
  - **Blue (00B0F0):** `[ally]...[/ally]`
- **Option to enable a bolded header row**
- **Copy BBCode output to clipboard**
- **Save output as a `.txt` file**

## Requirements
- Python 3.8+
- Required libraries in `requirements.txt` (install with `pip install`):
  ```sh
  pip -r requirements.txt
  ```

## Installation & Usage
### Prebuild .exe file
1. Double Click the `grepolis-table-converter.exe` in the `dist` directory
### Run python script directly
1. **Clone this repo and setup environment and libraries**
   
2. Run the script
   ```sh
   python grepolis-table-converter.py
   ```
3. Click "Select Excel File" to choose your spreadsheet.
4. The script automatically detects the table boundaries.
5. Adjust row/column settings if needed.
6. Enable or disable headline
7. Click "Convert to BBCode" to generate the formatted output.
8. Copy the BBCode or save it as a `.txt` file.

## File Output
- The generated BBCode follows this structure:
  ```
  [table]
  [**]Header1[||]Header2[/**]
  [*]Data1[|]Data2[/*]
  [/table]
  ```

## Notes
- Ensure your Excel file has a clear table structure without excessive empty rows/columns.
- The script reads from the **first sheet** in the Excel file.
- If a cell has a recognized color, it will be wrapped in the corresponding BBCode tag.

## License
This project is open-source under the MIT License.

## Author
Developed by Andreas