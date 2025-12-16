# Tableau PDF/PNG Export Tool v1.4.4

A desktop application with a graphical user interface (GUI) designed to automate the bulk export of views from Tableau Server or Tableau Cloud to either PDF or PNG format.

---

**[Screenshot Placeholder - Replace this with an image or GIF showing the tool's interface]**
*(Recommendation: Add a screenshot here to quickly show users what the tool looks like!)*

---

## Description

This tool streamlines the often tedious process of exporting multiple Tableau views. It offers two main modes: automating exports based on criteria defined in an Excel spreadsheet or exporting a selected set of views directly. It leverages the Tableau Server Client (TSC) library and provides options for filtering, parameterization, file naming, and output organization.

## Key Features

* **Graphical User Interface:** Easy-to-use interface built with PyQt5.
* **Two Export Modes:**
    * **Automate for a list:** Drive exports using an Excel (`.xlsx`) file, processing row by row.
    * **Export All Views Once:** Export a selected list of views directly, ignoring Excel data.
* **Flexible Automation (Excel Mode):**
    * **Excel Data Filtering:** Filter which rows in the Excel file are processed using column values.
    * **Conditional View Exclusion:** Exclude specific Tableau views from export based on data in the corresponding Excel row.
    * **Tableau View Filtering:** Apply a filter directly to the Tableau view based on an Excel column value before exporting.
    * **Parameter Overrides:** Override Tableau Parameter values using static text, values from Excel columns, or values selected in the Excel Data Filter section.
* **Output Customization:**
    * Export formats: **PDF** or **PNG**.
    * **File Naming:** Name output files based on the Tableau view name or dynamically using an Excel column value.
    * **Folder Organization:** Automatically organize output files into one or two levels of subfolders based on data in specified Excel columns.
    * **Numbering:** Optionally prefix output filenames with sequential numbers (e.g., `01_`, `02_`).
    * **Whitespace Trimming:** Optionally attempt to trim excess whitespace from the bottom of exported PDF/PNG files (requires `PyMuPDF` for PDF, `Pillow` for PNG).
* **View Management:**
    * Load view names directly from the specified Tableau workbook.
    * Globally include/exclude specific views from all export operations.
* **Configuration:**
    * Save and load all settings to/from `.ini` configuration files.
    * Quickly access recently used configuration files.
    * Configure Tableau connection details (Server/Cloud URL, Site ID, Personal Access Token) via a dedicated dialog.
* **User Experience:**
    * Real-time progress bar and detailed logging within the application.
    * Optional custom macOS-like theme.
    * Automatic check for newer versions on GitHub upon startup.
    * Desktop notifications on export completion, stoppage, or error (requires `plyer` library).

## Installation

There are two ways to use this tool:

**1. Using Pre-compiled Releases (Recommended)**

* Go to the [**Releases Page**](https://github.com/haythamsoufi/TableauPDF/releases) of this repository.
* Download the executable file (`.exe` for Windows, `.dmg`/`.zip` containing `.app` for macOS) appropriate for your operating system from the latest release.
* No Python installation is required. Extract the download if necessary and run the application.

**2. Running from Source Code**

* **Prerequisites:** Python 3.x installed on your system.
* **Clone:** Clone this repository:
    ```bash
    git clone [https://github.com/haythamsoufi/TableauPDF.git](https://github.com/haythamsoufi/TableauPDF.git)
    cd TableauPDF
    ```
* **Install Dependencies:** Install the required Python libraries:
    ```bash
    pip install -r requirements.txt
    ```
    *(Make sure the `requirements.txt` file exists and contains the necessary packages - see previous response).*
* **Required Assets:** Ensure the `icons` and `styles` folders (containing images and `macos_style.qss`) are present in the same directory as the script.
* **Run:** Execute the Python script:
    ```bash
    python "TableauPDF v1.4.4.py"
    ```
    *(Note: You might consider renaming the `.py` file to not include spaces or version numbers, e.g., `tableau_exporter_gui.py`)*

## Usage

1.  Launch the application (either the executable or via `python`).
2.  Click **"⚙️ Server Config"**: Enter your Tableau Server/Cloud URL, Site ID (leave blank for Default site), and your **Personal Access Token Name** and **Token Secret**. Click OK.
    * *Note: You need to generate a Personal Access Token (PAT) in your Tableau Server/Cloud account settings.*
3.  Enter the exact **Workbook Name** as it appears on Tableau.
4.  Click **"🔍 Test"** to connect, verify the workbook, and load the list of views it contains. A success message will appear if views are loaded.
5.  Select the **Export Mode**:
    * **"Automate for a list"**: Requires configuring the Excel File, Sheet Name, and potentially the Key Field, File Naming, Organize By options, and the Filtering & Conditional Logic section.
    * **"Export All Views Once"**: Simpler mode, exports views selected via the "Select Views" button directly. Ignores Excel settings and Filtering/Logic.
6.  Configure **Output Options**: Choose PDF/PNG, Output Folder, numbering, and trimming.
7.  Use **"📑 Select Views"** to globally exclude any views you *never* want to export.
8.  **Configure Logic (Excel Mode Only):**
    * Click **"➕ Filter"** to add rows that filter the *Excel* data before processing. Select a field and check the values to keep. Optionally check "Apply as Param" to use the first selected filter value as a Tableau Parameter override (Parameter name must be `FilterFieldNameParam`).
    * Click **"➕ Condition"** to add rows that *conditionally exclude* specific Tableau views based on Excel data for that row.
    * Click **"➕ Parameter"** to define Tableau Parameter overrides (can use static text or pull values from Excel columns).
9.  **(Optional) Save Configuration:** Click **"💾 Save Config"** to save your settings to an `.ini` file for later use.
10. Click **"▶️ Start Export"**. Monitor the progress bar and log area.
11. Click **"⏹️ Stop Export"** to gracefully interrupt the process if needed.

## Configuration File (`.ini`)

The tool allows saving and loading settings using `.ini` files. These files typically contain sections like:

* `[Server]`: URL, PAT details, Workbook Name.
* `[Paths]`: Excel file, Sheet Name, Output Folder, Tableau Filter Field.
* `[General]`: Export Mode, Format, Naming, Organization, Numbering, Global Exclusions.
* `[Filters]`: Details for each Excel filter row.
* `[Conditions]`: Details for each conditional view exclusion row.
* `[Parameters]`: Details for each Tableau parameter override row.

## Dependencies

This tool relies on the following Python libraries (install via `pip install -r requirements.txt`):

* `pandas`
* `requests`
* `tableauserverclient`
* `PyMuPDF`
* `Pillow`
* `PyQt5`
* `plyer`
* `packaging`

## Troubleshooting / Notes

* **Personal Access Tokens:** Ensure your PAT is valid and has the necessary permissions on Tableau Server/Cloud to view and export content.
* **Excel File Locked:** If you get errors reading the Excel file, ensure it is not open in Microsoft Excel or another program.
* **Permissions:** The tool needs write permissions for the specified Output Folder and the application data directory (`%APPDATA%\PDFExportTool` on Windows) for saving recent files.
* **Trimming:** PDF/PNG trimming quality depends on the complexity of the view and the accuracy of the libraries (PyMuPDF/Pillow) in detecting content boundaries. Review trimmed files.
* **Log File:** Check the `app.log` file (created in the same directory as the script/executable) for detailed debugging information.

## License

[**Specify Your License Here - e.g., MIT License**]

*(Recommendation: Choose a license like MIT or Apache 2.0 and add a `LICENSE` file to your repository).*