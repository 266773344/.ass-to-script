# Subtitle to DOCX Converter

A simple desktop application to convert subtitle files in `.ass` (Advanced SubStation Alpha) format to neatly formatted `.docx` (Microsoft Word) documents.

## Features

-   **Easy to Use:** A simple graphical user interface that requires no technical knowledge.
-   **Batch Processing:** Convert multiple `.ass` files in a single operation.
-   **Smart Formatting:**
    -   Automatically groups consecutive dialogue lines from the same speaker.
    -   Aligns all dialogue text for a clean, readable transcript.
-   **Standalone:** No need to install Python or any libraries. Just download and run.

## How to Use

1.  Navigate to the `dist` folder.
2.  Double-click the `Subtitle Converter.exe` file to launch the application.
3.  Click **"1. Select Subtitle Files (.ass)"** to choose one or more files.
4.  Click **"2. Select Output Folder"** to choose where to save the converted `.docx` files.
5.  Click **"3. Convert"** and wait for the process to complete.

The converted files will appear in the output folder you selected, with the same name as the original files but with a `.docx` extension.

---

## For Developers

If you want to run the script from source or modify it, you will need Python 3.x.

### Running from Source

1.  Clone this repository.
2.  Install the required packages:
    ```bash
    pip install -r requirements.txt
    ```
3.  Run the application:
    ```bash
    python main.py
    ```

### Building the Executable

To rebuild the standalone executable, you need `pyinstaller`.

1.  Install `pyinstaller`:
    ```bash
    pip install pyinstaller
    ```
2.  Run the build command from the project's root directory:
    ```bash
    pyinstaller --onefile --windowed --name="Subtitle Converter" main.py
    ```
The final executable will be located in the `dist` folder.
