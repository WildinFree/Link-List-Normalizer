# Link Normalizer (PowerShell Script)

This PowerShell script is a robust and user-friendly tool designed to process large lists of text entries, primarily focusing on extracting, normalizing, and deduplicating website domains. It provides a graphical user interface (GUI) for easy interaction and outputs the cleaned data into multiple structured formats.

## Features

* **Interactive Graphical User Interface (GUI):** The script presents a simple, intuitive window built with WPF (Windows Presentation Foundation) that allows users to easily select input files and output folders.
* **Intelligent Domain Extraction:** It employs a sophisticated regular expression (`combinedRegex`) to parse various input formats and accurately extract domain names. This includes:
    * Standard URLs (e.g., `https://www.example.com/path?query=1`, `http://blog.sub.domain.org`).
    * URLs without schemes (e.g., `www.example.net`, `sub.domain.co.uk`).
    * Raw domain names (e.g., `example.com`, `another-domain.info`).
    * Entries commonly found in `hosts` files (e.g., `0.0.0.0 example.com`, `127.0.0.1 somesite.net`).
    * Patterns from ad-blocking lists (e.g., `||example.com^`, `@@||goodsite.com`).
    * Domains embedded within certain path-like structures (e.g., `/path/to/domain.com/`).
    It intelligently removes common prefixes like `www.`, `ads--`, and leading numbers/hyphens from extracted domains to ensure consistent normalization.
* **High Performance & Scalability:**
    * **Asynchronous Processing:** Leverages PowerShell's `Start-ThreadJob` (for PowerShell 7+) or `RunspacePool` (for older versions) to process input lines concurrently across multiple threads, significantly speeding up the domain extraction and deduplication for large files.
    * **Memory Optimization:** Dynamically adjusts the processing `batchSize` based on available system memory, ensuring efficient resource utilization and preventing out-of-memory errors when handling massive input files.
    * **Streaming Input/Output:** Reads the input file line by line and writes output in chunks, minimizing the memory footprint for extremely large datasets.
* **Robust Deduplication:** Uses a `System.Collections.Generic.HashSet` to efficiently store and identify unique domains, ensuring that the final output contains no duplicates.
* **Comprehensive Output Formats:** Generates three distinct output files in a timestamped folder named after the input file:
    * **Plain Text (`.txt`):** A simple list of all unique, normalized domains, one per line.
    * **Comma-Separated Values (`.csv`):** A file containing each unique domain along with a count of how many times it appeared in the original input, useful for analysis.
    * **JSON (`.json`):** A structured JSON array of unique domains, including metadata like total unique domains and total duplicates, making it easy for other applications to consume.
* **Resume Capability:** The script saves its processing state (`conversion_state.xml`) during execution. If the script is interrupted, it can attempt to resume from the last saved state, avoiding reprocessing already handled lines.
* **Detailed Logging:** Provides comprehensive logging to a dedicated `debug.log` file (located in `%APPDATA%\HostsConverter\`) and an `error.txt` file for troubleshooting. It logs script initialization, UI interactions, processing progress, errors, and final statistics.
* **Real-time Feedback:** Displays status messages directly in the PowerShell console, indicating progress, successful operations, and any encountered issues.
* **Automatic Output Folder Management:** Creates a unique, timestamped output directory for each conversion run, preventing conflicts and keeping results organized.

## How to Use It

1.  **System Requirements:**
    * Windows Operating System (PowerShell is built-in).
    * PowerShell 5.1 or newer. PowerShell 7+ is recommended for optimal performance due to `Start-ThreadJob`.

2.  **Download the Script:**
    * Save the `Link Normalizer.ps1` file to a convenient location on your computer (e.g., `C:\Scripts\Link Normalizer.ps1`).

3.  **Prepare Your Input File:**
    * Create a plain text file (e.g., `my_links.txt`) containing the list of links or domains you wish to normalize. Place one entry per line.

4.  **Run the Script:**
    * Open **PowerShell** (you can search for "PowerShell" in the Start Menu).
    * Navigate to the directory where you saved the script using the `cd` command:
        ```powershell
        cd "C:\Scripts\" # Replace with your script's path
        ```
    * Execute the script:
        ```powershell
        .\Link Normalizer.ps1
        ```

5.  **Interact with the GUI:**
    * A small window titled "Hosts Converter" will appear.
    * **Input File:** Click the "Browse" button next to "Input File" and select your `my_links.txt` file.
    * **Output Folder:** Click the "Browse" button next to "Output Folder" and choose where you want the cleaned files to be saved. The script will remember your last chosen output folder.
    * **Convert:** Click the "Convert" button. The GUI window will close, and the script will begin processing in the PowerShell console.

6.  **Review Results:**
    * Once the script completes, it will display summary statistics in the PowerShell console and indicate the exact paths to the generated output files (`.txt`, `.csv`, `.json`) within the newly created timestamped folder.

## Output Files

For each conversion, the script creates a new, unique folder within your chosen output directory. This folder will be named in the format `[InputFileName] - YYYYMMDD` (e.g., `my_links - 20231027`). Inside, you will find:

* **`[InputFileName] - converted domains.txt`**: Contains a clean, sorted list of all unique domains found, with one domain per line.
* **`[InputFileName] - converted domains.csv`**: A CSV file with two columns: "Domain" and "DuplicateCount". This shows each unique domain and how many times it appeared in your original input file.
* **`[InputFileName] - converted domains.json`**: A JSON file containing an array of unique domains, along with metadata summarizing the total unique domains and total duplicates processed.

## Logging

The script maintains detailed logs to help you understand its operations and troubleshoot any issues:

* **Debug Log:** `%APPDATA%\HostsConverter\debug.log`
    * Records script initialization, UI interactions, file paths, batch processing details, and general progress.
* **Error Log:** `C:\Users\[YourUsername]\Documents\Programming Related Stuff\Chrome Extentions\Porn Blocker - Project Midnight Arrow\Blocker Resources\Powershell Output\Logs\error.txt` (Note: The path for the error log is hardcoded in the script and might need adjustment if you prefer a different location).
    * Captures specific error messages encountered during file operations, UI rendering, or processing.
