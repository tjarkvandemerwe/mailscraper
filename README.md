# R Script for Daily Outlook Email Extraction

## Overview

This R script automates the process of extracting emails received within a specified recent period (e.g., daily) from a Microsoft Outlook client installed on a Windows machine. Its primary purpose is to gather relevant email information (Sender, Subject, Received Time, Body) and format it into a text block suitable for feeding into a Large Language Model (LLM) for tasks like summarization or action item generation.

This script was developed as a Proof of Concept (POC) focusing on rapid setup using the `RDCOMClient` package.

## How it Works

The script leverages the `RDCOMClient` package in R, which allows interaction with Microsoft Office applications via the Component Object Model (COM) interface on Windows. It performs the following steps:

1.  **Connects** to the locally running/available Outlook application instance associated with the current Windows user.
2.  **Accesses** the specified email folder (defaults to the main Inbox of the primary account).
3.  **Retrieves** and sorts email items by received time (newest first).
4.  **Filters** items to include only those received on or after a configurable start date (e.g., today).
5.  **Extracts** key information for each filtered email:
    *   Sender Name
    *   Subject
    *   Received Time (correctly handles `COMDate` objects and converts to local time).
    *   Body Content:
        *   Attempts to get the plain text body (`$Body`).
        *   If plain text is unavailable/empty, attempts to get the HTML body (`$HTMLBody`).
        *   Uses the `rvest` package to parse the HTML and extract text content, stripping HTML tags.
        *   Handles cases where body content retrieval might fail (see Limitations).
6.  **Formats** the extracted data into a single text string, with clear delimiters between emails, ready for LLM processing.
7.  **Outputs** status messages to the console during execution.

## Requirements

*   **Operating System:** **Windows** (due to reliance on COM and `RDCOMClient`).
*   **Software:**
    *   R (tested with version 4.4.1)
    *   Microsoft Outlook (Desktop Client): Must be installed, configured with the target email account, and accessible by the user running the script.
*   **R Packages:**
    *   `RDCOMClient`: for connecting an retrieving outlook mail
    *   `lubridate`: date handling
    *   `rvest`: html parsing
    *   `dplyr`: data wrangling

    Install required packages in R:
    ```R
    devtools::install_github("omegahat/RDCOMClient")
    install.packages("lubridate")
    install.packages("rvest")
    install.packages("dplyr")
    ```

## Configuration

Adjust the following variables at the top of the `scrape_outlook.R` script:

*   `days_back`: Number of days back to include emails from (1 = today only since midnight).
*   `include_body`: Set to `TRUE` to extract email bodies, `FALSE` to skip (faster).
*   `outlook_folder_path`: Set to `NULL` for the default Inbox, or specify the path like `"Your Subfolder"` or `"Another Mailbox/Inbox"`. Use `/` as the separator. *Note: Folder names might be language-dependent if not using the default Inbox.*

## Usage

1.  Ensure all requirements are met (Windows, Outlook configured, R packages installed).
2.  Configure the variables in the script as needed.
3.  Run the script from RStudio or via the command line:
    ```bash
    Rscript scrape_outlook.R
    ```
4.  The script will print progress messages to the console.
5.  The extracted email data, formatted for an LLM, will be stored in the `llm_input_text` variable within the R environment and printed to the console at the end (if emails are found).

## Output

*   **Console Logs:** Status messages indicating folder access, email processing, warnings, and errors.
*   **Formatted Text:** A single string (`llm_input_text`) containing the concatenated details of the extracted emails, suitable for pasting into an LLM prompt or sending via an API. Example format:

    ```
    --- Email Start ---
    From: Sender Name
    Subject: Example Subject
    Received: YYYY-MM-DD HH:MM:SS TZ

    Body:
    This is the extracted email body content...
    --- Email End ---

    --- Email Start ---
    ...
    ```

## Limitations & Caveats

*   **Windows Only:** This script will *not* run on macOS or Linux.
*   **Requires Outlook Client:** Depends on a configured local Outlook installation. It does not interact directly with Exchange/Microsoft 365 via APIs independently of the client.
*   **Authentication:** Relies implicitly on the authentication context of the Outlook profile configured for the logged-in Windows user. No explicit credentials are handled in the script.
*   **Body Content Retrieval:** As discovered during development, for certain emails (particularly those containing complex embedded content like pasted screenshots), the Outlook COM interface (`$Body` and `$HTMLBody` properties accessed via `RDCOMClient`) may return empty content even if the email appears to have text/images in Outlook. In such cases, the script will correctly report `NA` for the body. This appears to be a limitation of the data exposed via COM for those specific items.
*   **Error Handling:** Basic `tryCatch` blocks are included, but error handling could be more robust for production use.
*   **Scheduling:** Running this automatically (e.g., via Windows Task Scheduler) requires careful configuration of the user account and may be unreliable if the user is not logged in or if Outlook has issues running automated tasks in the background.

## Potential Future Improvements

*   Migrate to using the **Microsoft Graph API** (e.g., via the `Microsoft365R` package) for a more robust, cross-platform solution that doesn't depend on the Outlook client.
*   Implement more sophisticated error handling and logging.
*   Add configuration via external files (e.g., YAML, JSON).
*   Integrate directly with an LLM API (e.g., using `httr2`).
*   Add options for handling attachments.

---

*Date: 2025-4-4*