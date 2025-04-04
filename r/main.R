# --- Configuration ---
# How many days back to check? 1 means only today's emails received since midnight.
days_back <- 1
# Set to TRUE to include email body, FALSE to exclude (faster, less data)
include_body <- TRUE
# Specify the folder path if not the main Inbox
# Examples: "Inbox", "Your Subfolder", "Another Mailbox/Inbox"
# Use NULL to default to the primary account's Inbox
outlook_folder_path <- NULL # Or e.g., "Inbox/MyProjectEmails"

# --- Dependencies ---
library(RDCOMClient)
library(dplyr)     # Optional: For easier data manipulation (if creating a data frame)
library(lubridate) # Optional: For easier date handling
library(rvest)

# --- Main Script ---

cat("Starting Outlook email scraping process...\n")

# Define today's date range
# We want emails received from midnight 'days_back' ago until now.
start_date <- Sys.Date() - days(days_back - 1) # Start of the period (e.g., today's midnight)
# Convert to character format suitable for Outlook filtering if needed,
# but often R-side filtering is easier for POC
# Example format: format(start_date, "%m/%d/%Y %H:%M %p")

cat("Targeting emails received on or after:", format(start_date, "%Y-%m-%d"), "\n")

extracted_data <- list() # Initialize a list to store email data

tryCatch({
  # 1. Create Outlook Application COM object
  outlookApp <- COMCreate("Outlook.Application")
  
  # 2. Get MAPI Namespace
  namespace <- outlookApp$GetNamespace("MAPI")
  
  # 3. Access the target folder
  targetFolder <- NULL
  if (is.null(outlook_folder_path) || outlook_folder_path == "Inbox") {
    cat("Accessing default Inbox folder...\n")
    # olFolderInbox = 6
    targetFolder <- namespace$GetDefaultFolder(6)
  } else {
    cat("Accessing custom folder path:", outlook_folder_path, "...\n")
    # Split path and navigate
    folder_parts <- strsplit(outlook_folder_path, "/")[[1]]
    currentFolder <- namespace$Folders(1)$Folders(folder_parts[1]) # Start from top level folder in default store
    if (length(folder_parts) > 1) {
      for (i in 2:length(folder_parts)) {
        currentFolder <- currentFolder$Folders(folder_parts[i])
      }
    }
    targetFolder <- currentFolder
  }
  
  if (is.null(targetFolder)) {
    stop("Could not access the specified folder.")
  }
  cat("Successfully accessed folder:", targetFolder$Name(), "\n")
  
  # 4. Get emails from the folder
  emails <- targetFolder$Items()
  cat("Total items in folder:", emails$Count(), "\n")
  
  # Optional: Use Outlook's Restrict method for efficiency (more complex)
  # filter_string <- paste0("[ReceivedTime] >= '", format(start_date, "%m/%d/%Y"), "'")
  # cat("Applying filter:", filter_string, "\n")
  # emails <- emails$Restrict(filter_string)
  # cat("Emails after filtering:", emails$Count(), "\n")
  # Note: Restrict can be faster but date/time formatting is tricky.
  # For POC, we will retrieve recent items and filter in R.
  
  # Sort emails by received time (descending) - makes processing recent ones first
  emails$Sort("[ReceivedTime]", TRUE)
  
  # 5. Iterate and Filter emails (R-side filtering for simplicity in POC)
  # Process up to a reasonable number (e.g., last 200) to avoid iterating huge inboxes
  num_to_check <- min(emails$Count(), 200)
  cat("Checking the latest", num_to_check, "emails for matches...\n")
  
  for (i in 1:num_to_check) {
    email <- emails$Item(i)
    
    # Ensure it's actually a mail item (folders can contain other types)
    # Check class using Name property of the Class property (less direct check)
    # Or use tryCatch around property access
    is_mail <- FALSE
    tryCatch({
      # Accessing a mail-specific property like Subject should work if it's mail
      subject_test <- email$Subject()
      is_mail <- TRUE
    }, error = function(e){
      # Not a mail item or error accessing property
    })
    
    if (!is_mail) next # Skip non-mail items
    
    # --- Enhanced Mail Item Check ---
    item_class <- tryCatch(email$Class(), error = function(e) NULL)
    if (is.null(item_class) || item_class != 43) { # olMail = 43
      next
    }
    
    # Extract Received Time and convert carefully
    received_time_com <- email$ReceivedTime()
    # print(paste("Raw Received Time Object Class:", class(received_time_com)))
    # if(inherits(received_time_com, "COMDate")) { print(paste("COMDate Value:", as.numeric(received_time_com))) }
    
    received_datetime <- NULL # Use POSIXct to preserve time and timezone info
    
    if(inherits(received_time_com, "POSIXct")){
      # If RDCOMClient already converted it correctly (less likely for COMDate)
      received_datetime <- received_time_com
      # Ensure it has the correct local timezone attribute for comparisons
      attr(received_datetime, "tzone") <- Sys.timezone()
      
    } else if (inherits(received_time_com, "Date")) {
      # If it came as a Date object, convert to POSIXct (start of day in local time)
      received_datetime <- as.POSIXct(received_time_com, tz = Sys.timezone())
      
    } else if (is.character(received_time_com)){
      # Attempt parsing if it's a string (unlikely here, but keep for robustness)
      received_datetime <- tryCatch(
        parse_date_time(received_time_com,
                        orders=c("mdy IMS p", "Ymd HMS", "Y-m-d H:M:S"), # Add more formats if needed
                        tz=Sys.timezone()), # Assume string represents local time if no offset specified
        error = function(e) NULL)
      
    } else if (inherits(received_time_com, "COMDate")) {
      # --- THIS IS THE KEY FIX ---
      # Handle COMDate explicitly using the OLE Automation epoch
      com_value <- as.numeric(received_time_com) # Get the numeric value (e.g., 45751.49)
      
      # OLE Automation Epoch Date: 1899-12-30
      # Important: Create the epoch as POSIXct UTC to avoid ambiguity
      ole_epoch_utc <- as.POSIXct("1899-12-30 00:00:00", tz = "UTC")
      
      # Calculate the actual datetime in UTC by adding the duration
      # Using lubridate::duration is safer than multiplying by 86400 due to potential leap seconds etc.
      received_datetime_utc <- ole_epoch_utc + lubridate::duration(com_value, "days")
      
      # Convert the calculated UTC time to the user's local time zone
      # This ensures comparisons with Sys.Date() work correctly
      received_datetime <- lubridate::with_tz(received_datetime_utc, tzone = Sys.timezone())
      # --- END OF KEY FIX ---
      
    } else {
      # Fallback attempt for other unexpected types
      cat("Warning: Unhandled/unexpected date type received:", class(received_time_com), "\n")
      # Try a generic conversion, but likely to fail or be wrong
      received_datetime <- tryCatch({
        temp_dt <- as.POSIXct(received_time_com)
        attr(temp_dt, "tzone") <- Sys.timezone() # Assume local if conversion works
        temp_dt
      }, error = function(e) NULL)
    }
    
    # Now derive the Date part (in local time) for filtering comparison
    received_date <- NULL
    if (!is.null(received_datetime) && !is.na(received_datetime)) {
      # as.Date respects the timezone attribute of the POSIXct object
      received_date <- as.Date(received_datetime)
    } else {
      cat("Warning: Could not parse date for email with subject:", email$Subject(), "\n")
      # Decide how to handle: skip this email, assign a default date, etc.
      # For now, we'll skip it in the filter by leaving received_date NULL
    }
    
    
    # Check if the email is within the desired date range
    if (!is.null(received_date) && received_date >= start_date) {
      
      email_subject <- tryCatch(email$Subject(), error=function(e) "?? Unknown Subject ??")
      cat("Processing email:", email_subject, "(Received:", format(received_datetime, "%Y-%m-%d %H:%M:%S %Z"), ")\n")
      
      # --- Retrieve Body Content (Using rvest for HTML - with DIAGNOSTICS) ---
      email_body_content <- NA_character_
      if (include_body) {
        cat("   - Attempting body retrieval...\n") # DIAGNOSTIC
        body_format <- tryCatch(email$BodyFormat(), error = function(e) {
          cat("   - Warning: Failed to get BodyFormat. Assuming Plain.\n") # DIAGNOSTIC
          1 # olFormatPlain = 1
        })
        cat("   - Detected BodyFormat:", body_format, "(1=Plain, 2=HTML, 3=RTF)\n") # DIAGNOSTIC
        
        # Attempt 1: Get Plain Text Body
        plain_body_text <- NULL
        email_body_content <- tryCatch({
          plain_body_text <- email$Body()
          if (!is.null(plain_body_text) && nzchar(trimws(plain_body_text))) {
            cat("   - Success: Found content in Plain Text Body.\n") # DIAGNOSTIC
            trimws(plain_body_text)
          } else {
            cat("   - Info: Plain Text Body was NULL or empty.\n") # DIAGNOSTIC
            NULL
          }
        }, error = function(e) {
          cat("   - Error retrieving Plain Text Body:", e$message, "\n") # DIAGNOSTIC
          NULL
        })
        
        # Attempt 2: If Plain Text failed/empty, try parsing HTMLBody with rvest
        if (is.null(email_body_content)) {
          cat("   - Attempting HTML Body retrieval and parsing...\n")
          html_body <- NULL # Initialize html_body variable
          html_body_raw <- NULL # Variable to store the raw result
          
          html_body_raw <- tryCatch(email$HTMLBody(), error = function(e) {
            cat("   - Error retrieving HTML Body:", e$message, "\n")
            NULL # Return NULL on error
          })
          
          # <<< --- ADD THIS DIAGNOSTIC LINE --- >>>
          cat("   - Raw HTMLBody result type:", typeof(html_body_raw), "Is NULL:", is.null(html_body_raw), "\n")
          if(!is.null(html_body_raw)) {cat("   - Raw HTMLBody (first 100 chars):", substr(html_body_raw, 1, 100), "\n")}
          
          # Now continue with the check using html_body_raw
          if (!is.null(html_body_raw) && nzchar(trimws(html_body_raw))) {
            html_body <- html_body_raw # Assign to html_body if valid
            cat("   - Success: Retrieved non-empty HTML Body (length:", nchar(html_body), "chars).\n")
            
            # ... (rest of the rvest parsing logic using html_body) ...
            
          } else {
            # This branch will now be taken if html_body_raw was NULL or empty
            cat("   - Info: HTML Body was NULL or empty (based on raw check).\n")
          }
        } # End HTML Body attempt
        
        
        # Final check and warning
        if (is.null(email_body_content) || !nzchar(email_body_content)) {
          cat("   - Warning: Final body content is NA or empty.\n") # DIAGNOSTIC
          email_body_content <- NA_character_ # Ensure it's NA if empty/failed
        } else {
          cat("   - Final: Storing extracted body content.\n") # DIAGNOSTIC
          # Optional truncation can be added here if needed
        }
      } else {
        cat("   - Skipping body retrieval (include_body is FALSE).\n") # DIAGNOSTIC
      }
      # --- End of Body Retrieval ---
      
      
      email_data <- list(
        Sender = tryCatch(email$SenderName(), error=function(e) NA_character_),
        Subject = email_subject,
        ReceivedTime = received_datetime,
        Body = email_body_content # Use the final extracted content (or NA)
      )
      extracted_data[[length(extracted_data) + 1]] <- email_data
      
    } else if (!is.null(received_date) && received_date < start_date) {
      # ... (rest of loop) ...
    }
    # Optional delay
    Sys.sleep(0.05)
    
  }
  
  cat("Finished processing emails. Found", length(extracted_data), "emails matching the criteria.\n")
  
}, error = function(e) {
  cat("An error occurred:\n")
  print(e$message)
  # Consider adding more robust error logging
}, finally = {
  # Clean up COM objects (Optional but good practice, R's GC usually handles it)
  # It's difficult to explicitly release with RDCOMClient, rely on R's garbage collection
  # rm(emails, targetFolder, namespace, outlookApp)
  # gc() # Trigger garbage collection
  cat("Outlook scraping process finished.\n")
})

# --- Output for LLM ---

# Option 1: Simple concatenated text block (good for direct LLM prompt)
llm_input_text <- ""
if (length(extracted_data) > 0) {
  for (email in extracted_data) {
    llm_input_text <- paste0(llm_input_text,
                             "--- Email Start ---\n",
                             "From: ", email$Sender, "\n",
                             "Subject: ", email$Subject, "\n",
                             "Received: ", format(email$ReceivedTime, "%Y-%m-%d %H:%M:%S"), "\n",
                             ifelse(include_body && !is.na(email$Body), paste0("\nBody:\n", trimws(email$Body), "\n"), ""),
                             "--- Email End ---\n\n")
  }
  # Print or save the text
  cat("\n--- Text for LLM ---\n")
  cat(llm_input_text)
  # writeLines(llm_input_text, "daily_emails_for_llm.txt")
} else {
  cat("\nNo emails found for the specified period to process.\n")
  llm_input_text <- "No new emails found for today."
}


# Option 2: Data Frame (useful for further R processing before LLM)
if (length(extracted_data) > 0) {
  # Need to handle potential NULLs or varying list structures carefully if using bind_rows
  # Safest is often a loop to build the data frame
  email_df <- do.call(rbind, lapply(extracted_data, function(e) {
    data.frame(
      Sender = e$Sender %||% NA_character_,
      Subject = e$Subject %||% NA_character_,
      ReceivedTime = e$ReceivedTime %||% as.POSIXct(NA),
      Body = ifelse(include_body, e$Body %||% NA_character_, NA_character_),
      stringsAsFactors = FALSE
    )
  }))
  
  # Define %||% inline if not using rlang/purrr
  `%||%` <- function(a, b) if (is.null(a) || length(a) == 0) b else a
  
  cat("\n--- Data Frame Summary ---\n")
  print(head(email_df))
  # You could save this df to CSV, etc.
  # write.csv(email_df, "daily_emails.csv", row.names = FALSE)
}
