#' Mailing Engine Class
#' 
#' The approach of the `$write_mails()` method is to send out mails via a SMTP
#' mailserver. A blind (bcc) copy is sent to the sender to have a check on
#' outgoing mails. The `$check_and_move()` method evaluates messages in the IMAP
#' account provided, augmenting the Excel document with information on outgoing
#' mail and moving messages to the "Sent" folder.
#' 
#' A viable alternative to sending mails via SMTP would be to draft mails
#' automatically. Draft mails (in the Drafts folder) would/could be sent out
#' automatically. As promising as respective functionality of the [Microsoft365R
#' package](https://CRAN.R-project.org/package=Microsoft365R) sounds, API
#' restrictions of organizations may inhibit this approach.
#' [RDCOMClient](https://github.com/omegahat/RDCOMClient) sounds like a
#' promising alternative (see
#' [stackoverflow](https://stackoverflow.com/questions/57811999/rdcomclient-create-write-mail-to-drafts-folder-of-specific-account)),
#' but this is a Windows only package.
#' 
#' @importFrom cli cli_alert_info cli_alert_success cli_alert_danger cli_alert_warning
#' @importFrom pbapply pblapply
#' @importFrom R6 R6Class
#' @importFrom openxlsx loadWorkbook read.xlsx getSheetNames
#' @importFrom Microsoft365R get_business_outlook
#' @importFrom mailR send.mail
#' @importFrom mRpostman configure_imap decode_mime_header
#' @importFrom rstudioapi askForPassword
#' @export
Mailing <- R6Class(
  classname = "Mailing",
  
  public = list(
    
    #' @field mailing_id Name of the mailing.
    #' @field wb Keeps xlsx worbook.
    #' @field sheet Name of the sheet with contact information.
    #' @field data A `data.frame` with content of sheet with contact information.
    #' @field tidcol Name of the column (length-one `character` vector) with tid.
    #' @field mailcol Name of the column (length-one `character` vector) with
    #'   Email addresses.
    #' @field outlook Object of class `ms_outlook` with login to outlook account.
    #' @field template The (loaded) template for Emails sent to respondents.
    #' @field from Sender of the mail.
    #' @field bcc BCC recipient.
    #' @field attachment Filename of a file to attach.
    #' @field smtp_server Mailout server.
    #' @field smtp_user Valid user for the mailout server.
    #' @field smtp_port Port of the smtp server to use.
    #' @field imap_url URL of the IMAP mail server.
    #' @field imap_user Username of the IMAP email account.
    #' @field header_style An openxlsx headerStyle for the column layout.
    mailing_id = NULL,
    wb = NULL,
    sheet = NULL,
    data = NULL,
    tidcol = NULL,
    mailcol = NULL,
    outlook = NULL,
    template = NULL,
    from = NULL,
    bcc = NULL,
    attachment = NULL,
    smtp_server = NULL,
    smtp_user = NULL,
    smtp_port = NULL,
    imap_url = NULL,
    imap_user = NULL,
    
    header_style = createStyle(
      fontSize = 12, fontColour = "#FFFFFF", halign = "center",
      fgFill = "#4F81BD", border = "TopBottom", borderColour = "#4F81BD"
    ),
    
    
    #' @param mailing_id The ID of the mailing. Used for new columns.
    #' @param xlsx_file The Excel file with sheet with contact information.
    #' @param sheet Name of the sheet (length-one `character` vector) with
    #'    contact information.
    #' @param mailcol Column of the `sheet` of the `xlsx_file` defining the 
    #'   mail address of the respondent.
    #' @param namecol Column with the name of the respondent (comma separated).
    #' @param funcol Column with the function of the respondent.
    #' @param gendercol Column with the gender of the respondent.
    #' @param template Filename of the template of the email to be sent. Needs
    #'   to be a plain text file. The content of this file will be loaded when
    #'   initializing the class.
    #' @param from Sender of the Email, can be something such as "Donald Duck
    #'   <donaldduck@@ducktown.org>".
    #' @param bcc BCC recipient of the Email, can be something such as "Donald
    #'   Duck <donaldduck@@ducktown.org>"
    #' @param attachment File to be attached. `NULL` (default) if no file shall
    #'   be attached.
    #' @param smtp_server SMTP server.
    #' @param smtp_user User for the SMTP server.
    #' @param smtp_port Port of the SMTP server.
    #' @param imap_url URL of the imap server.
    #' @param imap_user Username.
    #' @param header_style Default header style for new columns.
    initialize = function(mailing_id, xlsx_file, sheet, namecol = "Name", funcol = "Funktion", gendercol = "Geschlecht", tidcol = "tid", mailcol, template, from, bcc, attachment = NULL, smtp_server, smtp_user, smtp_port, imap_url, imap_user){
      stopifnot(
        is.character(mailing_id),
        length(mailing_id) == 1L,
        
        file.exists(xlsx_file),
        
        is.character(sheet),
        length(sheet) == 1L,
        sheet %in% getSheetNames(xlsx_file),
        
        is.character(mailcol),
        length(mailcol) == 1L,
        
        file.exists(template),
        
        is.character(from),
        length(from) == 1L,
        
        is.character(bcc),
        length(bcc) == 1L,
        
        is.character(attachment),
        length(attachment) == 1L,
        file.exists(attachment),
        
        is.character(smtp_server),
        length(smtp_server) == 1L,
        
        is.character(smtp_user),
        length(smtp_user) == 1L,
        
        is.numeric(smtp_port),
        length(smtp_port) == 1L
        
      )
      
      self$mailing_id <- mailing_id
      
      self$wb <- loadWorkbook(xlsxFile = xlsx_file)
      self$sheet <- sheet
      self$data <- read.xlsx(self$wb, sheet = sheet)

      stopifnot(tidcol %in% colnames(self$data))
      self$tidcol <- tidcol

      stopifnot(mailcol %in% colnames(self$data))
      self$mailcol <- mailcol
      
      self$template <- readLines(template)
      self$from <- from
      self$bcc <- bcc
      self$attachment <- attachment
      self$smtp_server <- smtp_server
      self$smtp_user <- smtp_user
      self$smtp_port <- smtp_port
      
      self$imap_url <- imap_url
      self$imap_user <- imap_user
      
      cli_alert_info("checking that required columns are available")
      stopifnot(namecol %in% colnames(self$data))
      stopifnot("Titel.(u.a..Prof.,.Dr.,.PhD)" %in% colnames(self$data))
      stopifnot(gendercol %in% colnames(self$data))

      cli_alert_info("splitting name into surname and forename")
      self$data$surname <- sapply(
        strsplit(self$data[[namecol]], "\\s*,\\s*"),
        `[[`,
        1L
      )
      self$data$forename <- sapply(
        strsplit(self$data[[namecol]], "\\s*,\\s*"),
        `[`,
        2L
      )
      
      cli_alert_info("generating salutation based on title")
      
      title_col <- self$data[["Titel.(u.a..Prof.,.Dr.,.PhD)"]]
      title_col <- gsub("^(.*?)\\s*$", "\\1", title_col)
      title_col <- gsub("^Prof. Dr$", "Prof. Dr.", title_col)
      title_col <- gsub("^Prof\\.$", "Prof. Dr.", title_col)
      title_col <- gsub("^Prof\\.\\s*Dr\\.\\s*Ing\\.$", "Prof. Dr.", title_col)
      title_col <- ifelse(title_col == "Dipl.-Ing.", NA, title_col)
      title_col <- gsub("^Dr. Dr.$", "Dr.", title_col)
      title_col <- gsub("^PD Dr.$", "Dr.", title_col)
      title_col <- ifelse(
        title_col == "Prof. Dr.",
        ifelse(self$data[[gendercol]] == "w", "Professorin", "Professor"),
        title_col
      )
      title_col <- ifelse(is.na(title_col), "", sprintf("%s ", title_col))
      
      salutation <- sprintf(
        ifelse(
          self$data[[gendercol]] == "w",
          "Sehr geehrte Frau %s%s",
          "Sehr geehrter Herr %s%s"
        ),
        title_col,
        self$data$surname
      )
      salutation <- gsub("\u00a0", " ", salutation)
      salutation <- gsub("\\s{2,}", " ", salutation)
      
      if (funcol %in% colnames(self$data)){
        self$data[[funcol]] <- ifelse(
          is.na(self$data[[funcol]]),
          "",
          self$data[[funcol]]
        )
        salutation <- ifelse(
          self$data[[funcol]] == "OB",
          sprintf(
            "%s%s",
            ifelse(
              self$data[[gendercol]] == "w",
              "Sehr geehrte Frau Oberb\u00fcrgermeisterin,<br/>",
              "Sehr geehrter Herr Oberb\u00fcrgermeister,<br/>"
            ),
            salutation
          ),
          salutation
        )
        salutation <- ifelse(
          self$data[[funcol]] == "BM",
          sprintf(
            "%s%s",
            ifelse(
              self$data[[gendercol]] == "w",
              "Sehr geehrte Frau B\u00fcrgermeisterin,<br/>",
              "Sehr geehrter Herr B\u00fcrgermeister,<br/>"),
            salutation
            ),
          salutation
        )
      } else {
        cli_alert_info("no column for adapting salutation for mayors")
      }
      
      if (any(is.na(salutation)))
        cli_alert_warning("found {length(is.na(salutation))} NA values for salutation")
      self$data$salutation <- salutation
      
      self$data$role <- ifelse(
        self$data[[gendercol]] == "w",
        "kommunale Mandatstr\u00e4gerin",
        "kommunalen Mandatstr\u00e4ger"
      )
      self$data$representative <- ifelse(
        self$data[[gendercol]] == "w",
        "Repr\u00e4sentantin",
        "Repr\u00e4sentant"
      )
      
      invisible(self)
    },
    
    #' @details Write mails.
    #' @param subject The subject of the mails to be sent.
    #' @param tls A `logical` value, whether to use TLS for the SMTP connection.
    #' @param dryrun A `logical` value.
    #' @param chunksize Size of chunks, an `integer` value.
    #' @param sleep Delay between mails in seconds, passed into `Sys.sleep()`.
    #' @param jitter Passed into `runif()`, a random delay added to delay sleep.
    #' @param wait Waiting time.
    #' @param personalize In the template of the mail to be sent, fields defined
    #'   by double angle brackets are assumed to be items for personalization.
    #'   Fields defined by the personalize vector are substituted by the
    #'   respective column of the parsed excel sheet.
    #' @param pwd Password for the mail account.
    write_mails = function(subject, personalize = c("salutation", "token"), tls = TRUE, dryrun = TRUE, chunksize = 10L, wait = 0.5, jitter = 1, sleep = 65, pwd = NULL){
      
      if (is.null(pwd))
        pwd <- rstudioapi::askForPassword("Please enter password for Email")
      
      smtp_data <- list(
        host.name = self$smtp_server,
        port = self$smtp_port,
        user.name = self$smtp_user,
        passwd = pwd,
        tls = tls
      )
      
      row_ids <- 1:nrow(self$data)
      f <- unlist(
        lapply(
          unique(ceiling(row_ids / chunksize)),
          rep,
          times = chunksize
        )
      )[row_ids]
      chunks <- split(self$data[[self$tidcol]], f = f)
      
      for (i in 1:length(chunks)){
        cli_alert_info("proceeding to chunk {i} of {length(chunks)}")
        chunk <- chunks[[i]]
        
        started <- Sys.time()
        
        for (id in chunk){
          
          case <- self$data[self$data[[self$tidcol]] == id,]
          if (nrow(case) != 1L)
            stop(sprintf("exactly one case required - not true for %d", id))
          
          mail <- self$template
          for (replace in personalize){
            mail <- gsub(sprintf("<<%s>>", replace), case[[replace]], mail)
          }
            
          body <- paste(mail, collapse = "")
          
          recipient <- strsplit(case[[self$mailcol]], split = "\\s")[[1]]
          recipient <- recipient[nchar(recipient) > 0L]
          recipient <- gsub("^\\s*(|;|,)(.*?)(|;|,|\\.)\\s*$", "\\2", recipient)
          
          # Remove all non-ASCII characters (including zero-width space \u200B)
          recipient <- unlist(lapply(
            recipient,
            function(r)
              if (Encoding(r) != "unknown"){
                iconv(r, Encoding(r), "ASCII", sub = "")
              } else {
                r
              }
          ))

          if (dryrun == TRUE){
            body <- paste(paste(recipient, collapse = "<br/>"), body, sep = "<br/>")
            recipient <- self$bcc
          }

          if (!is.null(recipient)){
            cli_alert_info("sending mail to: {paste(recipient, collapse = ' / ')}")
            worked <- try({
              send.mail(
                from = self$from,
                to = recipient,
                bcc = self$bcc,
                subject = subject,
                body = body,
                encoding = "utf-8",
                attach.files = self$attachment,
                smtp = smtp_data,
                authenticate = TRUE,
                html = TRUE
              )
            })
            if (is(worked) == "try-error"){
              cli_alert_danger("failed to send to: {paste(recipient, collapse = ' / ')}")
            }
          }
          wait_secs <- wait + runif(n = 1, max = jitter)
          cli_alert_info("delay before sending next mail: {wait_secs}")
          Sys.sleep(wait_secs)
          
        }
        duration <- format(Sys.time() - started)
        cli_alert_info("sent messages to {length(chunk)} recipients in {duration}")
        cli_alert_info("sleeping for {sleep} seconds")
        Sys.sleep(time = sleep)
      }
      
      cli_alert_success("mailing finished")
      invisible(self)
    },
    
    #' @details Move Mails sent from a specified mail address to a designated
    #'   folder
    #' @param sender Sender of the Email. Will be looked up in the FROM field
    #'   of the email.
    #' @param from Folder with mails to be moved.
    #' @param to Folder where to put the mails.
    #' @param move A `logical` value.
    #' @param pwd Password for the mail account.
    #' @param buffersize `integer` value passed into `configure_imap()` to avoid
    #'   issues when processing many messages.
    #' @param esearch `logical` value, to avoid issues with a large number of messages.
    #' @param verbose A `logical` value, so that you would see truncation.
    #' @importFrom lubridate dmy
    check_and_move = function(sender, from = "INBOX", to = "Sent", buffersize = 4194304, esearch = TRUE, verbose = TRUE, move = FALSE, pwd = NULL){
      
      if (is.null(pwd))
        pwd <- rstudioapi::askForPassword("Please enter password for Email")
      
      tmp_data <- read.xlsx(self$wb, sheet = self$sheet)
      con <- configure_imap(
        username = self$imap_user,
        password = pwd,
        url = self$imap_url
      )
      con$select_folder(name = from)
      
      mailout_col <- sprintf("%s_mailout", self$mailing_id)
      if (mailout_col %in% colnames(tmp_data)){
        mailout_col_index <- which(colnames(tmp_data) == mailout_col)
      } else {
        mailout_col_index <- ncol(tmp_data) + 1L
        writeData(
          wb = self$wb,
          sheet = self$sheet,
          x = c(mailout_col, rep("", times = nrow(tmp_data))),
          startCol = mailout_col_index,
          startRow = 1L,
          borderStyle = "none",
          headerStyle = self$header_style
        )
      }
      
      matches <- con$search_string(expr = sender, where = "FROM")
      pblapply(
        matches,
        function(i){
          header <- strsplit(
            con$fetch_header(i)[[sprintf("header%d", i)]],
            "\\r\\n"
          )[[1]]
          email_raw <- gsub(
            "^To:\\s*(.*?)$", "\\1",
            header[grep("^To:\\s", header)]
          )
          email <- gsub("^<(.*?)>$", "\\1", strsplit(email_raw, ",\\s*")[[1]])
          
          date_raw <- gsub("^Date:\\s(.*?)$", "\\1", grep("^Date:\\s", header, value = TRUE))
          date_time <- lubridate::parse_date_time(
            gsub("\\s*\\([^)]*\\)$", "", date_raw),
            orders = "a, d b Y H:M:S z"
          )
          date <- as.Date(date_time)
          time <- format(date_time, "%H:%M:%S")

          row_indices <- unique(unlist(sapply(
            email, function(m) grep(m, tmp_data[[self$mailcol]])
          )))
          if (length(row_indices) > 1L){
            cli_alert_warning("Multiple rows with Email: {email}")
          }
          for (row_index in row_indices){
            tmp_data <- read.xlsx(self$wb, sheet = self$sheet)
            new_cell_content <- paste(as.character(date), time, sep = " ")
            old_cell_content <- tmp_data[[mailout_col_index]][[row_index]]
            if (is.na(old_cell_content)) old_cell_content <- ""
            if (nchar(old_cell_content) > 0L){
              new_cell_content <- paste(old_cell_content, new_cell_content, sep = " // ")
            }
            writeData(
              wb = self$wb,
              sheet = self$sheet,
              x = new_cell_content,
              startCol = mailout_col_index,
              startRow = row_index + 1L,
              borderStyle = "none"
            )
          }
        }
      )
      if (move) con$move_msg(matches, to_folder = to)

      invisible(self)
    },
    
    #' @details Check for mail delivery failure, create respective column and move
    #'   mails to trash.
    #' @param trash Trash folder of the Mail account.
    #' @param move A `logical` value.
    #' @param pwd Password for the mail account.
    mail_delivery_failure = function(trash = "Gel&APY-schte Elemente", move = FALSE, pwd = NULL){
      
      if (is.null(pwd))
        pwd <- rstudioapi::askForPassword("Please enter password for Email")
      
      con <- configure_imap(
        username = self$imap_user,
        password = pwd,
        url = self$imap_url
      )
      con$select_folder(name = "INBOX")
      
      failed_mails_index <- con$search_string(
        expr = "The following addresses had permanent fatal errors",
        where = "BODY"
      )
      cli_alert_info("mails with fatal errors in INBOX: {length(failed_mails_index)}")
      
      tmp_data <- read.xlsx(self$wb, sheet = self$sheet)
      failed_col <- paste(self$mailing_id, "delivery_status", sep = "_")
      if (failed_col %in% colnames(tmp_data)){
        failed_col_no <- which(colnames(tmp_data) == failed_col)
        failed <- c(failed_col, tmp_data[[failed_col]])
      } else {
        failed <- c(failed_col, rep("", times = nrow(tmp_data)))
        failed_col_no <- ncol(tmp_data) + 1L
      }
      
      if (length(failed_mails_index) == 0) failed_mails_index <- c()
      if (length(failed_mails_index) > 0L){
        
        failed_mails <- unlist(lapply(
          failed_mails_index,
          function(i){
            body <- strsplit(con$fetch_body(i)[[sprintf("body%d", i)]], "\\r\\n")[[1]]
            err_ix <- grep("\\s+-+ The following addresses had permanent fatal errors\\s-+", body)
            email <- gsub("^<(.*?)>$", "\\1", body[err_ix[1] + 1])
            if (length(email) > 1) cli_alert_warning("Cannot extract mail: {email}")
            if (!grepl("@", email[1])) cli_alert_warning("Does not look like Email: {email[1]}")
            email[1]
          }
        ))

        row_indices <- sapply(
          failed_mails,
          function(m){
            row_index <- grep(m, tmp_data[[self$mailcol]])
            if (length(row_index) != 1L){
              cli_alert_warning("Cannot look up: {m}")
              return(NA)
            }
            row_index
          }
        )
        
        if (any(is.na(row_indices))){
          identified <- !is.na(row_indices)
          failed_mails <- failed_mails[identified]
          row_indices <- row_indices[identified]
        }
        
        for (i in 1L:length(row_indices)){
          if (failed[row_indices[i] + 1L] == "" || is.na(failed[row_indices[i] + 1L])){
            failed[row_indices[i] + 1L] <- failed_mails[i]
          } else {
            failed[row_indices[i] + 1L] <- paste(
              unique(
                c(failed_mails[i], strsplit(failed[row_indices[i] + 1L], " // ")[[1]])
              ),
              collapse = " // "
            )
          }
        }

        if (move) con$move_msg(
          na.omit(failed_mails_index[identified]),
          to_folder = trash
        )
      }
      
      writeData(
        wb = self$wb, sheet = self$sheet,
        x = failed,
        startCol = failed_col_no,
        startRow = 1L,
        borderStyle = "none",
        headerStyle = self$header_style
      )

      failed_mails
    }
  )
)
