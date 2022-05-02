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
    initialize = function(mailing_id, xlsx_file, sheet, mailcol, template, from, bcc, attachment = NULL, smtp_server, smtp_user, smtp_port, imap_url, imap_user){
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
      if (!"tid" %in% colnames(self$data))
        stop("presence of column 'tid' is mandatory yet missing")
      
      stopifnot(mailcol %in% colnames(self$data))
      self$mailcol <- mailcol
      
      # self$outlook <- get_business_outlook()
      self$template <- readLines(template)
      self$from <- from
      self$bcc <- bcc
      self$attachment <- attachment
      self$smtp_server <- smtp_server
      self$smtp_user <- smtp_user
      self$smtp_port <- smtp_port
      
      self$imap_url <- imap_url
      self$imap_user <- imap_user
      
      invisible(self)
    },
    
    #' @details Write mails.
    #' @param subject The subject of the mails to be sent.
    #' @param personalize In the template of the mail to be sent, fields defined
    #'   by double angle brackets are assumed to be items for personalization.
    #'   Fields defined by the personalize vector are substituted by the
    #'   respective column of the parsed excel sheet.
    write_mails = function(subject, personalize = c("salutation", "token"), dryrun = TRUE, chunksize = 10L, wait = 65){
      
      mail_passwd <- rstudioapi::askForPassword("Please enter password for Email")
      smtp_data <- list(
        host.name = self$smtp_server, port = self$smtp_port,
        user.name = self$smtp_user, passwd = mail_passwd
      )
      
      row_ids <- 1:nrow(self$data)
      f <- unlist(
        lapply(unique(ceiling(row_ids / chunksize)), rep, times = chunksize)
      )[row_ids]
      chunks <- split(self$data[["tid"]], f = f)
      
      for (i in 1:length(chunks)){
        message("PROCEEDING TO CHUNK ", i)
        chunk <- chunks[[i]]
        
        for (id in chunk){
          
          case <- subset(self$data, tid == id)
          if (nrow(case) != 1L)
            stop(sprintf("exactly one case required - not true for %d", id))
          
          mail <- self$template
          for (replace in personalize)
            mail <- gsub(sprintf("<<%s>>", replace), case[[replace]], mail)
          body <- paste(mail, collapse = "")
          
          recipient <- strsplit(case[[self$mailcol]], split = "\\s")[[1]]
          recipient <- recipient[nchar(recipient) > 0L]
          recipient <- gsub("^(\\s*|;|,)(.*?)(\\s*|;|,)$", "\\2", recipient)
          if (dryrun == TRUE){
            body <- paste(paste(recipient, collapse = "<br/>"), body, sep = "<br/>")
            recipient <- self$bcc
          }

          if (!is.null(recipient)){
            message(
              sprintf(
                "[%s] sending mail (tid %d): %s",
                format(Sys.time()), id, paste(recipient, collapse = " / ")
              )
            )
            worked <- try({
              send.mail(
                from = self$from, to = recipient, bcc = self$bcc,
                subject = subject, body = body, encoding = "utf-8",
                attach.files = self$attachment,
                smtp = smtp_data, authenticate = TRUE,
                html = TRUE
              )
            })
            if (is(worked) == "try-error"){
              message("FAIL: ", paste(recipient, collapse = " / "))
            }
          }
          Sys.sleep(0.5 + runif(1))
        }
        message(sprintf("... taking a %d second break ...", wait))
        Sys.sleep(time = wait)
      }
      
      message("*** MAILING FINISHED ***")
      invisible(self)
    },
    
    #' @details Move Mails sent from a specified mail address to a designated
    #'   folder
    #' @param sender Sender of the Email. Will be looked up in the FROM field
    #'   of the email.
    #' @param from Folder with mails to be moved.
    #' @param to Folder where to put the mails.
    #' @importFrom lubridate dmy
    check_and_move = function(sender, from = "INBOX", to = "Sent", move = FALSE){
      
      mail_passwd <- rstudioapi::askForPassword("Please enter password for Email")
      
      tmp_data <- read.xlsx(self$wb, sheet = self$sheet)
      con <- configure_imap(
        username = self$imap_user, password = mail_passwd,
        url = self$imap_url
      )
      con$select_folder(name = from)
      
      mailout_col <- sprintf("%s_mailout", self$mailing_id)
      if (mailout_col %in% colnames(tmp_data)){
        stop(sprintf("column %s already exists", mailout_col))
      } else {
        mailout_col_index <- ncol(tmp_data) + 1L
      }
      
      writeData(
        wb = self$wb, sheet = self$sheet,
        x = c(mailout_col, rep("", times = nrow(tmp_data))),
        startCol = mailout_col_index,
        startRow = 1L,
        borderStyle = "none",
        headerStyle = self$header_style
      )
      
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
          date <- lubridate::dmy(
            gsub("^(.*?).*?\\s\\d+:\\d{2}:\\d{2}\\s\\+\\d+$", "\\1", date_raw)
          )
          time <- gsub("^.*?.*?\\s(\\d+:\\d{2}:\\d{2})\\s\\+\\d+$", "\\1", date_raw)
          
          row_indices <- unique(unlist(sapply(
            email, function(m) grep(m, tmp_data[[self$mailcol]])
          )))
          if (length(row_indices) > 1){
            warning(sprintf("Multiple rows with Email: %s", email))
          }
          for (row_index in row_indices){
            writeData(
              wb = self$wb, sheet = self$sheet,
              x = paste(as.character(date), time, sep = " "),
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
    mail_delivery_failure = function(trash = "Gel&APY-schte Elemente", move = FALSE){
      
      mail_passwd <- rstudioapi::askForPassword("Please enter password for Email")
      
      con <- configure_imap(
        username = self$imap_user, password = mail_passwd,
        url = self$imap_url
      )
      
      con$select_folder(name = "INBOX")
      failed_mails_index <- con$search_string(
        expr = "Fehler bei der Nachrichtenzustellung",
        where = "BODY"
      )
      
      tmp_data <- read.xlsx(self$wb, sheet = self$sheet)
      failed <- c(
        paste(self$mailing_id, "delivery_status", sep = "_"),
        rep("", times = nrow(tmp_data))
      )

      if (length(failed_mails_index) > 0L){
        
        failed_mails <- unlist(lapply(
          failed_mails_index,
          function(i){
            body <- strsplit(con$fetch_body(i)[[sprintf("body%d", i)]], "\\r\\n")[[1]]
            email <- unique(gsub('^.*?"mailto:(.*?)".*?$', "\\1", grep("mailto:", body, value = TRUE)))
            if (length(email) == 0L){
              recipient_line <- grep("To: (&lt;|<).*?(&gt;|)", body, value = TRUE)
              email <- gsub("To:\\s(&lt;|<)(.*?)(&gt;|>)", "\\2", recipient_line)
              gr <- grep("grossstadtbefragung", email)
              if (length(gr) > 0) email <- email[-gr][1]
            }
            if (length(email) > 1) warning("Cannot extract mail: ", email)
            if (!grepl("@", email[1])) warning("Does not look like Email: ", email[1])
            email[1]
          }
        ))

        row_indices <- sapply(
          failed_mails,
          function(m){
            row_index <- grep(m, tmp_data[[self$mailcol]])
            if (length(row_index) != 1L){
              warning(sprintf("Cannot look up %s", m))
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

        failed[row_indices + 1L] <- failed_mails
        
        if (move) con$move_msg(failed_mails_index[identified], to_folder = trash)
      }
      
      writeData(
        wb = self$wb, sheet = self$sheet,
        x = failed,
        startCol = ncol(tmp_data) + 1L,
        startRow = 1L,
        borderStyle = "none",
        headerStyle = self$header_style
      )

      invisible(self)
    }
  )
)