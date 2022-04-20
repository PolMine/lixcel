#' Mailing Engine Class
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
    write_mails = function(subject, personalize = c("salutation", "token")){
      
      mail_passwd <- rstudioapi::askForPassword("Please enter password for Email")
      smtp_data <- list(
        host.name = self$smtp_server, port = self$smtp_port,
        user.name = self$smtp_user, passwd = mail_passwd
      )
      
      for (id in self$data[["tid"]]){
        
        case <- subset(self$data, tid == id)
        if (nrow(case) != 1L)
          stop(sprintf("exactly one case required - not true for %d", id))
        
        mail <- self$template
        for (replace in personalize)
          mail <- gsub(sprintf("<<%s>>", replace), case[[replace]], mail)
        body <- paste(mail, collapse = "</br>")
        
        send.mail(
          from = self$from, to = case[[self$mailcol]], bcc = self$bcc,
          subject = subject, body = body, encoding = "utf-8",
          attach.files = self$attachment,
          smtp = smtp_data, authenticate = TRUE,
          html = TRUE
        )
      }
      
      invisible(self)
    },
    
    #' @details Move Mails sent from a specified mail address to a designated
    #'   folder
    #' @param sender Sender of the Email. Will be looked up in the FROM field
    #'   of the email.
    #' @param from Folder with mails to be moved.
    #' @param to Folder where to put the mails.
    check_and_move = function(sender, from = "INBOX", to = "Sent"){
      
      mail_passwd <- rstudioapi::askForPassword("Please enter password for Email")
      
      tmp_data <- read.xlsx(self$wb, sheet = self$sheet)
      con <- configure_imap(
        username = self$imap_user, password = mail_passwd,
        url = self$imap_url
      )
      con$select_folder(name = from)
      
      mailout <- rep("", times = nrow(tmp_data))
      writeData(
        wb = self$wb, sheet = self$sheet,
        x = c(
          sprintf("%s_mailout", self$mailing_id),
          mailout
        ),
        startCol = ncol(tmp_data) + 1L,
        startRow = 1L,
        borderStyle = "none",
        headerStyle = self$header_style
      )
      
      matches <- con$search_string(expr = sender, where = "FROM")
      for (i in matches){
        header <- strsplit(
          con$fetch_header(i)[[sprintf("header%d", i)]],
          "\\r\\n"
        )[[1]]
        email <- gsub(
          "^To:\\s*.*?\\s*<(.*?)>$", "\\1",
          header[grep("^To:\\s", header)]
        )
        print(email)
        row_index <- grep(email, tmp_data[[self$mailcol]])
        if (length(row_index) != 1){
          warning(sprintf("cannot look up email: %s", email))
        } else {
          writeData(
            wb = self$wb, sheet = self$sheet, x = format(Sys.time()),
            startCol = ncol(tmp_data) + 1L, startRow = row_index + 1L,
            borderStyle = "none"
          )
        }
      }
      con$move_msg(matches, to_folder = to)

      invisible(self)
    },
    
    #' @details Check for mail delivery failure, create respective column and move
    #'   mails to trash.
    #' @param trash Trash folder of the Mail account.
    mail_delivery_failure = function(trash = "Gel&APY-schte Elemente"){
      
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

      if (length(failed) > 0L){
        
        failed_mails <- sapply(
          failed_mails_index,
          function(i){
            body <- strsplit(con$fetch_body(i)[[sprintf("body%d", i)]], "\\r\\n")[[1]]
            recipient_line <- grep("To: &lt;.*?&gt;", body, value = TRUE)
            gsub("To:\\s&lt;(.*?)&gt;", "\\1", recipient_line)
          }
        )

        row_indices <- sapply(
          failed_mails,
          function(m){
            row_index <- grep(m, tmp_data[[self$mailcol]])
            if (length(row_index) != 1L) stop(sprintf("Cannot look up %s", m))
            row_index
          }
        )

        failed[row_indices + 1L] <- "FAIL"
        
        con$move_msg(failed_mails_index, to_folder = trash)
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