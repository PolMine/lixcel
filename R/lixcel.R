ls_cols <- c("tid", "token", "completed", "usesleft")

#' Read and process LimeSurvey table with keys
#' 
#' @examples
#' fname <- system.file(package = "lixcel", "extdata", "csv", "tokens_01.csv")
#' litab <- read_limetab(file = fname)
#' summary(litab)
#' @importFrom tibble as_tibble
#' @importFrom utils read.csv
#' @export
#' @rdname lime
#' @param file File with LimeSurvey Data.
read_limetab <- function(file){
  df_raw <- read.csv(file = file, header = TRUE, sep = ",", quote = "\"")
  df <- df_raw[, ls_cols]
  df$completed <- ifelse(df$completed != "N", df$completed, NA)
  df$completed <- as.Date(df$completed)
  tbl <- as_tibble(df)
  class(tbl) <- c("limetab", class(tbl))
  tbl
}


#' @export
#' @importFrom tibble tibble
#' @rdname lime
#' @param object A `limetab` object.
#' @param ... Further arguments.
summary.limetab <- function(object, ...){
  tibble(
    no = length(which(is.na(object$completed))),
    yes = length(which(!is.na(object$completed)))
  )
}

#' @examples
#' excelfile <- system.file(package = "lixcel", "extdata", "xlsx", "contact_01.xlsx")
#' fname <- system.file(package = "lixcel", "extdata", "csv", "tokens_01.csv")
#' litab <- read_limetab(file = fname)
#' xlsx_sour <- tempfile(fileext = ".xlsx")
#' 
#' lixcel(
#'   excelfile = excelfile,
#'   sheet = "Grundgesagmtheit",
#'   lime = litab,
#'   destfile = xlsx_sour
#' )
#' 
#' fname <- system.file(package = "lixcel", "extdata", "csv", "tokens_02.csv")
#' litab <- read_limetab(file = fname)
#' xlsx_update <- tempfile(fileext = ".xlsx")
#' 
#' lixcel(
#'   excelfile = xlsx_sour,
#'   sheet = "Grundgesagmtheit",
#'   lime = litab,
#'   destfile = xlsx_update
#' )
#' @export
#' @rdname lime
#' @importFrom openxlsx loadWorkbook read.xlsx writeDataTable saveWorkbook
#' @importFrom openxlsx writeData getStyles createStyle addStyle freezePane protectWorksheet
#' @importFrom stats setNames na.omit
#' @importFrom dplyr left_join
#' @param excelfile An Excel file.
#' @param sheet Name of the sheet to augment.
#' @param lime A `limetab` object to feed in.
#' @param destfile Output Excel file.
#' @param wave Specification of wave, will be appended to the column names (tid,
#'   token, completed, usesleft).
#' @param mailcol The column in the Excel sheet with the email address. If 
#'   provided, the tokens are added only if an Email address is available.
lixcel <- function(excelfile, sheet, lime, wave = "", mailcol = NULL, destfile){
  
  cols <- paste(c("tid", "token", "completed", "usesleft" ), wave, sep = "_")
  names(cols) <- ls_cols
  
  wb <- loadWorkbook(file = excelfile)
  
  if (!sheet %in% names(wb)) stop("sheet not available")
  df <- read.xlsx(wb, sheet = sheet)
  
  if (nrow(lime) < nrow(df))
    stop("lime survey data has less rows than excel sheet")
  
  tokencol <- na.omit(df[[cols["token"]]])
  if (length(unique(tokencol)) < length(tokencol))
    stop("tokens are not unique")
  
  if (all(cols %in% colnames(df))){
    
    if (!all(na.omit(df[[cols["tid"]]]) %in% lime[["tid"]]))
      stop(
        "something's wrong: all IDs expected to be in workbook sheet - not true"
      )
    
    li_start <- which(colnames(df) == cols[1L])
    if (!all(
      colnames(df)[li_start:(li_start + length(cols) - 1L)] == cols
      )
    ){
      stop("order of column names not matching")
    }
    
    df_min <- df[, cols[c("tid", "token")]]
    colnames(df_min) <- c("tid", "token")
    df_min <- left_join(
      x = df_min,
      y = limetab,
      by = c("tid", "token")
    )
    
    for (x in c("completed", "usesleft")){
      writeData(
        wb = wb,
        sheet = sheet,
        x = df_min[[x]],
        startCol = which(colnames(df) == cols[x]),
        startRow = 2L
      )
    }


  } else {
    
    if (!is.null(mailcol)){
      has_mail <- !is.na(df[[mailcol]])
      cli_alert_info("assign tokens for {sum(has_mail)} with mail address")
      insert <- cols |>
        lapply(function(col) rep(NA, times = nrow(df))) |>
        as.data.frame()
      colnames(insert) <- cols
      insert[has_mail,] <- lime[1:sum(has_mail),]
    } else {
      insert <- lime
    }
    
    
    headerStyle <- createStyle(
      fontSize = 12, fontColour = "#FFFFFF", halign = "center",
      fgFill = "#4F81BD", border = "TopBottom", borderColour = "#4F81BD"
    )
      
    addStyle(
      wb = wb, sheet = sheet,
      style = headerStyle,
      rows = 1L, cols = 1L:nrow(df),
    )
    
    writeData(
      wb = wb, sheet = sheet,
      x = insert[1L:nrow(df),],
      startCol = ncol(df) + 1L,
      startRow = 1L,
      borderStyle = "none",
      headerStyle = headerStyle
    )
    
    protectWorksheet(
      wb = wb, sheet = sheet,
      protect = TRUE, password = "limesurvey",
      lockFormattingCells = FALSE, lockFormattingColumns = FALSE,
      lockInsertingColumns = TRUE, lockDeletingColumns = TRUE
    )

    for (col in 1L:ncol(df)){
      addStyle(
        wb = wb, sheet = sheet,
        style = createStyle(locked = FALSE),
        cols = col,
        rows = 2L:nrow(df),
        stack = TRUE
      )
    }

    freezePane(
      wb = wb,
      sheet = sheet,
      firstActiveRow = 2L
    )
  }

  saveWorkbook(
    wb = wb,
    file = destfile,
    overwrite = FALSE
  )
}

