ls_cols <- c("tid", "token", "completed", "usesleft")

#' Read and process LimeSurvey table with keys
#' 
#' @examples
#' fname <- system.file(package = "lime", "extdata", "csv", "tokens_01.csv")
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
#' excelfile <- system.file(package = "lime", "extdata", "xlsx", "contact_01.xlsx")
#' fname <- system.file(package = "lime", "extdata", "csv", "tokens_01.csv")
#' litab <- read_limetab(file = fname)
#' xlsx_sour <- tempfile(fileext = ".xlsx")
#' 
#' mix_workbook_n_lime(
#'   excelfile = excelfile,
#'   sheet = "Grundgesagmtheit",
#'   lime = litab,
#'   destfile = xlsx_sour
#' )
#' 
#' fname <- system.file(package = "lime", "extdata", "csv", "tokens_02.csv")
#' litab <- read_limetab(file = fname)
#' xlsx_update <- tempfile(fileext = ".xlsx")
#' 
#' mix_workbook_n_lime(
#'   excelfile = xlsx_sour,
#'   sheet = "Grundgesagmtheit",
#'   lime = litab,
#'   destfile = xlsx_update
#' )
#' @export
#' @rdname lime
#' @importFrom openxlsx loadWorkbook read.xlsx writeDataTable saveWorkbook
#'   writeData getStyles createStyle addStyle
#' @param excelfile An Excel file.
#' @param sheet Name of the sheet to augment.
#' @param lime A `limetab` object to feed in.
#' @param destfile Output Excel file.
lixcel <- function(excelfile, sheet, lime, destfile){
  
  wb <- loadWorkbook(file = excelfile)
  
  if (!sheet %in% names(wb)) stop("sheet not available")
  df <- read.xlsx(wb, sheet = sheet)

  if (nrow(lime) < nrow(df))
    stop("lime survey data has less rows than excel sheet")
  
  if (all(ls_cols %in% colnames(df))){
    
    if (!all(df[["tid"]] %in% lime[["tid"]]))
      stop("something's wrong: all IDs expected to be in workbook sheet - not true")
    
    li_start <- which(colnames(df) == ls_cols[1L])
    if (!all(
      colnames(df)[li_start:(li_start + length(ls_cols) - 1L)] == colnames(lime)
      )
    ){
      stop("order of column names not matching")
    }
    
    token_lime <- setNames(lime[["token"]], lime[["tid"]])
    
    if (!all(df[["token"]] == token_lime[as.character(df[["tid"]])]))
      stop("tokens do not match - stopping as things might have been mixed up!")
    
    for (x in c("completed", "usesleft")){
      writeData(
        wb = wb, sheet = sheet,
        x = setNames(lime[[x]], lime[["token"]])[df[["token"]]],
        startCol = which(colnames(df) == x),
        startRow = 2L
      )
    }


  } else {
    
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
      x = lime[1L:nrow(df),],
      startCol = ncol(df) + 1L,
      startRow = 1L,
      borderStyle = "none",
      headerStyle = headerStyle
    )
    
  }

  saveWorkbook(wb = wb, file = destfile, overwrite = FALSE)
  
}

