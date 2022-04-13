library(lime)

excelfile <- system.file(package = "lime", "extdata", "xlsx", "contact_01.xlsx")
fname <- system.file(package = "lime", "extdata", "csv", "tokens_01.csv")
litab <- read_limetab(file = fname)

mix_workbook_n_lime(
  excelfile = excelfile,
  sheet = "Grundgesagmtheit",
  lime = litab,
  destfile = "~/Lab/gitlab/vielfaltsstudie/inst/extdata/xlsx/contact_02.xlsx"
)
