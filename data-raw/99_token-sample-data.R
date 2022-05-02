library(password)
library(fs)

dir <- "/PATH/TO/TOKEN/DIR"
header <- "tid,firstname,lastname,email,emailstatus,token,language,validfrom,validuntil,invited,reminded,remindercount,completed,usesleft"
row <- '"1","","","","OK","WWy1TkN","de","","","N","N","0","N","1"'

txt <- paste(
  header, 
  paste(rep(row, times = 6000), collapse = "\n"),
  sep = "\n"
)

df <- read.csv(text = txt)
df$tid <- 1:nrow(df)
df$token <- unique(sapply(1:8000, function (i) password()))[1:6000]

write.csv(x = df, file = path(dir, "token_01.csv"), quote = TRUE, row.names = FALSE)
