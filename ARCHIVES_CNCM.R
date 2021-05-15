# Martin Gascon. Data Scientist @ Intertek 
setwd('~/Dropbox/a_intertek/R-macros/') 

# MODIFIABLE VARIABLES ##################################################

outdir<-"/Users/martin/Dropbox/Personal_use_only/varios/R-macros/Walter/"

# END MODIFIABLE VARIABLES ##################################################


# load some libraries
library(readxl)
library(Hmisc)
library(httr)
library(officer)
library(dplyr)
library(purrr)
library(stringr)
## ============================================= QUESTIONS ==================================================
      
# 1. Read from word document
# 2. Remove non-important content
# 3. Separate text into columns
# 4. Remove support columns and re-order
# 5. Save to CSV

# ============================================= RESPONSES ================================================  
      
      
# 1. Read from word document  ##################################################  

# read docx (object)
obj<-read_docx(paste(outdir, "ARCHIVES_CNCM.docx", sep = ""))

# read content from object (only text)
df <- data.frame(docx_summary(obj))

# rename rows sequentially
row.names(df)<-1:nrow(df)

# 2. Remove non-important content ##################################################  

# remove empty rows and those with paragraph in content_type
df<-subset(df,df$text !="")
df<-subset(df,df$content_type =="table cell")

# remove "Haut du formulaire", "Bas du formulaire"
df$text<-gsub(df$text, pattern = "Haut du formulaire",replacement = "")
df$text<-gsub(df$text, pattern = "Bas du formulaire",replacement = "")


# remove from 1st/31th rows the titles (16 mm and 35 mm)
df$text<-gsub(df$text, pattern = "LES BASES DE DONNEES DU CNCM EN BOBINES 16 ET 35 MM16 PELLICULE 16 MM)",replacement = "")
df<-df[-17,]

# rename rows consecutively again
row.names(df)<-1:nrow(df)

# remove (1) item in 35 mm because is repeated
df<-df[-18,]

# remove the rest of the columns and only leave the text
df<-df[4]


# 3. Separate text into columns ##################################################  

# add the format column
df$Format<-ifelse(as.numeric(row.names(df))<17,"16mm film","35mm film")

# Separate text into elements (using ":" as split) 
df$text1<-as.character(map(strsplit(df$text, split = ":"), 1)) # first correspond to "(1)mfn "
df$text2<-as.character(map(strsplit(df$text, split = ":"), 2)) # 2nd correspond to "Titre"


# Add original number
df$'Numéro(s) d’inventaire (original)'<-df$text1
df$'Numéro(s) d’inventaire (original)'<-gsub(df$'Numéro(s) d’inventaire (original)', pattern = "mfn",replacement = "")
df$'Numéro(s) d’inventaire (original)'<-gsub(df$'Numéro(s) d’inventaire (original)', pattern = "\\)",replacement = "")
df$'Numéro(s) d’inventaire (original)'<-gsub(df$'Numéro(s) d’inventaire (original)', pattern = "\\(",replacement = "")

# Add "Nouveau numéro suggéré par WALTER"
df$'Nouveau numéro suggéré par WALTER'<-paste("CNCM-",substr(df$Format,1,2),"-",sprintf("%03d",as.numeric(df$`Numéro(s) d’inventaire (original)`)),sep = "")

# now I know the number of character of text1 and text2, I can remove from text (+3 for 3x ":")
df$text<-substr(df$text, nchar(df$text1)+nchar(df$text2)+3, nchar(df$text))

# Add the title (let's separate the code and the title)
df$Title<-as.character(map(strsplit(df$text, split = "Code"), 1))

# Add the code (splitting by the word "Code"), remove ":" and change NULL by [AUCUN], trim whitespaces
df$Code<-as.character(map(strsplit(df$text, split = "Code"), 2))
df$Code<-gsub(df$Code, pattern = ":", replacement = "")
df$Code<-ifelse(df$Code=="NULL","[AUCUN]",df$Code)
df$Code<-str_trim(df$Code)


# 4. Remove support columns and re-order ###################################

df$text<-NULL
df$text1<-NULL
df$text2<-NULL
df<-df[c(3,2,4,5,1)]

# 5. Save to CSV ###########################################################
# requires to fix fileEncoding (UTF-8 does not work)
write.csv(df, paste(outdir, "ARCHIVES_CNCM.csv", sep = ""), row.names=FALSE)
