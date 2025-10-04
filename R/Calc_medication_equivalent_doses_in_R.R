library(openxlsx)
library(tidyverse)

# read and preproc sheets -----
# read input file using file.choose()
input.filename.full <- file.choose()
use <- read.xlsx(input.filename.full, detectDates = TRUE)

# read from the source file
eq <- read.xlsx("../source/251004medication_lists.xlsx",  # please update if needed
                sheet = "eq",
                detectDates = TRUE
)
annot <- read.xlsx("../source/251004medication_lists.xlsx",  # please update if needed
                   sheet = "annot",
                   detectDates = TRUE
)

# preproc eq sheet 
Name_CPpw <- eq %>% select(Name, CPpw)
Name_CPoldpw <- eq %>% select(Name, CPoldpw)
Name_IMPpw <- eq %>% select(Name, IMPpw)
Name_DZPpw <- eq %>% select(Name, DZPpw)
Name_BPDpw <- eq %>% select(Name, BPDpw)
# Name_ABCpw <- eq %>% select(Name, ABCpw)  # If you want to add eq dose, please delete comment-out "#" here.
Name_LDpw <- eq %>% select(Name, LDpw)


### How to update this program ###
# 1. Change annot and eq sheets in the Excel file of the source folder.
# 2. If you want to add drug names, you add columns "pname", "suggest", "aname", and "Name" in anoot sheet.
#    The original file composes of "pname" as trade names, "suggest" as equivalent Japanese Hiragana names, "aname" as often used in clinical settings.
#    Column "Name" lists standard drug name and all names listed in an input Excel file will be coverted to drug names in column "Name" (L74-111).
#    Therefore, value "Name" is mandatory and other three names are optional.
# 3. For eq sheet, you may add and revise "Name" and corresponding the power of drugs per the standard drug (e.g., chlorpromazine).
# 4. If you want to calc ABC equivalent doses, you may add the column before the column "LDpw".
#    In the program, you may delete comment-out "#" in L176-186, L205, L219, and L252 (type "Ctrl + Shirt + C").
### If you have any inquiry, please let the corresponding author Shinsuke Koike (skoike-tky@umin.ac.jp) ###


# preproc annot sheet
cname_Name <- annot %>% select(cname, Name)
pname_Name <- annot %>% select(pname, Name)
suggest_Name <- annot %>% select(suggest3, Name)
abbr_Name <- annot %>% select(aname, Name) %>%
  filter(!is.na(aname))
Name_Name <- annot %>% select(Name.orig = Name, Name = Name) %>%
  filter(!is.na(Name.orig))



# preproc "use" file for calc and output -----
# change vertical type for "use" data
use.gather <- use %>%
  mutate(
    c1 = paste(m1, d1, sep = "ZZZ"),
    c2 = if("m2" %in% colnames(use)) paste(m2, d2, sep = "ZZZ"),
    c3 = if("m3" %in% colnames(use)) paste(m3, d3, sep = "ZZZ"),
    c4 = if("m4" %in% colnames(use)) paste(m4, d4, sep = "ZZZ"),
    c5 = if("m5" %in% colnames(use)) paste(m5, d5, sep = "ZZZ"),
    c6 = if("m6" %in% colnames(use)) paste(m6, d6, sep = "ZZZ"),
    c7 = if("m7" %in% colnames(use)) paste(m7, d7, sep = "ZZZ"),
    c8 = if("m8" %in% colnames(use)) paste(m8, d8, sep = "ZZZ"),
    c9 = if("m9" %in% colnames(use)) paste(m9, d9, sep = "ZZZ"),
  ) %>%
  gather("drugno", "drugname_dose", starts_with("c")) %>%
  separate("drugname_dose", into = c("drugname", "drugdose"), sep = "ZZZ") %>%
  filter(!(drugname == "NA")) %>%
  select(ID, date, drugname, drugdose) %>%
  mutate(drugname = str_remove_all(drugname, "錠|散|顆粒|配合")) %>%  # for Japanese drug names
  fill(ID, date) %>%
  arrange(ID)

# add standard "Name" from use.gather$drugname
n <- vector()
for (i in 1:nrow(use.gather)) {
  for (j in 1:nrow(pname_Name)) {
    if (use.gather$drugname[i] == pname_Name$pname[j]) {
      n[i] <- pname_Name$Name[j]
      break
    }
  }
}

for (i in 1:nrow(use.gather)) {
  for (k in 1:nrow(suggest_Name)) {
    if (use.gather$drugname[i] == suggest_Name$suggest3[k]) {
      n[i] <- suggest_Name$Name[k]
      break
    }
  }
}

for (i in 1:nrow(use.gather)) {
  for (l in 1:nrow(abbr_Name)) {
    if (use.gather$drugname[i] == abbr_Name$aname[l]) {
      n[i] <- abbr_Name$Name[l]
      break
    }
  }
}

for (i in 1:nrow(use.gather)) {
  for (m in 1:nrow(Name_Name)) {
    if (use.gather$drugname[i] == Name_Name$Name.orig[m]) {
      n[i] <- Name_Name$Name[m]
      break
    }
  }
}

use.gather$Name <- n


# Calc power values for drug name "Name"
CP <- vector(length = nrow(use.gather))
for (i in 1:nrow(use.gather)) {
  for (l in 1:nrow(Name_CPpw)) {
    if (is.na(use.gather$Name[i]) || is.na(Name_CPpw$Name[l])) {
      next
    }
    if (use.gather$Name[i] == Name_CPpw$Name[l]) {
      CP[i] <- Name_CPpw$CPpw[l]
    }
  }
}

CPold <- vector(length = nrow(use.gather))
for (i in 1:nrow(use.gather)) {
  for (l in 1:nrow(Name_CPoldpw)) {
    if (is.na(use.gather$Name[i]) || is.na(Name_CPoldpw$Name[l])) {
      next
    }
    if (use.gather$Name[i] == Name_CPoldpw$Name[l]) {
      CPold[i] <- Name_CPoldpw$CPoldpw[l]
    }
  }
}

IMP <- vector(length = nrow(use.gather))
for (i in 1:nrow(use.gather)) {
  for (l in 1:nrow(Name_IMPpw)) {
    if (is.na(use.gather$Name[i]) || is.na(Name_IMPpw$Name[l])) {
      next
    }
    if (use.gather$Name[i] == Name_IMPpw$Name[l]) {
      IMP[i] <- Name_IMPpw$IMPpw[l]
    }
  }
}

DZP <- vector(length = nrow(use.gather))
for (i in 1:nrow(use.gather)) {
  for (l in 1:nrow(Name_DZPpw)) {
    if (is.na(use.gather$Name[i]) || is.na(Name_DZPpw$Name[l])) {
      next
    }
    if (use.gather$Name[i] == Name_DZPpw$Name[l]) {
      DZP[i] <- Name_DZPpw$DZPpw[l]
    }
  }
}

BPD <- vector(length = nrow(use.gather))
for (i in 1:nrow(use.gather)) {
  for (l in 1:nrow(Name_BPDpw)) {
    if (is.na(use.gather$Name[i]) || is.na(Name_BPDpw$Name[l])) {
      next
    }
    if (use.gather$Name[i] == Name_BPDpw$Name[l]) {
      BPD[i] <- Name_BPDpw$BPDpw[l]
    }
  }
}

# ABC <- vector(length = nrow(use.gather))
# for (i in 1:nrow(use.gather)) {
#   for (l in 1:nrow(Name_BPDpw)) {
#     if (is.na(use.gather$Name[i]) || is.na(Name_ABCpw$Name[l])) {
#       next
#     }
#     if (use.gather$Name[i] == Name_ABCpw$Name[l]) {
#       BPD[i] <- Name_ABCpw$ABCpw[l]
#     }
#   }
# }

LD <- vector(length = nrow(use.gather))
for (i in 1:nrow(use.gather)) {
  for (l in 1:nrow(Name_LDpw)) {
    if (is.na(use.gather$Name[i]) || is.na(Name_LDpw$Name[l])) {
      next
    }
    if (use.gather$Name[i] == Name_LDpw$Name[l]) {
      LD[i] <- Name_LDpw$LDpw[l]
    }
  }
}

use.gather$CPpw <- CP
use.gather$CPoldpw <- CPold
use.gather$IMPpw <- IMP
use.gather$DZPpw <- DZP
use.gather$BPDpw <- BPD
# use.gather$ABCpw <- ABC
use.gather$LDpw <- LD

# calc equivalent doses per drug and drugdose
use.eq <- use.gather %>%
  mutate(
    drugdose = as.numeric(drugdose), 
    CPeq = drugdose * 100 / CPpw,
    CPoldeq = drugdose * 100 / CPoldpw) %>%
  mutate(
    CPneweq = CPeq - CPoldeq,
    IMPeq = drugdose * 150 / IMPpw, 
    DZPeq = drugdose * 5 / DZPpw,
    BPDeq = drugdose * 2 / BPDpw, 
    # ABCeq = drugdose * 50 / ABCpw, 
    LDeq = drugdose * LDpw) %>%
  left_join(eq %>% select(Name, Notes) %>% filter(!is.na(Notes))) %>%
  select(-c(CPpw:LDpw))
use.eq[use.eq == Inf] <- NA

# add NoCalc = 1 if no eq values per drug
use.fixed <- use.eq %>%
  mutate_at(
    vars(CPeq:LDeq),
    ~ ifelse(is.na(.), 0.0, .)
  )

complete <- use.fixed %>%
  mutate(NoCalc = if_else(apply(use.fixed %>% select(CPeq:LDeq), 1, sum) == 0, 1, 0)) %>%
  mutate(NoList = if_else(is.na(Name), 1, 0)) %>%
  select(1:5, NoCalc, NoList, everything())



# write.xlsx -----
write.xlsx(complete, paste0("../output/", 
                            gsub("\\.xlsx$", "", basename(input.filename.full)),
                            "_bydrug.xlsx"))

grouped_complete <- complete %>% 
  group_by(ID, date) %>% 
  summarize(NoCalc = sum(NoCalc),
            CPeq = sum(CPeq),
            CPoldeq = sum(CPoldeq),
            CPneweq = sum(CPneweq),
            IMPeq = sum(IMPeq),
            DZPeq = sum(DZPeq),
            BPDeq = sum(BPDeq),
            # ABCeq = sum(ABCeq),
            LDeq = sum(LDeq)
            )

write.xlsx(grouped_complete, paste0("../output/", 
                                    gsub("\\.xlsx$", "", basename(input.filename.full)),
                                    "_byIDdate.xlsx"))
