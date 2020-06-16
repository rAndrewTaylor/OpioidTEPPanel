# Data processing for R-cmap program

require(tidyverse)

# import raw survey data
df_raw <- read_csv("data/survery_map_data.csv")

# import item list
df_item_list <- read_csv("data/statement_items.csv")

df_item_list <- df_item_list %>%
  mutate(Statement = str_replace_all(Statement, c("Opioid use disorder, untreated" = "Opioid use disorder untreated",
                                                  "Opioid use disorder, in treatment" = "Opioid use disorder in treatment",
                                                  "Type of opioid overdose \\(heroin, fentanyl, prescription, methadone\\)" = "Type of opioid overdose (heroin fentanyl prescription methadone)",
                                                  "opioid overdose with co-occurance of other substances \\(benzos, stimulants\\)" = "opioid overdose with co-occurance of other substances (benzos stimulants)",
                                                  "Stable, chronic opioid use" = "Stable chronic opioid use",
                                                  "Presenting for medical or surgical condition not related to opioid misuse \\(asthma exacerbation, acute appendicitis\\)" = "Presenting for medical or surgical condition not related to opioid misuse (asthma exacerbation acute appendicitis)",
                                                  "Presenting for mental health condition \\(depression, suicidal ideation\\)" = "Presenting for mental health condition (depression suicidal ideation)"
  )))
  

# create list/mapping of unique statements with IDs
statement_list <- df_raw %>%
  filter(Status == 'IP Address' & Progress=='100') %>%
  rowid_to_column("SorterId") %>%
  select(SorterId, Q11_0_GROUP:Q11_11_GROUP)

# pivot group data and replace names with statement IDs
statement_df <- statement_list %>%
  gather(key = "q_group", value = "Statement", Q11_0_GROUP:Q11_11_GROUP) %>%
  mutate(Statement = str_replace_all(Statement, c("Opioid use disorder, untreated" = "Opioid use disorder untreated",
                    "Opioid use disorder, in treatment" = "Opioid use disorder in treatment",
                    "Type of opioid overdose \\(heroin, fentanyl, prescription, methadone\\)" = "Type of opioid overdose (heroin fentanyl prescription methadone)",
                    "opioid overdose with co-occurance of other substances \\(benzos, stimulants\\)" = "opioid overdose with co-occurance of other substances (benzos stimulants)",
                    "Stable, chronic opioid use" = "Stable chronic opioid use",
                    "Presenting for medical or surgical condition not related to opioid misuse \\(asthma exacerbation, acute appendicitis\\)" = "Presenting for medical or surgical condition not related to opioid misuse (asthma exacerbation acute appendicitis)",
                    "Presenting for mental health condition \\(depression, suicidal ideation\\)" = "Presenting for mental health condition (depression suicidal ideation)"
  ))) %>%
  mutate(Statement = strsplit(Statement, ",")) %>% 
  unnest(Statement) %>%
  left_join(df_item_list) %>%
  select(-Statement) %>%
  group_by(SorterId, q_group) %>%
  mutate(row_id = row_number()) %>%
  ungroup() %>%
  spread(row_id, StatementID)

write.xlsx2(statement_df, "data/grouped_statments.xls")
  
  # 
  # group_by(ResponseId, q_group) %>%
  # summarise(Statements = list(StatementID)) %>%
  # arrange(ResponseId, q_group) %>%
  # spread(q_group, Statements) %>%
  # gather(key = "q_group", value = "Statements", Q11_0_GROUP:Q11_9_GROUP) %>%
  # arrange(ResponseId, q_group)

# create prioritization/feasibility data

statement_rank <- df_raw %>%
  filter(Status == 'IP Address' & Progress=='100') %>%
  rowid_to_column("SorterId") %>%
  select(SorterId, contains('#')) %>%
  gather(key = "StatementID", value = "rank", `Q5#1_1`:`Q5#2_70`) %>%
  mutate(StatementID = as.integer(substr(StatementID, 6, nchar(StatementID)))) %>%
  arrange(SorterId, StatementID) %>%
  mutate(category = ifelse(grepl("import", rank), "importance", "feasiblity")) %>%
  mutate(rank = as.integer(substr(rank, 0, 1))) %>%
  drop_na() %>%
  spread(category, rank)

write.xlsx2(statement_rank, "data/ranked_statments.xls")
  
# graph analysis
# create_edge_list





  