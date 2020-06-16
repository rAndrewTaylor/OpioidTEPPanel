require(tidyverse)
require(irr)

df <- read_csv("data/grouped_statments.csv")

df <- df %>%
  select(-X1) %>%
  gather(key = 'col_num', value='value', -q_group, -SorterId) %>%
  select(SorterId, q_group, value) %>%
  drop_na()

df <- df %>%
  left_join(df, by=c('SorterId', 'q_group')) %>%
  select(-q_group, -SorterId) %>%
  rename(source=value.x, target=value.y) %>%
  write_csv("data/paired_statements.csv")

df_ratings_feas <- ratings.dat %>%
  select(-Importance) %>%
  spread(key = UserID, value = Feasibility) %>%
  select(-StatementID, -`3`) %>%
  drop_na()

icc(ratings = df_ratings_feas, model='twoway')

df_ratings_imp <- ratings.dat %>%
  select(-Feasibility) %>%
  spread(key = UserID, value = Importance) %>%
  select(-StatementID, -`3`) %>%
  drop_na()

icc(ratings = df_ratings_imp, model='twoway')
k_df <- kappam.fleiss(df_ratings_imp, detail = TRUE)
kappam.fleiss(df_ratings_feas, detail = TRUE)
