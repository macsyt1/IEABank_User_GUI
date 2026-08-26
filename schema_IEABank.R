# ============================================================
# export_schema.R — dump the public schema to Excel
# One sheet per table, plus a _schema sheet listing every column.
# ============================================================

library(DBI)
library(RPostgres)
library(openxlsx)

# TRUE  -> sheets carry the current rows (useful as examples)
# FALSE -> sheets carry headers only (a blank template to fill in)
INCLUDE_DATA <- TRUE

con <- dbConnect(
  RPostgres::Postgres(),
  host     = Sys.getenv("SUPABASE_HOST"),
  port     = as.integer(Sys.getenv("SUPABASE_PORT")),
  dbname   = Sys.getenv("SUPABASE_DB"),
  user     = Sys.getenv("SUPABASE_RO_USER"),
  password = Sys.getenv("SUPABASE_RO_PWD"),
  sslmode  = "require"
)

on.exit(dbDisconnect(con))

tables <- dbGetQuery(
  con,
  "select table_name
     from information_schema.tables
    where table_schema = 'public'
      and table_type = 'BASE TABLE'
    order by table_name"
)$table_name

# Tables nobody filling in data should see
tables <- setdiff(tables, c("user_profiles"))

schema <- dbGetQuery(
  con,
  "select table_name,
          ordinal_position,
          column_name,
          data_type,
          is_nullable,
          column_default
     from information_schema.columns
    where table_schema = 'public'
    order by table_name, ordinal_position"
)

schema <- schema[schema$table_name %in% tables, ]

wb <- createWorkbook()

addWorksheet(wb, "_schema")
writeData(wb, "_schema", schema)
freezePane(wb, "_schema", firstRow = TRUE)
setColWidths(wb, "_schema", cols = 1:6, widths = "auto")

for (tbl_name in tables) {
  
  df <- dbGetQuery(con, sprintf('select * from public."%s"', tbl_name))
  
  if (!INCLUDE_DATA) {
    df <- df[0, , drop = FALSE]
  }
  
  sheet <- substr(tbl_name, 1, 31)
  
  addWorksheet(wb, sheet)
  writeData(wb, sheet, df)
  freezePane(wb, sheet, firstRow = TRUE)
  setColWidths(wb, sheet, cols = seq_along(df), widths = "auto")
}

saveWorkbook(
  wb,
  paste0("iea_item_bank_schema_", format(Sys.Date(), "%Y%m%d"), ".xlsx"),
  overwrite = TRUE
)