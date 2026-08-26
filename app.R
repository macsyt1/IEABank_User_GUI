# ============================================================
# app.R — IEA Item Bank Explorer
# ------------------------------------------------------------
# Public read-only visualizer for questionnaire items, scales,
# questionnaires, item selection cart, longitudinal detail modals,
# and Excel export.
# ============================================================

# ------------------------------------------------------------
# Libraries
# ------------------------------------------------------------
library(shiny)
library(bslib)
library(DBI)
library(RPostgres)
library(pool)
library(dplyr)
library(dbplyr)
library(tibble)
library(stringr)
library(DT)
library(openxlsx)
library(httr2)

# ------------------------------------------------------------
# App version
# ------------------------------------------------------------
APP_VERSION <- "0.1.0"

# ------------------------------------------------------------
# Institutional colors
# ------------------------------------------------------------
iea_blue <- "#0070b8"
iea_red  <- "#e2211c"

# ------------------------------------------------------------
# Conexión
# ------------------------------------------------------------
pool <- dbPool(
  RPostgres::Postgres(),
  host     = Sys.getenv("SUPABASE_HOST"),
  port     = as.integer(Sys.getenv("SUPABASE_PORT")),
  dbname   = Sys.getenv("SUPABASE_DB"),
  user     = Sys.getenv("SUPABASE_RO_USER"),
  password = Sys.getenv("SUPABASE_RO_PWD"),
  sslmode  = "require"
)

onStop(function() poolClose(pool))

# ------------------------------------------------------------
# Authentication
# ------------------------------------------------------------
# The public deploy runs as public_app and needs no login; anything
# else (internal_app, or a role we do not recognise) does. Deriving
# it from the credential means a missing variable fails closed.

SUPABASE_URL      <- Sys.getenv("SUPABASE_URL")
SUPABASE_ANON_KEY <- Sys.getenv("SUPABASE_ANON_KEY")

REQUIRE_LOGIN <- !grepl("^public_app", Sys.getenv("SUPABASE_RO_USER"))

supabase_sign_in <- function(email, password) {
  
  req <- request(paste0(SUPABASE_URL, "/auth/v1/token?grant_type=password")) |>
    req_method("POST") |>
    req_headers(
      apikey = SUPABASE_ANON_KEY,
      Authorization = paste("Bearer", SUPABASE_ANON_KEY)
    ) |>
    req_body_json(list(email = email, password = password))
  
  resp <- tryCatch(req_perform(req), error = function(e) NULL)
  
  if (is.null(resp) || resp_status(resp) >= 400) {
    return(NULL)
  }
  
  resp_body_json(resp)
}

get_user_profile <- function(user_id) {
  dbGetQuery(
    pool,
    "select user_id, first_name, last_name, role, active
       from public.user_profiles
      where user_id = $1",
    params = list(user_id)
  )
}

supabase_send_password_reset <- function(email) {
  
  reset_url <- "https://macsyt1.github.io/IEABank_Admin_GUI/reset-password.html"
  
  req <- request(paste0(SUPABASE_URL, "/auth/v1/recover")) |>
    req_url_query(redirect_to = reset_url) |>
    req_method("POST") |>
    req_headers(
      apikey = SUPABASE_ANON_KEY,
      Authorization = paste("Bearer", SUPABASE_ANON_KEY)
    ) |>
    req_body_json(list(email = email))
  
  resp <- tryCatch(req_perform(req), error = function(e) NULL)
  
  !is.null(resp) && resp_status(resp) < 400
}

login_ui <- function() {
  div(
    style = "
      max-width: 420px;
      margin: 90px auto;
      padding: 32px;
      border: 1px solid #ddd;
      border-radius: 12px;
      background: #fff;
      box-shadow: 0 2px 10px rgba(0,0,0,.08);
    ",
    
    div(
      style = "text-align:center; margin-bottom:24px;",
      img(src = "iea_logo.png", style = "height:52px; margin-bottom:14px;"),
      div(
        style = "font-size:1.4rem; font-weight:700; color:#54565A;",
        "IEABank Viewer"
      )
    ),
    
    textInput(
      "login_email",
      "Email",
      placeholder = "name@example.org"
    ),
    
    passwordInput(
      "login_password",
      "Password"
    ),
    
    actionButton(
      "login_btn",
      "Sign in",
      class = "btn-primary w-100"
    ),
    
    div(
      style = "text-align:center; margin-top:14px;",
      actionLink(
        "forgot_password_btn",
        "Forgot password?"
      )
    ),
    
    tags$script(HTML("
      $(document).on('keypress', '#login_email, #login_password', function(e) {
        if (e.which === 13) {
          $('#login_btn').click();
        }
      });
    ")),
    
    uiOutput("login_message")
  )
}

# ------------------------------------------------------------
# Helpers generales
# ------------------------------------------------------------

`%||%` <- function(a, b) if (is.null(a) || length(a) == 0) b else a

safe_chr <- function(x) {
  ifelse(is.na(x), "", as.character(x))
}

add_missing_columns <- function(df, cols) {
  for (col in cols) {
    if (!col %in% names(df)) {
      df[[col]] <- NA
    }
  }
  
  df
}

read_table_safe <- function(pool, table_name) {
  tryCatch(
    {
      tbl(pool, table_name) %>%
        collect()
    },
    error = function(e) {
      tibble()
    }
  )
}

filter_by_selected_items <- function(df, selected_item_uids) {
  if (nrow(df) == 0) {
    return(df)
  }
  
  if (!"item_uid" %in% names(df)) {
    return(df)
  }
  
  df %>%
    filter(item_uid %in% selected_item_uids)
}

show_added_notification <- function(n_items) {
  showNotification(
    paste(n_items, "item(s) added to the cart."),
    type = "message",
    duration = 2
  )
}

sort_years_desc <- function(x) {
  years <- unique(safe_chr(x))
  years <- years[years != ""]
  
  if (length(years) == 0) {
    return(character())
  }
  
  years_num <- suppressWarnings(as.numeric(years))
  
  if (all(!is.na(years_num))) {
    years[order(years_num, decreasing = TRUE)]
  } else {
    sort(years, decreasing = TRUE)
  }
}

sort_years_asc <- function(x) {
  rev(sort_years_desc(x))
}

datatable_all_items <- function(df, selection = "multiple") {
  datatable(
    df,
    rownames = FALSE,
    selection = selection,
    options = list(
      paging = FALSE,
      dom = "t",
      scrollX = TRUE,
      scrollY = "420px",
      ordering = TRUE
    )
  )
}

datatable_simple <- function(df, selection = "none") {
  datatable(
    df,
    rownames = FALSE,
    selection = selection,
    options = list(
      paging = FALSE,
      dom = "t",
      scrollX = TRUE,
      ordering = FALSE
    )
  )
}

storage_url <- function(path) {
  paste0(
    Sys.getenv("SUPABASE_URL"),
    "/storage/v1/object/public/item-examples/",
    path
  )
}

example_link <- function(path, label) {
  text <- ifelse(is.na(label) | label == "", basename(path), label)
  
  sprintf(
    '<a href="%s" target="_blank" rel="noopener">%s</a>',
    storage_url(path),
    text
  )
}

datatable_variants <- function(df) {
  datatable(
    df,
    rownames = FALSE,
    escape = FALSE,
    selection = "none",
    colnames = c(
      "Version",
      "Study",
      "Year",
      "Population",
      "Variable",
      "Scale(s)",
      "Example"
    ),
    options = list(
      paging = FALSE,
      dom = "t",
      scrollX = TRUE,
      ordering = FALSE
    )
  )
}

value_table_ui <- function(df) {
  if (nrow(df) == 0) {
    return(p(class = "text-muted", "Not available."))
  }
  
  tags$table(
    class = "table table-sm",
    tags$tbody(
      lapply(seq_len(nrow(df)), function(i) {
        tags$tr(
          tags$td(as.character(df[[1]][i])),
          tags$td(as.character(df[[2]][i]))
        )
      })
    )
  )
}

version_matrix_ui <- function(matrix_df, row_label = "Study") {
  if (nrow(matrix_df) == 0) {
    return(
      div(
        class = "empty-state",
        p(class = "text-muted", "No longitudinal information available.")
      )
    )
  }
  
  row_col <- names(matrix_df)[1]
  year_cols <- setdiff(names(matrix_df), row_col)
  
  div(
    class = "presence-table-wrapper",
    tags$table(
      class = "presence-table",
      tags$thead(
        tags$tr(
          tags$th(class = "presence-row-label", row_label),
          lapply(year_cols, function(y) {
            tags$th(class = "presence-year", y)
          })
        )
      ),
      tags$tbody(
        lapply(seq_len(nrow(matrix_df)), function(i) {
          tags$tr(
            tags$td(
              class = "presence-row-label",
              matrix_df[[row_col]][i]
            ),
            lapply(year_cols, function(y) {
              tags$td(
                class = "presence-cell",
                matrix_df[[y]][i]
              )
            })
          )
        })
      )
    )
  )
}

grid_details_button <- function(item_uid) {
  sprintf(
    paste0(
      '<button class="btn btn-sm btn-outline-primary" ',
      'onclick="event.stopPropagation(); ',
      "Shiny.setInputValue('grid_details_uid', '%s', {priority: 'event'})",
      '">Details &rarr;</button>'
    ),
    item_uid
  )
}

datatable_items_grid <- function(df) {
  datatable(
    df,
    rownames = FALSE,
    escape = FALSE,
    selection = "multiple",
    colnames = c(
      "Item UID",
      "Item",
      "Category",
      "Studies",
      "Years",
      "Versions",
      ""
    ),
    options = list(
      dom = "lrtip",
      pageLength = 25,
      lengthMenu = c(10, 25, 50, 100),
      scrollX = TRUE,
      autoWidth = FALSE,
      columnDefs = list(
        list(targets = 1, width = "40%"),
        list(targets = 5, className = "dt-center"),
        list(targets = 6, orderable = FALSE, searchable = FALSE, className = "dt-center")
      )
    )
  )
}

datatable_cart <- function(df) {
  datatable(
    df,
    rownames = FALSE,
    options = list(
      pageLength = 10,
      scrollX = TRUE
    )
  )
}

# ------------------------------------------------------------
# Standardization layer
# ------------------------------------------------------------

standardize_items <- function(df) {
  
  if (nrow(df) == 0) {
    return(df)
  }
  
  if ("wording_item" %in% names(df) && !"wording" %in% names(df)) {
    df <- df %>% rename(wording = wording_item)
  }
  
  if ("item_var" %in% names(df) && !"item_code" %in% names(df)) {
    df <- df %>% rename(item_code = item_var)
  }
  
  if ("varname" %in% names(df) && !"source_variable" %in% names(df)) {
    df <- df %>% rename(source_variable = varname)
  }
  
  if ("scale_description" %in% names(df) && !"scale" %in% names(df)) {
    df <- df %>% rename(scale = scale_description)
  }
  
  if ("type" %in% names(df) && !"item_type" %in% names(df)) {
    df <- df %>% rename(item_type = type)
  }
  
  if ("target" %in% names(df) && !"population" %in% names(df)) {
    df <- df %>% rename(population = target)
  }
  
  required_cols <- c(
    "item_admin_pk",
    "item_admin_id",
    "admin_id",
    "item_uid",
    "item_name",
    "category_id",
    "category_name",
    "item_code",
    "source_variable",
    "dataset_label",
    "wording",
    "wording_question",
    "wording_instruction",
    "wording_context",
    "wording_heading",
    "study",
    "phase",
    "year",
    "instrument",
    "population",
    "cycle",
    "section",
    "scale_id",
    "scale",
    "scale_uid",
    "scale_var",
    "scale_name",
    "scale_varname",
    "item_type",
    "response_format",
    "trend_status",
    "puf",
    "miss_id",
    "response_id"
  )
  
  df <- add_missing_columns(df, required_cols)
  
  df <- df %>%
    mutate(
      item_uid = if_else(
        is.na(item_uid) | item_uid == "",
        paste(
          safe_chr(admin_id),
          safe_chr(item_admin_id),
          safe_chr(item_code),
          sep = "_"
        ),
        safe_chr(item_uid)
      ),
      
      item_code = if_else(
        is.na(item_code) | item_code == "",
        safe_chr(source_variable),
        safe_chr(item_code)
      ),
      
      wording = if_else(
        is.na(wording) | wording == "",
        safe_chr(wording_question),
        safe_chr(wording)
      ),
      
      item_name = if_else(
        is.na(item_name) | item_name == "",
        safe_chr(wording),
        safe_chr(item_name)
      ),
      
      category_id = safe_chr(category_id),
      category_name = safe_chr(category_name),
      scale_uid = safe_chr(scale_uid),
      scale_var = safe_chr(scale_var),
      scale_name = safe_chr(scale_name),
      
      study = safe_chr(study),
      phase = safe_chr(phase),
      year = safe_chr(year),
      instrument = safe_chr(instrument),
      population = safe_chr(population),
      cycle = safe_chr(cycle),
      section = safe_chr(section),
      
      scale_id = safe_chr(scale_id),
      scale = safe_chr(scale),
      scale_varname = safe_chr(scale_varname),
      
      item_type = safe_chr(item_type),
      source_variable = safe_chr(source_variable),
      dataset_label = safe_chr(dataset_label),
      
      response_format = if_else(
        is.na(response_format) | response_format == "",
        safe_chr(response_id),
        safe_chr(response_format)
      ),
      
      trend_status = if_else(
        is.na(trend_status) | trend_status == "",
        "Not classified",
        safe_chr(trend_status)
      ),
      
      puf = if_else(is.na(puf), FALSE, as.logical(puf))
    )
  
  df
}

# ------------------------------------------------------------
# Data access layer
# ------------------------------------------------------------

get_items <- function(pool) {
  
  item_admin <- tbl(pool, "item_admin")
  
  admin_tbl <- tbl(pool, "admin") %>%
    select(
      admin_id,
      study,
      phase,
      year,
      instrument,
      target,
      cycle
    )
  
  scale_items <- tbl(pool, "scale_items") %>%
    left_join(
      tbl(pool, "scale_version"),
      by = "scale_id"
    ) %>%
    left_join(
      tbl(pool, "scale"),
      by = "scale_uid"
    ) %>%
    select(
      item_admin_pk,
      scale_uid,
      scale_var,
      scale_name,
      scale_id,
      scale_description,
      scale_varname
    )
  
  item_tbl <- tbl(pool, "item_id") %>%
    select(
      item_uid,
      item_name,
      category_id
    )
  
  category_tbl <- tbl(pool, "category") %>%
    select(
      category_id,
      category_name
    )
  
  df <- item_admin %>%
    left_join(
      admin_tbl,
      by = "admin_id"
    ) %>%
    left_join(
      scale_items,
      by = "item_admin_pk"
    ) %>%
    left_join(
      item_tbl,
      by = "item_uid"
    ) %>%
    left_join(
      category_tbl,
      by = "category_id"
    ) %>%
    collect()
  
  standardize_items(df)
}

get_item_history <- function(pool, selected_item_uids) {
  tibble()
}

get_item_variants <- function(pool, selected_item_uids) {
  tibble()
}

get_item_examples <- function(pool) {
  read_table_safe(pool, "item_example")
}

get_scale_membership <- function(pool, selected_items) {
  df <- read_table_safe(pool, "scale_items")
  
  if (nrow(df) == 0) {
    return(tibble())
  }
  
  df %>%
    left_join(
      read_table_safe(pool, "scale_version"),
      by = "scale_id"
    ) %>%
    filter(item_admin_pk %in% selected_items$item_admin_pk)
}

get_response_options <- function(pool, selected_items) {
  
  selected_response_ids <- selected_items %>%
    filter(!is.na(response_id), response_id != "") %>%
    distinct(response_id) %>%
    pull(response_id)
  
  if (length(selected_response_ids) == 0) {
    return(tibble())
  }
  
  response_values <- read_table_safe(pool, "value_scheme_value")
  
  if (nrow(response_values) == 0) {
    return(tibble())
  }
  
  if (!"response_id" %in% names(response_values)) {
    return(response_values)
  }
  
  response_values %>%
    filter(response_id %in% selected_response_ids)
}

get_missing_values <- function(pool, selected_items) {
  
  selected_miss_ids <- selected_items %>%
    filter(!is.na(miss_id), miss_id != "") %>%
    distinct(miss_id) %>%
    pull(miss_id)
  
  if (length(selected_miss_ids) == 0) {
    return(tibble())
  }
  
  miss_values <- read_table_safe(pool, "miss_scheme_value")
  
  if (nrow(miss_values) == 0) {
    return(tibble())
  }
  
  if (!"miss_id" %in% names(miss_values)) {
    return(miss_values)
  }
  
  miss_values %>%
    filter(miss_id %in% selected_miss_ids)
}

get_variables_and_data <- function(pool, selected_items) {
  selected_items %>%
    select(
      item_uid,
      item_name,
      category_id,
      category_name,
      item_admin_id,
      admin_id,
      study,
      phase,
      year,
      cycle,
      population,
      instrument,
      item_code,
      source_variable,
      dataset_label,
      puf
    ) %>%
    distinct()
}

derive_scales <- function(items) {
  
  if (nrow(items) == 0) {
    return(tibble())
  }
  
  items %>%
    filter(scale_uid != "") %>%
    group_by(scale_uid, scale_name) %>%
    summarise(
      varnames = paste(
        sort(unique(scale_varname[scale_varname != ""])),
        collapse = ", "
      ),
      studies = paste(sort(unique(study[study != ""])), collapse = ", "),
      years = paste(sort(unique(year[year != ""])), collapse = ", "),
      n_versions = n_distinct(scale_id),
      n_items = n_distinct(item_uid),
      .groups = "drop"
    ) %>%
    arrange(scale_uid)
}

derive_scale_admins <- function(items) {
  
  if (nrow(items) == 0) {
    return(tibble())
  }
  
  items %>%
    filter(scale_uid != "") %>%
    distinct(
      scale_uid,
      scale_id,
      scale_var,
      scale_varname,
      scale,
      admin_id,
      study,
      phase,
      year,
      cycle,
      population,
      instrument
    ) %>%
    arrange(scale_uid, year, scale_var)
}

derive_composition <- function(group_col, group_val, items, study_sel, population_sel, instrument_sel) {
  
  scope <- items %>%
    filter(
      study == study_sel,
      population == population_sel,
      instrument == instrument_sel
    )
  
  in_scale <- scope %>%
    filter(.data[[group_col]] == group_val)
  
  item_uids <- sort(unique(in_scale$item_uid))
  years <- sort_years_asc(scope$year)
  
  if (length(item_uids) == 0 || length(years) == 0) {
    return(tibble())
  }
  
  out <- data.frame(
    item_uid = item_uids,
    check.names = FALSE,
    stringsAsFactors = FALSE
  )
  
  for (yr in years) {
    out[[yr]] <- vapply(
      item_uids,
      function(iu) {
        cell <- in_scale %>%
          filter(item_uid == iu, year == yr)
        
        if (nrow(cell) > 0) {
          return(paste(sort(unique(cell$item_code)), collapse = ", "))
        }
        
        administered <- scope %>%
          filter(item_uid == iu, year == yr)
        
        if (nrow(administered) > 0) "\u2013" else ""
      },
      character(1)
    )
  }
  
  as_tibble(out)
}

derive_categories <- function(items) {
  
  if (nrow(items) == 0) {
    return(tibble())
  }
  
  items %>%
    filter(category_id != "") %>%
    group_by(category_id, category_name) %>%
    summarise(
      studies = paste(sort(unique(study[study != ""])), collapse = ", "),
      years = paste(sort(unique(year[year != ""])), collapse = ", "),
      n_items = n_distinct(item_uid),
      .groups = "drop"
    ) %>%
    arrange(category_id)
}

derive_category_admins <- function(items) {
  
  if (nrow(items) == 0) {
    return(tibble())
  }
  
  items %>%
    filter(category_id != "") %>%
    distinct(
      category_id,
      admin_id,
      study,
      phase,
      year,
      cycle,
      population,
      instrument
    )
}

derive_category_items <- function(cat_id, items) {
  
  items %>%
    filter(category_id == cat_id) %>%
    group_by(item_uid) %>%
    summarise(
      item_name = first(item_name),
      studies = paste(sort(unique(study[study != ""])), collapse = ", "),
      years = paste(sort(unique(year[year != ""])), collapse = ", "),
      n_versions = n_distinct(item_code),
      .groups = "drop"
    ) %>%
    arrange(item_uid) %>%
    mutate(details = grid_details_button(item_uid))
}

derive_questionnaires <- function(items) {
  
  if (nrow(items) == 0) {
    return(tibble())
  }
  
  items %>%
    distinct(
      study,
      phase,
      year,
      cycle,
      population,
      instrument,
      section,
      scale,
      item_uid
    ) %>%
    group_by(
      study,
      phase,
      year,
      cycle,
      population,
      instrument
    ) %>%
    summarise(
      n_items = n_distinct(item_uid),
      n_sections = n_distinct(section[section != ""], na.rm = TRUE),
      n_scales = n_distinct(scale[scale != ""], na.rm = TRUE),
      .groups = "drop"
    ) %>%
    mutate(
      questionnaire_uid = paste(
        safe_chr(study),
        safe_chr(cycle),
        safe_chr(population),
        safe_chr(instrument),
        sep = "_"
      ),
      questionnaire_uid = str_replace_all(questionnaire_uid, "[^A-Za-z0-9_]", "_")
    )
}

derive_item_versions <- function(uid, items, examples) {
  
  versions <- items %>%
    filter(item_uid == uid) %>%
    group_by(item_admin_id, admin_id) %>%
    summarise(
      across(
        c(
          item_code,
          study,
          phase,
          year,
          cycle,
          population,
          instrument,
          source_variable,
          dataset_label,
          wording,
          wording_question,
          wording_heading,
          wording_instruction,
          item_type,
          puf,
          response_id,
          miss_id
        ),
        first
      ),
      scale_varnames = paste(
        sort(unique(scale_varname[scale_varname != ""])),
        collapse = ", "
      ),
      scale_names = paste(
        sort(unique(scale[scale != ""])),
        collapse = "; "
      ),
      .groups = "drop"
    ) %>%
    arrange(year, item_code)
  
  if (nrow(examples) == 0) {
    versions$examples <- ""
    return(versions)
  }
  
  example_links <- examples %>%
    group_by(item_admin_id, admin_id) %>%
    summarise(
      examples = paste(example_link(path, label), collapse = " "),
      .groups = "drop"
    )
  
  versions %>%
    left_join(example_links, by = c("item_admin_id", "admin_id")) %>%
    mutate(examples = safe_chr(examples))
}

derive_item_version_matrix <- function(versions) {
  
  studies <- sort(unique(versions$study))
  studies <- studies[studies != ""]
  
  years <- sort_years_asc(versions$year)
  
  if (length(studies) == 0 || length(years) == 0) {
    return(tibble())
  }
  
  out <- data.frame(
    study = studies,
    check.names = FALSE,
    stringsAsFactors = FALSE
  )
  
  for (yr in years) {
    out[[yr]] <- vapply(
      studies,
      function(st) {
        cell <- versions %>%
          filter(study == st, year == yr)
        
        paste(sort(unique(cell$item_code)), collapse = ", ")
      },
      character(1)
    )
  }
  
  as_tibble(out)
}

# ------------------------------------------------------------
# Theme
# ------------------------------------------------------------

app_theme <- bs_theme(
  version = 5,
  primary = iea_blue,
  secondary = "#6C757D",
  base_font = font_google("Source Sans 3"),
  heading_font = font_google("Source Sans 3")
)

# ------------------------------------------------------------
# Branding
# ------------------------------------------------------------

app_brand <- div(
  class = "app-brand",
  tags$img(
    src = "iea_logo.png",
    class = "app-logo",
    alt = "IEA logo"
  ),
  div(
    class = "app-title",
    div(class = "app-title-main", "Item Bank Explorer")
  )
)

app_footer <- div(
  class = "app-footer",
  div(
    class = "app-footer-inner",
    div(
      class = "footer-line footer-version",
      tags$span(
        HTML("<strong>IEA Item Bank Explorer</strong> · Version ")
      ),
      actionLink(
        inputId = "version_click",
        label = APP_VERSION,
        class = "version-easter-egg"
      )
    ),
    div(
      class = "footer-line footer-copyright",
      HTML(
        "&copy; IEA, 2026 · International Association for the Evaluation of Educational Achievement"
      )
    )
  )
)

# ------------------------------------------------------------
# UI helpers
# ------------------------------------------------------------

page_header <- function(title, subtitle = NULL) {
  div(
    class = "page-header",
    h2(title),
    if (!is.null(subtitle)) {
      p(class = "text-muted", subtitle)
    }
  )
}

empty_state <- function(message) {
  div(
    class = "empty-state",
    p(class = "text-muted", message)
  )
}

scale_card <- function(scale_row) {
  
  scale_uid_safe <- str_replace_all(scale_row$scale_uid, "[^A-Za-z0-9_]", "_")
  
  card(
    class = "scale-card",
    card_body(
      div(
        class = "d-flex justify-content-between align-items-start gap-3",
        div(
          h5(scale_row$scale_name),
          p(
            class = "text-muted",
            paste(scale_row$scale_uid, "\u00b7", scale_row$varnames)
          )
        ),
        actionButton(
          inputId = paste0("add_scale_", scale_uid_safe),
          label = "+ Add",
          class = "btn btn-sm btn-primary"
        )
      ),
      tags$hr(),
      div(paste(scale_row$n_items, "items")),
      div(paste(scale_row$n_versions, "versions")),
      div(
        class = "small",
        paste(scale_row$studies, "\u00b7", scale_row$years)
      ),
      div(
        class = "mt-2",
        actionButton(
          inputId = paste0("details_scale_", scale_uid_safe),
          label = "View scale \u2192",
          class = "btn btn-sm btn-outline-primary"
        )
      )
    )
  )
}

category_card <- function(category_row) {
  
  category_id_safe <- str_replace_all(category_row$category_id, "[^A-Za-z0-9_]", "_")
  
  card(
    class = "scale-card",
    card_body(
      div(
        class = "d-flex justify-content-between align-items-start gap-3",
        div(
          h5(category_row$category_name),
          p(class = "text-muted", category_row$category_id)
        ),
        actionButton(
          inputId = paste0("add_category_", category_id_safe),
          label = "+ Add",
          class = "btn btn-sm btn-primary"
        )
      ),
      tags$hr(),
      div(paste(category_row$n_items, "items")),
      div(
        class = "small",
        paste(category_row$studies, "\u00b7", category_row$years)
      ),
      div(
        class = "mt-2",
        actionButton(
          inputId = paste0("details_category_", category_id_safe),
          label = "View category \u2192",
          class = "btn btn-sm btn-outline-primary"
        )
      )
    )
  )
}

questionnaire_card <- function(questionnaire_row) {
  
  questionnaire_uid_safe <- str_replace_all(
    questionnaire_row$questionnaire_uid,
    "[^A-Za-z0-9_]",
    "_"
  )
  
  card(
    class = "questionnaire-card",
    card_body(
      div(
        class = "d-flex justify-content-between align-items-start gap-3",
        div(
          h5(
            paste(
              questionnaire_row$study,
              questionnaire_row$year,
              questionnaire_row$population,
              questionnaire_row$instrument
            )
          ),
          p(class = "text-muted", "Complete questionnaire")
        ),
        actionButton(
          inputId = paste0("add_questionnaire_", questionnaire_uid_safe),
          label = "+ Add",
          class = "btn btn-sm btn-primary"
        )
      ),
      tags$hr(),
      div(paste(questionnaire_row$n_items, "items")),
      div(paste(questionnaire_row$n_sections, "sections")),
      div(paste(questionnaire_row$n_scales, "scales")),
      div(
        class = "mt-2",
        actionButton(
          inputId = paste0("details_questionnaire_", questionnaire_uid_safe),
          label = "View questionnaire →",
          class = "btn btn-sm btn-outline-primary"
        )
      )
    )
  )
}

# ------------------------------------------------------------
# UI
# ------------------------------------------------------------

main_ui <- function() page_navbar(
  id = "main_nav",
  title = app_brand,
  theme = app_theme,
  fillable = FALSE,
  footer = app_footer,
  
  header = tags$head(
    tags$style(HTML("
      :root {
        --iea-blue: #0070b8;
        --iea-red: #e2211c;
        --iea-bg: #F7F9FC;
        --iea-border: #E5E7EB;
        --iea-text: #1F2937;
        --iea-muted: #6B7280;
      }

      html,
      body {
        min-height: 100%;
      }

      body {
        background-color: var(--iea-bg);
        color: var(--iea-text);
        overflow-x: hidden;
      }

      .bslib-page-navbar {
        min-height: 100vh;
      }

      .bslib-page-navbar > .tab-content {
        padding-bottom: 2rem;
      }

      .tab-content {
        overflow: visible;
      }

      .navbar {
        border-bottom: 1px solid #D9DEE5;
        background-color: #FFFFFF !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03);
        padding-top: 0.6rem;
        padding-bottom: 0.6rem;
        min-height: 76px;
      }

      .navbar > .container-fluid {
        display: grid;
        grid-template-columns: auto minmax(0, 1fr) auto;
        align-items: center;
        column-gap: 3.5rem;
      }

      .navbar-brand {
        padding-top: 0;
        padding-bottom: 0;
        margin-right: 0;
        display: flex;
        align-items: center;
        min-height: 58px;
        grid-column: 1;
      }

      .app-brand {
        display: flex;
        align-items: center;
        gap: 1rem;
        min-height: 58px;
      }

      .app-logo {
        height: 48px;
        width: auto;
        display: block;
        flex: 0 0 auto;
      }

      .app-title {
        display: flex;
        align-items: center;
        justify-content: center;
        min-height: 48px;
        line-height: 1;
        margin: 0;
        padding: 0;
      }

      .app-title-main {
        font-weight: 700;
        font-size: 1.08rem;
        color: #222222;
        letter-spacing: 0.01em;
        white-space: nowrap;
        margin: 0;
        padding: 0;
        line-height: 1;
      }

      .navbar-toggler {
        grid-column: 3;
      }

      .navbar-collapse {
        grid-column: 2;
        display: flex !important;
        justify-content: center;
        align-items: center;
      }

      .navbar-nav {
        margin-left: auto;
        margin-right: auto;
        gap: 1.15rem;
        align-items: center;
      }

      .navbar-nav .nav-link {
        position: relative;
        font-weight: 600;
        color: #2F3A45 !important;
        padding-left: 0.35rem !important;
        padding-right: 0.35rem !important;
        border-radius: 0;
        background: transparent !important;
        white-space: nowrap;
      }

      .navbar-nav .nav-link:hover {
        color: var(--iea-red) !important;
      }

      .navbar-nav .nav-link.active {
        color: var(--iea-red) !important;
        font-weight: 700;
        background: transparent !important;
      }

      .btn-primary {
        background-color: var(--iea-blue);
        border-color: var(--iea-blue);
      }

      .btn-primary:hover {
        background-color: #005f9e;
        border-color: #005f9e;
      }

      .btn-outline-primary {
        color: var(--iea-blue);
        border-color: var(--iea-blue);
      }

      .btn-outline-primary:hover {
        background-color: var(--iea-blue);
        border-color: var(--iea-blue);
        color: #FFFFFF;
      }

      .btn-modal-action {
        margin-right: 0.4rem;
        margin-bottom: 0.4rem;
      }

      .page-header {
        padding: 1.5rem 0 1rem 0;
      }

      .page-header h2 {
        margin-bottom: 0.25rem;
        font-weight: 700;
        color: #1F2937;
      }

      .page-header h2::after {
        content: '';
        display: block;
        width: 44px;
        height: 3px;
        background: var(--iea-red);
        border-radius: 999px;
        margin-top: 0.55rem;
      }

      .filter-panel {
        background: #FFFFFF;
        border: 1px solid var(--iea-border);
        border-radius: 12px;
        padding: 1rem;
        position: sticky;
        top: 1rem;
        max-width: 100%;
      }

      .filter-panel h5 {
        color: #111827;
        font-weight: 700;
        margin-bottom: 1rem;
      }

      .item-card,
      .scale-card,
      .questionnaire-card {
        border: 1px solid var(--iea-border);
        border-radius: 14px;
        margin-bottom: 1rem;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03);
        background-color: #FFFFFF;
        overflow: hidden;
        max-width: 100%;
      }

      .item-card::before,
      .scale-card::before,
      .questionnaire-card::before {
        content: '';
        display: block;
        height: 4px;
        background: var(--iea-red);
      }

      .item-card:hover,
      .scale-card:hover,
      .questionnaire-card:hover {
        box-shadow: 0 6px 18px rgba(0,0,0,0.07);
        transform: translateY(-1px);
        transition: box-shadow 0.15s ease-in-out, transform 0.15s ease-in-out;
      }

      .item-card h5,
      .scale-card h5,
      .questionnaire-card h5 {
        color: #111827;
        font-weight: 700;
      }

      .item-wording {
        font-size: 1.05rem;
        margin-bottom: 0;
        color: #1F2937;
      }

      .badge {
        margin-right: 0.25rem;
        border: 1px solid var(--iea-border);
        font-weight: 600;
      }

      .badge.text-bg-light {
        background-color: #F8FAFC !important;
        color: #334155 !important;
      }

      .cart-summary {
        background: #FFFFFF;
        border: 1px solid var(--iea-border);
        border-radius: 14px;
        padding: 1rem;
        margin-bottom: 1rem;
        max-width: 100%;
      }

      .empty-state {
        background: #FFFFFF;
        border: 1px dashed #CBD5E1;
        border-radius: 14px;
        padding: 2rem;
        text-align: center;
      }

      .home-hero {
        background: #FFFFFF;
        border: 1px solid var(--iea-border);
        border-radius: 18px;
        padding: 2.25rem;
        margin-top: 1.5rem;
        margin-bottom: 1.5rem;
        position: relative;
        overflow: hidden;
      }

      .home-hero::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        width: 6px;
        height: 100%;
        background: var(--iea-red);
      }

      .home-hero h1 {
        color: #1F2937;
        font-weight: 750;
        max-width: 900px;
      }

      .home-hero .lead {
        color: #4B5563 !important;
      }

      .home-card {
        border: 1px solid var(--iea-border);
        border-radius: 14px;
        background-color: #FFFFFF;
        min-height: 180px;
      }

      .home-card h4 {
        color: var(--iea-blue);
        font-weight: 700;
      }

      a {
        color: var(--iea-blue);
      }

      .app-footer {
        position: relative;
        z-index: 1;
        margin-top: 2rem;
        border-top: 1px solid #D9DEE5;
        background: #FFFFFF;
        padding: 1rem 1.25rem;
        flex-shrink: 0;
      }

      .app-footer-inner {
        max-width: 1400px;
        margin: 0 auto;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        text-align: center;
        gap: 0.25rem;
        font-size: 0.92rem;
        color: var(--iea-muted);
      }

      .footer-line {
        width: 100%;
        text-align: center;
        line-height: 1.35;
      }

      .footer-version strong {
        color: #1F2937;
      }

      .footer-copyright {
        color: var(--iea-muted);
      }

      .version-easter-egg {
        color: var(--iea-muted) !important;
        text-decoration: none !important;
        cursor: default;
        font-weight: 500;
      }

      .version-easter-egg:hover,
      .version-easter-egg:focus,
      .version-easter-egg:active {
        color: var(--iea-muted) !important;
        text-decoration: none !important;
        outline: none !important;
        box-shadow: none !important;
      }

      .dataTables_wrapper {
        margin-bottom: 2rem;
        max-width: 100%;
      }

      .dataTables_wrapper .dataTables_length {
        display: none !important;
      }

      .shiny-bound-output {
        max-width: 100%;
      }

      .card {
        max-width: 100%;
      }

      .detail-modal h5 {
        font-weight: 700;
        color: #1F2937;
        margin-top: 0.5rem;
      }

      .detail-modal h6 {
        font-weight: 700;
        color: #374151;
        margin-top: 0.5rem;
      }

      .detail-modal p {
        margin-bottom: 0.35rem;
      }

      .detail-section {
        background: #F8FAFC;
        border: 1px solid var(--iea-border);
        border-radius: 12px;
        padding: 1rem;
        margin-bottom: 1rem;
      }

      .modal-title {
        font-weight: 700;
        color: #1F2937;
      }

      .modal-content {
        border-radius: 16px;
        border: 1px solid var(--iea-border);
      }

      .modal-header {
        border-bottom: 1px solid var(--iea-border);
      }

      .modal-footer {
        border-top: 1px solid var(--iea-border);
      }

      .modal-xl {
        max-width: 1180px;
      }

      .selectize-control {
        margin-bottom: 1rem;
      }

      .presence-table-wrapper {
        width: 100%;
        overflow-x: auto;
        border: 1px solid var(--iea-border);
        border-radius: 12px;
        background: #FFFFFF;
      }

      .presence-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.88rem;
      }

      .presence-table thead th {
        background: linear-gradient(#0070b8, #005f9e);
        color: #FFFFFF;
        font-weight: 700;
        text-align: center;
        padding: 0.4rem 0.45rem;
        white-space: nowrap;
        border-right: 1px solid rgba(255,255,255,0.25);
      }

      .presence-table tbody tr:nth-child(even) {
        background: #F1F7FD;
      }

      .presence-table tbody tr:nth-child(odd) {
        background: #FFFFFF;
      }

      .presence-row-label {
        text-align: left !important;
        min-width: 170px;
        max-width: 320px;
        white-space: nowrap;
        font-weight: 600;
      }

      .presence-table tbody td {
        padding: 0.35rem 0.45rem;
        border-bottom: 1px solid #E5E7EB;
        border-right: 1px solid #E5E7EB;
        vertical-align: middle;
      }

      .presence-cell {
        text-align: center;
        min-width: 48px;
      }

      .presence-dot {
        display: inline-block;
        width: 13px;
        height: 13px;
        border-radius: 50%;
        vertical-align: middle;
        border: 1px solid #CBD5E1;
      }

      .presence-exact {
        background: var(--iea-red);
        border-color: var(--iea-red);
      }

      .presence-variant {
        background: linear-gradient(to right, var(--iea-red) 50%, transparent 50%);
        border-color: var(--iea-red);
      }

      .presence-empty {
        background: transparent;
        border-color: transparent;
      }

      .presence-legend {
        padding: 0.65rem 0.75rem;
        font-size: 0.85rem;
        color: var(--iea-muted);
        border-top: 1px solid var(--iea-border);
        background: #F8FAFC;
      }

      @media (max-width: 991px) {
        .navbar {
          min-height: 68px;
        }

        .navbar > .container-fluid {
          display: flex;
          align-items: center;
          column-gap: 1rem;
        }

        .navbar-brand {
          min-height: 50px;
        }

        .navbar-collapse {
          justify-content: flex-start;
        }

        .navbar-nav {
          margin-left: 0;
          margin-right: 0;
          gap: 0;
          align-items: flex-start;
        }

        .app-brand {
          min-height: 50px;
          gap: 0.75rem;
        }

        .app-logo {
          height: 40px;
        }

        .app-title {
          min-height: 40px;
        }

        .app-title-main {
          font-size: 1rem;
        }

        .filter-panel {
          position: relative;
          top: auto;
          margin-bottom: 1rem;
        }
      }
    ")),
    tags$script(HTML("
      $(document).on('click', '#version_click', function(e) {
        e.preventDefault();

        window.__ieaVersionClicks = window.__ieaVersionClicks || 0;
        window.__ieaVersionClicks = window.__ieaVersionClicks + 1;

        if (window.__ieaVersionClicks >= 10) {
          window.__ieaVersionClicks = 0;
          window.open('https://www.iea.nl', '_blank');
        }
      });
    "))
  ),
  
  nav_panel(
    "Home",
    div(
      class = "home-hero",
      h1("Explore questionnaire items across studies, cycles and scales"),
      p(
        class = "lead text-muted",
        "Search, compare and export item metadata from IEA studies."
      ),
      div(
        class = "mt-3",
        actionButton("go_items", "Explore items", class = "btn btn-primary"),
        actionButton("go_scales", "Browse scales", class = "btn btn-outline-primary"),
        actionButton("go_categories", "Browse categories", class = "btn btn-outline-primary"),
        actionButton("go_questionnaires", "Browse questionnaires", class = "btn btn-outline-primary")
      )
    ),
    layout_columns(
      col_widths = c(4, 4, 4),
      card(
        class = "home-card",
        card_body(
          h4("Items"),
          p("Search individual items, wording, response options and metadata."),
          span(class = "badge rounded-pill text-bg-light", "Item-level")
        )
      ),
      card(
        class = "home-card",
        card_body(
          h4("Scales"),
          p("Browse constructs and groups of items across cycles."),
          span(class = "badge rounded-pill text-bg-light", "Scale-level")
        )
      ),
      card(
        class = "home-card",
        card_body(
          h4("Questionnaires"),
          p("Explore complete instruments, sections and questionnaire structures."),
          span(class = "badge rounded-pill text-bg-light", "Instrument-level")
        )
      )
    )
  ),
  
  nav_panel(
    "Items",
    page_header(
      "Items",
      "Search and filter questionnaire items across IEA studies."
    ),
    layout_columns(
      col_widths = c(3, 9),
      div(
        class = "filter-panel",
        h5("Filters"),
        uiOutput("filter_study_items_ui"),
        uiOutput("filter_year_items_ui"),
        uiOutput("filter_cycle_items_ui"),
        uiOutput("filter_population_items_ui"),
        uiOutput("filter_instrument_items_ui"),
        uiOutput("filter_scale_items_ui"),
        uiOutput("filter_type_items_ui"),
        uiOutput("filter_puf_items_ui"),
        actionButton(
          "reset_item_filters",
          "Reset filters",
          class = "btn btn-outline-secondary btn-sm"
        )
      ),
      div(
        uiOutput("items_count"),
        textInput(
          "search_items",
          NULL,
          placeholder = "Search within items, wording, variable, category, scale or section..."
        ),
        div(
          class = "mb-2",
          actionButton(
            "grid_add_items",
            "+ Add selected",
            class = "btn btn-sm btn-primary"
          )
        ),
        DTOutput("items_table")
      )
    )
  ),
  
  nav_panel(
    "Scales",
    page_header(
      "Explore Scales",
      "Browse constructs and item groups across studies and cycles."
    ),
    layout_columns(
      col_widths = c(3, 9),
      div(
        class = "filter-panel",
        h5("Filters"),
        uiOutput("filter_study_scales_ui"),
        uiOutput("filter_year_scales_ui"),
        uiOutput("filter_cycle_scales_ui"),
        uiOutput("filter_population_scales_ui"),
        uiOutput("filter_instrument_scales_ui"),
        actionButton(
          "reset_scale_filters",
          "Reset filters",
          class = "btn btn-outline-secondary btn-sm"
        )
      ),
      div(
        uiOutput("scales_count"),
        textInput(
          "search_scales",
          NULL,
          placeholder = "Search within scales, scale variables, study, cycle, target or instrument..."
        ),
        uiOutput("scales_results")
      )
    )
  ),
  
  nav_panel(
    "Categories",
    page_header(
      "Categories",
      "Browse the constructs items belong to, independently of administration."
    ),
    layout_columns(
      col_widths = c(3, 9),
      div(
        class = "filter-panel",
        h5("Filters"),
        uiOutput("filter_study_categories_ui"),
        uiOutput("filter_year_categories_ui"),
        uiOutput("filter_cycle_categories_ui"),
        uiOutput("filter_population_categories_ui"),
        uiOutput("filter_instrument_categories_ui"),
        actionButton(
          "reset_category_filters",
          "Reset filters",
          class = "btn btn-outline-secondary btn-sm"
        )
      ),
      div(
        uiOutput("categories_count"),
        textInput(
          "search_categories",
          NULL,
          placeholder = "Search within categories..."
        ),
        uiOutput("categories_results")
      )
    )
  ),
  
  nav_panel(
    "Questionnaires",
    page_header(
      "Explore Questionnaires",
      "Browse complete questionnaires and instruments."
    ),
    layout_columns(
      col_widths = c(3, 9),
      div(
        class = "filter-panel",
        h5("Filters"),
        uiOutput("filter_study_questionnaires_ui"),
        uiOutput("filter_year_questionnaires_ui"),
        uiOutput("filter_cycle_questionnaires_ui"),
        uiOutput("filter_population_questionnaires_ui"),
        uiOutput("filter_instrument_questionnaires_ui"),
        actionButton(
          "reset_questionnaire_filters",
          "Reset filters",
          class = "btn btn-outline-secondary btn-sm"
        )
      ),
      div(
        uiOutput("questionnaires_count"),
        textInput(
          "search_questionnaires",
          NULL,
          placeholder = "Search within questionnaires, study, cycle, target or instrument..."
        ),
        uiOutput("questionnaires_results")
      )
    )
  ),
  
  nav_panel(
    "Selection Cart",
    page_header(
      "Selection Cart",
      "Review selected unique items and export structured metadata."
    ),
    div(
      class = "cart-summary",
      uiOutput("cart_summary"),
      div(
        class = "mt-2",
        downloadButton(
          "download_excel",
          "Download Excel",
          class = "btn btn-primary"
        ),
        actionButton(
          "clear_cart",
          "Clear cart",
          class = "btn btn-outline-secondary"
        )
      )
    ),
    DTOutput("cart_table")
  )
)

ui <- if (REQUIRE_LOGIN) {
  page_fluid(
    theme = app_theme,
    padding = 0,
    uiOutput("app_ui")
  )
} else {
  main_ui()
}

# ------------------------------------------------------------
# Server
# ------------------------------------------------------------

server <- function(input, output, session) {
  
  current_user <- reactiveVal(
    if (REQUIRE_LOGIN) NULL else list(first_name = "", last_name = "", role = "public")
  )
  
  login_error <- reactiveVal(NULL)
  
  output$app_ui <- renderUI({
    if (is.null(current_user())) login_ui() else main_ui()
  })
  
  output$login_message <- renderUI({
    req(login_error())
    div(class = "text-danger mt-3", login_error())
  })
  
  observeEvent(input$forgot_password_btn, {
    showModal(
      modalDialog(
        title = "Reset password",
        p("Enter your email address. If it is registered, you will receive a password recovery link."),
        textInput("reset_email", "Email"),
        footer = tagList(
          modalButton("Cancel"),
          actionButton("send_reset_btn", "Send recovery email", class = "btn-primary")
        ),
        easyClose = FALSE
      )
    )
  })
  
  observeEvent(input$send_reset_btn, {
    
    email <- trimws(input$reset_email %||% "")
    
    if (!nzchar(email)) {
      showNotification("Enter your email address.", type = "error")
      return()
    }
    
    tryCatch(
      supabase_send_password_reset(email),
      error = function(e) NULL
    )
    
    removeModal()
    
    showNotification(
      "If this email address is registered, a password recovery link has been sent.",
      type = "message",
      duration = 8
    )
  })
  
  observeEvent(input$login_btn, {
    
    login_error(NULL)
    
    email <- trimws(input$login_email %||% "")
    password <- input$login_password %||% ""
    
    if (!nzchar(email) || !nzchar(password)) {
      login_error("Enter your email and password.")
      return()
    }
    
    auth <- supabase_sign_in(email, password)
    
    if (is.null(auth) || is.null(auth$user$id)) {
      login_error("Invalid email or password.")
      return()
    }
    
    profile <- get_user_profile(auth$user$id)
    
    if (nrow(profile) != 1) {
      login_error("This account has no IEABank profile.")
      return()
    }
    
    if (!isTRUE(profile$active[1])) {
      login_error("This account is inactive. Contact an administrator.")
      return()
    }
    
    current_user(as.list(profile[1, ]))
  })
  
  # ----------------------------------------------------------
  # Load data from Supabase
  # ----------------------------------------------------------
  
  items_data <- reactive({
    get_items(pool)
  })
  
  scales_data <- reactive({
    derive_scales(items_data())
  })
  
  scale_admins_data <- reactive({
    derive_scale_admins(items_data())
  })
  
  categories_data <- reactive({
    derive_categories(items_data())
  })
  
  category_admins_data <- reactive({
    derive_category_admins(items_data())
  })
  
  questionnaires_data <- reactive({
    derive_questionnaires(items_data())
  })
  
  examples_data <- reactive({
    get_item_examples(pool)
  })
  
  response_values_data <- reactive({
    read_table_safe(pool, "value_scheme_value")
  })
  
  missing_values_data <- reactive({
    read_table_safe(pool, "miss_scheme_value")
  })
  
  # ----------------------------------------------------------
  # Modal reactive stores
  # ----------------------------------------------------------
  
  modal_item <- reactiveVal(tibble())
  modal_item_versions <- reactiveVal(tibble())
  
  modal_scale_uid <- reactiveVal("")
  
  modal_scale_admins <- reactive({
    scale_admins_data() %>%
      filter(scale_uid == modal_scale_uid())
  })
  
  modal_scale_items <- reactive({
    req(input$scale_admin_selector)
    
    key <- strsplit(input$scale_admin_selector, "|", fixed = TRUE)[[1]]
    
    items_data() %>%
      filter(scale_id == key[1], admin_id == key[2]) %>%
      distinct(item_uid, .keep_all = TRUE) %>%
      arrange(item_uid)
  })
  
  modal_questionnaire_items <- reactiveVal(tibble())
  
  modal_questionnaire_filtered_items <- reactive({
    df <- modal_questionnaire_items()
    
    if (nrow(df) == 0) {
      return(df)
    }
    
    selected_scales <- input$questionnaire_scale_filter
    
    if (!is.null(selected_scales) && length(selected_scales) > 0) {
      df <- df %>%
        filter(scale %in% selected_scales)
    }
    
    df
  })
  
  modal_category_id <- reactiveVal("")
  
  modal_category_items <- reactive({
    derive_category_items(modal_category_id(), items_data())
  })
  
  modal_category_composition <- reactive({
    req(input$category_composition_scope)
    
    key <- strsplit(input$category_composition_scope, "|", fixed = TRUE)[[1]]
    
    derive_composition(
      group_col = "category_id",
      group_val = modal_category_id(),
      items = items_data(),
      study_sel = key[1],
      population_sel = key[2],
      instrument_sel = key[3]
    )
  })
  
  modal_scale_composition <- reactive({
    req(input$scale_composition_scope)
    
    key <- strsplit(input$scale_composition_scope, "|", fixed = TRUE)[[1]]
    
    derive_composition(
      group_col = "scale_uid",
      group_val = modal_scale_uid(),
      items = items_data(),
      study_sel = key[1],
      population_sel = key[2],
      instrument_sel = key[3]
    )
  })
  
  # ----------------------------------------------------------
  # Cart
  # ----------------------------------------------------------
  
  cart <- reactiveVal(tibble())
  
  add_items_to_cart <- function(new_items) {
    
    if (nrow(new_items) == 0) {
      showNotification(
        "No items to add.",
        type = "warning",
        duration = 2
      )
      return(NULL)
    }
    
    current <- cart()
    
    updated <- bind_rows(current, new_items) %>%
      standardize_items() %>%
      distinct(item_admin_id, admin_id, .keep_all = TRUE)
    
    n_added <- nrow(updated) - nrow(current)
    
    cart(updated)
    
    show_added_notification(max(n_added, 0))
  }
  
  # ----------------------------------------------------------
  # Detail modal helpers
  # ----------------------------------------------------------
  
  show_item_details_modal <- function(uid, all_items) {
    
    item <- all_items %>%
      filter(item_uid == uid) %>%
      slice(1)
    
    versions <- derive_item_versions(uid, all_items, examples_data())
    
    modal_item(all_items %>% filter(item_uid == uid))
    modal_item_versions(versions)
    
    showModal(
      modalDialog(
        title = paste("Item details:", uid),
        size = "xl",
        easyClose = TRUE,
        footer = tagList(
          actionButton(
            "modal_add_item_to_cart",
            "Add item to cart",
            class = "btn btn-primary"
          ),
          modalButton("Close")
        ),
        
        div(
          class = "detail-modal",
          
          div(
            class = "detail-section",
            h5(item$item_name),
            p(
              class = "text-muted",
              paste(uid, "\u00b7", item$category_id, "-", item$category_name)
            )
          ),
          
          div(
            class = "detail-section",
            h5("Variants"),
            DTOutput("item_variants_table")
          ),
          
          div(
            class = "detail-section",
            h5("Longitudinal item participation"),
            uiOutput("item_version_matrix")
          ),
          
          div(
            class = "detail-section",
            selectInput(
              "modal_version",
              "Select version",
              choices = setNames(
                paste(versions$item_admin_id, versions$admin_id, sep = "|"),
                paste(versions$item_admin_id, "-", versions$admin_id)
              ),
              width = "320px"
            ),
            uiOutput("modal_version_detail")
          )
        )
      )
    )
  }
  
  show_scale_details_modal <- function(scale_row, all_items) {
    
    modal_scale_uid(scale_row$scale_uid)
    
    admins <- all_items %>%
      filter(scale_uid == scale_row$scale_uid) %>%
      distinct(scale_id, scale_var, scale_varname, admin_id, study, year, population, instrument) %>%
      arrange(year, scale_var)
    
    scopes <- admins %>%
      distinct(study, population, instrument)
    
    showModal(
      modalDialog(
        title = paste("Scale details:", scale_row$scale_name),
        size = "xl",
        easyClose = TRUE,
        footer = modalButton("Close"),
        
        div(
          class = "detail-modal",
          
          div(
            class = "detail-section",
            h5(scale_row$scale_name),
            p(
              class = "text-muted",
              paste(
                scale_row$scale_uid,
                "\u00b7",
                scale_row$n_versions,
                "versions \u00b7",
                scale_row$n_items,
                "items"
              )
            )
          ),
          
          div(
            class = "detail-section",
            h5("Versions"),
            DTOutput("scale_versions_table")
          ),
          
          div(
            class = "detail-section",
            h5("Items in this scale"),
            selectInput(
              inputId = "scale_admin_selector",
              label = "Administration",
              choices = setNames(
                paste(admins$scale_id, admins$admin_id, sep = "|"),
                paste(admins$scale_varname, "-", admins$admin_id)
              ),
              width = "360px"
            ),
            actionButton(
              "modal_add_scale_all",
              "Add all items in this scale",
              class = "btn btn-primary btn-modal-action"
            ),
            actionButton(
              "modal_add_scale_selected",
              "Add selected items",
              class = "btn btn-outline-primary btn-modal-action"
            ),
            DTOutput("scale_items_table")
          ),
          
          div(
            class = "detail-section",
            h5("Longitudinal scale composition"),
            p(
              class = "text-muted",
              "Rows are item UIDs, columns are years. A cell shows the item version used that year; a dash means the item was administered but was not part of the scale."
            ),
            selectInput(
              inputId = "scale_composition_scope",
              label = "Study and instrument",
              choices = setNames(
                paste(scopes$study, scopes$population, scopes$instrument, sep = "|"),
                paste0(scopes$study, " \u00b7 ", scopes$population, scopes$instrument)
              ),
              width = "360px"
            ),
            uiOutput("scale_composition_matrix")
          )
        )
      )
    )
  }
  
  show_category_details_modal <- function(category_row, all_items) {
    
    modal_category_id(category_row$category_id)
    
    scopes <- all_items %>%
      filter(category_id == category_row$category_id) %>%
      distinct(study, population, instrument)
    
    showModal(
      modalDialog(
        title = paste("Category details:", category_row$category_name),
        size = "xl",
        easyClose = TRUE,
        footer = modalButton("Close"),
        
        div(
          class = "detail-modal",
          
          div(
            class = "detail-section",
            h5(category_row$category_name),
            p(
              class = "text-muted",
              paste(
                category_row$category_id,
                "\u00b7",
                category_row$n_items,
                "items \u00b7",
                category_row$studies,
                "\u00b7",
                category_row$years
              )
            )
          ),
          
          div(
            class = "detail-section",
            h5("Items in this category"),
            actionButton(
              "modal_add_category_all",
              "Add all items in this category",
              class = "btn btn-primary btn-modal-action"
            ),
            DTOutput("category_items_table")
          ),
          
          div(
            class = "detail-section",
            h5("Longitudinal category composition"),
            p(
              class = "text-muted",
              "Rows are item UIDs, columns are years. A cell shows the item version used that year."
            ),
            selectInput(
              inputId = "category_composition_scope",
              label = "Study and instrument",
              choices = setNames(
                paste(scopes$study, scopes$population, scopes$instrument, sep = "|"),
                paste0(scopes$study, " \u00b7 ", scopes$population, scopes$instrument)
              ),
              width = "360px"
            ),
            uiOutput("category_composition_matrix")
          )
        )
      )
    )
  }
  
  show_questionnaire_details_modal <- function(questionnaire_row, all_items) {
    
    questionnaire_items <- all_items %>%
      filter(
        study == questionnaire_row$study,
        year == questionnaire_row$year,
        cycle == questionnaire_row$cycle,
        population == questionnaire_row$population,
        instrument == questionnaire_row$instrument
      ) %>%
      distinct(item_uid, .keep_all = TRUE) %>%
      arrange(section, scale, item_uid)
    
    modal_questionnaire_items(questionnaire_items)
    
    scale_choices <- questionnaire_items %>%
      filter(!is.na(scale), scale != "") %>%
      distinct(scale) %>%
      arrange(scale) %>%
      pull(scale)
    
    showModal(
      modalDialog(
        title = paste(
          "Questionnaire details:",
          questionnaire_row$study,
          questionnaire_row$year,
          questionnaire_row$population,
          questionnaire_row$instrument
        ),
        size = "xl",
        easyClose = TRUE,
        footer = modalButton("Close"),
        
        div(
          class = "detail-modal",
          
          layout_columns(
            col_widths = c(6, 6),
            
            div(
              class = "detail-section",
              h5("Questionnaire"),
              p(strong("Study: "), questionnaire_row$study),
              p(strong("Year: "), questionnaire_row$year),
              p(strong("Cycle: "), questionnaire_row$cycle)
            ),
            
            div(
              class = "detail-section",
              h5("Instrument"),
              p(strong("Target: "), questionnaire_row$population),
              p(strong("Instrument: "), questionnaire_row$instrument),
              p(strong("Items: "), questionnaire_row$n_items),
              p(strong("Sections: "), questionnaire_row$n_sections),
              p(strong("Scales: "), questionnaire_row$n_scales)
            )
          ),
          
          div(
            class = "detail-section",
            h5("Items in this questionnaire"),
            p(
              class = "text-muted",
              "Filter by one or more scales, select rows to add specific items, or add the complete questionnaire."
            ),
            actionButton(
              "modal_add_questionnaire_all",
              "Add complete questionnaire",
              class = "btn btn-primary btn-modal-action"
            ),
            actionButton(
              "modal_add_questionnaire_selected",
              "Add selected items",
              class = "btn btn-outline-primary btn-modal-action"
            ),
            selectizeInput(
              inputId = "questionnaire_scale_filter",
              label = "Filter by scale",
              choices = scale_choices,
              selected = NULL,
              multiple = TRUE,
              options = list(
                placeholder = "Select one or more scales; leave empty to show all items"
              )
            ),
            DTOutput("questionnaire_items_table")
          )
        )
      )
    )
  }
  
  # ----------------------------------------------------------
  # Modal outputs
  # ----------------------------------------------------------
  
  output$item_variants_table <- renderDT({
    datatable_variants(
      modal_item_versions() %>%
        mutate(pop_instrument = paste0(population, instrument)) %>%
        select(
          item_admin_id,
          study,
          year,
          pop_instrument,
          source_variable,
          scale_varnames,
          examples
        )
    )
  })
  
  output$item_version_matrix <- renderUI({
    version_matrix_ui(
      derive_item_version_matrix(modal_item_versions())
    )
  })
  
  selected_version <- reactive({
    req(input$modal_version)
    
    key <- strsplit(input$modal_version, "|", fixed = TRUE)[[1]]
    
    modal_item_versions() %>%
      filter(item_admin_id == key[1], admin_id == key[2])
  })
  
  output$modal_version_detail <- renderUI({
    v <- selected_version()
    
    responses <- response_values_data() %>%
      filter(response_id == v$response_id)
    
    missings <- missing_values_data() %>%
      filter(miss_id == v$miss_id)
    
    tagList(
      layout_columns(
        col_widths = c(6, 6),
        
        div(
          h6("Wording"),
          p(strong("Heading: "), v$wording_heading),
          p(strong("Question: "), v$wording_question),
          p(strong("Instruction: "), v$wording_instruction),
          p(strong("Item: "), v$wording)
        ),
        
        div(
          h6("Administration"),
          p(strong("Administration: "), v$admin_id),
          p(strong("Variable: "), v$source_variable),
          p(strong("Dataset label: "), v$dataset_label),
          p(strong("Item type: "), v$item_type),
          p(strong("PUF: "), as.character(v$puf)),
          p(strong("Scale(s): "), ifelse(v$scale_names == "", "Not assigned", v$scale_names))
        )
      ),
      
      layout_columns(
        col_widths = c(6, 6),
        
        div(
          h6(paste("Response scheme:", v$response_id)),
          value_table_ui(responses %>% select(value, label))
        ),
        
        div(
          h6(paste("Missing scheme:", v$miss_id)),
          value_table_ui(missings %>% select(value, category))
        )
      ),
      
      actionButton(
        "modal_add_version_to_cart",
        paste("+ Add", v$item_admin_id, "-", v$admin_id),
        class = "btn btn-sm btn-primary"
      )
    )
  })
  
  output$category_items_table <- renderDT({
    datatable_items_grid(
      modal_category_items() %>%
        mutate(category = "") %>%
        select(item_uid, item_name, category, studies, years, n_versions, details)
    )
  })
  
  output$category_composition_matrix <- renderUI({
    version_matrix_ui(
      modal_category_composition(),
      row_label = "Item UID"
    )
  })
  
  output$scale_versions_table <- renderDT({
    datatable_simple(
      modal_scale_admins() %>%
        mutate(pop_instrument = paste0(population, instrument)) %>%
        select(
          scale_var,
          scale_id,
          scale_varname,
          scale,
          study,
          year,
          pop_instrument
        )
    )
  })
  
  output$scale_composition_matrix <- renderUI({
    version_matrix_ui(
      modal_scale_composition(),
      row_label = "Item UID"
    )
  })
  
  output$scale_items_table <- renderDT({
    df <- modal_scale_items()
    
    if (nrow(df) == 0) {
      return(datatable_simple(tibble(message = "No items available.")))
    }
    
    df %>%
      select(
        item_uid,
        item_code,
        source_variable,
        wording,
        section,
        item_type,
        trend_status,
        puf
      ) %>%
      datatable_all_items(selection = "multiple")
  })
  
  output$questionnaire_items_table <- renderDT({
    df <- modal_questionnaire_filtered_items()
    
    if (nrow(df) == 0) {
      return(datatable_simple(tibble(message = "No items available for the selected scale filter.")))
    }
    
    df %>%
      select(
        section,
        scale,
        scale_varname,
        item_uid,
        item_code,
        source_variable,
        wording,
        item_type,
        trend_status,
        puf
      ) %>%
      datatable_all_items(selection = "multiple")
  })
  
  # ----------------------------------------------------------
  # Modal add buttons
  # ----------------------------------------------------------
  
  observeEvent(input$modal_add_version_to_cart, {
    v <- selected_version()
    
    add_items_to_cart(
      modal_item() %>%
        filter(item_admin_id == v$item_admin_id, admin_id == v$admin_id)
    )
  })
  
  observeEvent(input$modal_add_item_to_cart, {
    item <- modal_item()
    
    if (nrow(item) > 0) {
      add_items_to_cart(item)
    }
  })
  
  observeEvent(input$modal_add_category_all, {
    add_items_to_cart(
      items_data() %>%
        filter(category_id == modal_category_id())
    )
  })
  
  observeEvent(input$modal_add_scale_all, {
    add_items_to_cart(modal_scale_items())
  })
  
  observeEvent(input$modal_add_scale_selected, {
    df <- modal_scale_items()
    selected_rows <- input$scale_items_table_rows_selected
    
    if (is.null(selected_rows) || length(selected_rows) == 0) {
      showNotification(
        "Select at least one item from the scale table.",
        type = "warning",
        duration = 2
      )
      return(NULL)
    }
    
    add_items_to_cart(df[selected_rows, , drop = FALSE])
  })
  
  observeEvent(input$modal_add_questionnaire_all, {
    add_items_to_cart(modal_questionnaire_items())
  })
  
  observeEvent(input$modal_add_questionnaire_selected, {
    df <- modal_questionnaire_filtered_items()
    selected_rows <- input$questionnaire_items_table_rows_selected
    
    if (is.null(selected_rows) || length(selected_rows) == 0) {
      showNotification(
        "Select at least one item from the questionnaire table.",
        type = "warning",
        duration = 2
      )
      return(NULL)
    }
    
    add_items_to_cart(df[selected_rows, , drop = FALSE])
  })
  
  # ----------------------------------------------------------
  # Navigation buttons from Home
  # ----------------------------------------------------------
  
  observeEvent(input$go_items, {
    nav_select("main_nav", "Items")
  })
  
  observeEvent(input$go_scales, {
    nav_select("main_nav", "Scales")
  })
  
  observeEvent(input$go_categories, {
    nav_select("main_nav", "Categories")
  })
  
  observeEvent(input$go_questionnaires, {
    nav_select("main_nav", "Questionnaires")
  })
  
  # ----------------------------------------------------------
  # Dynamic filter UI: Items
  # ----------------------------------------------------------
  
  output$filter_study_items_ui <- renderUI({
    choices <- sort(unique(na.omit(items_data()$study)))
    selectInput("filter_study_items", "Study", choices = c("All", choices))
  })
  
  output$filter_year_items_ui <- renderUI({
    choices <- sort(unique(na.omit(items_data()$year)))
    selectInput("filter_year_items", "Year", choices = c("All", choices))
  })
  
  output$filter_cycle_items_ui <- renderUI({
    choices <- sort(unique(na.omit(items_data()$cycle)))
    selectInput("filter_cycle_items", "Cycle", choices = c("All", choices))
  })
  
  output$filter_population_items_ui <- renderUI({
    choices <- sort(unique(na.omit(items_data()$population)))
    selectInput("filter_population_items", "Target", choices = c("All", choices))
  })
  
  output$filter_instrument_items_ui <- renderUI({
    choices <- sort(unique(na.omit(items_data()$instrument)))
    selectInput("filter_instrument_items", "Instrument", choices = c("All", choices))
  })
  
  output$filter_scale_items_ui <- renderUI({
    choices <- sort(unique(na.omit(items_data()$scale)))
    selectInput("filter_scale_items", "Scale", choices = c("All", choices))
  })
  
  output$filter_type_items_ui <- renderUI({
    choices <- sort(unique(na.omit(items_data()$item_type)))
    selectInput("filter_type_items", "Item type", choices = c("All", choices))
  })
  
  output$filter_puf_items_ui <- renderUI({
    selectInput(
      "filter_puf_items",
      "Public use file",
      choices = c("All", "TRUE", "FALSE")
    )
  })
  
  # ----------------------------------------------------------
  # Filtered items
  # ----------------------------------------------------------
  
  filtered_items <- reactive({
    df <- items_data()
    
    if (!is.null(input$filter_study_items) && input$filter_study_items != "All") {
      df <- df %>% filter(study == input$filter_study_items)
    }
    
    if (!is.null(input$filter_year_items) && input$filter_year_items != "All") {
      df <- df %>% filter(year == input$filter_year_items)
    }
    
    if (!is.null(input$filter_cycle_items) && input$filter_cycle_items != "All") {
      df <- df %>% filter(cycle == input$filter_cycle_items)
    }
    
    if (!is.null(input$filter_population_items) && input$filter_population_items != "All") {
      df <- df %>% filter(population == input$filter_population_items)
    }
    
    if (!is.null(input$filter_instrument_items) && input$filter_instrument_items != "All") {
      df <- df %>% filter(instrument == input$filter_instrument_items)
    }
    
    if (!is.null(input$filter_scale_items) && input$filter_scale_items != "All") {
      df <- df %>% filter(scale == input$filter_scale_items)
    }
    
    if (!is.null(input$filter_type_items) && input$filter_type_items != "All") {
      df <- df %>% filter(item_type == input$filter_type_items)
    }
    
    if (!is.null(input$filter_puf_items) && input$filter_puf_items != "All") {
      df <- df %>% filter(as.character(puf) == input$filter_puf_items)
    }
    
    if (!is.null(input$search_items) && nzchar(input$search_items)) {
      q <- tolower(input$search_items)
      
      df <- df %>%
        filter(
          grepl(q, tolower(item_code), fixed = TRUE) |
            grepl(q, tolower(item_uid), fixed = TRUE) |
            grepl(q, tolower(item_name), fixed = TRUE) |
            grepl(q, tolower(category_id), fixed = TRUE) |
            grepl(q, tolower(category_name), fixed = TRUE) |
            grepl(q, tolower(wording), fixed = TRUE) |
            grepl(q, tolower(source_variable), fixed = TRUE) |
            grepl(q, tolower(scale), fixed = TRUE) |
            grepl(q, tolower(scale_varname), fixed = TRUE) |
            grepl(q, tolower(section), fixed = TRUE) |
            grepl(q, tolower(dataset_label), fixed = TRUE)
        )
    }
    
    df
  })
  
  items_grid <- reactive({
    filtered_items() %>%
      group_by(item_uid) %>%
      summarise(
        item_name = first(item_name),
        category = first(category_name),
        studies = paste(sort(unique(study[study != ""])), collapse = ", "),
        years = paste(sort(unique(year[year != ""])), collapse = ", "),
        n_versions = n_distinct(item_code),
        .groups = "drop"
      ) %>%
      arrange(item_uid) %>%
      mutate(details = grid_details_button(item_uid))
  })
  
  output$items_count <- renderUI({
    h5(paste(nrow(items_grid()), "items found"))
  })
  
  output$items_table <- renderDT({
    datatable_items_grid(items_grid())
  })
  
  selected_grid_uids <- reactive({
    items_grid()$item_uid[input$items_table_rows_selected]
  })
  
  observeEvent(input$grid_add_items, {
    add_items_to_cart(
      items_data() %>%
        filter(item_uid %in% selected_grid_uids())
    )
  })
  
  observeEvent(input$grid_details_uid, {
    show_item_details_modal(input$grid_details_uid, items_data())
  })
  
  observeEvent(input$reset_item_filters, {
    updateSelectInput(session, "filter_study_items", selected = "All")
    updateSelectInput(session, "filter_year_items", selected = "All")
    updateSelectInput(session, "filter_cycle_items", selected = "All")
    updateSelectInput(session, "filter_population_items", selected = "All")
    updateSelectInput(session, "filter_instrument_items", selected = "All")
    updateSelectInput(session, "filter_scale_items", selected = "All")
    updateSelectInput(session, "filter_type_items", selected = "All")
    updateSelectInput(session, "filter_puf_items", selected = "All")
    updateTextInput(session, "search_items", value = "")
  })
  
  # ----------------------------------------------------------
  # Dynamic filter UI: Scales
  # ----------------------------------------------------------
  
  output$filter_study_scales_ui <- renderUI({
    choices <- sort(unique(na.omit(scale_admins_data()$study)))
    selectInput("filter_study_scales", "Study", choices = c("All", choices))
  })
  
  output$filter_year_scales_ui <- renderUI({
    choices <- sort(unique(na.omit(scale_admins_data()$year)))
    selectInput("filter_year_scales", "Year", choices = c("All", choices))
  })
  
  output$filter_cycle_scales_ui <- renderUI({
    choices <- sort(unique(na.omit(scale_admins_data()$cycle)))
    selectInput("filter_cycle_scales", "Cycle", choices = c("All", choices))
  })
  
  output$filter_population_scales_ui <- renderUI({
    choices <- sort(unique(na.omit(scale_admins_data()$population)))
    selectInput("filter_population_scales", "Target", choices = c("All", choices))
  })
  
  output$filter_instrument_scales_ui <- renderUI({
    choices <- sort(unique(na.omit(scale_admins_data()$instrument)))
    selectInput("filter_instrument_scales", "Instrument", choices = c("All", choices))
  })
  
  # ----------------------------------------------------------
  # Filtered scales
  # ----------------------------------------------------------
  
  filtered_scales <- reactive({
    admins <- scale_admins_data()
    
    if (!is.null(input$filter_study_scales) && input$filter_study_scales != "All") {
      admins <- admins %>% filter(study == input$filter_study_scales)
    }
    
    if (!is.null(input$filter_year_scales) && input$filter_year_scales != "All") {
      admins <- admins %>% filter(year == input$filter_year_scales)
    }
    
    if (!is.null(input$filter_cycle_scales) && input$filter_cycle_scales != "All") {
      admins <- admins %>% filter(cycle == input$filter_cycle_scales)
    }
    
    if (!is.null(input$filter_population_scales) && input$filter_population_scales != "All") {
      admins <- admins %>% filter(population == input$filter_population_scales)
    }
    
    if (!is.null(input$filter_instrument_scales) && input$filter_instrument_scales != "All") {
      admins <- admins %>% filter(instrument == input$filter_instrument_scales)
    }
    
    df <- scales_data() %>%
      filter(scale_uid %in% admins$scale_uid)
    
    if (!is.null(input$search_scales) && nzchar(input$search_scales)) {
      q <- tolower(input$search_scales)
      
      matching_uids <- admins %>%
        filter(
          grepl(q, tolower(scale), fixed = TRUE) |
            grepl(q, tolower(scale_id), fixed = TRUE) |
            grepl(q, tolower(scale_varname), fixed = TRUE)
        ) %>%
        pull(scale_uid)
      
      df <- df %>%
        filter(
          grepl(q, tolower(scale_uid), fixed = TRUE) |
            grepl(q, tolower(scale_name), fixed = TRUE) |
            grepl(q, tolower(varnames), fixed = TRUE) |
            grepl(q, tolower(studies), fixed = TRUE) |
            grepl(q, tolower(years), fixed = TRUE) |
            scale_uid %in% matching_uids
        )
    }
    
    df
  })
  
  output$scales_count <- renderUI({
    h5(paste(nrow(filtered_scales()), "scales found"))
  })
  
  output$scales_results <- renderUI({
    df <- filtered_scales()
    
    if (nrow(df) == 0) {
      return(empty_state("No scales found."))
    }
    
    tagList(
      lapply(seq_len(nrow(df)), function(i) {
        scale_card(df[i, ])
      })
    )
  })
  
  observeEvent(input$reset_scale_filters, {
    updateSelectInput(session, "filter_study_scales", selected = "All")
    updateSelectInput(session, "filter_year_scales", selected = "All")
    updateSelectInput(session, "filter_cycle_scales", selected = "All")
    updateSelectInput(session, "filter_population_scales", selected = "All")
    updateSelectInput(session, "filter_instrument_scales", selected = "All")
    updateTextInput(session, "search_scales", value = "")
  })
  
  observe({
    df <- scales_data()
    
    if (nrow(df) == 0) {
      return(NULL)
    }
    
    lapply(seq_len(nrow(df)), function(i) {
      local({
        scale_row <- df[i, ]
        scale_uid_safe <- str_replace_all(scale_row$scale_uid, "[^A-Za-z0-9_]", "_")
        
        observeEvent(input[[paste0("add_scale_", scale_uid_safe)]], {
          add_items_to_cart(
            items_data() %>%
              filter(scale_uid == scale_row$scale_uid)
          )
        }, ignoreInit = TRUE)
        
        observeEvent(input[[paste0("details_scale_", scale_uid_safe)]], {
          show_scale_details_modal(scale_row, items_data())
        }, ignoreInit = TRUE)
      })
    })
  })
  
  # ----------------------------------------------------------
  # Dynamic filter UI: Categories
  # ----------------------------------------------------------
  
  output$filter_study_categories_ui <- renderUI({
    choices <- sort(unique(na.omit(category_admins_data()$study)))
    selectInput("filter_study_categories", "Study", choices = c("All", choices))
  })
  
  output$filter_year_categories_ui <- renderUI({
    choices <- sort(unique(na.omit(category_admins_data()$year)))
    selectInput("filter_year_categories", "Year", choices = c("All", choices))
  })
  
  output$filter_cycle_categories_ui <- renderUI({
    choices <- sort(unique(na.omit(category_admins_data()$cycle)))
    selectInput("filter_cycle_categories", "Cycle", choices = c("All", choices))
  })
  
  output$filter_population_categories_ui <- renderUI({
    choices <- sort(unique(na.omit(category_admins_data()$population)))
    selectInput("filter_population_categories", "Target", choices = c("All", choices))
  })
  
  output$filter_instrument_categories_ui <- renderUI({
    choices <- sort(unique(na.omit(category_admins_data()$instrument)))
    selectInput("filter_instrument_categories", "Instrument", choices = c("All", choices))
  })
  
  # ----------------------------------------------------------
  # Filtered categories
  # ----------------------------------------------------------
  
  filtered_categories <- reactive({
    admins <- category_admins_data()
    
    if (!is.null(input$filter_study_categories) && input$filter_study_categories != "All") {
      admins <- admins %>% filter(study == input$filter_study_categories)
    }
    
    if (!is.null(input$filter_year_categories) && input$filter_year_categories != "All") {
      admins <- admins %>% filter(year == input$filter_year_categories)
    }
    
    if (!is.null(input$filter_cycle_categories) && input$filter_cycle_categories != "All") {
      admins <- admins %>% filter(cycle == input$filter_cycle_categories)
    }
    
    if (!is.null(input$filter_population_categories) && input$filter_population_categories != "All") {
      admins <- admins %>% filter(population == input$filter_population_categories)
    }
    
    if (!is.null(input$filter_instrument_categories) && input$filter_instrument_categories != "All") {
      admins <- admins %>% filter(instrument == input$filter_instrument_categories)
    }
    
    df <- categories_data() %>%
      filter(category_id %in% admins$category_id)
    
    if (!is.null(input$search_categories) && nzchar(input$search_categories)) {
      q <- tolower(input$search_categories)
      
      df <- df %>%
        filter(
          grepl(q, tolower(category_id), fixed = TRUE) |
            grepl(q, tolower(category_name), fixed = TRUE)
        )
    }
    
    df
  })
  
  output$categories_count <- renderUI({
    h5(paste(nrow(filtered_categories()), "categories found"))
  })
  
  output$categories_results <- renderUI({
    df <- filtered_categories()
    
    if (nrow(df) == 0) {
      return(empty_state("No categories found."))
    }
    
    tagList(
      lapply(seq_len(nrow(df)), function(i) {
        category_card(df[i, ])
      })
    )
  })
  
  observeEvent(input$reset_category_filters, {
    updateSelectInput(session, "filter_study_categories", selected = "All")
    updateSelectInput(session, "filter_year_categories", selected = "All")
    updateSelectInput(session, "filter_cycle_categories", selected = "All")
    updateSelectInput(session, "filter_population_categories", selected = "All")
    updateSelectInput(session, "filter_instrument_categories", selected = "All")
    updateTextInput(session, "search_categories", value = "")
  })
  
  observe({
    df <- categories_data()
    
    if (nrow(df) == 0) {
      return(NULL)
    }
    
    lapply(seq_len(nrow(df)), function(i) {
      local({
        category_row <- df[i, ]
        category_id_safe <- str_replace_all(category_row$category_id, "[^A-Za-z0-9_]", "_")
        
        observeEvent(input[[paste0("add_category_", category_id_safe)]], {
          add_items_to_cart(
            items_data() %>%
              filter(category_id == category_row$category_id)
          )
        }, ignoreInit = TRUE)
        
        observeEvent(input[[paste0("details_category_", category_id_safe)]], {
          show_category_details_modal(category_row, items_data())
        }, ignoreInit = TRUE)
      })
    })
  })
  
  # ----------------------------------------------------------
  # Dynamic filter UI: Questionnaires
  # ----------------------------------------------------------
  
  output$filter_study_questionnaires_ui <- renderUI({
    choices <- sort(unique(na.omit(questionnaires_data()$study)))
    selectInput("filter_study_questionnaires", "Study", choices = c("All", choices))
  })
  
  output$filter_year_questionnaires_ui <- renderUI({
    choices <- sort(unique(na.omit(questionnaires_data()$year)))
    selectInput("filter_year_questionnaires", "Year", choices = c("All", choices))
  })
  
  output$filter_cycle_questionnaires_ui <- renderUI({
    choices <- sort(unique(na.omit(questionnaires_data()$cycle)))
    selectInput("filter_cycle_questionnaires", "Cycle", choices = c("All", choices))
  })
  
  output$filter_population_questionnaires_ui <- renderUI({
    choices <- sort(unique(na.omit(questionnaires_data()$population)))
    selectInput("filter_population_questionnaires", "Target", choices = c("All", choices))
  })
  
  output$filter_instrument_questionnaires_ui <- renderUI({
    choices <- sort(unique(na.omit(questionnaires_data()$instrument)))
    selectInput("filter_instrument_questionnaires", "Instrument", choices = c("All", choices))
  })
  
  # ----------------------------------------------------------
  # Filtered questionnaires
  # ----------------------------------------------------------
  
  filtered_questionnaires <- reactive({
    df <- questionnaires_data()
    
    if (!is.null(input$filter_study_questionnaires) && input$filter_study_questionnaires != "All") {
      df <- df %>% filter(study == input$filter_study_questionnaires)
    }
    
    if (!is.null(input$filter_year_questionnaires) && input$filter_year_questionnaires != "All") {
      df <- df %>% filter(year == input$filter_year_questionnaires)
    }
    
    if (!is.null(input$filter_cycle_questionnaires) && input$filter_cycle_questionnaires != "All") {
      df <- df %>% filter(cycle == input$filter_cycle_questionnaires)
    }
    
    if (!is.null(input$filter_population_questionnaires) && input$filter_population_questionnaires != "All") {
      df <- df %>% filter(population == input$filter_population_questionnaires)
    }
    
    if (!is.null(input$filter_instrument_questionnaires) && input$filter_instrument_questionnaires != "All") {
      df <- df %>% filter(instrument == input$filter_instrument_questionnaires)
    }
    
    if (!is.null(input$search_questionnaires) && nzchar(input$search_questionnaires)) {
      q <- tolower(input$search_questionnaires)
      
      df <- df %>%
        filter(
          grepl(q, tolower(study), fixed = TRUE) |
            grepl(q, tolower(year), fixed = TRUE) |
            grepl(q, tolower(cycle), fixed = TRUE) |
            grepl(q, tolower(population), fixed = TRUE) |
            grepl(q, tolower(instrument), fixed = TRUE) |
            grepl(q, tolower(questionnaire_uid), fixed = TRUE)
        )
    }
    
    df
  })
  
  output$questionnaires_count <- renderUI({
    h5(paste(nrow(filtered_questionnaires()), "questionnaires found"))
  })
  
  output$questionnaires_results <- renderUI({
    df <- filtered_questionnaires()
    
    if (nrow(df) == 0) {
      return(empty_state("No questionnaires found."))
    }
    
    tagList(
      lapply(seq_len(nrow(df)), function(i) {
        questionnaire_card(df[i, ])
      })
    )
  })
  
  observeEvent(input$reset_questionnaire_filters, {
    updateSelectInput(session, "filter_study_questionnaires", selected = "All")
    updateSelectInput(session, "filter_year_questionnaires", selected = "All")
    updateSelectInput(session, "filter_cycle_questionnaires", selected = "All")
    updateSelectInput(session, "filter_population_questionnaires", selected = "All")
    updateSelectInput(session, "filter_instrument_questionnaires", selected = "All")
    updateTextInput(session, "search_questionnaires", value = "")
  })
  
  observe({
    df <- questionnaires_data()
    
    if (nrow(df) == 0) {
      return(NULL)
    }
    
    lapply(seq_len(nrow(df)), function(i) {
      local({
        questionnaire_row <- df[i, ]
        questionnaire_uid_safe <- str_replace_all(
          questionnaire_row$questionnaire_uid,
          "[^A-Za-z0-9_]",
          "_"
        )
        
        observeEvent(input[[paste0("add_questionnaire_", questionnaire_uid_safe)]], {
          selected_items <- items_data() %>%
            filter(
              study == questionnaire_row$study,
              year == questionnaire_row$year,
              cycle == questionnaire_row$cycle,
              population == questionnaire_row$population,
              instrument == questionnaire_row$instrument
            ) %>%
            distinct(item_uid, .keep_all = TRUE)
          
          add_items_to_cart(selected_items)
        }, ignoreInit = TRUE)
        
        observeEvent(input[[paste0("details_questionnaire_", questionnaire_uid_safe)]], {
          show_questionnaire_details_modal(questionnaire_row, items_data())
        }, ignoreInit = TRUE)
      })
    })
  })
  
  # ----------------------------------------------------------
  # Cart outputs
  # ----------------------------------------------------------
  
  output$cart_summary <- renderUI({
    n <- nrow(cart())
    
    if (n == 0) {
      p("No items selected yet.", class = "text-muted")
    } else {
      h5(
        paste(
          n_distinct(cart()$item_uid),
          "items /",
          n,
          "versions selected"
        )
      )
    }
  })
  
  output$cart_table <- renderDT({
    df <- cart()
    
    if (nrow(df) == 0) {
      return(
        datatable_cart(
          tibble(
            item_code = character(),
            wording = character(),
            study = character(),
            year = character(),
            cycle = character(),
            population = character(),
            instrument = character(),
            scale = character()
          )
        )
      )
    }
    
    df %>%
      select(
        item_uid,
        item_code,
        item_name,
        category_id,
        category_name,
        wording,
        source_variable,
        study,
        phase,
        year,
        cycle,
        population,
        instrument,
        scale,
        scale_varname,
        item_type,
        dataset_label,
        puf
      ) %>%
      datatable_cart()
  })
  
  observeEvent(input$clear_cart, {
    cart(tibble())
  })
  
  # ----------------------------------------------------------
  # Excel export
  # ----------------------------------------------------------
  
  output$download_excel <- downloadHandler(
    filename = function() {
      paste0("iea_item_bank_selection_", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      
      selected <- cart() %>%
        standardize_items()
      
      selected_item_uids <- selected$item_uid
      
      item_history <- get_item_history(pool, selected_item_uids)
      
      item_variants <- get_item_variants(pool, selected_item_uids)
      
      scale_membership <- get_scale_membership(
        pool,
        selected
      )
      
      response_options <- get_response_options(
        pool,
        selected
      )
      
      missing_values <- get_missing_values(
        pool,
        selected
      )
      
      variables_and_data <- get_variables_and_data(
        pool,
        selected
      )
      
      export_metadata <- tibble(
        field = c(
          "export_date",
          "source",
          "app_version",
          "number_of_selected_items"
        ),
        value = c(
          as.character(Sys.Date()),
          "IEA Item Bank Explorer",
          APP_VERSION,
          as.character(n_distinct(selected$item_uid))
        )
      )
      
      wb <- createWorkbook()
      
      addWorksheet(wb, "Selected_Items")
      writeData(wb, "Selected_Items", selected)
      
      addWorksheet(wb, "Item_History")
      writeData(wb, "Item_History", item_history)
      
      addWorksheet(wb, "Item_Variants")
      writeData(wb, "Item_Variants", item_variants)
      
      addWorksheet(wb, "Scale_Membership")
      writeData(wb, "Scale_Membership", scale_membership)
      
      addWorksheet(wb, "Response_Options")
      writeData(wb, "Response_Options", response_options)
      
      addWorksheet(wb, "Missing_Values")
      writeData(wb, "Missing_Values", missing_values)
      
      addWorksheet(wb, "Variables_and_Data")
      writeData(wb, "Variables_and_Data", variables_and_data)
      
      addWorksheet(wb, "Export_Metadata")
      writeData(wb, "Export_Metadata", export_metadata)
      
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}

# ------------------------------------------------------------
# Run app
# ------------------------------------------------------------

shinyApp(ui, server)