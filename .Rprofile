source("renv/activate.R")

# IEABank schema compatibility layer (2026-09-18).
# The User GUI keeps its stable internal field names while Supabase
# uses the definitive IEABank schema from student_v2_2_teacher_v1_0.xlsx.
tbl <- function(src, from, ...) {
  if (!is.character(from) || length(from) != 1L) {
    return(dplyr::tbl(src, from, ...))
  }

  requested <- from
  mapped <- switch(
    requested,
    item_id = "item",
    category = "item_category",
    value_scheme = "response_scheme",
    value_scheme_value = "response_scheme_values",
    response_scheme_value = "response_scheme_values",
    miss_scheme_value = "miss_scheme_values",
    scale_items = "scale_item",
    requested
  )

  out <- dplyr::tbl(src, mapped, ...)

  if (identical(requested, "admin")) {
    out <- dplyr::mutate(
      out,
      instrument = .data$instrument_name,
      target = .data$pop,
      cycle = .data$cycle_num
    )
  }

  if (identical(requested, "item_admin")) {
    out <- dplyr::mutate(
      out,
      item_admin_id = .data$item_id,
      puf = .data$is_puf_quest
    )
  }

  if (requested %in% c("item_translation", "stimulus_admin", "value_scheme_translation")) {
    out <- dplyr::mutate(out, puf = .data$is_puf)
  }

  out
}
