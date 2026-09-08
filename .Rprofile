source("renv/activate.R")

# IEABank schema compatibility layer (2026-09).
# Keep the User GUI stable while the database uses the new canonical names.
tbl <- function(src, from, ...) {
  if (!is.character(from) || length(from) != 1L) {
    return(dplyr::tbl(src, from, ...))
  }

  requested <- from
  mapped <- switch(
    requested,
    item_id = "item",
    value_scheme = "response_scheme",
    value_scheme_value = "response_scheme_value",
    requested
  )

  out <- dplyr::tbl(src, mapped, ...)

  if (identical(requested, "item_admin")) {
    out <- dplyr::mutate(
      out,
      item_admin_id = .data$item_id,
      puf = .data$is_puf
    )
  }

  if (requested %in% c("item_translation", "stimulus_admin", "value_scheme_translation")) {
    out <- dplyr::mutate(out, puf = .data$is_puf)
  }

  out
}
