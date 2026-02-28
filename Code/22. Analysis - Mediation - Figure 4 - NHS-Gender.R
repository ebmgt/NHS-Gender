# This file is available at https://github.com/ebmgt/NHS-Gender
# Authors: Robert Badgett; rbadgett@kumc.edu
# Permission: GNU GPLv3 https://choosealicense.com/licenses/gpl-3.0/
# Last edited 2026-01-12

# Startup -----
## Set working directory -----
if (Sys.getenv("RSTUDIO") != "1"){
  args <- commandArgs(trailingOnly = FALSE)
  script_path <- sub("--file=", "", args[grep("--file=", args)])  
  script_path <- dirname(script_path)
  setwd(script_path)
}else{
  setwd(dirname(rstudioapi::getSourceEditorContext()$path))
  #ScriptsDir <- paste(getwd(),'/Scripts',sep='')
}
getwd()

## Include common code -----
source("00. common code - NHS-Gender - 2026-02-22.R")

# Packages ------
library(betareg)
library(statmod)

# Functions -----
function_squeeze01 <- function(x, eps = 1e-6) {
  x <- as.numeric(x)
  x <- pmin(pmax(x, eps), 1 - eps)
  x
}

function_p2star <- function(p) {
  ifelse(is.na(p), "",
         ifelse(p < .001, "***",
                ifelse(p < .01,  "**",
                       ifelse(p < .05,  "*", ""))))
}

##  Functions, other ------
function_extract_coef <- function(fit, term) {
  if (is.null(fit)) return(c(est = NA_real_, p = NA_real_))
  co <- summary(fit)$coefficients$mean
  if (!(term %in% rownames(co))) return(c(est = NA_real_, p = NA_real_))
  c(est = unname(co[term, "Estimate"]),
    p   = unname(co[term, "Pr(>|z|)"]))
}


# _____________________________-----
# Parameters -----
#* Select Data sources ----
data.source <- "All"
#* Select Outcome sources ----
Outcome.source <- 'Burned_out_rate'
Outcome.source.label <- gsub("_", " ", Outcome.source)
#* Select Religion.reference ----
Religion.reference <- 'No religion'

outcome <- "Burned_out_rate"
x_focal <- "Female_rate"

# _____________________________-----
# Data grab -----

#wb.temp <- loadWorkbook("../Data/regdat - 2025-06-16.xlsx")
wb.temp <- loadWorkbook("../Data/data_trusts_well-being - organisational results.xlsx")
regdat <- read.xlsx (wb.temp, sheet = 1, startRow = 1, colNames = TRUE, na.strings = "NA", detectDates = TRUE)
regdat$Religion_None_Withheld <- rowSums(regdat[, c("Religion_Decline_answer_rate", "Religion_None_rate")], na.rm = TRUE) / 100

nrow(regdat)

## Ensure rates are 0 to 1. New 2025-03-07 -----
cat(green$bold("\nChecking for rates between 0 and 1\n\n"))
rate_cols <- grep("rate", names(regdat), ignore.case = TRUE, value = TRUE)
rate_maxes <- sapply(regdat[, rate_cols, drop = FALSE], max, na.rm = TRUE)

cols_to_adjust <- names(rate_maxes[rate_maxes > 1.5])

if (length(cols_to_adjust) > 0) {
  cat(crayon::red$bold("Converting these rate columns from percent to proportion:\n"))
  cat(paste(cols_to_adjust, collapse = ", "), "\n\n")
  
  before <- sapply(regdat[, cols_to_adjust, drop = FALSE], max, na.rm = TRUE)
  regdat[, cols_to_adjust] <- regdat[, cols_to_adjust] / 100
  after  <- sapply(regdat[, cols_to_adjust, drop = FALSE], max, na.rm = TRUE)
  
  cat(crayon::green$bold("Max before -> after (first 10):\n"))
  print(head(cbind(before, after), 10))
} else {
  cat(crayon::green$bold("All rate columns already look like proportions (max <= 1.5).\n"))
}

## vars_needed -----
# Used until 2026-02-18
suppressors <- c("Emotional_manager_rate",
                 "Emotional_public_rate",
                 "Emotional_colleagues_rate",
                 "Time_demands_meet_rate"
)
vars_needed <- c(outcome, x_focal, suppressors)

# Used starting 2026-02-18
vars_x <- grep("rate", names(regdat), value = TRUE)
vars_x <- vars_x[!grepl("Burned_out|Outcome|Stress|response|sometimes|ever", vars_x)]
vars_x <- vars_x[!grepl("Female", vars_x)]
# Remove psychological intermediate states
vars_x <- vars_x[!grepl("exhausting|Frustration|Autonomy|Membership|Mastery", vars_x)]

vars_needed <- c(outcome, x_focal, vars_x)

# _________________________________________________________-----
# MacKinnon diagrams -----
# PURPOSE:
#   1) Multipanel MacKinnon-style diagrams (linear beta-regression paths; easy to interpret)
#   2) Separate spline plots to visualize Female_rate nonlinearity vs Burned_out_rate
#
# NOTES (recommended):
# - MacKinnon diagrams here are "linear-path" decompositions using beta regression coefficients.
# - Spline plots are shown separately so the diagram remains readable and faithful to MacKinnon.
# - This is structural decomposition of associations (not causal mediation) unless design supports causality.
#
# IMPORTANT (required):
# - This section is self-contained. Do NOT paste older versions of these functions below.
# - There must be ONLY ONE definition each of:
#     get_mackinnon_coeffs(), pretty_label_NHS(), function_coef_style(), function_fmt_p(),
#     arrow_box_to_box(), function_draw_mackinnon_panel()
#   Otherwise, later duplicates will overwrite earlier functions.

# 1) Setup data ------
set.seed(20260112)

outcome <- "Burned_out_rate"
x_focal <- "Female_rate"

vars_x <- grep("rate", names(regdat), value = TRUE)
vars_x <- vars_x[!grepl("Burned_out|Outcome|Stress|response|sometimes|ever", vars_x)]
vars_x <- vars_x[!grepl("Female", vars_x)]
vars_x <- vars_x[!grepl("Religion_None_rate", vars_x)]
# Remove psychological intermediate states
vars_x <- vars_x[!grepl("exhausting|Frustration|Autonomy|Membership|Mastery", vars_x)]

suppressors <- vars_x

vars_needed <- c(outcome, x_focal, vars_x)

dat <- na.omit(regdat[, vars_needed])
dat <- as.data.frame(dat)

stopifnot(all(vars_needed %in% names(dat)))
for (v in vars_needed) stopifnot(is.numeric(dat[[v]]))

cat("\n--- Complete-case N ---\n")
cat("N =", nrow(dat), "\n\n")

# 1a) Ensure (0,1) for beta regression (recommended) ------
# NOTE: shrink transform applied only if needed.
shrink01 <- function(y) {
  n <- length(y)
  (y * (n - 1) + 0.5) / n
}
for (v in c(outcome, vars_needed)) {
  if (any(dat[[v]] <= 0 | dat[[v]] >= 1)) {
    cat("NOTE:", v, "has 0/1 values; applying shrink transform for betareg.\n")
    dat[[v]] <- shrink01(dat[[v]])
  }
}

# 2) MacKinnon related functions------

function_get_mackinnon_coeffs <- function(data,
                                          mediator,
                                          independent,
                                          outcome,
                                          covars = NULL) {
  
  # Allow unquoted names or strings
  m_var <- if (is.character(mediator)) mediator else deparse(substitute(mediator))
  x_var <- if (is.character(independent)) independent else deparse(substitute(independent))
  y_var <- if (is.character(outcome)) outcome else deparse(substitute(outcome))
  
  # Build RHS terms
  rhs_xc  <- c(x_var, covars)
  rhs_xc  <- rhs_xc[!is.na(rhs_xc) & nzchar(rhs_xc)]
  rhs_xcS <- paste(rhs_xc, collapse = " + ")
  
  rhs_xmc  <- c(x_var, m_var, covars)
  rhs_xmc  <- rhs_xmc[!is.na(rhs_xmc) & nzchar(rhs_xmc)]
  rhs_xmcS <- paste(rhs_xmc, collapse = " + ")
  
  # Formulas (MacKinnon notation)
  f_tau <- as.formula(paste(y_var, "~", rhs_xcS))    # total effect:   Y ~ X (+C)
  f_m   <- as.formula(paste(m_var, "~", rhs_xcS))    # a-path:         M ~ X (+C)
  f_y   <- as.formula(paste(y_var, "~", rhs_xmcS))   # b + direct:     Y ~ X + M (+C)
  
  # Fits (beta regression; mean submodel coefficients)
  fit_tau <- betareg::betareg(f_tau, data = data)
  fit_m   <- betareg::betareg(f_m,   data = data)
  fit_y   <- betareg::betareg(f_y,   data = data)
  
  sm_tau <- summary(fit_tau)$coefficients$mean
  sm_m   <- summary(fit_m)$coefficients$mean
  sm_y   <- summary(fit_y)$coefficients$mean
  
  coefs <- list(
    # roles (store explicitly so we can validate later)
    x_var = x_var,
    y_var = y_var,
    m_med = m_var,
    
    # MacKinnon paths (logit-mean scale by default)
    tau_est   = unname(sm_tau[x_var, "Estimate"]),
    tau_p     = unname(sm_tau[x_var, "Pr(>|z|)"]),
    
    alpha_est = unname(sm_m[x_var, "Estimate"]),
    alpha_p   = unname(sm_m[x_var, "Pr(>|z|)"]),
    
    beta_est  = unname(sm_y[m_var, "Estimate"]),
    beta_p    = unname(sm_y[m_var, "Pr(>|z|)"]),
    
    taup_est  = unname(sm_y[x_var, "Estimate"]),
    taup_p    = unname(sm_y[x_var, "Pr(>|z|)"]),
    
    # strength metrics (on the same coefficient scale)
    delta_tau     = unname(sm_y[x_var, "Estimate"]) - unname(sm_tau[x_var, "Estimate"]),  # τ′ − τ
    abs_delta_tau = abs(unname(sm_y[x_var, "Estimate"]) - unname(sm_tau[x_var, "Estimate"]))
  )
  
  return(coefs)
}


function_draw_mackinnon <- function(mediator,
                                    independent = Female_rate,
                                    outcome     = Burned_out_rate,
                                    data        = dat,
                                    covars      = NULL,
                                    coefs       = NULL,
                                    warn_delay  = 0.6) {
  
  # Resolve names (unquoted or strings)
  m_var <- if (is.character(mediator)) mediator else deparse(substitute(mediator))
  x_var <- if (is.character(independent)) independent else deparse(substitute(independent))
  y_var <- if (is.character(outcome)) outcome else deparse(substitute(outcome))
  
  # Validate provided coefs (if any)
  roles_match <- FALSE
  if (!is.null(coefs) && is.list(coefs)) {
    has_roles <- all(c("x_var", "y_var", "m_med") %in% names(coefs))
    if (has_roles) {
      roles_match <- identical(coefs$x_var, x_var) &&
        identical(coefs$y_var, y_var) &&
        identical(coefs$m_med, m_var)
    }
  }
  
  # If roles don't match or coefs missing, recompute
  recomputed <- FALSE
  if (is.null(coefs) || !roles_match) {
    if (!is.null(coefs) && !roles_match) recomputed <- TRUE
    coefs <- function_get_mackinnon_coeffs(
      data        = data,
      mediator    = m_var,
      independent = x_var,
      outcome     = y_var,
      covars      = covars
    )
  }
  
  # Draw using your existing renderer (uses pretty_label_NHS internally)
  function_draw_mackinnon_panel(
    coefs   = coefs,
    x_label = x_var,
    y_label = y_var
  )
  
  # If we had to recompute due to role mismatch, write a warning onto panel
  if (recomputed) {
    usr <- par("usr")
    # top-left-ish overlay
    text(x = usr[1] + 0.2, y = usr[4] - 0.3,
         labels = "Note: roles changed; coefficients recomputed",
         adj = c(0, 1), cex = 0.85, col = "red", font = 2)
    Sys.sleep(warn_delay)
  }
  
  invisible(coefs)
}
# COSMETIC GOALS (implemented):
#  - Smaller boxes; bottom boxes farther apart; center-to-center arrows; larger fonts.
#  - Remove underscores from node/mediator labels.
#  - Alpha/beta labels placed near arrow midpoints; sign-colored per your rule.
function_draw_mackinnon_panel <- function(coefs,
                                          x_label = "Female_rate",
                                          y_label = "Burned_out_rate") {
  
  plot.new()
  plot.window(xlim = c(0, 10), ylim = c(0, 10))
  
  # For debugging and diagnosis
  # box(col = "red", lwd = 2)
  
  y_master <- 1
  
  # ---- Layout (recentered for wider boxes)
  # -1 below pushes the plot down
  xX <- 1.6; yX <- 4.4 - y_master
  xY <- 8.4; yY <- 4.4 - y_master
  xM <- 5.0; yM <- 7.9 - y_master
  
  # Box geometry (smaller)
  bw <- 1.55
  bh <- 0.62
  
  # Line weights
  box_lwd <- 1.6
  arr_lwd <- 1.6
  
  # Boxes
  rect(xX-bw, yX-bh, xX+bw, yX+bh, border = "black", lwd = box_lwd)
  rect(xM-bw, yM-bh, xM+bw, yM+bh + 1, border = "black", lwd = box_lwd)
  rect(xY-bw, yY-bh, xY+bw, yY+bh, border = "black", lwd = box_lwd)
  
  # Node labels
  cex_node <- 1.25
  text(xX, yX, pretty_label_NHS(x_label), cex = cex_node)
  text(xY, yY, pretty_label_NHS(y_label), cex = cex_node)
  
  # Mediator labels
  m_raw   <- coefs$m_med
  message("DEBUG m_med = ", paste(coefs$m_med, collapse = " | "))
  m_title <- pretty_label_NHS(m_raw)
  m_box <- pretty_label_NHS(vapply(m_raw, transform_label, character(1)))
  
  cex_m <- 1.10
  yM_center <- (((yM - bh) + (yM + bh + 1)) / 2)  # yM + 0.5
  text(xM, yM_center, m_box, cex = cex_m)
  
  # Arrows
  arrow_box_to_box(xX, yX, xM, yM, bw, bh, length = 0.10, lwd = arr_lwd)
  arrow_box_to_box(xM, yM, xY, yY, bw, bh, length = 0.10, lwd = arr_lwd)
  arrow_box_to_box(xX, yX, xY, yY, bw, bh, length = 0.10, lwd = arr_lwd)
  
  # Midpoints for label placement
  mid_XM_x <- (xX + xM) / 2
  mid_XM_y <- (yX + yM) / 2
  mid_MY_x <- (xM + xY) / 2
  mid_MY_y <- (yM + yY) / 2
  
  # Title -----
  # This needs recoding to work with y_master
  mtext(m_title, line = - 2, cex = 1.15, font = 2)
  
  # Alpha and beta labels
  dx_char <- strwidth("X", cex = 1.05, units = "user")
  
  alpha_style <- function_coef_style(coefs$alpha_est)
  alpha_txt <- paste0("alpha = ", signif(coefs$alpha_est, 3),
                      "\n(p = ", function_fmt_p(coefs$alpha_p), ")")
  text(x = mid_XM_x - 0.8 - dx_char, y = mid_XM_y + 0.5,
       labels = alpha_txt, cex = 1.05,
       col = alpha_style$col, font = alpha_style$font)
  
  beta_style <- function_coef_style(coefs$beta_est)
  beta_txt <- paste0("beta = ", signif(coefs$beta_est, 3),
                     "\n(p = ", function_fmt_p(coefs$beta_p), ")")
  text(x = mid_MY_x + 0.8 + dx_char, y = mid_MY_y + 0.5,
       labels = beta_txt, cex = 1.05,
       col = beta_style$col, font = beta_style$font)
  
  # Direction of change (NOTE: this is now correct)
  coefs$direction <- ifelse(
    abs(coefs$taup_est) > abs(coefs$tau_est), "↑",
    ifelse(abs(coefs$taup_est) < abs(coefs$tau_est), "↓", "—")
  )
  
  # Bottom path label (tau and tau') -----
  text(5.0, 4.0 - y_master,
       paste0("c' = ", signif(coefs$taup_est, 3), " ",
              "(p c ' = ", function_fmt_p(coefs$tau_p), ")")
      ,cex = 0.98, font = 2,
  )
  
  # Bottom block: LEFT-ALIGNED (for all panels) to the midpoint of the Female box -----
  bottom_txt <- paste0(
    "Direct effect: c' = ", signif(coefs$taup_est, 3),
    " (p = ", function_fmt_p(coefs$taup_p), ")\n",
    "Total effect:  c  = ", signif(coefs$tau_est, 3),
    " (p = ", function_fmt_p(coefs$tau_p), ")\n",
    "Magnitude change: ", coefs$direction
  )
  
  # Place and left-align at Female-box -----
  # If want align at box's midpoint
  #text(x = xX, y = 2.45, labels = bottom_txt, cex = 0.98, adj = 0)
  # If want alignment at 1/4 of box width from the left edge
  x_bottom <- xX - 0.5 * bw
  text(x = x_bottom, y = 3 - y_master, labels = bottom_txt, cex = 0.98, adj = c(0, 1))
  
  show_interpretation <- grepl(
    "Harassment",
    m_title,
    ignore.case = TRUE
  )
  
  # Interpretation -----
  if (isTRUE(show_interpretation)) {
    
    interpretation_raw <- "Higher female representation is associated with lower reported harassment."
    interpretation_body <- paste(strwrap(interpretation_raw, width = 50), collapse = "\n")
    
    x_int <- xM                          # current
    # x_int <- ((xX + bw) + (xY - bw)) / 2    # corridor midpoint (recommended)
    # x_int <- xM - 0.35                    # slight left shift option
    
    y_int <- 3 - y_master # Same Y as the bottom block
    
    # Calculate a gap based on 1 line of text
    line_gap <- strheight("A", cex = 0.98, units = "user") * 1.5
    
    # 2. Bold header: Top-aligned to match the other block
    text(
      x = x_int,
      y = y_int,
      labels = "Possible interpretation:",
      cex = 0.98,
      font = 2,
      adj = c(0, 1) # Top-aligned
    )
    
    # 3. Body: Subtract the gap to move it DOWN
    text(
      x = x_int,
      y = y_int - line_gap, # MINUS moves it down
      labels = interpretation_body,
      cex = 0.98,
      font = 1,
      adj = c(0, 1) # Top-aligned
    )
  }
}


# 3a) transform_label for intervenor text inside the M box ONLY
transform_label <- function(label) {
  
  message("--- LAUNCHING: ", label)
  
  # 1. Check Exclusions (Total Exit)
  if (label %in% c("Female_rate", "Burned_out_rate")) return(label)
  
  # 2. Stage 1: Substitutions
  if (grepl("emotional", label, ignore.case = TRUE)) {
    label <- gsub("emotional", "Harassment_by", label, ignore.case = TRUE)
  }
  if (grepl("discrimination", label, ignore.case = TRUE)) {
    label <- gsub("discrimination", "discrimination_by", label, ignore.case = TRUE)
  }
  
  # 3. Stage 2: Line Breaks (UPDATING variable, NOT returning)
  if (grepl("_by", label, fixed = TRUE)) {
    label <- gsub("_by", "_by\n", label, fixed = TRUE)
  } else if (grepl("_rate", label, fixed = TRUE)) {
    label <- gsub("_rate", "_\nrate", label, fixed = TRUE)
  }
  
  # 4. Stage 3: Append the suffix (Now everyone who didn't exit at Step 1 gets this)
  label <- paste0("Intervening variable:\n", label)
  
  # 5. Final Reporting & Return
  message("   >> FINAL RESULT: ", gsub("\n", " [NL] ", label))
  return(label)
}

# 3b) Arrow from box-center to box-center, touching box edges (requested)
arrow_box_to_box <- function(x1, y1, x2, y2, bw, bh, length = 0.10, lwd = 1.6) {
  dx <- x2 - x1
  dy <- y2 - y1
  if (dx == 0 && dy == 0) return(invisible(NULL))
  
  tx1 <- if (dx != 0) bw / abs(dx) else Inf
  ty1 <- if (dy != 0) bh / abs(dy) else Inf
  t1  <- min(tx1, ty1)
  
  tx2 <- if (dx != 0) bw / abs(dx) else Inf
  ty2 <- if (dy != 0) bh / abs(dy) else Inf
  t2  <- min(tx2, ty2)
  
  xs <- x1 + dx * t1
  ys <- y1 + dy * t1
  xe <- x2 - dx * t2
  ye <- y2 - dy * t2
  
  arrows(xs, ys, xe, ye, length = length, lwd = lwd)
}

# RESTART HERE -----
# 5) Build coefficients and plot multipanel MacKinnon ------
suppressors_4 <- c("Emotional_colleagues_rate", "Time_demands_meet_rate",
                   "Patient_facing_rate", "Older_rate")

mack_list <- lapply(suppressors_4, function(m) {
  function_get_mackinnon_coeffs(
    data        = dat,
    mediator    = m,
    independent = x_focal,
    outcome     = outcome,
    covars      = NULL
  )
})

names(mack_list) <- suppressors_4

# 6) Plot -----
op <- par(no.readonly = TRUE)
on.exit(par(op), add = TRUE)

# Give room for top title row and bottom notes row
plot.new()
par(
  mfrow = c(2, 2),
  # mar  = c(1.0, 0.5, 2.2, 0.5),   # inner margins for each panel
  mar  = c(0.2, 0.5, 1.2, 0.5), 
  # OUTER margins: bottom, left, top, right
  #oma  = c(3.0, 0.5, 2.2, 0.5) 
  oma  = c(3.8, 0.5, 3.2, 0.5) 
)

## * panels -----
for (m in suppressors_4) {
  function_draw_mackinnon_panel(mack_list[[m]], x_label = x_focal, y_label = outcome)
  }

##* outer margins (oma) -----
# Ensure outer margins (oma) are large enough to hold the text
# c(bottom, left, top, right) - increase bottom and top as needed
par(oma = c(6, 1, 5, 1)) 

# --- Top row title (outer, left-aligned) ---
raw_text <- "Figure 4. Mediation pathway diagrams for four example covariates identified through random forest screening (Table 2) and asssessed in the pathway analysis (Table 3)."

# 2. Wrap the combined text
# Adjust width based on your device size
wrapped_caption <- str_wrap(raw_text, width = 170) 

# 3. Use a single mtext call
# We use font = 1 (plain). If you want the whole thing bold, use font = 2.
mtext(
  wrapped_caption,
  side = 3, 
  outer = TRUE, 
  line = 2,     # Adjusted to bring it closer to the plot
  adj = 0,      # Left aligned
  cex = 1.0, 
  font = 1      # Plain text ensures no weird vertical jumping
)

# --- Bottom row notes (outer, left-aligned) ---
# Use increasing line numbers to stack from top to bottom
mtext(
  expression(bold("Notes:")),
  side = 1, outer = TRUE, line = 1, adj = 0
)
mtext(
  "Each figure is a MacKinnon mediation diagram using beta-regression models (logit link), estimated on the logit-mean scale.",
  side = 1, outer = TRUE, line = 2, adj = 0
)
mtext(
  #"NOT SURE I HAVE THE COLORS INTUITIVE AND CORRECT. TRICKY AS HIGHER TIME, LOWER BURNOUT",
  "",
  side = 1, outer = TRUE, line = 3, adj = 0, cex = 0.8
)

mtext(
  paste0("rbadgett@kumc.edu (best at 1200 px width) ", Sys.Date()),
  side = 1, outer = TRUE, line = 4, adj = 1, cex = 0.8
)

# Reset layout (optional: keep par settings if plotting more)
par(mfrow = c(1, 1), oma = c(0,0,0,0))

#______________________________________ -----

