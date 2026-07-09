#===============================================================================
# Generic Export Function
# Exports any number of tables to both Excel and Word
#===============================================================================
pacman::p_load(openxlsx, officer, flextable, tibble)
export_results <- function(tables,
                           excel_file = "Results.xlsx",
                           word_file  = "Results.docx",
                           doc_title  = "Analysis Results") {
  #-----------------------------------------------------------------------------
  # Internal function
  # Converts row names into first column if necessary
  #-----------------------------------------------------------------------------
  prepare_table <- function(x){
    x  <- as.data.frame(x)
    rn <- rownames(x)
    if(!is.null(rn)){
      if(any(rn != as.character(seq_len(nrow(x))))){
        x <- tibble::rownames_to_column(x,
                                        var = "Variable")
      }
    }
    rownames(x) <- NULL
    x
  }
  #-----------------------------------------------------------------------------
  # Excel
  #-----------------------------------------------------------------------------
  wb <- createWorkbook()
  #-----------------------------------------------------------------------------
  # Word
  #-----------------------------------------------------------------------------
  doc <- read_docx()
  doc <- body_add_par(
    doc,
    doc_title,
    style = "heading 1"
  )
  #-----------------------------------------------------------------------------
  # Loop over all tables
  #-----------------------------------------------------------------------------
  for(i in seq_along(tables)){
    sheet_name <- names(tables)[i]
    tab <- prepare_table(tables[[i]])
    ## Excel
    addWorksheet(wb, sheet_name)
    writeData(
      wb,
      sheet = sheet_name,
      x     = tab
    )
    ## Word
    doc <- body_add_par(
      doc,
      sheet_name,
      style = "heading 2"
    )
    doc <- body_add_flextable(
      doc,
      flextable(tab)
    )
    doc <- body_add_par(doc, "")
  }
  #-----------------------------------------------------------------------------
  # Save
  #-----------------------------------------------------------------------------
  saveWorkbook(
    wb,
    excel_file,
    overwrite = TRUE
  )
  print(
    doc,
    target = word_file
  )
  cat("\n=========================================\n")
  cat("Export completed successfully.\n")
  cat("Excel :", excel_file,"\n")
  cat("Word  :", word_file,"\n")
  cat("=========================================\n")
}
#===============================================================================
# Examples:
#-------------------------------------------------------------------------------
# Obtain model outputs
#-------------------------------------------------------------------------------
#ssm_table <- my.table(ssm_trt)
#me_table  <- me.ssm(ssm_trt)
#-------------------------------------------------------------------------------
# Export everything
#-------------------------------------------------------------------------------
#export_results(
#  tables = list(
#    "Descriptive Statistics"     = desc_stats,
#    "Variance Inflation Factors" = vif_table,
#    "Sample Selection Model"     = ssm_table,
#    "Marginal Effects"           = me_table
#  ),
#  excel_file = "Econometric_Results.xlsx",
#  word_file  = "Econometric_Results.docx",
#  doc_title  = "Econometric Analysis Results"
#)
#-------------------------------------------------------------------------------
# Another example:
#export_results(
#  tables = list(
#    "Summary"            = summary_table,
#    "Correlation Matrix" = cor_table,
#    "ANOVA"              = anova_table,
#    "Regression 1"       = reg1,
#    "Regression 2"       = reg2,
#    "Marginal Effects"   = me_table,
#    "Forecast Accuracy"  = forecast_table
#  ),
#  excel_file = "Project2.xlsx",
#  word_file  = "Project2.docx"
#)
#===============================================================================