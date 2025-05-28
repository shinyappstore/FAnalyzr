library(shiny)
library(shinydashboard)
library(haven)
library(readxl)
library(foreign)
library(psych)
library(GPArotation)
library(ggplot2)
library(shinythemes)
library(DT)
library(readr)
library(semTools)
library(plotly)
library(officer)
library(flextable)
library(lavaan)
library(semPlot)
library(dplyr)

ui <- dashboardPage(
  dashboardHeader(title = "FAnalyzr"),
  dashboardSidebar(
    sidebarMenu(
      menuItem("Exploratory Factor Analysis", tabName = "efa", icon = icon("search")),
      menuItem("Confirmatory Factor Analysis", tabName = "cfa", icon = icon("check-circle")),
      menuItem("Regression Table", tabName = "regression", icon = icon("table"))
    )
  ),
  dashboardBody(
    # Custom CSS for white background and black text in all output areas
    tags$head(
      tags$style(HTML("
        /* Main background */
        .content-wrapper, .right-side {
          background-color: white !important;
        }
        
        /* All output areas */
        .shiny-output, 
        pre, 
        .dataTables_wrapper, 
        .table, 
        .verbatim-text-output,
        .dt-container {
          color: #000000 !important;
          background-color: white !important;
        }
        
        /* Tab boxes */
        .nav-tabs-custom, 
        .nav-tabs-custom > .tab-content {
          background-color: white !important;
        }
        
        /* Text output */
        .shiny-text-output {
          color: #000000 !important;
          background-color: white !important;
        }
        
        /* DataTables */
        table.dataTable {
          color: #000000 !important;
        }
        
        /* Plot areas */
        .shiny-plot-output {
          background-color: white !important;
        }
      "))
    ),
    
    tabItems(
      # EFA Tab
      tabItem(tabName = "efa",
              fluidPage(
                theme = shinytheme("flatly"),
                titlePanel("Exploratory Factor Analysis"),
                sidebarLayout(
                  sidebarPanel(
                    fileInput("file_efa", "Upload Dataset (CSV, Excel, SAV, DTA)", 
                              accept = c(".csv", ".xlsx", ".xls", ".sav", ".dta")),
                    uiOutput("varSelect_efa"),
                    numericInput("nfactors", "Number of Factors to Extract:", value = 2, min = 1),
                    selectInput("fm", "Extraction Method:", 
                                choices = c("minres", "ml", "pa", "wls", "gls", "uls", "principal")),
                    selectInput("rotate", "Rotation Method:", 
                                choices = c("none", "varimax", "promax", "oblimin", "simplimax", "quartimin", "geominQ")),
                    selectInput("reliability", "Select Reliability Method:", 
                                choices = c("Cronbach Alpha", "McDonald Omega")),
                    actionButton("analyze", "Run EFA", class = "btn btn-primary"),
                    downloadButton("downloadAllEFA", "Download All EFA Results (Word)")
                  ),
                  mainPanel(
                    tabsetPanel(
                      tabPanel("Scree", plotlyOutput("screePlot"), br(), downloadButton("downloadScree", "Download Scree Plot")),
                      tabPanel("FA Parallel", 
                               verbatimTextOutput("parallelInterpretation"),
                               plotOutput("faParallelPlot"), 
                               br(), 
                               downloadButton("downloadParallel", "Download Parallel Plot")
                      ),
                      tabPanel("KMO & Bartlett's Test", verbatimTextOutput("kmoBartlett")),
                      tabPanel("Variance Explained", verbatimTextOutput("varianceTable")),
                      tabPanel("EFA Results", 
                               verbatimTextOutput("efaResults"), 
                               br(),
                               downloadButton("downloadEFAWord", "Download EFA Results (Word)")
                      ),
                      tabPanel("Reliability", verbatimTextOutput("reliabilityResult"), downloadButton("downloadDiagram", "Download Factor Diagram")),
                      tabPanel("Correlation Matrix", 
                               DTOutput("corMatrix"),
                               tags$div(
                                 style = "margin-top: 10px; font-style: italic; font-size: 14px; color: #000000;",
                                 "*p < 0.05, **p < 0.01, ***p < 0.001"
                               )
                      )
                    )
                  )
                )
              )
      ),
      
      # CFA Tab
      tabItem(tabName = "cfa",
              fluidPage(
                theme = shinytheme("flatly"),
                titlePanel("Confirmatory Factor Analysis"),
                sidebarLayout(
                  sidebarPanel(
                    fileInput("file_cfa", "Upload Dataset (CSV, Excel, SAV, DTA)", 
                              accept = c(".csv", ".xlsx", ".xls", ".sav", ".dta")),
                    uiOutput("varSelect_cfa"),
                    textAreaInput("modelText", "Model (lavaan syntax)", height = "250px",
                                  placeholder = "e.g.,\n ind60 =~ x1 + x2 + x3\n dem60 =~ y1 + y2 + y3\n dem60 ~ ind60"),
                    selectInput("estimator", "Select Estimator:",
                                choices = c("ML" = "ML", "MLR" = "MLR", "WLSMV" = "WLSMV"),
                                selected = "MLR"),
                    actionButton("runModel", "Run CFA", class = "btn btn-success"),
                    downloadButton("downloadAllCFA", "Download All CFA Results (Word)")
                  ),
                  mainPanel(
                    tabsetPanel(
                      tabPanel("Model Summary", verbatimTextOutput("modelSummary")),
                      tabPanel("Standardized Estimates", verbatimTextOutput("standardized")),
                      tabPanel("Unstandardized Estimates", verbatimTextOutput("unstandardized")),
                      tabPanel("Modification Indices", DTOutput("modIndices")),
                      tabPanel("Composite Reliability & Convergent Validity", verbatimTextOutput("reliabilityResults")),
                      tabPanel("SEM Plot", 
                               downloadButton("downloadPlot", "Download Plot (PNG)"),
                               plotOutput("semPlot", height = "600px")
                      ),
                      tabPanel("Standardized Residual Covariances", 
                               verbatimTextOutput("residCov")
                      ),
                      tabPanel("Discriminant Validity", 
                               tableOutput("discriminantValidity")
                      )
                    )
                  )
                )
              )
      ),
      
      # Regression Table Tab
      tabItem(tabName = "regression",
              fluidPage(
                theme = shinytheme("flatly"),
                titlePanel("Regression Table"),
                sidebarLayout(
                  sidebarPanel(
                    fileInput("file_reg", "Upload Dataset (CSV, Excel, SAV, DTA)", 
                              accept = c(".csv", ".xlsx", ".xls", ".sav", ".dta")),
                    uiOutput("varSelect_reg"),
                    textAreaInput("regModelText", "Regression Model (lavaan syntax)", height = "150px",
                                  placeholder = "e.g.,\n y ~ x1 + x2 + x3"),
                    selectInput("regEstimator", "Select Estimator:",
                                choices = c("ML" = "ML", "MLR" = "MLR", "WLSMV" = "WLSMV"),
                                selected = "MLR"),
                    actionButton("runRegModel", "Run Regression", class = "btn btn-info"),
                    downloadButton("downloadAllReg", "Download All Regression Results (Word)")
                  ),
                  mainPanel(
                    DTOutput("regressionTable")
                  )
                )
              )
      )
    ),
    tags$footer(
      div(
        style = "text-align: center; padding: 10px; font-size: 14px; color: #555555;",
        "For suggestions or assistance with funding a Heroku plan, please contact me:",
        br(),
        tags$a(href = "mailto:mudassiribrahim30@gmail.com", "mudassiribrahim30@gmail.com")
      )
    )
  )
)

server <- function(input, output, session) {
  # Shared data reactive
  data_shared <- reactive({
    # Check which file input was used
    if (!is.null(input$file_efa)) {
      file <- input$file_efa
    } else if (!is.null(input$file_cfa)) {
      file <- input$file_cfa
    } else if (!is.null(input$file_reg)) {
      file <- input$file_reg
    } else {
      return(NULL)
    }
    
    req(file)
    ext <- tools::file_ext(file$name)
    switch(ext,
           csv = read_csv(file$datapath),
           xlsx = read_excel(file$datapath),
           xls = read_excel(file$datapath),
           sav = read_sav(file$datapath),
           dta = read_dta(file$datapath),
           validate("Unsupported file type")
    )
  })
  
  # Consistent variable selection UI for all tabs
  output$varSelect_efa <- renderUI({
    req(data_shared())
    selectInput("vars_efa", "Select Variables:", choices = names(data_shared()), multiple = TRUE)
  })
  
  output$varSelect_cfa <- renderUI({
    req(data_shared())
    selectInput("vars_cfa", "Select Variables:", choices = names(data_shared()), multiple = TRUE)
  })
  
  output$varSelect_reg <- renderUI({
    req(data_shared())
    selectInput("vars_reg", "Select Variables:", choices = names(data_shared()), multiple = TRUE)
  })
  
  # Selected data for each tab
  selectedData_efa <- reactive({
    req(input$vars_efa)
    data_shared()[, input$vars_efa, drop = FALSE]
  })
  
  selectedData_cfa <- reactive({
    req(input$vars_cfa)
    data_shared()[, input$vars_cfa, drop = FALSE]
  })
  
  selectedData_reg <- reactive({
    req(input$vars_reg)
    data_shared()[, input$vars_reg, drop = FALSE]
  })
  
  # EFA Server Logic
  efa_result <- reactiveVal(NULL)
  parallel_result <- reactiveVal(NULL)
  
  observeEvent(input$analyze, {
    req(selectedData_efa())
    kmo <- KMO(selectedData_efa())
    bartlett <- cortest.bartlett(selectedData_efa())
    
    output$kmoBartlett <- renderPrint({
      cat("Kaiser-Meyer-Olkin (KMO) Test: ", round(kmo$MSA, 3), "\n")
      cat("Bartlett's Test of Sphericity: ", round(bartlett$p.value, 3), "\n")
    })
    
    # Correlation matrix with asterisks
    cor_test <- corr.test(selectedData_efa())
    r <- round(cor_test$r, 3)
    p <- cor_test$p
    
    stars <- matrix("", nrow = nrow(p), ncol = ncol(p))
    stars[p < 0.001] <- "***"
    stars[p < 0.01 & p >= 0.001] <- "**"
    stars[p < 0.05 & p >= 0.01] <- "*"
    
    r_formatted <- matrix("", nrow = nrow(r), ncol = ncol(r))
    for (i in 1:nrow(r)) {
      for (j in 1:ncol(r)) {
        r_formatted[i, j] <- paste0(formatC(r[i, j], format = "f", digits = 3), stars[i, j])
      }
    }
    
    rownames(r_formatted) <- rownames(r)
    colnames(r_formatted) <- colnames(r)
    
    output$corMatrix <- renderDT({
      datatable(as.data.frame(r_formatted), options = list(pageLength = 10, scrollX = TRUE))
    })
    
    # Scree Plot
    output$screePlot <- renderPlotly({
      ev <- eigen(cor(selectedData_efa(), use = "pairwise.complete.obs"))
      scree_data <- data.frame(Factor = 1:length(ev$values), Eigenvalue = ev$values)
      plot_ly(scree_data, x = ~Factor, y = ~Eigenvalue, type = 'scatter', mode = 'lines+markers',
              marker = list(size = 8)) %>%
        layout(title = "Scree Plot", xaxis = list(title = "Factor"), yaxis = list(title = "Eigenvalue"))
    })
    
    output$downloadScree <- downloadHandler(
      filename = function() { "scree_plot.png" },
      content = function(file) {
        ev <- eigen(cor(selectedData_efa(), use = "pairwise.complete.obs"))
        scree_data <- data.frame(Factor = 1:length(ev$values), Eigenvalue = ev$values)
        p <- ggplot(scree_data, aes(x = Factor, y = Eigenvalue)) +
          geom_point() + geom_line() +
          labs(title = "Scree Plot", x = "Factor", y = "Eigenvalue") +
          theme_minimal() + theme(plot.margin = margin(20, 20, 20, 20))
        ggsave(file, plot = p, width = 8, height = 6)
      }
    )
    
    # FA Parallel with interpretation
    output$faParallelPlot <- renderPlot({
      parallel <- fa.parallel(selectedData_efa(), fa = "fa", n.iter = 20, show.legend = TRUE)
      parallel_result(parallel)
    })
    
    output$parallelInterpretation <- renderPrint({
      req(parallel_result())
      parallel <- parallel_result()
      cat("Parallel Analysis Results:\n")
      cat("Suggested number of factors based on eigenvalues:", parallel$nfact, "\n")
      cat("Suggested number of components based on eigenvalues:", parallel$ncomp, "\n\n")
      cat("Interpretation:\n")
      cat("The parallel analysis suggests extracting", parallel$nfact, 
          "factors when the actual data eigenvalues are greater than the corresponding", 
          "percentiles of the random data eigenvalues.\n")
      cat("This is a more robust method than the traditional 'eigenvalue > 1' rule.\n")
    })
    
    output$downloadParallel <- downloadHandler(
      filename = function() { "fa_parallel_plot.png" },
      content = function(file) {
        png(file, width = 1000, height = 800)
        parallel <- fa.parallel(selectedData_efa(), fa = "fa", n.iter = 20, show.legend = TRUE)
        dev.off()
      }
    )
    
    efa <- fa(selectedData_efa(), nfactors = input$nfactors, rotate = input$rotate, fm = input$fm)
    efa_result(efa)
    
    output$efaResults <- renderPrint({
      print(efa, digits = 3)
    })
    
    output$varianceTable <- renderPrint({
      cat("Total Variance Explained:\n")
      print(round(efa$Vaccounted, 3))
    })
    
    output$reliabilityResult <- renderPrint({
      if (input$reliability == "Cronbach Alpha") {
        print(psych::alpha(selectedData_efa()), digits = 3)
      } else {
        print(omega(selectedData_efa(), nfactors = input$nfactors), digits = 3)
      }
    })
    
    output$downloadDiagram <- downloadHandler(
      filename = function() { "factor_diagram.png" },
      content = function(file) {
        png(file, width = 1000, height = 800)
        fa.diagram(efa_result())
        dev.off()
      }
    )
  })
  
  # Download all EFA results in Word
  output$downloadAllEFA <- downloadHandler(
    filename = function() { "all_efa_results.docx" },
    content = function(file) {
      req(efa_result())
      doc <- read_docx()
      
      # Add title
      doc <- body_add_par(doc, "Exploratory Factor Analysis Results", style = "heading 1")
      
      # Add analysis details
      doc <- body_add_par(doc, paste("Extraction Method:", input$fm), style = "Normal")
      doc <- body_add_par(doc, paste("Rotation Method:", input$rotate), style = "Normal")
      doc <- body_add_par(doc, paste("Number of Factors:", input$nfactors), style = "Normal")
      
      # Add KMO and Bartlett's test
      kmo <- KMO(selectedData_efa())
      bartlett <- cortest.bartlett(selectedData_efa())
      doc <- body_add_par(doc, "KMO and Bartlett's Test", style = "heading 2")
      doc <- body_add_par(doc, paste("KMO Measure of Sampling Adequacy:", round(kmo$MSA, 3)), style = "Normal")
      doc <- body_add_par(doc, paste("Bartlett's Test p-value:", round(bartlett$p.value, 3)), style = "Normal")
      
      # Add parallel analysis interpretation
      if (!is.null(parallel_result())) {
        parallel <- parallel_result()
        doc <- body_add_par(doc, "Parallel Analysis Results", style = "heading 2")
        doc <- body_add_par(doc, paste("Suggested number of factors:", parallel$nfact), style = "Normal")
      }
      
      # Add factor loadings
      efa <- efa_result()
      doc <- body_add_par(doc, "Factor Loadings", style = "heading 2")
      ft <- flextable(round(efa$loadings[1:nrow(efa$loadings), ], 3))
      doc <- body_add_flextable(doc, ft)
      
      # Add variance explained
      doc <- body_add_par(doc, "Variance Explained", style = "heading 2")
      variance_table <- as.data.frame(round(efa$Vaccounted, 3))
      ft_variance <- flextable(variance_table)
      doc <- body_add_flextable(doc, ft_variance)
      
      # Add reliability
      doc <- body_add_par(doc, "Reliability Analysis", style = "heading 2")
      if (input$reliability == "Cronbach Alpha") {
        alpha_result <- psych::alpha(selectedData_efa())
        doc <- body_add_par(doc, paste("Cronbach's Alpha:", round(alpha_result$total$raw_alpha, 3)), style = "Normal")
      } else {
        omega_result <- omega(selectedData_efa(), nfactors = input$nfactors)
        doc <- body_add_par(doc, paste("McDonald's Omega Total:", round(omega_result$omega.tot, 3)), style = "Normal")
        doc <- body_add_par(doc, paste("McDonald's Omega Hierarchical:", round(omega_result$omega_h, 3)), style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  
  # CFA Server Logic
  modelResults <- eventReactive(input$runModel, {
    req(input$modelText, selectedData_cfa(), input$estimator)
    sem(model = input$modelText, data = selectedData_cfa(), estimator = input$estimator, std.lv = TRUE)
  })
  
  output$modelSummary <- renderPrint({
    req(modelResults())
    cat("Model Summary:\n")
    print(summary(modelResults(), fit.measures = TRUE))
    cat("\nFit Measures:\n")
    print(fitMeasures(modelResults()))
  })
  
  output$standardized <- renderPrint({
    req(modelResults())
    parameterEstimates(modelResults(), standardized = TRUE)[, c("lhs", "op", "rhs", "est", "std.all")]
  })
  
  output$unstandardized <- renderPrint({
    req(modelResults())
    parameterEstimates(modelResults(), standardized = FALSE)[, c("lhs", "op", "rhs", "est")]
  })
  
  output$modIndices <- renderDT({
    req(modelResults())
    mod_ind <- modindices(modelResults()) %>%
      mutate_if(is.numeric, ~ round(., 4))
    datatable(mod_ind, options = list(pageLength = 10, scrollX = TRUE))
  })
  
  output$reliabilityResults <- renderPrint({
    req(modelResults())
    
    reliability_vals <- tryCatch({
      semTools::reliability(modelResults())
    }, error = function(e) e)
    
    ave_vals <- tryCatch({
      semTools::AVE(modelResults())
    }, error = function(e) e)
    
    htmt_vals <- tryCatch({
      semTools::htmt(modelResults())
    }, error = function(e) e)
    
    omega_result <- tryCatch({
      psych::omega(selectedData_cfa(), warnings = FALSE)
    }, error = function(e) e)
    
    cat("Composite Reliability & Convergent Validity Measures:\n\n")
    
    cat("Composite Reliability (CR):\n")
    print(reliability_vals)
    
    cat("\nAverage Variance Extracted (AVE):\n")
    print(ave_vals)
    
    cat("\nHeterotrait-Monotrait Ratio (HTMT):\n")
    print(htmt_vals)
    
    cat("\nOmega Coefficients (Total and Hierarchical):\n")
    print(omega_result)
  })
  
  # Standardized Residual Covariances
  output$residCov <- renderPrint({
    req(modelResults())
    resid_cov <- resid(modelResults(), type = "standardized")$cov
    resid_cov <- round(resid_cov, 3)
    print(resid_cov)
  })
  
  # Discriminant Validity Calculation
  output$discriminantValidity <- renderTable({
    req(modelResults())
    
    # Get the standardized solution
    std_solution <- standardizedSolution(modelResults())
    
    # Extract factor correlations
    factor_cors <- std_solution %>%
      filter(op == "~~" & lhs != rhs & lhs %in% unique(std_solution$lhs[std_solution$op == "=~"])) %>%
      select(lhs, rhs, est.std) %>%
      rename(Factor1 = lhs, Factor2 = rhs, Correlation = est.std)
    
    # Calculate AVE for each factor
    loadings <- std_solution %>%
      filter(op == "=~") %>%
      group_by(lhs) %>%
      summarise(AVE = mean(est.std^2)) %>%
      rename(Factor = lhs)
    
    # Create discriminant validity table
    disc_validity <- factor_cors %>%
      left_join(loadings, by = c("Factor1" = "Factor")) %>%
      rename(AVE1 = AVE) %>%
      left_join(loadings, by = c("Factor2" = "Factor")) %>%
      rename(AVE2 = AVE) %>%
      mutate(Sqrt_AVE1 = sqrt(AVE1),
             Sqrt_AVE2 = sqrt(AVE2),
             Discriminant_Valid = ifelse(abs(Correlation) < pmin(Sqrt_AVE1, Sqrt_AVE2), "Yes", "No")) %>%
      select(Factor1, Factor2, Correlation, AVE1, AVE2, Discriminant_Valid)
    
    disc_validity <- disc_validity %>%
      mutate(across(where(is.numeric), ~ round(., 3)))
    
    disc_validity
  }, rownames = FALSE)
  
  # Download all CFA results in Word
  output$downloadAllCFA <- downloadHandler(
    filename = function() { "all_cfa_results.docx" },
    content = function(file) {
      req(modelResults())
      doc <- read_docx()
      
      # Add title
      doc <- body_add_par(doc, "Confirmatory Factor Analysis Results", style = "heading 1")
      
      # Add model specification
      doc <- body_add_par(doc, "Model Specification", style = "heading 2")
      doc <- body_add_par(doc, input$modelText, style = "Normal")
      
      # Add fit measures
      fit <- fitMeasures(modelResults())
      doc <- body_add_par(doc, "Model Fit Measures", style = "heading 2")
      fit_table <- data.frame(
        Measure = c("Chi-square", "df", "p-value", "CFI", "TLI", "RMSEA", "SRMR"),
        Value = round(c(fit["chisq"], fit["df"], fit["pvalue"], 
                        fit["cfi"], fit["tli"], fit["rmsea"], fit["srmr"]), 3)
      )
      ft_fit <- flextable(fit_table)
      doc <- body_add_flextable(doc, ft_fit)
      
      # Add standardized estimates
      std_est <- parameterEstimates(modelResults(), standardized = TRUE)[, c("lhs", "op", "rhs", "est", "std.all")]
      doc <- body_add_par(doc, "Standardized Estimates", style = "heading 2")
      ft_std <- flextable(std_est)
      doc <- body_add_flextable(doc, ft_std)
      
      # Add reliability and validity
      doc <- body_add_par(doc, "Reliability and Validity", style = "heading 2")
      
      if (!inherits(try(semTools::reliability(modelResults())), "try-error")) {
        rel <- semTools::reliability(modelResults())
        ave <- semTools::AVE(modelResults())
        
        rel_table <- data.frame(
          Factor = names(rel["omega",]),
          Omega = round(as.numeric(rel["omega",]), 3),
          AVE = round(as.numeric(ave), 3)
        )
        ft_rel <- flextable(rel_table)
        doc <- body_add_flextable(doc, ft_rel)
      }
      
      # Add discriminant validity
      disc_validity <- output$discriminantValidity()
      if (!is.null(disc_validity)) {
        doc <- body_add_par(doc, "Discriminant Validity", style = "heading 2")
        ft_disc <- flextable(disc_validity)
        doc <- body_add_flextable(doc, ft_disc)
      }
      
      print(doc, target = file)
    }
  )
  
  output$semPlot <- renderPlot({
    req(modelResults())
    semPaths(modelResults(), what = "std", layout = "tree", edge.label.cex = 1.2, 
             sizeMan = 6, sizeLat = 8, nCharNodes = 0, edge.color = "black", 
             mar = c(5, 5, 5, 5), fade = FALSE, rotation = 2, label.cex = 1.2, 
             edge.width = 2, style = "ram", nodeWidth = 3, edge.arrow.size = 0.5)
  })
  
  output$downloadPlot <- downloadHandler(
    filename = function() {
      paste("sem_plot", ".png", sep = "")
    },
    content = function(file) {
      png(file, width = 1200, height = 900)
      semPaths(modelResults(), what = "std", layout = "tree", edge.label.cex = 1.2, 
               sizeMan = 6, sizeLat = 8, nCharNodes = 0, edge.color = "black", 
               mar = c(5, 5, 5, 5), fade = FALSE, rotation = 2, label.cex = 1.2, 
               edge.width = 2, style = "ram", nodeWidth = 3, edge.arrow.size = 0.5)
      dev.off()
    }
  )
  
  # Regression Table Server Logic
  regResults <- eventReactive(input$runRegModel, {
    req(input$regModelText, selectedData_reg(), input$regEstimator)
    sem(model = input$regModelText, data = selectedData_reg(), estimator = input$regEstimator)
  })
  
  output$regressionTable <- renderDT({
    req(regResults())
    params <- parameterEstimates(regResults(), standardized = TRUE)
    effects_table <- params[params$op == "~", c("lhs", "op", "rhs", "est", "se", "z", "pvalue")]
    colnames(effects_table) <- c("Dependent", "Operator", "Predictor", "Estimate", "Std. Error", "Z-value", "P-value")
    
    effects_table$Estimate <- round(effects_table$Estimate, 4)
    effects_table$`Std. Error` <- round(effects_table$`Std. Error`, 4)
    effects_table$`Z-value` <- round(effects_table$`Z-value`, 4)
    effects_table$`P-value` <- round(effects_table$`P-value`, 4)
    
    datatable(effects_table, options = list(pageLength = 15, scrollX = TRUE))
  })
  
  # Download all Regression results in Word
  output$downloadAllReg <- downloadHandler(
    filename = function() { "all_regression_results.docx" },
    content = function(file) {
      req(regResults())
      doc <- read_docx()
      
      # Add title
      doc <- body_add_par(doc, "Regression Analysis Results", style = "heading 1")
      
      # Add model specification
      doc <- body_add_par(doc, "Model Specification", style = "heading 2")
      doc <- body_add_par(doc, input$regModelText, style = "Normal")
      
      # Add results table
      params <- parameterEstimates(regResults(), standardized = TRUE)
      effects_table <- params[params$op == "~", c("lhs", "op", "rhs", "est", "se", "z", "pvalue", "std.all")]
      colnames(effects_table) <- c("Dependent", "Operator", "Predictor", "Estimate", "Std. Error", "Z-value", "P-value", "Std. Estimate")
      
      effects_table <- effects_table %>%
        mutate(across(where(is.numeric), ~ round(., 3)))
      
      doc <- body_add_par(doc, "Regression Coefficients", style = "heading 2")
      ft <- flextable(effects_table)
      doc <- body_add_flextable(doc, ft)
      
      # Add model fit if available
      if (!is.null(fitMeasures(regResults()))) {
        fit <- fitMeasures(regResults())
        doc <- body_add_par(doc, "Model Fit Measures", style = "heading 2")
        fit_table <- data.frame(
          Measure = c("Chi-square", "df", "p-value", "CFI", "TLI", "RMSEA", "SRMR"),
          Value = round(c(fit["chisq"], fit["df"], fit["pvalue"], 
                          fit["cfi"], fit["tli"], fit["rmsea"], fit["srmr"]), 3)
        )
        ft_fit <- flextable(fit_table)
        doc <- body_add_flextable(doc, ft_fit)
      }
      
      print(doc, target = file)
    }
  )
}

shinyApp(ui, server)
