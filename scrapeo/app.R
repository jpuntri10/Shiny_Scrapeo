
library(shiny)
library(RSelenium)
library(stringr)
library(writexl)
library(readxl)
library(DT)
library(xml2)
library(rvest)

selenium_ok <- function(host = "localhost", port = 4444) {
  ok <- FALSE
  try({
    drv <- remoteDriver(remoteServerAddr = host, port = port, browserName = "chrome")
    ok <- TRUE
  }, silent = TRUE)
  ok
}

limpiar_valor <- function(x) {
  if (is.na(x) || length(x) == 0) return(NA_character_)
  trimws(sub("^.*?:", "", x))
}

find_ruc_input <- function(remDr) {
  campo <- try(remDr$findElement(using = "xpath",
                                 value = "//input[@type='text' and contains(@placeholder,'Ingrese RUC')]"),
               silent = TRUE)
  if (inherits(campo, "try-error")) {
    campo <- try(remDr$findElement(using = "css selector", value = "input.form-control"), silent = TRUE)
  }
  if (inherits(campo, "try-error")) {
    campo <- remDr$findElement(using = "xpath", value = "//input[@type='text' and not(@disabled)]")
  }
  campo
}

find_buscar_button <- function(remDr) {
  btn <- try(remDr$findElement(using = "xpath",
                               value = "//button[contains(normalize-space(.),'Buscar')] | //input[@type='submit' or contains(@value,'Buscar')]"),
             silent = TRUE)
  if (inherits(btn, "try-error")) {
    btns <- remDr$findElements(using = "css selector",
                               value = "input[type='submit'], button[type='submit'], input[value*='Buscar'], button")
    if (length(btns) > 0) btn <- btns[[1]] else stop("No se encontró el botón 'Buscar'.")
  }
  btn
}

wait_for_nav <- function(remDr, oldUrl, timeout_secs = 12) {
  t0 <- Sys.time()
  repeat {
    newUrl <- tryCatch(remDr$getCurrentUrl()[[1]], error = function(e) oldUrl)
    if (!identical(newUrl, oldUrl)) break
    if (as.numeric(difftime(Sys.time(), t0, units = "secs")) > timeout_secs) break
    Sys.sleep(0.4)
  }
}

parse_items <- function(remDr) {
  items <- remDr$findElements(using = "css selector", value = ".list-group-item")
  if (length(items) == 0) return(NULL)
  
  get_text <- function(el, css) {
    x <- try(el$findElement(using = "css selector", value = css), silent = TRUE)
    if (inherits(x, "try-error")) return("")
    t <- try(x$getElementText()[[1]], silent = TRUE)
    if (inherits(t, "try-error")) return("") else return(trimws(t))
  }
  
  labels <- sapply(items, function(it) {
    txt <- get_text(it, ".list-group-item-heading")
    if (txt == "") txt <- get_text(it, ".col-sm-5")
    txt
  })
  values <- sapply(items, function(it) {
    txt <- get_text(it, ".list-group-item-text")
    if (txt == "") txt <- get_text(it, ".col-sm-7")
    txt
  })
  data.frame(label = labels, value = values, stringsAsFactors = FALSE)
}

extract_first <- function(lbl_pattern, pairs_df) {
  hit <- which(grepl(lbl_pattern, pairs_df$label, ignore.case = TRUE))
  if (length(hit)) pairs_df$value[hit[1]] else NA_character_
}

click_representantes_btn <- function(remDr) {
  btn_rep <- try(remDr$findElement(using = "xpath",
                                   value = "//button[contains(@class,'btnInfRepLeg') or contains(.,'Representante')]"),
                 silent = TRUE)
  if (!inherits(btn_rep, "try-error")) {
    try(btn_rep$clickElement(), silent = TRUE)
    Sys.sleep(1.2)
  } else {
    try(remDr$executeScript("var f = document.forms['formRepLeg']; if (f) { f.submit(); }"), silent = TRUE)
    Sys.sleep(1.2)
  }
}

consulta_representantes <- function(remDr) {
  click_representantes_btn(remDr)
  page_src <- try(remDr$getPageSource()[[1]], silent = TRUE)
  if (inherits(page_src, "try-error") || is.null(page_src)) return(NULL)
  doc <- try(read_html(page_src), silent = TRUE)
  if (inherits(doc, "try-error")) return(NULL)
  
  tablas <- html_elements(doc, "table")
  if (length(tablas) == 0) return(NULL)
  
  target_tbl <- NULL
  for (t in tablas) {
    headers <- html_text(html_elements(t, "thead th"))
    if (length(headers) == 0) next
    has_cols <- all(c("Documento", "Nro. Documento", "Nombre", "Cargo", "Fecha Desde") %in% trimws(headers))
    if (has_cols) { target_tbl <- t; break }
  }
  if (is.null(target_tbl)) target_tbl <- html_element(doc, "div.table-responsive table.table")
  if (is.null(target_tbl)) return(NULL)
  
  rows <- html_elements(target_tbl, "tbody tr")
  if (length(rows) == 0) return(NULL)
  
  datos <- lapply(rows, function(r) {
    vals <- html_elements(r, "td") |> html_text()
    vals <- trimws(vals)
    length(vals) <- 5
    vals
  })
  
  mat <- do.call(rbind, datos)
  df  <- as.data.frame(mat, stringsAsFactors = FALSE)
  colnames(df) <- c("Documento", "Nro_Documento", "Nombre", "Cargo", "Fecha_Desde")
  df
}

consulta_un_ruc <- function(remDr, ruc) {
  remDr$navigate("https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/FrameCriterioBusquedaWeb.jsp")
  campo <- find_ruc_input(remDr)
  campo$clearElement()
  campo$sendKeysToElement(list(ruc))
  btn <- find_buscar_button(remDr)
  oldUrl <- remDr$getCurrentUrl()[[1]]
  try(btn$clickElement(), silent = TRUE)
  wait_for_nav(remDr, oldUrl, timeout_secs = 12)
  
  pairs_df <- parse_items(remDr)
  
  # Extraer fecha explícitamente por XPath
  fecha_inicio_val <- NA_character_
  fecha_node <- try(remDr$findElement(using = "xpath",
                                      value = "//h4[contains(.,'Fecha de Inicio de Actividades')]/following::p[1]"),
                    silent = TRUE)
  if (!inherits(fecha_node, "try-error")) {
    fecha_inicio_val <- fecha_node$getElementText()[[1]]
  }
  
  if (is.null(pairs_df) || nrow(pairs_df) == 0) {
    df <- data.frame(
      Numero_RUC               = ruc,
      Razon_Social             = NA_character_,
      Tipo_Contribuyente       = NA_character_,
      Tipo_Documento           = NA_character_,
      Estado                   = NA_character_,
      Condicion                = NA_character_,
      Actividad                = NA_character_,
      Fecha_Inicio_Actividades = fecha_inicio_val,
      Anios_Creacion           = NA_real_,
      stringsAsFactors         = FALSE
    )
    return(list(df = df, representantes = NULL))
  }
  
  numero_ruc_text <- extract_first("^\\s*Número\\s+de\\s+RUC", pairs_df)
  razon_social <- NA_character_
  if (!is.na(numero_ruc_text)) {
    partes <- strsplit(numero_ruc_text, "-")[[1]]
    if (length(partes) > 1) razon_social <- trimws(partes[2])
  }
  numero_ruc <- {
    m <- regexpr("\\b\\d{11}\\b", ifelse(is.na(numero_ruc_text), "", numero_ruc_text))
    if (m[1] > 0) substr(numero_ruc_text, m[1], m[1] + attr(m, "match.length") - 1) else ruc
  }
  
  tipo_contribuyente <- limpiar_valor(extract_first("^\\s*Tipo\\s+Contribuyente", pairs_df))
  tipo_documento     <- limpiar_valor(extract_first("^\\s*Tipo\\s*de\\s*Documento", pairs_df))
  estado             <- limpiar_valor(extract_first("^\\s*Estado\\s+del\\s+Contribuyente", pairs_df))
  condicion          <- limpiar_valor(extract_first("^\\s*Condición\\s+del\\s+Contribuyente", pairs_df))
  actividad          <- limpiar_valor(extract_first("^\\s*Actividad", pairs_df))
  
  fecha_inicio <- fecha_inicio_val
  fecha_inicio_date <- try(as.Date(fecha_inicio, format = "%d/%m/%Y"), silent = TRUE)
  if (inherits(fecha_inicio_date, "try-error")) fecha_inicio_date <- NA
  anios_creacion <- if (!is.na(fecha_inicio_date)) {
    as.numeric(difftime(Sys.Date(), fecha_inicio_date, units = "days")) %/% 365
  } else NA_real_
  
  df <- data.frame(
    Numero_RUC               = numero_ruc,
    Razon_Social             = razon_social,
    Tipo_Contribuyente       = tipo_contribuyente,
    Tipo_Documento           = tipo_documento,
    Estado                   = estado,
    Condicion                = condicion,
    Actividad                = actividad,
    Fecha_Inicio_Actividades = fecha_inicio,
    Anios_Creacion           = anios_creacion,
    stringsAsFactors         = FALSE
  )
  
  rep_df <- consulta_representantes(remDr)
  list(df = df, representantes = rep_df)
}

# -------------------- UI --------------------
ui <- fluidPage(
  titlePanel("Consulta RUC - SUNAT"),
  tabsetPanel(
    tabPanel("Por RUC",
             sidebarLayout(
               sidebarPanel(
                 textInput("ruc", "Ingrese RUC (11 dígitos):", value = "10424410981"),
                 actionButton("consultar", "Consultar"),
                 br(), br(),
                 downloadButton("descargar", "Descargar Excel")
               ),
               mainPanel(DTOutput("tabla"), br(), tags$h4("Representantes Legales"), DTOutput("tabla_rep"))
             )),
    tabPanel("Por Excel (lote)",
             sidebarLayout(
               sidebarPanel(
                 fileInput("archivo", "Suba archivo (.xlsx o .csv) con columna RUC",
                           accept = c(".xlsx", ".xls", ".csv")),
                 checkboxInput("dedup", "Eliminar RUCs duplicados", TRUE),
                 numericInput("nmax", "Máximo a procesar (0 = todos)", value = 0, min = 0),
                 actionButton("procesar", "Procesar lote"),
                 br(), br(),
                 downloadButton("descargar_lote", "Descargar Excel consolidado")
               ),
               mainPanel(DTOutput("tabla_lote"), br(), tags$h4("Representantes (Lote)"), DTOutput("tabla_reps_lote"))
             ))
  )
)

# -------------------- Server --------------------
server <- function(input, output, session) {
  resultado_ind <- reactiveVal(NULL)
  representantes_ind <- reactiveVal(NULL)
  
  observeEvent(input$consultar, {
    req(input$ruc)
    if (!selenium_ok()) {
      showModal(modalDialog("Selenium no disponible", easyClose = TRUE))
      return(invisible())
    }
    withProgress(message = "Consultando en SUNAT…", value = 0, {
      remDr <- NULL
      on.exit({ if (!is.null(remDr)) try(remDr$close(), silent = TRUE) }, add = TRUE)
      remDr <- remoteDriver(remoteServerAddr = "localhost", port = 4444, browserName = "chrome",
                            extraCapabilities = list(chromeOptions = list(args = c("--disable-gpu","--window-size=1280,900"))))
      remDr$open()
      remDr$setTimeout(type = "implicit", milliseconds = 7000)
      res <- consulta_un_ruc(remDr, input$ruc)
      resultado_ind(res$df)
      representantes_ind(res$representantes)
    })
  })
  
  output$tabla <- renderDT({ req(resultado_ind()); datatable(resultado_ind(), rownames = FALSE) })
  output$tabla_rep <- renderDT({
    reps <- representantes_ind()
    if (is.null(reps) || nrow(reps) == 0) datatable(data.frame(Mensaje = "Sin representantes"), rownames = FALSE)
    else datatable(reps, rownames = FALSE)
  })
  
  output$descargar <- downloadHandler(
    filename = function() sprintf("resultado_ruc_%s.xlsx", input$ruc),
    content = function(file) {
      wb <- list(Consulta = resultado_ind())
      reps <- representantes_ind()
      if (!is.null(reps) && nrow(reps) > 0) wb$Representantes <- reps
      write_xlsx(wb, file)
    }
  )
  
  resultado_lote <- reactiveVal(NULL)
  reps_lote <- reactiveVal(NULL)
  
  observeEvent(input$procesar, {
    req(input$archivo)
    if (!selenium_ok()) {
      showModal(modalDialog("Selenium no disponible", easyClose = TRUE))
      return(invisible())
    }
    ext <- tolower(tools::file_ext(input$archivo$name))
    df_in <- if (ext %in% c("xlsx", "xls")) read_excel(input$archivo$datapath) else read.csv(input$archivo$datapath)
    if (!("RUC" %in% names(df_in))) {
      showModal(modalDialog("El archivo debe contener la columna 'RUC'.", easyClose = TRUE))
      return(invisible())
    }
    rucs <- as.character(df_in$RUC)
    rucs <- sapply(rucs, function(x) {
      m <- regexpr("\\d{11}", gsub("\\D", "", as.character(x)))
      if (m[1] > 0) substr(gsub("\\D", "", as.character(x)), m[1], m[1] + 10) else NA_character_
    }, USE.NAMES = FALSE)
    rucs <- rucs[!is.na(rucs)]
    if (input$dedup) rucs <- unique(rucs)
    if (input$nmax > 0) rucs <- head(rucs, input$nmax)
    
    withProgress(message = "Procesando lote…", value = 0, {
      remDr <- NULL
      on.exit({ if (!is.null(remDr)) try(remDr$close(), silent = TRUE) }, add = TRUE)
      remDr <- remoteDriver(remoteServerAddr = "localhost", port = 4444, browserName = "chrome",
                            extraCapabilities = list(chromeOptions = list(args = c("--disable-gpu","--window-size=1280,900"))))
      remDr$open()
      remDr$setTimeout(type = "implicit", milliseconds = 7000)
      
      out_rows <- list()
      reps_list <- list()
      for (i in seq_along(rucs)) {
        incProgress(i / length(rucs), detail = paste("Consultando", rucs[i]))
        Sys.sleep(0.8)
        res <- consulta_un_ruc(remDr, rucs[i])
        out_rows[[i]] <- res$df
        if (!is.null(res$representantes) && nrow(res$representantes) > 0) {
          reps_tmp <- res$representantes
          reps_tmp <- cbind(RUC = res$df$Numero_RUC[1], reps_tmp)
          reps_list[[length(reps_list) + 1]] <- reps_tmp
        }
      }
      resultado_lote(do.call(rbind, out_rows))
      reps_lote(if (length(reps_list) > 0) do.call(rbind, reps_list) else NULL)
    })
  })
  
  output$tabla_lote <- renderDT({ req(resultado_lote()); datatable(resultado_lote(), rownames = FALSE) })
  output$tabla_reps_lote <- renderDT({
    reps <- reps_lote()
    if (is.null(reps) || nrow(reps) == 0) datatable(data.frame(Mensaje = "Sin representantes"), rownames = FALSE)
    else datatable(reps, rownames = FALSE)
  })
  
  output$descargar_lote <- downloadHandler(
    filename = function() paste0("resultado_rucs_", format(Sys.time(), "%Y%m%d_%H%M"), ".xlsx"),
    content = function(file) {
      wb <- list(Consulta = resultado_lote())
      reps <- reps_lote()
      if (!is.null(reps) && nrow(reps) > 0) wb$Representantes <- reps
      write_xlsx(wb, file)
    }
  )
}

shinyApp(ui, server)


