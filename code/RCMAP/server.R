#runApp('RCMAPdir',display.mode="no", launch.browser=FALSE, port=2197)

require(shiny)
library("smacof")
library("readxl")
library("xtable")
library("DT")
library("rtf")
library("tools")
source("helper.R")

# change this if the size of the input file is too large for shiny to handle
options(shiny.maxRequestSize=100*1024^2)

eps <<- matrix(c(0,0.02,-0.02,0,0,-0.02,0.02,0),ncol=2,byrow = T)
eps <<- rbind(eps, eps%*%matrix(c(1,1,-1,1),ncol=2)*sqrt(2)/2)

reportChoices <- c("Sorters", "Statements", "Clusters")
analysisChoices <- c("ANOVA", "Tukey")
settingsChoices <- c("Number of Clusters", "Cluster Names", "Save Project", 
                     "Save Report")
homedir <<- Sys.getenv("HOME")

plotChoices <<- c()

# return character vector of choices for a given panel
getChoices <<- function(panel) {
  switch(panel,
         "Plots" = plotChoices,
         "Reports" = reportChoices,
         "Analysis" = analysisChoices,
         "Settings" = settingsChoices)
}

shinyServer(function(input, output, session) {
  
  observe({
    if (is.null(input$reallyQuit))
      return(NULL)
    if (input$reallyQuit) stopApp() # stop shiny
  })
  
  ###################################################################
  # show the SIDE-BAR: if no file has been selected, show the menu to
  # choose an input file. Otherwise, show the full sidebar for an
  # open project. The side-bar content depends on which tab is
  # selected (Summary, Plots, etc.)
  ###################################################################
  output$sidebar <- renderUI({
    if (is.null(input$inputFile)) {
      sidebar <- list("inputFile", label="Select a concept mapping file", 
                      accept=c(".xls",".xlsx",".RData"))
      do.call(fileInput, sidebar)
    } else {
      switch (input$conditionedPanels,
              'Summary' = {
                sidebar <- list(tags$a(href = "", "Load a different data set"))
              },
              'Plots' = {
                sidebar <- list("plotType",  "Plot Type", plotChoices)
                if (is.null(input$plotType)) {
                  tagList(do.call(radioButtons, sidebar))
                } else {
                  sidebar <- list("plotType", "Plot type", plotChoices,
                                  selected = input$plotType)
                  slct <- switch(input$plotType,
                                 "Point Map (MDS)" = uiOutput("chooseUnivariateRating"),
                                 "Clusters" = tagList(uiOutput("chooseClusterPlot"),
                                                      uiOutput("chooseUnivariateRating")),
                                 "Univariate Rating - by Statement" = tagList(uiOutput("chooseUnivariateRating"),
                                                                              uiOutput("demographicSelection")),
                                 "Univariate Rating - by Cluster" = tagList(uiOutput("chooseUnivariateRating"),
                                                                            uiOutput("demographicSelection")),
                                 "Bar Chart" = tagList(uiOutput("chooseBarChartRatings"),
                                                       uiOutput("demographicSelection")),
                                 "Pattern Matching" = tagList(uiOutput("chooseBivariateRating1"),
                                                              uiOutput("chooseBivariateRating2"),
                                                              uiOutput("demographicSelection")),
                                 "GoZone" = tagList(uiOutput("chooseClusterCheckbox"),
                                                    actionButton("selectAllClusters","Select All"),
                                                    actionButton("unselectAllClusters","Unselect All"),
                                                    uiOutput("chooseBivariateRating1"),
                                                    uiOutput("chooseBivariateRating2"),
                                                    uiOutput("demographicSelection")),
                                 "Dendrogram" = ""
                  )                  
                  tagList(do.call(radioButtons,sidebar),hr(),slct,hr(),
                          uiOutput("includeInReport"))
                  
                }
              },
              'Reports' = {
                sidebar = list("reportType",  "Report Type", reportChoices)
                if (is.null(input$reportType)) {
                  tagList(do.call(radioButtons,sidebar))
                } else {
                  sidebar = list("reportType", "Report type", reportChoices,selected=input$reportType)
                  slct = switch(input$reportType,
                                "Sorters" = uiOutput("demographicSelection"),
                                "Statements" = tagList(uiOutput("chooseUnivariateRating"),
                                                       uiOutput("demographicSelection"),
                                                       uiOutput("sortBy"),
                                                       uiOutput("byCluster")),
                                "Clusters" = tagList(uiOutput("chooseUnivariateRating"),
                                                     uiOutput("demographicSelection"))
                  )                  
                  tagList(do.call(radioButtons,sidebar),hr(),slct,hr(),
                          uiOutput("includeInReport"))
                }
              },
              'Analysis' = {
                sidebar = list("analysisType",  "Test", analysisChoices)
                if (is.null(input$analysisType)) {
                  tagList(do.call(radioButtons,sidebar))
                } else {
                  sidebar = list("analysisType", "Test", analysisChoices,selected=input$analysisType)
                  slct = switch(input$analysisType,
                                "ANOVA" = tagList(uiOutput("chooseUnivariateRating"),
                                                  uiOutput("demographicSelection")),
                                "Tukey" = tagList(uiOutput("chooseUnivariateRating"),
                                                  uiOutput("demographicSelection")))                  
                  tagList(do.call(radioButtons,sidebar),hr(),slct,hr(),
                          uiOutput("includeInReport"))
                }
              },
              'Users' = {
                sidebar = list("Select Users",p(),
                               uiOutput("demographicSelection"),
                               actionButton("resetGroup","Reset"),
                               textInput("newGroupName","Enter name of new group"),
                               actionButton("createNewGroup","Create Group"),
                               actionButton("removeGroup","Remove Group")
                )
              },
              'Settings' = {
                sidebar = list("setting",  "Setting", settingsChoices)
                if (is.null(input$setting)) {
                  sidebar = tagList(do.call(radioButtons,sidebar))
                } else {
                  sidebar = list("setting", "Setting", settingsChoices,selected=input$setting)
                  slct = switch(input$setting,
                                "Number of Clusters" = uiOutput("chooseNoCluster"),
                                "Cluster Names" = list(uiOutput("changeClusterNames"),
                                                       uiOutput("getClusterName"),
                                                       actionButton("changeName","Change"),
                                                       uiOutput("clusterNameError")),
                                "Save Project" = list(uiOutput("projectDescription"),
                                                      uiOutput("projectName"),
                                                      actionButton("saveProject","Save") #,downloadButton("downloadProject","Download Project")
                                ),
                                "Save Report" = list(uiOutput("reportName"),
                                                     checkboxInput("editablePlots", "Editable plots (Windows only)", value=FALSE),
                                                     downloadButton("downloadReport","Download Report"))
                  )
                  sidebar = list(do.call(radioButtons,sidebar),hr(),slct)
                }
                
              }
      )
    }
  })   
  
  output$showDoc = renderUI({
    if (is.null(input$file1)) {
      p(helpText(a("About R-CMap", href="Documentation.html", target="_blank")))
    }
  })
  
  
  #######################################################
  # rendering functions
  
  ######
  ## functions for rendering the side-bar
  
  # Choose rating variables for statistical comparisons
  output$chooseUnivariateRating <- renderUI({
    vars <- ratingVariables()
    if (is.null(input$univariateRating)){
      dflt <- vars[1]
    } else {
      dflt <- input$univariateRating
    }
    selectInput("univariateRating", "Rating variable", vars,selected=dflt)
  })
  
  output$chooseBivariateRating1 <- renderUI({
    vars <- ratingVariables()
    if (is.null(input$bivariateRating1)){
      dflt <- vars[1]
    } else {
      dflt <- input$bivariateRating1
    }
    menuLabel <- "Rating variable"
    if (input$conditionedPanels == "Plots") {
      if (input$plotType == "GoZone") {
        menuLabel <- paste(menuLabel,"(horizontal axis)")
      } else if (grepl("Pattern Matching", input$plotType)) {
        menuLabel <- paste(menuLabel,"(left-hand side)")
      }
    }
    selectInput("bivariateRating1", menuLabel, vars,selected=dflt)
  })
  
  output$chooseBivariateRating2 <- renderUI({
    vars <- ratingVariables()
    if (is.null(input$bivariateRating2)){
      dflt <- vars[2]
    } else {
      dflt <- input$bivariateRating2
    }
    menuLabel <- "Rating variable"
    if (input$conditionedPanels == "Plots") {
      if (input$plotType == "GoZone") {
        menuLabel <- paste(menuLabel,"(vertical axis)")
      } else if (grepl("Pattern Matching", input$plotType)) {
        menuLabel <- paste(menuLabel,"(right-hand side)")
      }
    }
    selectInput("bivariateRating2", menuLabel, vars,selected=dflt)
  })
  
  # select subgroups (which were created based on demographic criteria)
  output$demographicSelection <- renderUI({
    cmap <- ReadSavedData()
    val <- input$demographicSelection
    selectInput("demographicSelection", "Group",
                choices=names(cmap$demographicGroups), selected=val)    
  })
  
  ## plotting options - selected in the side-bar
  
  # choose the type of a cluster plot (rays/polygons)
  output$chooseClusterPlot <- renderUI({
    if (is.null(input$clusterplot)){
      dflt <- "Rays"
    } else {
      dflt <- input$clusterplot
    }
    selectInput("clusterplot", "Type", c("Rays","Polygons"), selected=dflt)
  })
  
  
  # choose up to four rating variables to display in a bar-chart
  output$chooseBarChartRatings <- renderUI({
    ratingVars <- ratingVariables()
    lapply(1:min(length(ratingVars),4),function(i) {
      inputName <- paste0("chooseRating",i)
      currentSelect <- input[[inputName]]
      if (is.null(currentSelect))
        dflt <- switch(i,ratingVars[1],"None")
      else
        dflt <- currentSelect
      selectInput(inputName,paste("Choose rating variable",i),
                  choices=c("None",ratingVars),selected=dflt)
    })
  })
  
  output$quit = renderUI({
    if (input$Quit) {
      checkboxInput("reallyQuit", "Confirm quit")
    }
  })
  
  # checkboxes for selecting which clusters to plot in the gozone
  output$chooseClusterCheckbox <- renderUI({
    input$showGoZone
    input$selectAllClusters
    clNames <- as.character(getClusters())
    n.clust <- length(clNames)
    if (!openedGoZone) {
      dflt <- 1:n.clust
      openedGoZone <<- TRUE
    } else {
      dflt <- input$showGoZone
    }
    choices <- 1:n.clust
    names(choices) <- clNames
    checkboxGroupInput("showGoZone","Select Clusters", choices, selected=dflt)
  })
  
  # set the value of the "include in report" checkbox in every plot
  output$includeInReport <- renderUI({
    input$conditionedPanels
    input$plotType
    input$reportType
    input$analysisType
    input$chooseNoCluster
    input$clusterplot
    input$univariateRating
    input$chooseRating1
    input$chooseRating2
    input$chooseRating3
    input$chooseRating4
    input$bivariateRating1
    input$bivariateRating2
    input$demographicSelection
    goZoneIndices()
    input$sortBy
    input$byCluster
    isolate({
      cmap <- ReadSavedData()
      if (is.null(getSavedObject(input))) {
        dflt <- FALSE
      } else {
        dflt <- TRUE
      }  
      checkboxInput("includeInReport", "Include in report", value=dflt)
    })
  })
  
  
  ## Reports options
  
  # How to sort the statements report
  output$sortBy <- renderUI({
    if(is.null(input$sortBy)) {
      srt <- "Statement ID"
    } else {
      srt <- input$sortBy
    }
    selectInput("sortBy", "Sort by", c("Statement ID","Mean (descending)"),
                selected=srt)
  })
  
  # whether to output by cluster, or all the statements in one table
  output$byCluster <- renderUI({
    if(is.null(input$byCluster)) {
      bycls <- FALSE
    } else {
      if(input$byCluster == FALSE) {
        bycls <- FALSE
      }
      else {
        bycls <- TRUE
      }
    }
    checkboxInput("byCluster","By clusters",value = bycls)
  })
  
  
  ## settings
  # select number of clusters, change/set their names
  
  output$chooseNoCluster <- renderUI({
    n.phrases <- get.nsorted()
    if (is.null(input$chooseNoCluster)){
      dflt <- 3
      if (length(getClusters()) > 0) 
        dflt <- length(getClusters())
    } else {
      dflt <- input$chooseNoCluster
    }
    if (!is.null(input$chooseNoCluster) && input$chooseNoCluster!=ReadSavedData()$n.clust)
      openedGoZone <<- FALSE
    SetNumClust(dflt)
    selectInput("chooseNoCluster", "Number of clusters", 2:min(20,n.phrases),selected=dflt)
  })
  
  output$changeClusterNames <- renderUI({
    input$chooseNoCluster
    input$changeName
    clusters <- getClusters()
    if (is.null(input$changeClusterNames)) {
      dflt <- clusters[1]
    } else if (!(input$changeClusterNames %in% clusters)) {
      dflt <- input$getClusterName
    } else {
      dflt <- input$changeClusterNames
    }
    selectInput("changeClusterNames", "Cluster", clusters, selected=dflt)
  })
  
  output$getClusterName <- renderUI({
    input$chooseNoCluster
    input$changeClusterNames
    clusters <- getClusters()
    if (is.null(input$changeClusterNames)) {
      dflt <- clusters[1]
    } else {
      dflt <- input$changeClusterNames
    }
    textInput("getClusterName", "Cluster Name", dflt)
  })
  
  output$clusterNameError <- renderUI({
    input$chooseNoCluster
    input$getClusterName
    div(clusterNameError(), style="color:red")
  })
  
  # set project name and description
  
  output$projectDescription <- renderUI({
    textInput("projectDescription", "Description")    
  })
  
  output$projectName <- renderUI({
    textInput("projectName", "Save as")    
  })
  
  
  # set the file name for the report
  output$reportName <- renderUI({
    textInput("reportName", "Save as")    
  })
  
  
  ###################################################################
  # show the main panel
  ###################################################################
  
  output$conditionedPanels = renderUI({
    if (is.null(input$inputFile))  {
      tabs <- tabPanel(title="New Project", value="NewProject")
      do.call(tabsetPanel, list(tabs,id="conditionedPanels"))
    } else {
      tabstrs <- list("Summary", "Plots", "Reports", "Analysis", "Users",
                      "Settings")
      tabs <- base::lapply(tabstrs, tabPanel)
      for (i in 1:length(tabstrs)) {
        tabs[[i]][[2]]$'data-value' <- tabstrs[[i]]
      }
      tabs$id <- "conditionedPanels"
      tabs$selected <- input$conditionedPanels
      switch(input$conditionedPanels,
             NewProject = { tabs$selected = "Summary"
             do.call(tabsetPanel, tabs)},
             Summary = { 
               list(do.call(tabsetPanel, tabs), 
                    div(style = "overflow:auto; width:100%;", 
                        do.call(uiOutput,list('showProjectInfo')))) },
             Plots = { tagList(do.call(tabsetPanel, tabs), 
                               p(""),
                               do.call(textOutput,list("mouseOverInfo")),
                               #                               div(style = "width:auto; height:auto;object-fit: cover;
                               #                                   min-width=800px;min-height=600px;",
                               do.call(plotOutput,list('plot',width="900px",height="700px",
                                                       hover="plot_hover")),
                               do.call(tableOutput,list('showPlotInfo'))) },
             Reports = { tagList(do.call(tabsetPanel, tabs), 
                                 p(""),
                                 do.call(uiOutput,list('summarytablesheader')),
                                 p(""),
                                 do.call(uiOutput,list('summarytables'))) },
             Analysis = { tagList(do.call(tabsetPanel, tabs),
                                  p(""),
                                  do.call(uiOutput,list('analysistablesheader')),
                                  p(""),
                                  do.call(uiOutput,list('analysistables'))) },
             Users = {
               tagList(do.call(tabsetPanel, tabs),
                       do.call(DT::dataTableOutput,
                               list('showDemographicData')))
             },
             Settings = { tagList(do.call(tabsetPanel, tabs),
                                  do.call(plotOutput,list('plotDendrogram',width="900px",height="700px")),
                                  do.call(tableOutput,list('showStatementsInCluster')),br(),
                                  do.call(tableOutput,list('showSuggestedLables'))) },
             do.call(tabsetPanel, tabs)
      )
    }
  })
  
  
  #################
  # rendering the main panel
  
  # show a quick summary of the content of the input file
  output$showProjectInfo <- renderUI({
    info <- showProjectInfo()
    tags$div(tags$br(),sprintf("  Data File: %s",info[1]),tags$br(),sprintf("  Number of raters: %s  ",info[2]),
             tags$br(),sprintf("  Number of statements: %s  ",info[3]),tags$br(),
             sprintf("  Description: %s  ",info[4]),tags$br(),
             sprintf("  Rating variables: %s  ",info[5]))
  })
  
  # when we plot the MDS graph, the user can see the statement
  # and its average rating when the mouse is over a point in the graph
  output$mouseOverInfo <- renderText({
    if ((input$plotType %in% c("Bar Chart", "Pattern Matching")))
      return(NULL)
    if (is.null(input$plot_hover$x))
      paste0("Put the mouse over a point to see the corresponding statement")
    else {
      cmap <- ReadSavedData()
      rtg <- rep("",cmap$n.phrases)
      switch(input$plotType,
             "GoZone" = {
               cmap$bivratingcol[1] = 
                 which(colnames(cmap$ratings.dat) == input$bivariateRating1)
               cmap$bivratingcol[2] = 
                 which(colnames(cmap$ratings.dat) == input$bivariateRating2)               
               # cmap$clustIndices = clustIndices
               dat <- cmap$sort.dat
               if (is.null(input$demographicSelection))
                 cmap$demographicGroup <- "All"
               else
                 cmap$demographicGroup <- input$demographicSelection
               cmap$selectedGroup <- cmap$demographicGroups[[cmap$demographicGroup]]
               if (cmap$demographicGroup == "All")
                 cmap$selectedGroup$rows <- 1:length(cmap$demographics.dat$SorterID)
               selected <- cmap$demographics.dat$SorterID[cmap$selectedGroup$rows]
               ratings.dat <- cmap$ratings.dat[which(cmap$ratings.dat$UserID %in% selected), ]
               ratings1 <- tapply(ratings.dat[[cmap$bivratingcol[1]]],INDEX=ratings.dat[[2]],FUN=mean,na.rm=T)
               ratings2 <- tapply(ratings.dat[[cmap$bivratingcol[2]]],INDEX=ratings.dat[[2]],FUN=mean,na.rm=T)    
               
               dsts2 <- (input$plot_hover$x-ratings1)^2 + (input$plot_hover$y-ratings2)^2
             },
             "Dendrogram" = {
               dsts2 <- (input$plot_hover$x - 1:length(cmap$fit.WardClust$order))^2
               dsts2[cmap$fit.WardClust$order] = dsts2
             },
             {
               if (is.null(input$demographicSelection)) {
                 selected <- 1:length(cmap$demographics.dat$SorterID)
               }
               else {
                 if (input$demographicSelection == "All")
                   selected <- 1:length(cmap$demographics.dat$SorterID)
                 else
                   selected <- (cmap$demographicGroups[[input$demographicSelection]])$rows
               }
               raters <- cmap$demographics.dat$SorterID[selected]
               dat <- cmap$ratings.dat[cmap$ratings.dat[[1]] %in% raters,]
               dsts2 <- (cmap$x-input$plot_hover$x)^2 + (cmap$y-input$plot_hover$y)^2
               ratings <- tapply(dat[[input$univariateRating]],
                                 INDEX=dat[[2]],FUN=mean,na.rm=T)   
               rtg <- paste0(" (average rating=",sprintf("%2.3f",ratings),")")
             }
      )
      mindst <- which.min(dsts2)[1]
      paste0(cmap$statement.dat$StatementID[mindst]," : ", 
             cmap$statement.dat$Statement[mindst],rtg[mindst])
    }
  })
  
  ##############
  # create the different plots (some functions in the helper file)
  
  output$plot <- renderPlot({
    input$plotType
    input$changeClusterNames
    input$chooseNoCluster
    input$getClusterName
    if (is.null(input$plotType)) {
      ptype <- "Point Map (MDS)"
    } else {
      ptype <- input$plotType
    }
    switch(ptype,
           "Point Map (MDS)" = { ShowMDSPlot() },
           "Clusters" = {
             if (is.null(input$clusterplot))
               return(NULL)
             if(input$clusterplot == "Rays") {
               ShowClusterPlot()
             } else {
               ShowClusterPolygonsPlot()
             }
           },
           "Univariate Rating - by Statement" = {
             if (is.null(input$univariateRating) | is.null(input$demographicSelection))
               return(NULL)
             PlotRatingsLadder(input$univariateRating, input$demographicSelection)
           },
           "Univariate Rating - by Cluster" = {
             if (is.null(input$univariateRating) | is.null(input$demographicSelection))
               return(NULL)
             ShowClusterRankingsPlotShades(input$univariateRating, input$demographicSelection)
           },
           "Bar Chart" = {
             chooseRatingVars = c(input$chooseRating1,input$chooseRating2,
                                  input$chooseRating3,input$chooseRating4)
             if (is.null(chooseRatingVars) | is.null(input$demographicSelection))
               return(NULL)
             BarChart(chooseRatingVars, input$demographicSelection)
           },
           "Pattern Matching" = {
             if (is.null(input$bivariateRating1) | is.null(input$demographicSelection))
               return(NULL)
             ClusterLadderGraph(input$bivariateRating1,input$bivariateRating2,input$demographicSelection)
           },
           "GoZone" = {
             if(length(goZoneIndices()) == 0)
               return(NULL)
             GoZoneGraphClusterList(goZoneIndices(),input$bivariateRating1,
                                    input$bivariateRating2,input$demographicSelection)
           },
           "Dendrogram" = { ShowDendrogram("full") }
    )
  })
  
  # in case of a GoZone, show additional information in a table format
  inccolnames <- FALSE
  output$showPlotInfo <- renderTable({
    tab <- table(c())
    if (input$plotType=="GoZone") {
      if (!is.null(input$showGoZone)) {
        vars <- getClusters()
        dflt <- vars[goZoneIndices()]
        tab <- data.frame(toupper(letters[1:length(dflt)]),dflt)
        colnames(tab) <- c("Label","Cluster Name")
        inccolnames <- TRUE
      }
    }
    tab
  },include.rownames=FALSE,include.colnames=inccolnames)
  
  
  #######
  # the summary tab header
  output$summarytablesheader <-  renderUI({
    input$changeClusterNames
    input$chooseNoCluster
    input$getClusterName
    if (is.null(input$reportType)) {
      rtype <- "Sorters"
    } else {
      rtype <- input$reportType
    }
    switch(rtype,
           "Sorters" = { 
             if (is.null(input$demographicSelection))
               return(NULL)
             hdr <- sprintf("Sorters from group %s",input$demographicSelection)
           },
           "Statements" = {
             if (is.null(input$univariateRating))
               return(NULL)
             hdr <- sprintf("Summary of statements, with %s rating statistics for group %s",input$univariateRating,input$demographicSelection)
           },
           "Clusters" = {
             if (is.null(input$univariateRating))
               return(NULL)
             hdr <- sprintf("Summary of clusters, with %s rating statistics for group %s",input$univariateRating,input$demographicSelection)
           }
    )    
    hdr
  })
  
  # the analysis tab header
  output$analysistablesheader <-  renderUI({
    if (is.null(input$analysisType)) {
      ttype <- "ANOVA"
    } else {
      ttype <- input$analysisType
    }
    switch(ttype,
           "ANOVA" = { 
             hdr <- sprintf("Analysis of Variance for group %s - %s",
                            input$demographicSelection,input$univariateRating)
           },
           "Tukey" = {
             hdr <- sprintf("Comparing clusters (Tukey's test) for group %s - %s ",input$demographicSelection,input$univariateRating)
           }
    )    
    hdr
  })
  
  # show the tables with the F-test result, or Tukey's test
  output$analysistables <-  renderTable({
    input$changeClusterNames
    input$chooseNoCluster
    input$getClusterName
    if (is.null(input$univariateRating) | is.null(input$demographicSelection))
      return(NULL)
    if (is.null(input$analysisType)) {
      ttype <- "ANOVA"
    } else {
      ttype <- input$analysisType
    }
    lbl <- clusterLabels()
    switch(ttype,
           "ANOVA" = { 
             dframe <- rating.Ftest.cluster(
               cluster.id=lbl,
               col=input$univariateRating,
               demographicGroup=input$demographicSelection)
           },
           "Tukey" = {
             dframe <- rating.TukeyHSDtest.cluster(
               cluster.id=lbl,
               col=input$univariateRating,
               demographicGroup=input$demographicSelection)
           }
    )    
    colnames(dframe) <- gsub("\\."," ",colnames(dframe))    
    dframe
  },include.rownames=TRUE,include.colnames=TRUE)
  
  output$PlotBridgingIndex <- renderPlot({
    input$plotType
    input$chooseNoCluster
    input$getClusterName
    ShowBridgingIndex()
  })
  
  # plot a dendrogram of the clusters
  output$plotDendrogram <- renderPlot({
    input$plotType
    input$changeClusterNames
    input$chooseNoCluster
    input$getClusterName
    input$showDemographicData_rows_all
    if (is.null(input$setting)) {
      stype <- "Number of Clusters"
    } else {
      stype <- input$setting
    }
    switch(stype,
           "Number of Clusters" = {
             if (is.null(input$chooseNoCluster)){
               return(NULL)
             } else {
               dflt = input$chooseNoCluster
             }
             if (!is.null(input$chooseNoCluster) && input$chooseNoCluster!=ReadSavedData()$n.clust)
               openedGoZone <<- FALSE
             SetNumClust(dflt)
             ShowDendrogram("full")
           },
           # note: program seems to run fine when the five lines below are commented out. The last three lines are carried out by observeEvent
           "Cluster Names" = {              if (is.null(input$setting) || is.null(input$changeClusterNames))
             return(NULL)
             if(input$changeName > 0 #&& is.null(clusterNameError())
             ) {
               nclname <- newName()
             }
             plotCluster(input$getClusterName,input$changeClusterNames)   
           },
           "Save Project" = {
             if (is.null(input$projectName))
               return(NULL)
             if(input$saveProject > 0) {
               saved <- saveProject()
             }  
           }
    )
  })
  
  # in the settings tab, show the statements in a selected cluster
  output$showStatementsInCluster <- renderTable({
    input$changeClusterNames  
    input$setting
    if (is.null(input$setting) || is.null(input$changeClusterNames) || input$setting!="Cluster Names" || !(input$changeClusterNames %in% getClusters()))
      return(NULL)
    info <- showInfo(input$setting,input$changeClusterNames)
    info
  }, 
  include.rownames=FALSE,
  caption = "Statements in cluster",
  caption.placement = getOption("xtable.caption.placement", "top"), 
  caption.width = getOption("xtable.caption.width", NULL))
  
  # in the settings tab, show suggested labels
  output$showSuggestedLables <- renderTable({
    input$changeClusterNames  
    input$setting
    if (is.null(input$setting) || is.null(input$changeClusterNames) || input$setting!="Cluster Names" || !(input$changeClusterNames %in% getClusters()))
      return(NULL)
    info = showMoreInfo(input$setting,input$changeClusterNames) # problem: this is updating before input$changeClusterNames
    info
  }, include.rownames=FALSE,
  caption = "Suggested cluster names",
  caption.placement = getOption("xtable.caption.placement", "top"), 
  caption.width = getOption("xtable.caption.width", NULL))
  
  
  #######################################################
  # observe blocks
  
  # read the input data (Excel file, and an RData file) from 
  # the fileInput dialog
  observe({
    input$inputFile
    if (is.null(input$inputFile))
      return(NULL)
    OpenProject(input$inputFile)
  })
  
  # select all clusters in gozone plot
  observeEvent(input$selectAllClusters, {
    openedGoZone <<- FALSE
  })
  
  # unselect all clusters in gozone plot
  observeEvent(input$unselectAllClusters, {
    vars <- getClusters()
    updateCheckboxGroupInput(session,"showGoZone",choices=vars,selected=c())
  }) 
  
  # choose which plots should be included in the report
  observe({
    input$includeInReport
    isolate({
      if (!is.null(input$includeInReport)) {
        savedObject <- getSavedObject(input)
        if (input$includeInReport & is.null(savedObject)) {
          saveObject(input)
          updateCheckboxInput(session,"includeInReport",value=TRUE)
        }
        if (!input$includeInReport & !is.null(savedObject)) {
          removeSavedObject(input)
          updateCheckboxInput(session,"includeInReport",value=FALSE)
        }
      }
    })
  })
  
  # functions to create subgroups by demographic criteria
  observe({
    input$createNewGroup
    input$resetGroup
    input$demographicSelection
    
    cmap <- ReadSavedData()
    if (is.null(input$demographicSelection))
      searchCols <- NULL
    else
      searchCols <- cmap$demographicGroups[[input$demographicSelection]]$searchCols
    
    output$showDemographicData <- DT::renderDataTable({
      demographicData("all", input$demographicSelection)
    }, filter=list(position='top'),options=list(searchCols = searchCols,paging = FALSE,searching = TRUE,server=F))
  })
  
  observeEvent(input$createNewGroup, {
    cmap <- ReadSavedData()
    if (nchar(input$newGroupName) > 0 && input$newGroupName!="All") {
      len <- 1
      settingsList <- list(NULL) # need to include NULL as first entry
      for (setting in input$showDemographicData_search_columns) {
        entry <- NULL
        if (nchar(setting) > 0) {
          entry <- list(search = setting)
        }
        settingsList[[len + 1]] = entry
        len <- len + 1
      }
      settingsList
      cmap$demographicGroups[[input$newGroupName]] = 
        list(searchCols = settingsList, rows = input$showDemographicData_rows_current)
      with(cmap,save(description,statement.dat,clustname,n.clust,excel.file,excel.file.name,reportObjects,demographicGroups,
                     sort.dat, inc.mat, group.sim.mat,gsm, 
                     fit.MDS, x, y, stress, n.indiv, n.phrases, fit.WardClust,
                     ratings.dat,colpoints,coltext,showStatementsVal,
                     ratings.scales.dat, mins, maxs, demographics.dat, critvar1, critvar2,
                     file="activesession.RData"))
      updateSelectInput(session,"demographicSelection",
                        choices=names(cmap$demographicGroups),selected=input$newGroupName)
    }
  })
  
  observeEvent(input$removeGroup, {
    cmap <- ReadSavedData()
    if (input$demographicSelection != "All") { # don't let user delete the "All" group
      cmap$demographicGroups[[input$demographicSelection]] = NULL
      with(cmap,save(description,statement.dat,clustname,n.clust,
                     excel.file,excel.file.name,reportObjects,demographicGroups,
                     sort.dat, inc.mat, group.sim.mat,gsm, 
                     fit.MDS, x, y, stress, n.indiv, n.phrases, fit.WardClust,
                     ratings.dat,colpoints,coltext,showStatementsVal,
                     ratings.scales.dat, mins, maxs, demographics.dat, critvar1, critvar2,
                     file="activesession.RData"))
      updateSelectInput(session,"demographicSelection",
                        choices=names(cmap$demographicGroups),selected="All")
      # defaults back to "All"
    }
  })
  
  # output$summarytables is surrounded by observe so that renderTable does not throw an error
  # when switching report types after 'by cluster' is checked
  observe({
    input$byCluster
    input$reportType
    input$chooseNoCluster
    input$changeClusterNames
    input$univariateRating
    input$demographicSelection
    input$saveProject
    
    output$summarytables <-  renderTable({
      input$changeClusterNames
      input$chooseNoCluster
      #input$getClusterName
      #input$saveProject
      isolate({
        if (is.null(input$reportType)) {
          rtype = "Sorters"
        } else {
          rtype = input$reportType
        }
      })
      lbl = clusterLabels()
      switch(rtype,
             "Sorters" = { 
               if (is.null(input$demographicSelection))
                 return(NULL)
               df = get.summary(input$demographicSelection)
             },
             "Statements" = {
               if (is.null(input$univariateRating) | is.null(input$demographicSelection))
                 return(NULL)
               df <- rating.summary(col=input$univariateRating,demographicGroup=input$demographicSelection)
               if (!is.null(input$sortBy)){
                 if (input$sortBy == "Mean (descending)") {
                   df = df[order(df[,3],decreasing =TRUE),]
                 }
                 
                 if (input$byCluster) {
                   df = df[order(df[,5],decreasing =FALSE),]
                 } 
               }
             },
             "Clusters" = {
               if (is.null(input$univariateRating))
                 return(NULL)
               df <- rating.summary.cluster(cluster.id=lbl,col=input$univariateRating,demographicGroup=input$demographicSelection)
             }
      )    
      colnames(df) <- gsub("\\."," ",colnames(df))
      df
    },add.to.row=add.to.row.args(input$byCluster,isolate(input$reportType),input$univariateRating,input$demographicSelection),
    include.rownames=FALSE,include.colnames=TRUE)
  })
  
  
  goZoneIndices <- reactive({
    # input$showGoZone is a character vector; this expression returns an 
    # integer vector and should be used in its place outside of 
    # output$chooseClusterCheckbox
    as.integer(input$showGoZone)
  })
  
  observeEvent(input$changeName,newName())
  
  observeEvent(input$setting, {
    if (input$setting=="Cluster Names") {
      clusters = getClusters()
      updateSelectInput(session,"changeClusterNames",choices=clusters,selected=clusters[1])
    }
  })
  
  openedGoZone <- FALSE
  
  
  newName <- reactive({
    input$changeName
    isolate({
      if (is.null(clusterNameError())) {
        updateClusterName(input$chooseNoCluster,input$changeClusterNames,input$getClusterName)
        input$getClusterName
      }
    })
  })
  
  # error message displayed when an invalid cluster name is entered in getClusterName
  clusterNameError <- reactive({
    input$chooseNoCluster
    #input$changeClusterNames
    input$getClusterName
    clusters = getClusters()
    isolate({
      if (!(is.null(input$getClusterName)) && input$getClusterName=="") {
        return("Error: cluster name cannot be blank")
      } else if (!is.null(input$changeClusterNames) && input$getClusterName!=input$changeClusterNames && input$getClusterName %in% clusters) {
        return("Error: cluster name already in use")
      }
      return(NULL)
    })
  })
  
  saveProject <- reactive({
    input$saveProject 
    input$projectName
    isolate({
      saveSession(input$projectName,input$projectDescription)
      ReloadProject()
      updateTabsetPanel(session, "conditionedPanels", selected = "Summary")
      return("saved")
    })
  })
  
  # handler for downloading the project file
  output$downloadProject <- downloadHandler(
    filename = function() {paste0(input$projectName,".RData")},
    content = function(file) {
      cm = ReadSavedData()
      cm$filename = input$projectName
      cm$description = input$projectDescription
      save(description,filename,filename,statement.dat,clustname,n.clust, excel.file.name, reportObjects,demographicGroups,
           sort.dat, inc.mat, group.sim.mat,gsm, 
           fit.MDS, x, y, stress, n.indiv, n.phrases, fit.WardClust,
           ratings.dat,colpoints,coltext,showStatementsVal,
           ratings.scales.dat, mins, maxs, demographics.dat, critvar1, critvar2,
           file="activesession.RData", envir=cm)
      save(description,filename,filename,statement.dat,clustname,n.clust, excel.file.name, reportObjects,demographicGroups,
           sort.dat, inc.mat, group.sim.mat,gsm, 
           fit.MDS, x, y, stress, n.indiv, n.phrases, fit.WardClust,
           ratings.dat,colpoints,coltext,showStatementsVal,
           ratings.scales.dat, mins, maxs, demographics.dat, critvar1, critvar2,
           file=file, envir=cm)
    })
  
  # handler for downloading the report (saved as .doc file)
  output$downloadReport <- downloadHandler(
    filename = function() { paste0(input$reportName,".doc") },
    content = function(file) { createReport(input$editablePlots, file) }
  )
  
})
