shinyUI(fluidPage(theme = "bootstrap.css",
                  tags$h1( tags$style(HTML("
                                     @import url('//fonts.googleapis.com/css?family=Lobster|Cabin:400,700');
                                     h1 {
                                     font-family: 'Lobster', cursive;
                                     font-weight: 500; line-height: 1.1;
                                     margin-left: 0px; color: #984B43;
                                     background-color: #f7eccc;
                                     }"))),
  pageWithSidebar(
    headerPanel("R-CMap - Group Concept Mapping"),
    sidebarPanel(
      uiOutput('sidebar'),
                 hr(),
      fluidRow(
        column(width=6, "",checkboxInput("Quit", "Quit")),
        column(width=6, "",uiOutput('quit'))),
      uiOutput('showDoc'),
      width=3),
    mainPanel(uiOutput('conditionedPanels'))
  )))

