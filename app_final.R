library(readxl)
library(caret)
library(dplyr)
library(ggplot2)
library(shiny)
library(mice) 
library(DT)
library(tibble)

data<- read_excel("~/Downloads/Window_Manufacturing.xlsx")

realBlanks<- c("Window Type", "Glass Supplier", "Glass Supplier Location")
for (field in realBlanks) {
  data[[field]][is.na(data[[field]])] <- "NA"
}

#coercing columns
data$`Window Type`<- as.factor(data$`Window Type`)
data$`Glass Supplier`<- as.factor(data$`Glass Supplier`)
data$`Glass Supplier Location`<- as.factor(data$`Glass Supplier Location`)

colnames(data)<- c("Breakage_Rate", "Window_Size", "Glass_Thickness","Ambient_Temp", "Cut_Speed", "Edge_Deletion_Rate", 
                "Spacer_Distance", "Window_Color", "Window_Type", "Glass_Supplier", "Silicon_Viscosity", "Glass_Supplier_Location")

#fill in empty values
imputedValues <- mice(data=data
                      , seed=2016     # keep to replicate results
                      , method="cart" # model you want to use
                      , m=1           # Number of multiple imputations
                      , maxit = 1     # number of iterations
)

# impute the missing values in our data.frame
data <- mice::complete(imputedValues,1) # completely fills in the missing
unprocessed_data<- data

# dummy variables
dummies <- dummyVars(Breakage_Rate ~ ., data )  # create dummies for Xs
d_encoded <- data.frame(predict(dummies, newdata = data)) # actually creates the dummies
names(d_encoded) <- gsub("\\.", "", names(d_encoded))          # removes dots from col names
data<- cbind(data$Breakage_Rate, d_encoded)       
names(data)[1]<- "Breakage_Rate"

# Compute the correlation matrix
corr_matrix <- cor(data, use = "pairwise.complete.obs")

#Find correlated variables with correlation higher than 80% (or lower than -80%)
high_corr <- findCorrelation(corr_matrix, cutoff = 0.8)

# Remove the highly correlated variables from the dataset
data <- data[, -high_corr]

#remove linear combos
Breakage_Rate <- data$Breakage_Rate

# create a column of 1s. This will help identify all the right linear combos
data <- cbind(rep(1, nrow(data)), data[2:ncol(data)])
names(data)[1] <- "ones"

# identify the columns that are linear combos
comboInfo <- findLinearCombos(data)
comboInfo

# remove columns identified that led to linear combos
data <- data[, -comboInfo$remove]

# remove the "ones" column in the first column
data <- data[, c(2:ncol(data))]

# Add the target variable back to our data.frame
data <- cbind(Breakage_Rate, data)

nzv <- nearZeroVar(data, saveMetrics = TRUE)
head(nzv, 20)
#remove Window_TypeNA and Glass_Supplier_LocationNA
data$Window_TypeNA<-NULL
data$Glass_Supplier_LocationNA<- NULL

set.seed(1234) # set a seed so you can replicate your results

inTrain <- createDataPartition(y = data$Breakage_Rate,   # outcome variable
                               p = .80,   # % of training data you want
                               list = F)
# create your partitions
train <- data[inTrain,]  # training data set
test <- data[-inTrain,]  # test data set


# Define UI for application
ui <- fluidPage(
  
  # App title
  titlePanel("Window Breakage Rate Analysis"),
  
  # Sidebar layout with input and output definitions
  sidebarLayout(
    sidebarPanel(
      p("Use this dashboard to analyze factors affecting window breakage rates.")
    ),
    
    # Main panel for displaying outputs
    mainPanel(
      tabsetPanel(
        
        # First tab: Problem Framing
        tabPanel("Problem Framing",
                 h3("Business Problem"),
                 p("The window manufacturing company wants to reduce the breakage rate of windows during production and shipping."),
                 p("The end-users for this application are window manufacturing technicians and process engineers. Their goal is to identify factors that affect breakage rates and adjust processes accordingly to reduce defects."),
                 p("Constraints include available data from manufacturers and customers, as well as the need to balance cost efficiency with product quality."),
                 p("Expected benefits: Reducing breakage rates by 5% could save the company thousands in rework costs and improve customer satisfaction."),
                 
                 h3("Analytics Problem Framing"),
                 p("The problem is reformulated into an analytics problem by analyzing how different variables influence the breakage rate. We are using three types of analytics:"),
                 tags$ul(
                   tags$li("Descriptive Analytics: Summarize the current and historical breakage rates across different manufacturing conditions."),
                   tags$li("Predictive Analytics: Predict future breakage rates based on manufacturing settings using regression models."),
                   tags$li("Prescriptive Analytics: Suggest process adjustments to minimize breakage.")
                 ),
                 p("We assume that all data is accurate and that each variable may have some influence on the breakage rate. Key relationships include the effect of glass thickness and window type on breakage rates."),
                 p("Key success metrics: Reduction in breakage rate, improved consistency in production output.")
        ),
        
        # Second tab: Data
        tabPanel("Data",
                 h3("Data Dictionary"),
                 DT::dataTableOutput("dataDictTable"),
                 h3("Sample Data"),
                 DT::dataTableOutput("windowDataTable")
        ),
        
        #Third tab: Descriptive Analytics
        tabPanel("Descriptive Analytics",
                 fluidRow(
                   column(6, plotOutput("plot1")),
                   column(6, plotOutput("plot2"))
                 ),
                 fluidRow(
                   column(6, plotOutput("plot3")),
                   column(6, plotOutput("plot4"))
                 ),
                 fluidRow(
                   column(12, plotOutput("plot5"))
          
        )
      ),
      
      #Fourth tab: Predictive Analytics
      tabPanel("Predictive Analytics",
               sidebarLayout(
                 sidebarPanel(
                   actionButton("run_model", "Run Linear Regression Model"),
                   checkboxInput("show_test", "Show Test Data Fit", value = TRUE)
                 ),
                 mainPanel(
                   h3("Estimated Coefficients and p-values"),
                   tableOutput("model_summary"),   # Table for coefficients and p-values
                   h3("Model Fit on Training Data"),
                   plotOutput("train_fit_plot"),   # Plot for model fit on training data
                   conditionalPanel(
                     condition = "input.show_test == true",
                     h3("Model Fit on Test Data"),
                     plotOutput("test_fit_plot")   # Plot for model fit on testing data
                   )
                 )
               )
      ),
      
      #Fifth tab: Prescriptive Analytics
      tabPanel("Prescriptive Analytics",
               sidebarLayout(
                 sidebarPanel(
                   h4("Set Non-Controllable Variables"),
                   sliderInput("Ambient_Temp", "Ambient Temperature", 
                               min = min(data$Ambient_Temp), max = max(data$Ambient_Temp), value = mean(data$Ambient_Temp)),
                   sliderInput("Glass_Thickness", "Glass Thickness", 
                               min = min(data$Glass_Thickness), max = max(data$Glass_Thickness), value = mean(data$Glass_Thickness)),
                   sliderInput("Window_Size", "Window Size", 
                               min = min(data$Window_Size), max = max(data$Window_Size), value = mean(data$Window_Size)),
                   h4("Optimization Controls"),
                   actionButton("optimize", "Run Optimization")
                 ),
                 mainPanel(
                   h4("Optimization Results"),
                   textOutput("min_Breakage_Rate"),
                   textOutput("values")
                 )
               )
      )
    )
  )
)
)

# Define server logic
server <- function(input, output) {
  
  # Data dictionary (based on the loaded data dictionary)
  data_dict <- data.frame(
    Variable = c("Breakage Rate",	"Window Size","Glass thickness", "Ambient Temp",	"Cut speed",	"Edge Deletion rate",	
                 "Spacer Distance",	"Window color",	"Window Type",	"Glass Supplier",	"Silicon Viscosity",	
                 "Glass Supplier Location"),
    `Unit of Measure` = c('Numerical', 'Numerical', 'Numerical', 'Numerical', 'Numerical', 'Numerical', 'Numerical', 'Numerical',
                          'Categorical','Categorical', 'Numerical', 'Categorical'),
    `Predictive Model Perspective` = c('Target', 'Input', 'Input', 'Input', 'Input', 'Input', 'Input', 'Input', 'Input', 'Input', 'Input', 'Input'),
    `Process Perspective` = c('Manufacturer', 'Customer spec', 'Customer spec', 'Manufacturer', 'Manufacturer', 'Manufacturer',
                              'Manufacturer', 'Customer spec', 'Customer spec', 'Manufacturer', 'Manufacturer', 'Manufacturer')
  )
  
  # Render the Data Dictionary
  output$dataDictTable <- DT::renderDataTable({
    datatable(data_dict)
  })
  
  # Load the actual window data (assuming it's in the environment)
  window_data <- data
  
  # Render a sample of the actual data
  output$windowDataTable <- DT::renderDataTable({
    datatable(head(window_data, 10))
  })
  
  #Plots
  
  output$plot1 <- renderPlot({
    ggplot(unprocessed_data, aes(x = Window_Type)) +
      geom_bar(fill = "lightblue") +
      labs(title = "Frequency of Window Types",
           x = "Window Type",
           y = "Count") +
      theme_minimal()
  })
  
  output$plot2 <- renderPlot({
    ggplot(unprocessed_data, aes(x = Glass_Supplier)) +
      geom_bar(fill = "lightgreen") +
      labs(title = "Frequency of Glass Suppliers",
           x = "Glass Supplier",
           y = "Count") +
      theme_minimal()
  })
  
  output$plot3 <- renderPlot({
    ggplot(unprocessed_data, aes(x = Glass_Supplier_Location)) +
      geom_bar(fill = "salmon") +
      labs(title = "Frequency of Glass Supplier Locations",
           x = "Glass Supplier Location",
           y = "Count") +
      theme_minimal()
  })
  
  output$plot4<-renderPlot({
    library(reshape2)
    
    cor_matrix <- cor(unprocessed_data[, sapply(unprocessed_data, is.numeric)])
    melted_cor <- melt(cor_matrix)
    
    ggplot(melted_cor, aes(Var1, Var2, fill = value)) +
      geom_tile() +
      scale_fill_gradient2(low = "blue", high = "red", mid = "white", 
                           midpoint = 0, limit = c(-1, 1), space = "Lab", 
                           name="Correlation") +
      theme_minimal() +
      labs(title = "Correlation Matrix Heatmap",
           x = "Variables",
           y = "Variables") +
      theme(axis.text.x = element_text(angle = 45, hjust = 1))
  })
  
  
  # Fit the model when the button is clicked
  observeEvent(input$run_model, {
    # Fit a linear regression model on training data
    lm_model <- lm(Breakage_Rate ~ ., data = train)
    
    # Render the table with the estimated coefficients and p-values
    output$model_summary <- renderTable({
      summary(lm_model)$coefficients %>%
        as.data.frame() %>%
        rownames_to_column(var = "Variable") %>%
        rename(Estimate = Estimate, Std_Error = `Std. Error`, t_value = `t value`, p_value = `Pr(>|t|)`) %>%
        mutate(Estimate = round(Estimate, 4),
               p_value = round(p_value, 4))
    })
    
    # Plot the model fit on training data
    output$train_fit_plot <- renderPlot({
      ggplot(train, aes(x = Window_Size, y = Breakage_Rate)) +
        geom_point() +
        geom_smooth(method = "lm", formula = y ~ x, se = FALSE, color = "blue") +
        labs(title = "Model Fit: Breakage Rate vs Window Size (Training Data)",
             x = "Window Size",
             y = "Breakage Rate") +
        theme_minimal()
    })
    
    # Plot the model fit on testing data if checkbox is selected
    output$test_fit_plot <- renderPlot({
      ggplot(test, aes(x = Window_Size, y = Breakage_Rate)) +
        geom_point() +
        geom_smooth(method = "lm", formula = y ~ x, se = FALSE, color = "green") +
        labs(title = "Model Fit: Breakage Rate vs Window Size (Testing Data)",
             x = "Window Size",
             y = "Breakage Rate") +
        theme_minimal()
    })
  })
  
  model<- lm(Breakage_Rate~Window_Size+Glass_Thickness+Ambient_Temp+Cut_Speed+Edge_Deletion_Rate+Spacer_Distance+Silicon_Viscosity, data=data)
  coefs <- coefficients(model)
  # Optimization function that runs when 'optimize' button is clicked
  observeEvent(input$optimize, {
    lps.model <- make.lp(nrow=0, ncol=8)
    set.type(lps.model, columns=1:8, type="real")
    lp.control(lps.model, sense="min")
    set.objfn(lps.model, obj=c(coefs[1], coefs[2], coefs[3], coefs[4], coefs[5], coefs[6], coefs[7], coefs[8]))
    add.constraint(lps.model, c(0, 0, 0, 0, 1, 0, 0, 0), ">=", min(data$Cut_Speed)) # Min Cut Speed
    add.constraint(lps.model, c(0, 0, 0, 0, 0, 1, 0, 0), ">=", min(data$Edge_Deletion_Rate))   # Min Edge Deletion Rate
    add.constraint(lps.model, c(0, 0, 0, 0, 0, 0, 1, 0), ">=", min(data$Spacer_Distance))   # Min Spacer Distance
    add.constraint(lps.model, c(0, 0, 0, 0, 0, 0, 0, 1), ">=", min(data$Silicon_Viscosity))   # Min Silicon Viscosity
    
    add.constraint(lps.model, c(0, 0, 0, 0, 1, 0, 0, 0), "<=", max(data$Cut_Speed))   # Max Cut Speed
    add.constraint(lps.model, c(0, 0, 0, 0, 0, 1, 0, 0), "<=", max(data$Edge_Deletion_Rate))   # Max Edge Deletion Rate
    add.constraint(lps.model, c(0, 0, 0, 0, 0, 0, 1, 0), "<=", max(data$Spacer_Distance))   # Max Spacer Distance
    add.constraint(lps.model, c(0, 0, 0, 0, 0, 0, 0, 1), "<=", max(data$Silicon_Viscosity))     # Max Silicon Viscosity
    
    add.constraint(lps.model, c(1, 0, 0, 0, 0, 0, 0, 0), "=", 1)
    add.constraint(lps.model, c(0, 1, 0, 0, 0, 0, 0, 0), "=", input$Window_Size)
    add.constraint(lps.model, c(0, 0, 1, 0, 0, 0, 0, 0), "=", input$Glass_Thickness)
    add.constraint(lps.model, c(0, 0, 0, 1, 0, 0, 0, 0), "=", input$Ambient_Temp)
    
    solve(lps.model)
    output$min_Breakage_Rate <- renderText({get.objective(lps.model)})
    output$values <-renderText({get.variables(lps.model)})
   
  })
  
}

# Run the application 
shinyApp(ui = ui, server = server)

