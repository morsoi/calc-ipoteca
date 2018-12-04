
library(shiny)
library(shinydashboard)
library(dplyr)
library(crosstalk)
library(DT)
library(data.table)
library(sqldf)
library(shinyFiles)
library(xlsx)
ui<- fluidPage( sidebarPanel(h3('Аннуитентная ипотека'),
                             numericInput("kr", label = h5('Сумма кредита (млн.)'), value = 10),
                             numericInput("pr", label = h5('Процентная ставка'), value = 9.75),
                             h6('Меняем срок или платеж, второе значение оставляем 0'),
                             numericInput("srok", label = h5('Срок (лет)'), value = 10),
                             numericInput("pay", label = h5('Ежемесячный платеж (тыс. руб.)'), value = 0),
                             textOutput("text1"),
                             dateInput("d1", label = h5('Дата начала'),format = "dd-mm-yyyy", value = '2019-01-01'),
                             numericInput("kom", label = h5('Ежемесячная комиссия (руб.)'), value = 0),
                             h5('Страховка, раз в год'),
                             numericInput("pr_ins", label = h5('Процент страхового взноса'), value = 0),
                             downloadButton("downloadData", "Сохранить"
                                          ,style="color: #fff; background-color: #337ab7; border-color: #2e6da4")
) ,
mainPanel(
  h4('Сводная по кредиту'),
  DT::dataTableOutput("x2"),
  h4('Сводная по кредиту с учетом доп. взносов'),
  DT::dataTableOutput("x3"),
  h4('График доп. взносов'),
  column(6,offset = 6, 
         HTML('<div class="btn-group" role="group" aria-label="Basic example">'),
         actionButton(inputId = "Add_row_head",label = "Добавить платеж"),
         actionButton(inputId = "Del_row_head",label = "Удалить платеж"),
         HTML('</div>')),
  
  DT::dataTableOutput("x0"),
  h3('График платежей'),
  DT::dataTableOutput("x1")
)
)

######################
server <- function(input, output) {
  
  add.months= function(date,n) seq(date, by = paste (n, "months"), length = 2)[2]
  val.catch.error <- function(x){ tryCatch(as.Date(x, '%Y-%m-%d'), error=function(e) NA) } 
  add.day= function(date,n) seq(date, by = paste (n, "day"), length = 2)[2]
  NF <- function(x) { return(format(x,big.mark=",",trim=TRUE,scientific=FALSE)) } 

  ###############################  
  dataInput <- reactive({  
    d1<- input$d1
    kr<- input$kr*1000000
    pr<- input$pr/100
    srok<- input$srok
    pr_ins<- input$pr_ins/100
    
    pay <- if(srok==0){ input$pay*1000} else 0
    
    pay<-  if (pay ==0 ) {
      (kr*pr/12)/(1-(1/(1+pr/12)^(srok*12)))
    } else pay
    
    pay<-if (pay <= kr*pr/12) {
      kr*pr/12} else pay 
    
    srok<- if (srok ==0) {
      log(((pay)/(pay - kr*pr/12)),base=(1+pr/12))/12
    } else srok
    
    
    pereplata <- (srok*12*pay - kr)
    
    kom <- input$kom  
    
    d2<-if(srok>0){ add.months(d1, srok*12)} else d1
    d2 <-if(is.na(d2)){d1} else d2
    
    dt0 <- data.frame(data=d1,kredit=kr,pereplat=kr*pr/12
                      ,pay=pay ,ostatok=kr +kr*pr/12 - pay
                      ,ins=kr*(1+ pr )*pr_ins)
    pay1<-pay
    i<-1
    while (min(dt0$ostatok)>0) {
      i<-i+1
      kr1<- as.numeric(dt0[i-1,5])
      data<-add.months(dt0[i-1,1], 1)
      
      if(kr1 +kr1*pr/12 - pay1<0) {
        pay1<-(kr1 +kr1*pr/12 )
      }
      ins<- if (i%%12==0) { kr1*(1+ pr )*pr_ins } else 0
      
      dt0[i,]<-list(data,kr1,kr1*pr/12,pay1 ,kr1 +kr1*pr/12 - pay1,ins)
    }
    ins <- sum(dt0$ins)
    dt<-data.table(Start=d1,End=d2,kr,pr,srok,pay,pereplata,kom,pr_ins,ins)
    
  })
  ##
  output$text1 <- renderText({
    t<-dataInput()
    kr<- t$kr
    pr<- t$pr
    pay <- t$pay
    pereplat=round(kr*pr/12,0)
    
    if (pereplat>input$pay) { paste('Увеличь сумму платежа до ', pereplat)
    }  else  ''
    
  } ) 
  ###  
  dataInput2<- reactive({  
    t<-dataInput()
    data<- t$Start
    kr<- t$kr
    pr<- t$pr
    srok<- t$srok
    pay <- t$pay
    kom <- t$kom
    pr_ins<-t$pr_ins
    
    # x$dt$data<-as.Date(format(x$dt$data, "%Y-%m-01"),"%Y-%m-%d")
    
    dt1<- x$dt%>% group_by(data = x$dt$data) %>%
      summarise(add_pay = sum(add_pay),add_pay1=sum(add_pay1)) 
    
    fin<- add.months(as.Date(data,"%Y-%m-%d"), srok*12+1) 
    m<-data.frame(m1=seq(data,fin,by='month'),m2=seq(add.months(data, 1),add.months(fin, 1),by='month') )
    dt1<-sqldf("SELECT dt1.data,dt1.add_pay,dt1.add_pay1
               , m.m1 as m1
               FROM dt1 inner JOIN m ON dt1.data >= m.m1 and dt1.data < m.m2
               --where dt1.add_pay + dt1.add_pay1>0
               ")
    
    #############
    m<-data
    
    add_pay1<- if(nrow(dt1[which(dt1$data==data & dt1$add_pay1>0),] )>0) {
      sum(subset(dt1[which(dt1$data==data),],select =c("add_pay1")))
    } else 0
    
    add_pay<- if(nrow(dt1[which(dt1$data==data & dt1$add_pay>0),] )>0) {
      sum(subset(dt1[which(dt1$data==data ),],select =c("add_pay")))
    } else 0
    
    kr<- if(add_pay+add_pay1>0) { kr-add_pay-add_pay1 } else kr
    
    pay <-  if (add_pay1 >0)  {
      ((kr)*pr/12)/(1-(1/(1+pr/12)^(srok*12)))
    } else pay
    
    d<-as.integer(add.months(data, 1) - data)
    
    if(nrow(dt1[which(dt1$m==m & (dt1$add_pay1 +dt1$add_pay) >0),] )>0){
      
      dt <- data.frame(data,m=data,kredit=kr,pereplat=kr*pr/(12*d)
                       ,pay=pay/d,ostatok=kr +kr*pr/(12*d) - pay/d 
                       ,add_pay=add_pay ,add_pay1,ins=kr*(1+ pr )*pr_ins,t=1)
    } else {
      dt <- data.frame(data,m=data,kredit=kr,pereplat=kr*pr/(12)
                       ,pay=pay ,ostatok=kr +kr*pr/(12) - pay 
                       ,add_pay=0 ,add_pay1=0,ins=kr*(1+ pr )*pr_ins,t=0)
    }
    #dt<-dt[-30,]
    
    i<-1
    #min(dt$ostatok)>0
    while (min(dt$ostatok)>0) {
      
      i<-i+1
      kr1<- as.numeric(dt[i-1,6])
      
      pay<-if(dt[i-1,"t"]==1) {dt[i-1,"pay"]*d} else dt[i-1,"pay"] 
      
      m <- if(dt[i-1,"t"]==1 & add.day(dt[i-1,1],1)< add.months(dt[i-1,2], 1)) {dt[i-1,2] } else add.months(dt[i-1,2], 1)
      d<-as.integer(add.months(m, 1) - m) 
      if(nrow(dt1[which(dt1$m==m & (dt1$add_pay+dt1$add_pay1)>0),] )>0) { 
        
        data<- if(dt[i-1,"t"]==1){ add.day(dt[i-1,1], 1)} else m
        
        add_pay1<- if(nrow(dt1[which(dt1$data==data & dt1$add_pay1>0),] )>0) {
          sum(subset(dt1[which(dt1$data==data),],select =c("add_pay1")))
        } else 0
        
        add_pay<- if(nrow(dt1[which(dt1$data==data & dt1$add_pay>0),] )>0) {
          sum(subset(dt1[which(dt1$data==data),],select =c("add_pay")))
        } else 0
        kr1<-kr1 -add_pay
        if (add_pay1>0 )  {
          srok <- log(((pay)/(pay - kr1*pr/12)),base=(1+pr/12))/12
          kr1<-kr1 - add_pay1 
          pay <- ((kr1)*pr/(12))/(1-(1/(1+pr/12)^(srok*12)))
        } 
        
        if(kr1 +kr1*pr/(12*d) <=pay/d) {
          pay<-(kr1 + kr1*pr/(12*d))*d 
        }
        ostatok<-if(kr1 +kr1*pr/(12*d) >pay/d) { 
          kr1 +kr1*pr/(12*d) - pay/d } else  0
        
        ins<- if (as.integer(add.months(dt[1,"data"], 12) - data)%%365==0) { kr1*(1+ pr )*pr_ins } else 0
        
        dt[i,]<-list(data,m,kr1,kr1*pr/(12*d),pay/d ,ostatok
                     ,add_pay ,add_pay1,ins,t=1)
      }  else {
        data<-add.months(dt[i-1,2], 1) 
        m <- data
        d<-as.integer(add.months(dt[i-1,2], 1) - dt[i-1,2])
        pay<-if(dt[i-1,"t"]==1) {dt[i-1,"pay"]*d} else dt[i-1,"pay"]
        
        if(kr1 +kr1*pr/(12) <=pay) {
          pay<-(kr1 + kr1*pr/(12) ) 
        }
        ostatok<-if(kr1 +kr1*pr/(12) > pay) { 
          kr1 +kr1*pr/(12) - pay }  else  0
        
        ins<- if (nlevels(cut(seq(dt[1,2],data, by="months"),"months"))%%12==0) { kr1*(1+ pr )*pr_ins } else 0
        
        dt[i,]<-list(data,m,kr1,kr1*pr/(12),pay ,ostatok,0 ,0,ins,0) 
      }
    }
    dt<- dt%>% group_by (data=m)%>%
      summarise(kredit = max(kredit),pereplat=sum(pereplat),pay=sum(pay),ostatok=min(ostatok)
                ,add_pay = sum(add_pay),add_pay1 = sum(add_pay1),ins= sum(ins) )
  })
  ######################
  x <- reactiveValues()
  dt<- data.frame(data=as.Date('2019-01-01','%Y-%m-%d'),add_pay=0,add_pay1=0)
  x$dt<-dt
  
  observeEvent(input$d1,{
    x$dt<-data.frame(data=as.Date(input$d1,'%Y-%m-%d'),add_pay=0,add_pay1=0)
  })
  #####################
  output$x0 <- DT::renderDataTable({
  dt<-x$dt
    DT::datatable(dt, editable = TRUE
                   , colnames = c('Дата внесения', 'Взнос на сокращение срока', 'Взнос на сокращение платежа')) %>% 
      formatRound(2:3,0) %>% formatDate(1, "toLocaleDateString")
  } )
  
  
  observeEvent(input$Add_row_head,{
     new_row<-data.frame(data=add.months(max(x$dt$data), 1),add_pay=as.numeric(tail(x$dt["add_pay"], n=1)),add_pay1=0)
    x$dt<-rbind(x$dt,new_row)
  })
  
  observeEvent(input$Del_row_head,{
    
    if (length(input$x0_rows_selected)) {
      x$dt <- x$dt[-input$x0_rows_selected, ]
    } else 
      x$dt
    
  }
  )
  ##################
  observeEvent(input$x0_cell_edit, {
    info = input$x0_cell_edit
    str(info)
    i = info$row
    j = info$col
    #v = info$value
    # problem starts here
    v = if(j==1 &  is.na(val.catch.error(info$value))){as.Date(input$d1,'%Y-%m-%d')} else info$value
    x$dt[i, j] <- isolate(DT::coerceValue(v, x$dt[i, j]))
  })  
  ######################Detalka
  output$x1 <-  DT::renderDataTable(server = FALSE,{
    
    t<-dataInput()
    kom <- t$kom
    
    dt<-dataInput2()
    dt$pay<-dt$pay + kom
    dt<-dt %>% 
      mutate_if(is.numeric, round, digits=0)
    x$x1<-dt
    pay<- 
    datatable(dt%>%
                bind_rows(summarise(dt, data=max(data),kredit=max(kredit)
                                    , pereplat=round(sum(pereplat),0)
                                    , pay=if(sum(add_pay1)>0) {round(mean(dt$pay),0) } else round(max(dt$pay),0)
                                    ,ostatok=0
                                    ,add_pay=round(sum(add_pay),0),add_pay1=round(sum(add_pay1),0),ins=round(sum(ins),0)))
                , colnames = c('Месяц', 'Сумма кредита', 'Переплата', 'Платеж', 'Остаток','Взнос на сокращение срока', 'Взнос на сокращение платежа','Страховка')
              ,extensions = "Buttons", options = list(
                dom = 'Bfrtip',
                buttons = 
                  list(list(
                    extend = 'collection',
                    buttons = list(  list(extend = "csv", filename = "Ipoteka")
                                   , list(extend = "excel", filename = "Ipoteka")
                                   , list(extend = "pdf", filename = "Ipoteka")),
                    text = 'Сохранить таблицу'
                  )))
              ) %>% formatRound(2:8,0) %>% formatDate(1, "toLocaleDateString") %>% 
      formatStyle(0,target = "row", fontWeight = styleEqual(nrow(dt)+1, "bold"))
  } 
  )  
  
  ##
  output$x2 <- renderDataTable({
    
    dt<-dataInput()
    dt$kom<-dt$kom*dt$srok*12
  dt<-dt[,-"pr_ins"]
  dt<-dt %>% 
    mutate_if(is.numeric, round, digits=0)
  x$x2<-dt
    datatable(dt,options = list(dom = 't'),rownames = FALSE#, server = FALSE
              , colnames = c('Начало', 'Конец', 'Кредит', 'Процент', 'Срок','Платеж','Переплата','Комиссия','Страховка')
    ) %>% formatRound(3,0)%>% formatRound(6:9,0) %>% formatPercentage(4,2) %>% formatDate(1:2, "toLocaleDateString")
    
  }#,  server = FALSE, options = list(dom = 'tip')
  # ,caption = htmltools::tags$caption(
  #   style = 'caption-side: bottom; text-align: center;',   'Сводная по кредиту '  )
  
  )
  
  #############
  output$x3 <- renderDataTable({
    
    dt<-dataInput2()
    dt1<-dataInput()
    
    Delta_month_srok <- if(sum(dt$add_pay + dt$add_pay1)>0){ (dt1$End-max(dt$data))/30 } else 0
    Delta_Pereplata_rub = if (sum(dt$add_pay + dt$add_pay1)>0 & dt1$pereplat - sum(dt$pereplat)>=0) {
      (dt1$pereplat - sum(dt$pereplat))} else 0 
    
    Pereplata<-  if(sum(dt$add_pay + dt$add_pay1)>0){ sum(dt$pereplat) } else dt1$pereplat
    
    pp<-if(dt1$ins - sum(dt$ins)>0){dt1$ins - sum(dt$ins)} else 0
   
     sale<- if(sum(dt$add_pay + dt$add_pay1)>0){
      Delta_Pereplata_rub + dt1$kom*Delta_month_srok + pp } else 0

    dt<-data.frame(Start=min(dt$data),End=max(dt$data),Delta_month_srok
                   ,Pereplata=Pereplata
                   ,Delta_Pereplata_rub
                   ,Kom = dt1$kom * nrow(dt)
                   ,ins = sum(dt$ins)
                   ,Add_pay=sum(dt$add_pay + dt$add_pay1)
                   ,sale=sale)
    dt<-dt %>% 
      mutate_if(is.numeric, round, digits=0)
    x$x3<-dt
    datatable(dt,options = list(dom = 't') ,rownames = FALSE
              , colnames = c('Начало', 'Конец', 'Сокращение срока (мес)', 'Переплата', 'Разница в переплате'
                             ,'Комиссия','Страховка','Сумма доп. взносов','Экономия')
    ) %>% formatRound(3,1)%>% formatRound(4:9,0) %>% formatDate(1:2, "toLocaleDateString")
         
  })
  ############## Save
  
  output$downloadData <- downloadHandler(
    filename = function() {
      (paste("Ipoteka.xlsx"))
    },
    content = function(file) {

    dt1 <- x$x1
    dt1$data <- format(dt1$data,'%d.%m.%Y')
    dt2 <- x$x2
    dt2$Start <- format(dt2$Start,'%d.%m.%Y')
    dt2$End <- format(dt2$End,'%d.%m.%Y')
    dt3 <- x$x3
    dt3$Start <- format(dt3$Start,'%d.%m.%Y')
    dt3$End <- format(dt3$End,'%d.%m.%Y')
    ##################
    wb<-createWorkbook("xlsx")
     TITLE_STYLE <- CellStyle(wb)+ Font(wb,  heightInPoints=16, 
                                       color="blue", isBold=TRUE, underline=1)
    SUB_TITLE_STYLE <- CellStyle(wb) + 
      Font(wb,  heightInPoints=14,
           isItalic=TRUE, isBold=FALSE)
    # Styles for the data table row/column names
    TABLE_ROWNAMES_STYLE <- CellStyle(wb) + Font(wb, isBold=TRUE)
    TABLE_COLNAMES_STYLE <- CellStyle(wb) + Font(wb, isBold=TRUE) +
      Alignment(wrapText=TRUE, horizontal="ALIGN_CENTER") +
      Border(color="black", position=c("TOP", "BOTTOM"), 
             pen=c("BORDER_THIN", "BORDER_THICK")) 
    # Create a new sheet in the workbook
    sheet <- createSheet(wb, sheetName = "Ипотека")
    xlsx.addTitle<-function(sheet, rowIndex, title, titleStyle){
      rows <-createRow(sheet,rowIndex=rowIndex)
      sheetTitle <-createCell(rows, colIndex=1)
      setCellValue(sheetTitle[[1,1]], title)
      setCellStyle(sheetTitle[[1,1]], titleStyle)
    }
    # Add title
    xlsx.addTitle(sheet, rowIndex=1, title="Расчет ипотеки с доп платежами",
                  titleStyle = TITLE_STYLE)
    # Add sub title
    xlsx.addTitle(sheet, rowIndex=2, 
                  title="Сводная",
                  titleStyle = SUB_TITLE_STYLE)
    # Add sub title
    xlsx.addTitle(sheet, rowIndex=5, 
                  title="Сводная с учетом доп взносов",
                  titleStyle = SUB_TITLE_STYLE)
    xlsx.addTitle(sheet, rowIndex=9, 
                  title="График доп платежей",
                  titleStyle = SUB_TITLE_STYLE)
    # Add a table into a worksheet
    addDataFrame(dt2[1,], sheet, startRow=3, startColumn=1,
                 colnamesStyle = TABLE_COLNAMES_STYLE,
                 row.names=FALSE)
    # Change column width
    addDataFrame(dt3[1,], sheet, startRow=6, startColumn=1, 
                 colnamesStyle = TABLE_COLNAMES_STYLE,
                 row.names=FALSE)
    addDataFrame(dt1, sheet, startRow=10, startColumn=1, 
                 colnamesStyle = TABLE_COLNAMES_STYLE,
                 rownamesStyle = TABLE_ROWNAMES_STYLE)
  ##############

    saveWorkbook(wb,file)
    }
    )

}
#########

shinyApp(ui = ui, server = server)
