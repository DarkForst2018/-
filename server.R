#辅助运行脚本
library(openxlsx)
library(readxl)
library(tidyverse)
library(sampling)

dt <- read_xlsx("e:/rproject/shiny/测评对象抽样/size.xlsx")%>%
    mutate(抽样向量=if_else(抽样类别==1,15,10))

size <- dt$抽样向量

size_tb <- as_tibble(size)


server <- function(input, output) {
    
    #size_tb <- reactive({
        #size_tb
   # })
    output$size_tb <- renderTable({
        size_tb
    })
    
    #抽样名单文档载入
    data <- reactive({
        if(input$文档载入判定1=="使用测试文档"){
            dt <- read_xlsx("e:/rproject/shiny/data/2021崂山艺术项目/崂山区学生艺术测评模拟名单.xlsx")
            
            
        }else{
            req(input$upload1)
            dt <- read_xlsx(paste0(input$upload1$datapath))
        }
    })
    
    #行列数量
    output$data_num <- renderText({
        paste0("列数:",ncol(data()),"-","行数:",nrow(data()))
    })
    
    #待抽样数据显示
    output$dt <- renderTable({
        head(data())
    })
    
    #学段数量统计
    #学校数量统计
    #学生数量统计
    
    
    
    #待抽样学校名单
    output$school_list <- renderTable({
        data()%>%
            distinct(学校,学段)
        
    })
    
    #待抽样学校数据统计
    output$school_num <- renderTable({
        data()%>%
            distinct(学校,学段)%>%
            group_by(学段)%>%
            count(name="学校数量")
        
    })
    
    #分层抽样------------------------------------------------------------
    #分层抽样变量设置
    observeEvent(data(), {
        choices <- unique(names(data()))
        updateSelectizeInput(inputId = "stratified_var", choices = choices,selected=choices[1])
    })  
    
    #统计各分层人数
    
    stratified_num <- reactive({
        req(input$stratified_var)
        if(input$文档载入判定1=="使用测试文档"){
            dt <- read_xlsx("e:/rproject/shiny/data/2021崂山艺术项目/崂山区学生艺术测评模拟名单.xlsx")%>%
                group_by(across(all_of(input$stratified_var)))%>%
                count(name = "人数")
            
            
        }else{
            req(input$upload1)
            dt <- read_xlsx(paste0(input$upload1$datapath))%>%
                group_by(across(all_of(input$stratified_var)))%>%
                count(name = "人数")
        }
    })
    
    
    output$stratified_num <- renderTable({
     
        stratified_num()
            
        
    })
    
    #分层变量统计数据下载
    #抽样名单下载
    output$down_strata <- downloadHandler(
        filename = function() {
            paste("分层变量统计数据",".xlsx",sep="")
        },
        content = function(file) {
            write.xlsx(stratified_num(), file, row.names = FALSE)
        }
    )
    
    #分层抽样设置动态菜单
    output$抽样方法 <- renderUI({
        if (input$分层抽样方法 == "自定义法") {
            textInput("自定义抽样人数","自定义抽样人数向量",value=NULL)
        } else if (input$分层抽样方法 == "每层抽取固定人数" ) {
            numericInput("sc_sample_size", '抽样人数',value=30)
        } else {
            #tableOutput("size_tb")
        }
    
    })
    
    
    #分层抽样执行
    sampling_result <- reactive({
        项目抽样人数 <- input$sc_sample_size
        测评项目数量 <- 1
        分层变量 <- c("学校","性别")
        分层变量数量 <- length(分层变量)
        学校数量 <- length(unique(data()$学校))
        sc_sample_size <- rep(项目抽样人数/分层变量数量,学校数量*测评项目数量*分层变量数量)
        #sc_sample_size
        sub <- strata(data(),stratanames = 分层变量,size=size,method="srswor")
        
        getdata(data(),sub)%>%as_tibble()%>%
            rownames_to_column(var="序号")%>%
            mutate(序号=as.numeric(序号))
        
    })
    
    output$sampling_result <- renderTable({
        
        head(sampling_result())
        
    })
    
    #分层抽样结果表格下载
    output$down1 <- downloadHandler(
        filename = function() {
            paste("分层抽样结果表格", ".xlsx")
        },
        content = function(file) {
            write.xlsx(sampling_result(), file, row.names = FALSE)
        }
    )
    
    #获取抽样名单
    sample_list_raw <- reactive({
        sampling_result()%>%
            #删除抽样信息
            select(-(ID_unit:Stratum))
        
        
    })
    
    #抽样名单展示
    output$sampling_list <- renderTable({
        
        head(sample_list_raw())
        
    })
    
    #抽样名单下载
    output$down2 <- downloadHandler(
        filename = function() {
            paste("分层抽样学生名单",".xlsx",sep="")
        },
        content = function(file) {
            write.xlsx(sample_list_raw(), file, row.names = FALSE)
        }
    )
    
    #抽样名单数据处理--------------------------------------------------------
    #抽样名单文档载入
    data2 <- reactive({
        if(input$文档载入判定2=="使用测试文档"){
            dt <- read_xlsx("e:/rproject/shiny/data/2021崂山艺术项目/分层抽样学生名单.xlsx")
        }else{ 
            req(input$upload2)
            dt <- read_xlsx(paste0(input$upload2$datapath))
        }
    })
    
    #文档预览
    output$data2 <- renderTable({
        
        head(data2())
        
    })
    
    #各校名单分割及导出
    data_sc <- reactive({
        #req(input$upload2)
        #参数设置
        #待测学校数量
        sc_num <- length(unique(data2()$学校))
        #导出文档路径
        
        for(i in 1:sc_num) {
            #提取学校名称
            sc_name_list <- unique(data2()$学校)
            sc_name <- paste0(sc_name_list[[i]])
            new_art <- data2()%>%
                filter(学校==sc_name)
            
            #表格格式设置
            style <- createStyle(
                fontName = "华文中宋",halign = "center", valign = "center",
                textDecoration = "Bold",fontSize=10.5,wrapText=TRUE,fontColour="black"
            )
            
            openxlsx::write.xlsx(new_art, 
                                 file = paste(input$分割名单导出路径, 
                                              sc_name,"_抽测名单_艺术测评",".xlsx",sep=""),
                                 borders = "rows",tableStyle = "TableStyleMedium1",asTable = TRUE,
                                 headerStyle = style)
        }
    })
    #各校名单导出
    observeEvent(input$go, {
        data_sc()
        
    })
    
    #各校名单批量格式刷新
    observeEvent(input$go2, {
        
        input_path <- input$抽样名单导入路径
        
        output_path <- input$抽样名单格式刷新路径
        #获取文件名(注：应保证文件夹内仅有excel文档)
        files <- dir(path = input_path)
        
        #参数设置
        wb_name <-files 
        #工作表命名
        ws_name1 <- "sheet1"
        
        #格式设置
        字体 <- "华文中宋"
        表格样式 <- "TableStylemedium1"
        数字格式 <- "0"
        缩放比例 <- 78
        
        #运行格式刷
        #全部表格载入列表
        for (i in 1:length(files)){
            
            #对象表格载入
            dt<- 
                read_xlsx(paste0(input_path,wb_name[i]))
            
            #创建工作簿
            wb <- createWorkbook()
            
            #添加工作表
            addWorksheet(wb, ws_name1,tabColour="#4F81BD")
            
            #表头格式设置
            headerStyle <- createStyle(
                fgFill = "#4F81BD", halign = "CENTER", ,valign = "center",textDecoration = "Bold",
                border = "Bottom", fontColour = "white",fontSize = 14,borderColour = "black",
                fontName=字体,wrapText=TRUE)
            
            #主体格式设置
            bodyStyle <- createStyle(
                halign = "CENTER",valign = "center",border = "Bottom", fontColour = "black",fontSize = 12,
                fontName=字体,textDecoration = "Bold",borderColour = "black",wrapText=TRUE,
                borderStyle="medium")
            
            #数据写入
            writeDataTable(wb, 1, dt, startRow = 1, startCol = 1,headerStyle = headerStyle,tableStyle = 表格样式)
            
            #表格风格写入
            addStyle(wb, sheet = 1, bodyStyle, rows = 2:(nrow(dt)+1), cols = 1:(ncol(dt)), gridExpand = TRUE)
            
            #设置行高
            setRowHeights(wb, 1, rows = 2:(nrow(dt)+1), heights = 30) 
            
            #设置列宽
            setColWidths(wb,  1, cols = 1:(ncol(dt)), widths = 15)
            #数字样式设置
            numstyle <- createStyle(
                halign = "CENTER",valign = "center",border = "Bottom", fontColour = "black",fontSize = 12,
                fontName="华文中宋",textDecoration = "Bold",borderColour = "black",wrapText=TRUE,
                borderStyle="medium",numFmt = 数字格式)
            
            #添加数字风格
            addStyle(wb, 1, style = numstyle, rows = 1:(nrow(dt)+1), cols = 1:(ncol(dt)), gridExpand = TRUE)
            
            #打印设置
            #portrait为纵向，landscape为横向 
            pageSetup(wb, 1 , orientation="portrait",scale = 缩放比例, printTitleRows = 1, 
                      top=0.5 , bottom= 1, left= 0.64, right = 0.64, header=0.76, footer=0.76) 
            
            #数据导出
            saveWorkbook(wb,paste(output_path,wb_name[i]))
            
        }
    })
    
    #学生帐号生成----------------------------------------------------------------------
    #学校信息上传
    ac_data1 <- reactive({
        if(input$文档载入判定3=="使用测试文档"){
            dt <- read_xlsx("e:/rproject/shiny/data/2021崂山艺术项目/2021年崂山区参测学校信息表_20210913.xlsx")
            
        }else{
            req(input$upload3)
            dt <- read_xlsx(paste0(input$upload3$datapath))
        }
    })
    
    #待抽样数据显示
    output$ac_data1 <- renderTable({
        head(ac_data1())
    })
    
    #学生名单上传------------------------------------------------------------------
    ac_data2 <- reactive({
        if(input$文档载入判定4=="使用测试文档"){
            dt <- read_xlsx("e:/rproject/shiny/data/2021崂山艺术项目/分层抽样学生名单.xlsx")
            
        }else{
            req(input$upload4)
            dt <- read_xlsx(paste0(input$upload4$datapath))
        }
    })
    
    #待抽样数据显示
    output$ac_data2 <- renderTable({
        head(ac_data2())
    })
    
    ##提取抽样学校名单
    act <- reactive({
        #参数设置--------------------------
        #文档名称设置
        sc_info <- ac_data1()
        dt_raw <-  ac_data2()
        
        项目编号 <- input$password_3
        密码最小值 <- input$password_min
        密码最大值 <- input$password_max
        学校帐号位数 <- 6
        帐号结束序号 <- 35
        #学校帐号生成--------------------------
        set.seed(245)
        #建立学校帐号信息
        school_account <- sc_info%>%
            #建立项目编号
            mutate(项目编号=项目编号)%>%
            #生成密码
            mutate(密码=runif(nrow(sc_info),min=密码最小值,max=密码最大值))%>%
            #去除密码小数点
            mutate(密码=round(密码,0))%>%
            #学校id文本化
            mutate(学校ID=formatC(学校ID, flag = '0', width = 3))%>%
            #生成学校帐号
            unite(学校编号,项目编号,学校ID,sep="",remove=FALSE)%>%
            select(-项目编号)
        
        account <- dt_raw%>%
            #匹配学校帐号信息
            left_join(select(school_account,-序号),by=(c("学校","学段")))%>%
            rename(学生序号=序号)%>%
            #序号格式转换
            mutate(学生序号=formatC(学生序号, flag = '0', width = 2))%>%
            mutate(学生字段="学生")%>%
            unite("学生姓名",学生字段,学生序号,sep="",remove=FALSE)%>%
            select(-学生字段)%>%
            #生成学生账号
            unite("帐号",学校编号,学生序号,sep="")%>%
            arrange(帐号)
        
        #生成帐号导入信息表
        account_input <- account %>%
            #添加序号
            rownames_to_column(var="ID")%>%
            #增加新变量
            mutate(身份证=帐号,性别=0,入学年级=0,班级=0,民族=0,
                      出生日期=0,户口所在地="青岛市")%>%
            #重命名
            rename(学籍号码=帐号,学校名称=学校)%>%
            #二次格式调整
            select(ID,学生姓名,学籍号码,身份证,性别,入学年级,班级,民族,出生日期,户口所在地,
                   密码,学校ID,学校名称)
    })
    
    #学生帐号信息表预览
    output$act <- renderTable({
        head(act())
    })
    
    #分层抽样结果表格下载
    output$account_down <- downloadHandler(
        filename = function() {
            paste("学生帐号信息表", ".xlsx")
        },
        content = function(file) {
            write.xlsx(act(), file, row.names = FALSE)
        }
    )
}