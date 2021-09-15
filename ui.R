## ui.R ##
library(shinydashboard)
library(openxlsx)
library(readxl)
library(tidyverse)
library(sampling)


#函数设置
#文档载入判定
radio1 <- function(id){
 radioButtons(id,label="是否使用测试文档",choices=c("使用测试文档","使用上传数据"),selected="使用测试文档")
}



#UI设置
header <- dashboardHeader(title = "测评对象抽样")

sidebar <- dashboardSidebar(
    tags$head(tags$style(HTML('
         .main-header .logo {
        font-family: "Georgia", Times, "Times New Roman", serif;
        font-weight: bold;
        font-size: 24px;
      }
      '))),
    sidebarMenu(
        
        menuItem("待抽样数据上传", tabName = "sampling", icon = icon("th")),
        menuItem("分层抽样", tabName = "分层抽样", icon = icon("th")),
        menuItem("抽样名单数据处理", tabName = "抽样名单处理", icon = icon("th")),
        menuItem("学生帐号生成", tabName = "学生帐号生成", icon = icon("th")))
)

body <- dashboardBody(
    # Also add some custom CSS to make the title background area the same
    # color as the rest of the header.
    tags$head(tags$style(HTML('
         .main-header .logo {
        font-family: "Georgia", Times, "Times New Roman", serif;
        font-weight: bold;
        font-size: 24px;
      }
      '))),
    tabItems(
      
        # First tab content
        tabItem(tabName = "sampling",
                fluidRow(
                    valueBox("学段数量", 0, icon = icon("list"),color = "blue"),
                    valueBox("学生数量", 0, icon = icon("list"),color = "purple"),
                    valueBox("学校数量", 0, icon = icon("list"),color = "yellow"),
                    box(solidHeader=TRUE,status="primary",collapsible = TRUE,
                        title="是否使用默认文档",
                        radio1("文档载入判定1")
                    ),
                    box(solidHeader=TRUE,status="primary",collapsible = TRUE,
                        title="待抽样数据上传",
                        fileInput("upload1",NULL,buttonLabel = "文件上传", accept = ".xlsx")
                    ),
                    
                    box(solidHeader=TRUE,status="primary",collapsible = TRUE,
                        title="原始数据展示",
                        tableOutput("data_num"),
                        tableOutput("dt")
                    ),
                    
                    
                    box(solidHeader=TRUE,status="primary",collapsible = TRUE,
                        title = "待抽样学校名单",
                        tableOutput("school_list")
                    ),
                    
                    box(solidHeader=TRUE,status="primary",collapsible = TRUE,
                        title = "待抽样学校数量统计",
                        tableOutput("school_num")
                    )
                    #fluidRow  
                )
                #tabItem        
        ),
        # Second tab content
        tabItem(tabName = "分层抽样",
                fluidRow(
                    box(solidHeader=TRUE,status="primary",collapsible = TRUE,
                        title = "分层抽样",
                        selectizeInput("stratified_var", '分层变量', choices = NULL,multiple = TRUE),
                        #下载分层抽样表格
                        downloadButton("down_strata", "分层变量统计数据下载"),
                        #统计各校分层人数
                        tableOutput("stratified_num"),
                        
                    ),
                    
                    box(solidHeader=TRUE,status="primary",collapsible = TRUE,
                        title = "分层抽样人数设置",
                        #分层抽样方法选择
                        radioButtons("分层抽样方法",label="分层抽样方法",
                                     choices=c("自定义法","每层抽取固定人数","后端载入"),selected="自定义法"),
                        #抽样人数设置
                        uiOutput("抽样方法"),
                        #textInput("自定义抽样人数","自定义抽样人数向量",value=NULL),
                        #numericInput("sc_sample_size", '抽样人数',value=30),
                        
                        
                    ),
                    
                    box(solidHeader=TRUE,status="primary",collapsible = TRUE,
                        title = "运行分层抽样",
                        h2("注:暂保留后端设置并运行"),
                        actionButton("samp", '运行分层抽样'),
                        downloadButton("down1", "抽样结果下载"),
                        tableOutput("sampling_result"),
                        
                        
                    ),
                    
                    box(solidHeader=TRUE,status="primary",collapsible = TRUE,
                        title = "抽样名单展示及下载",
                        tableOutput("sampling_list"),
                        downloadButton("down2", "抽样名单下载"),
                        
                    ),
                    #fluidRow
                )
                #tabItem        
        ),
        tabItem(tabName = "抽样名单处理",
                fluidRow(
                    box(solidHeader=TRUE,status="primary",collapsible = TRUE,
                        title="是否使用默认文档",
                        radio1("文档载入判定2"),
                    ),
                    
                    box(solidHeader=TRUE,status="primary",collapsible = TRUE,
                        title = "已抽样名单数据上传",
                        fileInput("upload2",NULL,buttonLabel = "文件上传", accept = ".xlsx"),
                        h2("文档预览"),
                        tableOutput("data2"),
                    ),
                    box(solidHeader=TRUE,status="primary",collapsible = TRUE,
                        title = "各校名单分割及导出",
                        #导出文档路径
                        textInput("分割名单导出路径","导出路径设置",value="e:/rproject/shiny/data/2021崂山艺术项目/各校学生名单/"),
                        actionButton("go", "名单分割及导出"),
                        #downloadButton("down3", "各校抽样名单下载"),
                    ),
                    box(solidHeader=TRUE,status="primary",collapsible = TRUE,
                        title = "抽样名单批量格式刷",
                        #导出文档路径
                        textInput("抽样名单导入路径","导入路径设置",
                                  value="e:/rproject/shiny/data/2021崂山艺术项目/各校学生名单/"),
                        textInput("抽样名单格式刷新路径","导出路径设置",
                                  value="e:/rproject/shiny/data/2021崂山艺术项目/各校学生名单_格式刷新/"),
                        actionButton("go2", "抽样名单批量格式刷新"),
                    ),
                    #fluidRow
                )
                
                #tabItem       
        ),
        tabItem(tabName = "学生帐号生成",
                fluidRow(
                    box(solidHeader=TRUE,status="primary",collapsible = TRUE,
                        title = "学校信息表上传",
                        fileInput("upload3",NULL,buttonLabel = "文档上传", accept = ".xlsx"),
                        radio1("文档载入判定3"),
                        #数据预览
                        tableOutput("ac_data1"),
                    ),
                    fluidRow(
                        box(solidHeader=TRUE,status="primary",collapsible = TRUE,
                            title = "学生名单上传",
                            fileInput("upload4",NULL,buttonLabel = "文件上传", accept = ".xlsx"),
                            radio1("文档载入判定4"),
                            #数据预览
                            tableOutput("ac_data2"),
                        ),
                        box(solidHeader=TRUE,status="primary",collapsible = TRUE,
                            title = "学生帐号生成",
                            numericInput("password_3", '项目编号',value=211),
                            numericInput("password_min", '密码最小值',value=100000),
                            numericInput("password_max", '密码最大值',value=999999),
                            #学生帐号信息表下载
                            downloadButton("account_down", "学生帐号信息表"),
                            #数据预览
                            tableOutput("act"),
                        )
                        
                    ),
                    #fluidRow
                )
                #tabItem
        )
        
        #tabItems
    )
)
ui <- dashboardPage(skin = "blue",header,sidebar,body)