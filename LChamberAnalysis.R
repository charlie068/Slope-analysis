### Initialise ####
## set Folder
setwd("~/Teagasc/31082018")

##Set libraries
if (TRUE){
  library(data.table)
  library(readxl)
  library(ggplot2)
  library(plotrix)
  library("writexl")
  library("Hmisc") 
  library("Rmisc")
  library("scales")
  library("multcompView")
  library("officer")
  library("magrittr")
  library("reshape2")
  library("rvg")
  library("dplyr")
  library("ggpubr")
  library("moments")
  library("rcompanion")
  library("officer")
  library("magrittr")
  #library("rvg")
}

##Load excel File
growth<-read_excel("GROWTH MEASURMENTS020818.xlsx")
growth<-data.table(growth)
##melt the File and Format it in data.table
growth<-melt(growth, id=c('Variety', 'Date', 'Row', 'Chamber'))
names(growth)[6] <- "Height"
growth<-data.table(growth)

growth<-growth[Variety!="Bar"]

growth<-na.omit(growth)

### Growth Graph ####

#remove outliers
growth_out <- growth[,.SD[Height < (quantile(Height,0.75)+1.5*IQR(Height)) & Height > (quantile(Height,0.25)-1.5*IQR(Height))],by=c("Variety", "Chamber", "Date")]
sub_growth_L<-growth_out[Chamber=='L',]

##get average per time a per variety
growth_average<- as.data.table(growth_out[, .('Mean'=mean(Height, na.rm = TRUE),'Std'=std.error(Height, na.rm=TRUE)), by = list(Variety, Date, Chamber)])

##plot means as bar graph
ggplot(data=growth_average[Variety!='Bar'&Chamber!='R'], aes(x=Date, y=Mean), group=Chamber) + geom_line(aes(linetype=Chamber)) + geom_point() + 
  facet_grid(Chamber ~ Variety) +
  geom_errorbar(data=growth_average[Variety!='Bar'&Chamber!='R'], mapping=aes(x=Date, ymin=Mean+Std, ymax=Mean-Std))+
  theme(axis.text.x = element_text(angle = 90, hjust = 1)) +
  ggtitle("Grass Height Average") +
  xlab("Date") + ylab("Mean(Height)")

#Write excel file with growth average
#write_excel_csv(growth_average, 'growthAverage.csv')

##plot means as bar graph in one graph (two colours)
ggplot(data=growth_average[Variety!='Bar'&Chamber!='R'], aes(x=Date, y=Mean)) + geom_line(size=1,aes(colour=Chamber)) +
  geom_point(shape=1,size=2,aes(colour=Chamber)) + 
  facet_grid(. ~ Variety) +
  geom_errorbar(data=growth_average[Variety!='Bar'&Chamber!='R'], mapping=aes(x=Date, ymin=Mean+Std, ymax=Mean-Std, colour=Chamber))+
  theme(axis.text.x = element_text(angle = 90, hjust = 1)) +
  ggtitle("Grass Height Average") +
  xlab("Date") + ylab("Mean(Height)")


##plot means as line graph (one colour) on one plot
ggplot(data=growth_average[Variety!='Bar'&Chamber!='R'], aes(x=Date, y=Mean)) + geom_line(size=1, aes(linetype=Chamber)) +
  geom_point(shape=1,size=2) + 
  facet_grid(. ~ Variety) +
  geom_errorbar(data=growth_average[Variety!='Bar'&Chamber!='R'], mapping=aes(x=Date, ymin=Mean+Std, ymax=Mean-Std))+
  theme(axis.text.x = element_text(angle = 90, hjust = 1)) +
  ggtitle("Grass Height Average") +
  xlab("Date") + ylab("Mean(Height)")



##plot means as bar graph in one graph (two colours)
ggtile<-paste("Grass Height Chamber L")
growth_chM<-growth_average[Variety!='Bar'&Chamber=="L"]
ggplot(growth_chM, aes(x=Date, y=Mean)) + geom_line(size=1,aes(colour=Variety)) +
  geom_point(shape=1,size=2,aes(colour=Variety)) + 
  #facet_grid(. ~ Variety) +
  geom_errorbar(growth_chM, mapping=aes(x=Date, ymin=Mean+Std, ymax=Mean-Std, colour=Variety),width=0.2)+
  theme(axis.text.x = element_text(angle = 90, hjust = 1)) +
  ggtitle(ggtile) +
  xlab("Date") + ylab("Mean(Height)")+
  theme_bw(base_size=20) +
  theme(plot.title = element_text(vjust=-0.6,hjust = 0.5, face="bold", size=20))+
  theme(panel.grid.major = element_blank(), panel.grid.minor = element_blank())+
  scale_y_continuous(limits=c(0,35)) 


#####          TWO WAY ANOVA   ###### BETWEEN 2 DATES and    ###############


sub_growth_out<-growth_out[(Variety!="Bar")&!as.character(Date)=="2018-04-09"&Chamber!="R", ] ##remove bar and chamber R from outliers removed table


list_date<-unique(sub_growth_out$Date)

###remove outliers 
#sub_growth_out <- sub_growth1[,.SD[Height < (quantile(Height,0.75)+1.5*IQR(Height)) & Height > (quantile(Height,0.25)-1.5*IQR(Height))],by=c("Variety", "Chamber", "Date")]


##Create table
date<-unique(sub_growth_out$Date)
ldate<-length(date)

#### RATIO ANALYSIS BETWEEN TWO DATES  
#choose dates
doc <- read_pptx()  # create pptx  
for (a in 1:(ldate-1)){
  for (b in 2:ldate){
    
    
    date1<-a
    date2<-b
    date_chosen1<-as.character(date[date1])
    date_chosen2<-as.character(date[date2])
    
    #define data.table 
    resume <- data.table(Variety1=character(),
                         Variety2=character(),
                         value=numeric(), 
                         significance=numeric(), 
                         stringsAsFactors=FALSE) 
    ###Choose chamber sub_growth_out
    sub_growth_date<-sub_growth_out[(as.character(Date)==date_chosen1|as.character(Date)==date_chosen2)&Chamber=="L"]  
    list_variety<-unique(sub_growth_date$Variety)
    lenvar=length(list_variety)   #number of variety
    #do two way between each variety
    
    ggtitl<-paste(date_chosen1,date_chosen2,sep="-")
    print(ggtitl)
    
    for (i in 1:lenvar){
      for (j in 1:lenvar){
        if (i!=j){
          print(paste(list_variety[i],list_variety[j],sep="/"))
          sub_growth_2way<-sub_growth_date[(Variety==list_variety[i]|Variety==list_variety[j])]
          aov_growth_2way = aov(Height~Variety*Date,sub_growth_2way)  #do the analysis of variance
          resume <-rbind(resume,list(as.character(list_variety[i]),as.character(list_variety[j]),as.numeric(unlist(summary(aov_growth_2way)[[1]][5][1])[3]),(as.numeric(unlist(summary(aov_growth_2way)[[1]][5][1])[3]))))
        }
      }
      
    }
    #View(resume)
    #summary(aov_growth_2way)
    #list((as.character(paste(list_variety[i-1],list_variety[j],sep="_"))),as.numeric(unlist(summary(aov_growth_2way)[[1]][5][1])[3]),(unlist(summary(aov_growth_2way)[[1]][5][1])[3])<0.05)
    ##write pvalues of 2WAY-ANOVA in xl
    write_xlsx(resume, paste(ggtitl,"L_twoway.xlsx",sep="_"  ))
    resume[,value:=NULL]
    resumeC<-cast(resume, Variety1~Variety2)
    write_xlsx(resumeC, paste(ggtitl,"L_twowayC1640.xlsx",sep="_"  ))
    
    #Make Graph
    # Calculate error bars
    tgc <- summarySE(sub_growth_date, measurevar="Height", groupvars=c("Variety","Date"))
    doc<-doc %>% 
      add_slide(layout = "Title and Content", master = "Office Theme") # add slide
    doc<-doc %>% 
      ph_with_text(type = "title", str = ggtitl) # add title
    
    #Make graph
    p<-ggplot(data=tgc, aes(x=Variety, y=Height, fill=as.factor(Date), group=as.factor(Date))) +
      geom_bar(colour="black", stat="identity", position = "dodge2") +
      geom_errorbar(aes(ymax = Height + se, ymin = Height-se, group=as.factor(Date)),
                    position = position_dodge(width = 0.9), width = 0.25)+
      ggtitle(ggtitl)+
      theme(panel.grid.major = element_blank(), panel.grid.minor = element_blank()) +
      labs(x = 'Variety', y = 'Growth (cm)') +
      theme(axis.text=element_text(size=12),
            axis.title=element_text(size=14,face="bold"))+
      #theme(title = element_text(face="bold", size=15)) +
      theme(plot.title = element_text(size=20,face="bold"))+
      scale_fill_brewer(palette="Set1")+
      scale_y_continuous(limits=c(0,20))+
      theme_bw(base_size=20) +
      theme(plot.title = element_text(vjust=-0.6,hjust = 0.5, face="bold", size=20))+
      theme(panel.grid.major = element_blank(), panel.grid.minor = element_blank()) 
    print(p)
    # scale_fill_manual(values=c("#999999", "#E69F00", "#56B4E9")) #+ 
    #theme(legend.position="none")
    
    doc <- doc %>% ph_with_vg_at(code= print(p),left=1, top=2, width=6,height=4 )
    #doc <- ph_with_vg_at(doc, code = barplot(1:5, col = 2:6),
    #    left = 1, top = 2, width = 6, height = 4)
    
    
  }}


print(doc, target="twoway_date_variety082018_L.pptx" )
### END TWO WAY per date and variety




### STATS 1way anova ####



sub_growth_out<-growth_out[(Variety!="Bar")&!as.character(Date)=="2018-04-09"&Chamber!="R", ] ##remove bar and chamber R from outliers removed table


list_date<-unique(sub_growth_out$Date)

###remove outliers  (already remvoed in the beginning)
#sub_growth_out <- sub_growth1[,.SD[Height < (quantile(Height,0.75)+1.5*IQR(Height)) & Height > (quantile(Height,0.25)-1.5*IQR(Height))],by=c("Variety", "Chamber", "Date")]


##Create table
date<-unique(sub_growth_out$Date)
ldate<-length(date)

#####  ANALYSIS BETWEEN TWO DATES  
#choose dates
#  doc =pptx() # create pptx     


list_variety<-unique(sub_growth_date$Variety)
lenvar=length(list_variety)   #number of variety



for (a in 1:(ldate-1)){
  b <- (a+1)
  
  
  date1<-a
  date2<-b
  date_chosen1<-as.character(date[date1])
  date_chosen2<-as.character(date[date2])
  
  ###Choose Date
  sub_growth_date<-sub_growth_out[(as.character(Date)==date_chosen1|as.character(Date)==date_chosen2)&Chamber=="M"]  
  
  #do two way between each variety
  
  ggtitl<-paste(date_chosen1,date_chosen2,sep="_")
  print(ggtitl)
  
  ### ANOVA ONE WAY per see   ####
  model=lm(sub_growth_date$Height~sub_growth_date$Date)
  
  #ANOVA=aov(model)
  #summary(ANOVA)
  
  
  # Tukey test to study each pair of treatment :
  # TUKEY <- TukeyHSD(x=ANOVA, 'sub_growth_date$Variety', conf.level=0.95)
  #(TUKEY)
  # View(resume)
  ##write pvalues of 2WAY-ANOVA in xl
  
  ggtitl<-paste(date_chosen1,date_chosen2,sep="-")
  print(ggtitl)
  #do ttest between each date for each variet
  
  
  resume <- data.table(Interaction=character(),
                       value=numeric(), 
                       significance=logical(), 
                       stringsAsFactors=FALSE)
  
  
  for (v in 1:lenvar){
    
    print(list_variety[v])
    results_ttest<-t.test(sub_growth_date[Variety==list_variety[v]&as.character(Date)==date_chosen1,Height],sub_growth_date[Variety==list_variety[v]&as.character(Date)==date_chosen2,Height])
    results_ttest[3]
    resume<-rbind(resume,list(as.character(list_variety[v]),as.numeric(results_ttest[3]),(results_ttest[3])<0.05))
  }
  
  
  write_xlsx(resume, paste(ggtitl,"ttest.xlsx",sep="_"  ))
}  


#### GROWTH between two dates  ####

######### STATS 1way anova ##


sub_growth_out<-growth_out[(Variety!="Bar")&!as.character(Date)=="2018-04-09"&Chamber!="R", ] ##remove bar and chamber R from outliers removed table



##Create table
date<-unique(sub_growth_out$Date)
ldate<-length(date)


#choose dates



list_variety<-unique(sub_growth_date$Variety)
lenvar=length(list_variety)   #number of variety

#declare ppt 
#doc =pptx() # create pptx    
#Loop per date
for (a in 1:(ldate-1)){
  for (b in 2:ldate){
    
    
    date1<-a
    date2<-b
    date_chosen1<-as.character(date[date1])
    date_chosen2<-as.character(date[date2])
    
    ###Choose Date
    sub_growth_date<-sub_growth_out[(as.character(Date)==date_chosen1|as.character(Date)==date_chosen2)&Chamber=="M"]  
    
    #do two way between each variety
    
    ggtitl<-paste(date_chosen1,date_chosen2,sep="_")
    print(ggtitl)
    
    
    
    resume <- data.table(Interaction=character(),
                         value=numeric(), 
                         stringsAsFactors=FALSE)
    for (v in 1:lenvar){
      
      print(list_variety[v])
      growthratediff<-sub_growth_date[Variety==list_variety[v]&as.character(Date)==date_chosen2,mean(Height)]-sub_growth_date[Variety==list_variety[v]&as.character(Date)==date_chosen1,mean(Height)]
      growthratediff<-growthratediff/(sub_growth_date[Variety==list_variety[v]&as.character(Date)==date_chosen1,mean(Height)])
      growthratediff<-(growthratediff/as.numeric(date[date2]-date[date1]))*100
      resume<-rbind(resume,list(as.character(list_variety[v]),as.numeric(growthratediff)))
    }
    
    ##plot means as bar graph in one graph (two colours)
    ggtile<-paste("Difference of Height Average between two dates")
    growth_diff_average<-resume
    names(growth_diff_average)=c("Variety","Height")
    p<-ggplot(growth_diff_average, aes(x=Variety, y=Height, fill=Variety)) +
      geom_bar(stat="identity") +
      ggtitle(ggtitl)+
      theme(panel.grid.major = element_blank(), panel.grid.minor = element_blank()) +
      labs(x = 'Variety', y = 'Growth (% per day)') +
      theme(axis.text=element_text(size=12),
            axis.title=element_text(size=14,face="bold"))+
      #theme(title = element_text(face="bold", size=15)) +
      theme(plot.title = element_text(size=20,face="bold"))+
      #scale_fill_brewer(palette="Set1")+
      scale_y_continuous(limits=c(0,5))+
      theme_bw(base_size=20) +
      theme(plot.title = element_text(vjust=-0.6,hjust = 0.5, face="bold", size=20))+
      theme(panel.grid.major = element_blank(), panel.grid.minor = element_blank()) 
    print(p)
    
    
    #doc<-addSlide(doc,"Title and Content") # add slide
    #doc<-addTitle(doc,ggtitl) # add title
    #doc <- addPlot(doc, fun=function() print(p) ,vector.graphic =TRUE )  # add plot
    
    
  }}
#writeDoc(doc, "growth_diff3005b.pptx" )



### END of Graph based on growth diff




