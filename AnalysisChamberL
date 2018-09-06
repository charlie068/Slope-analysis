### Initialise ####
## set Folder
setwd("~/Teagasc/31082018")

##Set libraries
#install.packages("ggpubr")
if (TRUE){
  library(data.table)
  library(readxl)
  library(ggplot2)
  library("ggpubr")
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
  library('reshape2')
  library('reshape')
  library('FSA')
  library('car')
}

##Load excel File
growth<-read_excel("GROWTH MEASURMENTS020818.xlsx")

growth<-data.table(growth)
##melt the File and Format it in data.table
growth<-melt(growth, id=c('Variety', 'Date', 'Row', 'Chamber'))
names(growth)[6] <- "Height"
growth<-data.table(growth)

growth<-growth[Chamber!='R']

growth<-na.omit(growth)

### Growth Graph ####'

#remove outliers
growth_out <- growth[,.SD[Height < (quantile(Height,0.75)+1.5*IQR(Height)) & Height > (quantile(Height,0.25)-1.5*IQR(Height))],by=c("Variety", "Chamber", "Date")]
growth_out2<-growth_out[, "days" := round(as.numeric(difftime((growth_out$Date),strptime("2018-02-06", format = "%Y-%m-%d"), units = ("days"))))]


sub_growth_L<-growth_out2[Chamber=='L',]
sub_growth_M<-growth_out2[Chamber=='M',]

##get average per time a per variety
growth_average<- as.data.table(growth_out2[, .('Mean'=mean(Height, na.rm = TRUE),'Std'=std.error(Height, na.rm=TRUE)), by = list(Variety, Date,days, Chamber)])
growth_average_L<-growth_average[Chamber=='L']
growth_average_M<-growth_average[Chamber=='M']


growth_dummy<-sub_growth_L

#growth_dummy$vardum=0
#unique(sub_growth_L$Variety)
growth_dummy$vardum <-ifelse(sub_growth_L$Variety=='Again', 0,
                             ifelse(sub_growth_L$Variety=='Agreen', 1, 
                                    ifelse(sub_growth_L$Variety=='Gl', 2,
                                           ifelse(sub_growth_L$Variety=='Jan', 3,
                                                  ifelse( sub_growth_L$Variety=='Lp', 4,
                                                          ifelse(sub_growth_L$Variety=='Ma', 5,
                                                                 ifelse(sub_growth_L$Variety=='Mi', 6,
                                                                        ifelse(sub_growth_L$Variety=='Mu', 7,
                                                                               ifelse(sub_growth_L$Variety=='Ro', 8,
                                                                                      ifelse(sub_growth_L$Variety=='Tw', 9,
                                                                                             ifelse(sub_growth_L$Variety=='Bar', 10,-1)))))))))))
#growth_dummy<-growth_dummy[, "days" := as.numeric(difftime(strptime(as.character(growth_dummy$Date), format = "%Y-%m-%d"),strptime("2018-02-06", format = "%Y-%m-%d"), units = ("days")))] 
#write_xlsx(growth_dummy, "growthLforRb.xlsx")
########## end initialisation   #######'

#######################################'
##### ANALYSES ########################
#######################################'


##plot last time point Chamber L
ggplot(growth_average_L[as.character(Date)=="2018-06-15"],aes(Variety,Mean,fill=Variety)) + 
  geom_bar( stat = "summary", fun.y = "mean")+
  geom_errorbar(data=growth_average_L[as.character(Date)=="2018-06-15"], mapping=aes(x=Variety, ymin=Mean+Std, ymax=Mean-Std))+
  theme(axis.text.x = element_text(angle = 90, hjust = 1)) +
  ggtitle("Grass Height Average L") +
  xlab("Date") + ylab("Mean(Height)")


##plot means as bar graph
ggplot(data=growth_average_L, aes(x=Date, y=Mean), group=Chamber) + geom_line(aes(linetype=Chamber)) + geom_point() + 
  facet_grid(Chamber ~ Variety) +
  geom_errorbar(data=growth_average_L, mapping=aes(x=Date, ymin=Mean+Std, ymax=Mean-Std))+
  theme(axis.text.x = element_text(angle = 90, hjust = 1)) +
  ggtitle("Grass Height Average") +
  xlab("Date") + ylab("Mean(Height)")

ggplot(data=growth_average_L, aes(x=days, y=Mean), group=Chamber) + geom_line(aes(linetype=Chamber)) + geom_point() + 
  facet_grid(Chamber ~ Variety) +
  geom_errorbar(data=growth_average_L, mapping=aes(x=days, ymin=Mean+Std, ymax=Mean-Std))+
  theme(axis.text.x = element_text(angle = 90, hjust = 1)) +
  ggtitle("Grass Height Average") +
  xlab("Date") + ylab("Mean(Height)")

##plot means as bar graph
ggplot(data=growth_average_L, aes(x=days, y=Mean), group=Chamber) + geom_line(aes(linetype=Variety)) + geom_point() + 
  #facet_grid(Chamber ~ Variety) +
  geom_errorbar(data=growth_average_L, mapping=aes(x=days, ymin=Mean+Std, ymax=Mean-Std))+
  theme(axis.text.x = element_text(angle = 90, hjust = 1)) +
  ggtitle("Grass Height Average") +
  xlab("Date") + ylab("Mean(Height)")


#####################################'
#### One-way Anova per date #########
#####################################'

### FUNCTION for labels ###
generate_label_df <- function(TUKEY, variable){
  
  # Extract labels and factor levels from Tukey post-hoc 
  Tukey.levels <- TUKEY[[variable]][,4]
  Tukey.labels <- data.frame(multcompLetters(Tukey.levels)['Letters'])
  
  #I need to put the labels in the same order as in the boxplot :
  Tukey.labels$treatment=rownames(Tukey.labels)
  Tukey.labels=Tukey.labels[order(Tukey.labels$treatment) , ]
  return(Tukey.labels)
}

date<-unique(sub_growth_L$days)
ldate<-length(date)


#choose dates
doc <- read_pptx()  # create pptx  

for (a in 1:(ldate)){

    date1<-a
    date_chosen1<-as.character(date[date1])
    doc<-doc %>% 
      add_slide(layout = "Title and Content", master = "Office Theme") # add slide
    doc<-doc %>% 
      ph_with_text(type = "title", str = date_chosen1) # add title

 
    #Change data.frame here
      data<-sub_growth_L[days==date_chosen1] #growth_diff
    
    ### ANOVA ONE WAY
      model=lm(data$Height~data$Variety)
      #summary(model)
      ANOVA=aov(model)
      summary(ANOVA)
    #test residuals
      qqnorm(model$res)
      plot(model$fitted,model$res,xlab="Fitted",ylab="Residuals")
      hist(residuals(model), col="darkgray")
    #sub_growth_L_aver<-sub_growth_L[,.("Mean"=mean(Height),na.rm = TRUE),by=c("Date","Variety")]
    #ggplot(sub_growth_L_aver) +
    #  aes(x = Date, y = Mean, color = Variety) +
    #  geom_line(aes(group = Variety)) +
    #  geom_point()
    
    # Tukey test to study each pair of treatment :
    TUKEY <- TukeyHSD(x=ANOVA, 'data$Variety', conf.level=0.95)
    (TUKEY)
    #generate labels using function
    labels<-generate_label_df(TUKEY , "data$Variety")
    
    names(labels)<-c('Letters','Variety')#rename columns for merging
    
    #yvalue<-aggregate(Height~Variety, data=data, mean)# obtain letter position for y axis using means
    yvalue<-data[,quantile(Height,0.75),by=Variety]
    final<-merge(labels,yvalue) #merge dataframes
    #final[,3]<-final[,3]*1.1
    names(final)[3]<-"aver"
    sub_titl<-date_chosen1
    sub_titl<-"One-Way ANOVA"
    
    
    p1<-ggplot(data, aes(x = Variety, y = Height)) +
      geom_blank() +
      geom_boxplot()+
      theme_bw() +
      theme(panel.grid.major = element_blank(), panel.grid.minor = element_blank()) +
      labs(x = 'Variety', y = 'Growth (cm)') +
      ggtitle(sub_titl)+
      #ggtitle(expression(atop(bold(paste(sub_titl)), atop(italic("(Height)"), "")))) +
      theme(plot.title = element_text(hjust = 0.5, face='bold')) +
      #annotate(geom = "rect", xmin = 1.5, xmax = 4.5, ymin = -Inf, ymax = Inf, alpha = 0.6, fill = "grey90") +
      geom_boxplot(fill = '#99ff99', stat = "boxplot") +
      geom_text(data = final, aes(x = Variety, y = aver*1.1, label = Letters),vjust=0,hjust=0, nudge_x=0.1, size=5) +
      #geom_vline(aes(xintercept=4.5), linetype="dashed") +
      theme(plot.title = element_text(vjust=-0.6, face="bold", size=20)) 
    
    #Bar plot
    yvalue_mean<-data[,.("aver"=mean(Height),"se"=std.error(Height)),by=Variety]
    setkey(yvalue_mean,Variety)
    labels<-data.table(labels)
    labels<-labels[order(Variety)]
    setkey(labels,Variety)
    setcolorder(labels, c("Variety", "Letters"))
    final2<-labels[yvalue_mean,all=TRUE] #merge dataframes Full outer join
    
    
   p2<-ggplot(final2, aes(x = Variety, y = aver, fill=Variety)) +
      geom_blank() +
      theme_bw() +
      theme(panel.grid.major = element_blank(), panel.grid.minor = element_blank()) +
      labs(x = 'Variety', y = 'Growth ratio') +
      ggtitle(sub_titl)+
      geom_bar(stat = "identity")+
      geom_text(data = final2, aes(x = Variety, y = aver+se, label = Letters, vjust=0, hjust=0.5),nudge_y=1,size=5) +
      theme(plot.title = element_text(vjust=-0.6, face="bold", size=20))+
      theme(plot.title = element_text(hjust = 0.5, face='bold')) +
      geom_errorbar(aes(Variety,ymin=aver-se,ymax=aver+se),width=0.5)+
      scale_y_continuous(limits=c(0,20))
   
   doc <- doc %>% ph_with_vg_at(code= print(p1),left=0, top=1.5, width=9,height=3 )
   doc <- doc %>% ph_with_vg_at(code= print(p2),left=0, top=4.5, width=9,height=3 )
}

print(doc, target="onewayanova_Lb.pptx" )


#####################################'
#### Two-way Anova for growth #######
#####################################'
library(emmeans)
library(car)
library(multcomp)
library(lme4)
library(lmerTest)
#create a datatable with dummy variable for each speecies
data<-growth_dummy
interaction.plot(x.factor = data$days, 
                 trace.factor = data$Variety,
                 response = data$Height)

xyplot(Height ~ days | Variety, groups = ~ vardum, data = data,
       type = "o", layout=c(4, 4))
#for our data
#library(optimx)
grass.lmer <- lmer((Height) ~ Variety * days +(1|vardum), data = data, REML = FALSE, 
                   control = lmerControl(optimizer ="Nelder_Mead"))

grass.lst <- lstrends (grass.lmer, ~ Variety, var = "days")
cld (grass.lst)

slope.mdl <- lstrends (grass.lmer, ~ Variety, var = "days")
summary(slope.mdl)

grass.mdl<-lm(Height ~ Variety * days +(1|vardum), data = data)
summary(grass.mdl)

##plot means as bar graph 
ggplot(data=growth_average_L, aes(x=days, y=Mean)) + 
  theme_bw() +
  #geom_point(shape=1,size=2) + 
  #facet_grid(. ~ Variety) +
  stat_smooth(aes(colour=Variety), method = "lm", se = FALSE)+  
  theme(axis.text.x = element_text(angle = 90, hjust = 1)) +
  ggtitle("Grass Height Average") +
  xlab("Days") + ylab("Mean(Height)")

slope.mdl[1]
str(slope.mdl)
a<-data.table(slope.mdl[1])
str(a)














