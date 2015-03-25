#QUESTION 1:::::::::::::::::::::::::::::::::::
#---------------------------------------------
#Install a package name "xlsx" for working with excel files
install.packages("xlsx")
require(xlsx)
filePath <- "G:\\fastModel\\tournament.xlsx" #specify the path to the excel file

# read all spreadsheets to different data frames
fmChrono <- read.xlsx(filePath, sheetName = "CHRONO", startRow = 6)
fmAlpha <- read.xlsx(filePath, sheetName = "ALPHA")
fm17U <- read.xlsx(filePath, sheetName = "17U")
fm16U <- read.xlsx(filePath, sheetName = "16U")
fm15U <- read.xlsx(filePath, sheetName = "15U")
fm14U <- read.xlsx(filePath, sheetName = "14U")

#In the first spreadsheet CHRONO, we have team names with scores, so I am 
#splitting this columns to get names separate from scores 
fmChrono1 <- data.frame(do.call('rbind', strsplit(as.character(fmChrono$HOME),' - ',fixed=TRUE)))
fmChrono2 <- data.frame(do.call('rbind', strsplit(as.character(fmChrono$AWAY),' - ',fixed=TRUE)))

#deleting all unwanted extra columns 
fm14U$NA. <- NULL
fm14U$NA..1 <- NULL
fm16U$NA. <- NULL
fm16U$NA..1 <- NULL
fm17U$NA. <- NULL
fm17U$NA..1 <- NULL
fm17U$NA..2 <- NULL
fm17U$NA..3 <- NULL
fm17U$NA..4 <- NULL
fm17U$NA..5 <- NULL

#combining U14, U15, U16, U17 spread sheets for 
#getting all teams in one data frame 
new_df <- rbind(fm14U, fm15U, fm16U, fm17U)

#combining all sheets to get participating team list
team_names <- rbind(new_df["HOME"], 
                    setNames(new_df["AWAY"],names(new_df["HOME"])), 
                    setNames(fmChrono1["X1"],names(new_df["HOME"])), 
                    setNames(fmChrono2["X1"],names(new_df["HOME"])),
                    setNames(fmAlpha["HOME.TEAM.BOLDED"],names(new_df["HOME"])),
                    setNames(fmAlpha["HOME.TEAM.BOLDED.1"],names(new_df["HOME"])))
team_names_sort <- team_names[order(team_names$HOME),]

#getting unique teams from entire list
participants<-c(unique(team_names))
teams <- as.data.frame(participants)
teams <- teams[order(teams$HOME) , ]
teams <- as.data.frame(teams)
#giving identification number to all teams after ordering
teams$Num.Id <- as.numeric(as.factor(teams$teams))

View(teams)

#wriring these data frame to excel sheet defined at perticular location
output.location1 <- "G:\\fastModel\\outAllTeams.xlsx"
write.xlsx(x = teams, file = output.location1, 
           sheetName = "ALLTEAMS", row.names = FALSE)


#QUESTION 2:::::::::::::::::::::::::::::::::::::
#---------------------------------------------

fmChrono <- cbind(fmChrono,HOME.TEAM=fmChrono1$X1)
fmChrono <- cbind(fmChrono,AWAY.TEAM=fmChrono2$X1)

#Removing NULL columns from the data frame
fmChrono$NA. <- NULL
fmChrono$NA..1 <- NULL
fmChrono$NA..2 <- NULL
fmChrono$NA..3 <- NULL
fmChrono$NA..4 <- NULL
fmChrono$NA..5 <- NULL
fmChrono$NA..6 <- NULL

#Getting Date and Time  
fmChrono <- cbind(fmChrono, DATE=as.Date(fmChrono$TIME))
fmChrono <- cbind(fmChrono,TIMEING=as.character.POSIXt(fmChrono$TIME, "%H:%M:%S"))
Hometeams <- teams
colnames(Hometeams) <- c("HOME.TEAM", "HOME.TEAM.ID")
Awayteams <- teams
colnames(Awayteams) <- c("AWAY.TEAM", "AWAY.TEAM.ID")

datetime <- data.frame(DAY=c("THUR","FRI", "SAT"), DATE1=c("July, 10","July, 11", "July, 12"))
#Getting Home team Ids
fmChrono<- merge(fmChrono, Hometeams, by="HOME.TEAM")
#Getting Away team Ids
fmChrono<- merge(fmChrono, Awayteams, by="AWAY.TEAM")
fmChrono<- merge(fmChrono, datetime, by="DAY")

#Creating final dataframe
ALLMATCHES <- data.frame(DATE=fmChrono$DATE1, TIME=fmChrono$TIMEING, COURT.NUMBER=fmChrono$COURT, HOME.TEAM.ID=fmChrono$HOME.TEAM.ID, AWAY.TEAM.ID=fmChrono$AWAY.TEAM.ID, DAY=fmChrono$DAY)

#wriring these data frame to excel sheet defined at perticular location
output.location2 <- "G:\\fastModel\\outAllMatches.xlsx"
write.xlsx(x = ALLMATCHES, file = output.location2, 
           sheetName = "ALLMATCHES", row.names = FALSE)
