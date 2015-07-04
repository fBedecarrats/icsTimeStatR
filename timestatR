# TIMESTATS

# A R script to analyse events generated in an ics file
# Designed for time management with Thunderbird

# CC-BY-SA, contributions welcomed
# By: Florent Bedecarrats, ...
# Includes snippet from Jason Bryer
# Also inspired by Kay Cichini (widgets) and Roger Peng (regex)

# CONTENT
  # Setup
    # Load options
    # Let user adapt options
    # Save option upfate
  # Retrive and prepare calendar
    # Select folder
    # List files
    # Prepare files
  # Parse document
  # Analyse data
  # Let user adjust analysis parameters
  # Output and save

# SETUP
# Install needed packages
list.of.packages <- c("tcltk", "hwriter", "xlsx")
new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
if(length(new.packages)) install.packages(new.packages)
lapply(list.of.packages, require, character.only=T)

# Set options
if (file.exists("TimeStats_options.Rdata")) {
  load("TimeStats_options.Rdata")
  } else {
    default.folder <- ""
    start.stats <- "01/01/2013"
    end.stats <- "31/12/2013"
    hours.full.day <- "8"
    days.off <- c(6,7)
  }

# Let user adapt options folder
calendar.folder <- tk_choose.dir(default = default.folder, 
                                 caption = "Specify your calendar folder") 
setwd(calendar.folder)

# Setting other options
# Begin fun by J Bryer (https://gist.github.com/jbryer/3342915)
# modified to include defaults. Modifs commented
varEntryDialog <- function(vars,
                           labels = vars,
                           defaults = vars, # added to JBryer model
                           fun = rep(list(as.character), length(vars)),
                           title = 'Variable Entry',
                           prompt = NULL) {
  require(tcltk)
  stopifnot(length(vars) == length(labels), length(labels) == length(fun))
  
  # Create a variable to keep track of the state of the dialog window:
  # done = r0; If the window is active
  # done = 1; If the window has been closed using the OK button
  # done = 2; If the window has been closed using the Cancel button or destroyed
  done <- tclVar(0)
  
  tt <- tktoplevel()
  tkwm.title(tt, title)
  entries <- list()
  tclvars <- list()
  
  # Capture the event "Destroy" (e.g. Alt-F4 in Windows) and when this happens,
  # assign 2 to done.
  tkbind(tt,"<Destroy>",function() tclvalue(done)<-2)
  for(i in seq_along(vars)) {
    tclvars[[i]] <- tclVar("")
    entries[[i]] <- tkentry(tt, textvariable=tclvars[[i]])
  }
  # Begin section added to JBryer model
  set.defaults <- function() {
    for(i in seq_along(entries)) {
      tclvalue(tclvars[[i]]) <<- defaults[i]
    }
  }
  set.defaults()
  # End section added to JBryer model
  doneVal <- as.integer(tclvalue(done))
  results <- list()
  
  reset <- function() {
    for(i in seq_along(entries)) {
      tclvalue(tclvars[[i]]) <<- ""
    }
  }
  reset.but <- tkbutton(tt, text="Reset", command=reset)
  cancel <- function() {
    tclvalue(done) <- 2
  }
  cancel.but <- tkbutton(tt, text='Cancel', command=cancel)
  submit <- function() {
    for(i in seq_along(vars)) {
      tryCatch( {
        results[[vars[[i]]]] <<- fun[[i]](tclvalue(tclvars[[i]]))
        tclvalue(done) <- 1
      },
      error = function(e) { tkmessageBox(message=geterrmessage()) },
      finally = { }
      )
    }
  }
  submit.but <- tkbutton(tt, text="Submit", command=submit)
  if(!is.null(prompt)) {
    tkgrid(tklabel(tt,text=prompt), columnspan=3, pady=10)
  }
  for(i in seq_along(vars)) {
    tkgrid(tklabel(tt, text=labels[i]), entries[[i]], pady=10, padx=10, columnspan=4)
  }
  tkgrid(submit.but, cancel.but, reset.but, pady=10, padx=10, columnspan=3)
  tkfocus(tt)
  
  # Do not proceed with the following code until the variable done is non-zero.
  # (But other processes can still run, i.e. the system is not frozen.)
  tkwait.variable(done)
  if(tclvalue(done) != 1) {
    results <- NULL
  }
  tkdestroy(tt)
  return(results)
}

if(FALSE) { #Test the dialog
  vals <- varEntryDialog(vars=c('Variable1', 'Variable2'))
  str(vals)
  vals <- varEntryDialog(vars=c('Var1', 'Var2'),
                         labels=c('Enter an integer:', 'Enter a string:'),
                         fun=c(as.integer, as.character))
  str(vals)
  #Add a custom validation function
  vals <- varEntryDialog(vars=c('Var1'),
                         labels=c('Enter an integer between 0 and 10:'),
                         fun=c(function(x) {
                           x <- as.integer(x)
                           if(x >= 0 & x <= 10) {
                             return(x)
                           } else {
                             stop("Why didn't you follow instruction!\nPlease enter a number between 0 and 10.")
                           }
                         } ))
  str(vals)
  #Return a list
  vals <- varEntryDialog(vars=c('Var1'),
                         labels=c('Enter a comma separated list of something:'),
                         fun=c(function(x) {
                           return(strsplit(x, split=','))
                         }))
  vals$Var1
  str(vals)
} 
# End fun by J Bryer

options <- varEntryDialog(vars = c("start.stats", "end.stats"),
                          labels = c("Date beginning dd/mm/yyyy", "Date end dd/mm/yyyy"),
                          defaults = c(start.stats, end.stats),
                          fun = c(as.character, as.character))

# Save options
default.folder <- calendar.folder
start.stats <- as.character(options[1])
end.stats <- as.character(options[2])
hours.full.day <- as.numeric(options[3])
days.off <- as.numeric(options[4])
save(default.folder, start.stats, end.stats, hours.full.day, days.off, 
     file = "TimeStats_options.Rdata")

# Identify all ics files in folder and merge them
calendar.files <- list.files(path = calendar.folder, pattern = "\\.ics$")
calendar.paths <- paste(calendar.folder, calendar.files, sep = "/")

# Identify all ics files in folder and merge them with identifier
for (i in seq_along(calendar.paths)){
  start <- "BEGIN:VEVENT"
  id <- as.character(calendar.files[i])
  loc <- as.character(calendar.paths[i]) 
  separ <- "\t"
  calendar.transit <- readChar(loc, nchars=1e7)
  calendar.transit <- gsub(start, paste0(start, separ, id, separ, sep = ""), 
                           calendar.transit) 
  calendar.raw <- paste(ifelse(exists("calendar.raw"), calendar.raw, ""), 
                        calendar.transit, sep = "", collapse = "\t")
                        
}


calendar.raw <- gsub("\r\n", "\t", calendar.raw)
calendar.raw <- gsub("\tBEGIN:VEVENT", "\nBEGIN:VEVENT",calendar.raw)
calendar.raw <- gsub("\tBEGIN:VCALENDAR", "\nBEGIN:VCALENDAR",calendar.raw)

# Work around to load these events (didn't manage otherwise)
writeLines(calendar.raw,"calendar.txt", sep = "\n")
remove(calendar.raw)
calendar <- as.data.frame(readLines("calendar.txt"))
colnames(calendar) <- c("raw")
calendar$full <- as.character(calendar$raw)
#calendar <- as.data.frame(calendar[-1,])

# Removes non-events
calendar <- subset(calendar, grepl("BEGIN:VEVENT",calendar$raw))

# Extract Title
calendar$title <- regexec("\tSUMMARY:.*?\t",calendar$full)
calendar$title <- regmatches(calendar$full,calendar$title)
calendar$title <- unlist({calendar$title[sapply(calendar$title, length)==0] <- NA; calendar$title})
calendar$title <- gsub("\tSUMMARY:|\t","",calendar$title)

# Extract Category
calendar$category <- regexec("\tCATEGORIES:.*?\t", calendar$full)
calendar$category <- regmatches(calendar$full, calendar$category)
calendar$category <- unlist({calendar$category[sapply(calendar$category, length)==0] <- NA; calendar$category})
calendar$category <- gsub("\tCATEGORIES:|\t", "", calendar$category)


# Extract caldendarname
calendar$calname <- regexec("^BEGIN:VEVENT\t.*?\t", calendar$full)
calendar$calname <- regmatches(calendar$full, calendar$calname)
calendar$calname <- unlist({calendar$calname[sapply(calendar$calname, length)==0] <- NA; calendar$calname})
calendar$calname <- gsub("^BEGIN:VEVENT\t|.ics\t", "", calendar$calname)

# Extract begin time
calendar$begin.icstime <- regexec("\tDTSTART;TZID=.*?\t", calendar$full)
calendar$begin.icstime <- regmatches(calendar$full, calendar$begin.icstime)
calendar$begin.icstime <- unlist({calendar$begin.icstime[sapply(calendar$begin.icstime, length)==0] <- NA; calendar$begin.icstime})                                                             
calendar$begin.icstime <- gsub("\tDTSTART;TZID=|(.*?):|\t", "", calendar$begin.icstime)
calendar$begin.posixct <- ifelse(is.na(calendar$begin.icstime), NA, as.POSIXct(as.character(calendar$begin.icstime), format = "%Y%m%dT%H%M%S"))
calendar$begin.date <- as.Date(calendar$begin.icstime, format = "%Y%m%d")

# Extract end time
calendar$end.icstime <- regexec("\tDTEND;TZID=.*?\t", calendar$full)
calendar$end.icstime <- regmatches(calendar$full, calendar$end.icstime)
calendar$end.icstime <- unlist({calendar$end.icstime[sapply(calendar$end.icstime, length)==0] <- NA; calendar$end.icstime})
calendar$end.icstime <- gsub("\tDTEND;TZID=|(.*?):|\t", "", calendar$end.icstime)
calendar$end.posixct <- ifelse(is.na(calendar$end.icstime), NA, as.POSIXct(as.character(calendar$end.icstime), format = "%Y%m%dT%H%M%S"))
calendar$end.date <- as.Date(calendar$end.icstime, format = "%Y%m%d")

# Convert dates as POSIXct (seconds since 01/01/1970) and calculate duration
calendar$duration.hours <- (calendar$end.posixct - calendar$begin.posixct) / 3600

# Keeps only events in the user screening range
calendar <- subset(calendar, begin.date >= as.Date(start.stats, format = "%d/%m/%Y") &
                     end.date <= as.Date(end.stats, format = "%d/%m/%Y"),
                   select = c(-raw, -full))

# TimeStats
calendar$category <- factor(calendar$category, exclude = NULL)
timestats <- with(calendar, tapply(duration.hours, list(category, calname), sum))


# Print html
p = openPage("timestat.html", head = "Total task durations (hours) by category")
hwrite(timestats, p)
closePage(p)
browseURL("timestat.html")

row.names(timestats) <- ifelse(is.na(row.names(timestats)), "Undefined", row.names(timestats)) 
write.xlsx(timestats, "timestat.xlsx", showNA = FALSE)
