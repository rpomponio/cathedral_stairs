library(data.table)
library(readxl)

metadat <- fread("Metadata_AllStairTrackers.tsv")
n.sheets <- nrow(metadat)
base.time <- as.POSIXct("1899-12-31 00:00:00", tz="UTC")

# begin with 2011-2012 in unique format
times <- data.table()
for (j in metadat[1, StartingSheet]:11){
  dat.ij <- data.table(read_xlsx(
    file.path("Data", metadat[1, FileName]),
    sheet=j,
    range=metadat[1, DataRange],
    col_types=c("date", "text", "numeric", "text", "skip",
                "date", "numeric", rep("date", 12), "skip", "skip")))
  rep.cols <- colnames(dat.ij)[grepl("\\d+", colnames(dat.ij))]
  times.ij <- dat.ij[, lapply(.SD, function(x) as.numeric(x - base.time)),
                     .SDcols=rep.cols]
  colnames(times.ij) <- paste0("Rep", 1:length(rep.cols))
  times.ij[, EntryYear:=metadat[1, EntryYear]]
  times.ij[, Date:=as.IDate(dat.ij$Date)]
  times.ij[, Weight:=dat.ij$`Pack Wt. (lbs.)`]
  times.ij[, Week:=dat.ij$Workout]
  times.ij[, StudentID:=paste0(metadat[1, EntryYear], "-", 100 + j)]
  times <- rbindlist(list(times, times.ij), fill=TRUE)
}

# iterate over remaining spreadsheets with consistent formats
for (i in 2:nrow(metadat)){
  for (j in metadat[i, StartingSheet]:30){
    skip.sheet <- FALSE
    dat.ij <- tryCatch({
      data.table(read_xlsx(
        file.path("Data", metadat[i, FileName]),
        sheet=j,
        range=metadat[i, DataRange],
        col_types=c("numeric", "date", "numeric", "text", "numeric", "skip",
                    "date", rep("date", 14), rep("skip", ifelse(i==4, 3, 2)))))},
      error=function(e) { skip.sheet <<- TRUE})
    if(skip.sheet || nrow(dat.ij)==0) { next }
    rep.cols <- colnames(dat.ij)[grepl("Rep \\d+", colnames(dat.ij))]
    times.ij <- dat.ij[, lapply(.SD, function(x) as.numeric(x - base.time)),
                       .SDcols=rep.cols]
    colnames(times.ij) <- paste0("Rep", 1:length(rep.cols))
    times.ij[, EntryYear:=metadat[i, EntryYear]]
    times.ij[, Date:=as.IDate(dat.ij$Date)]
    times.ij[, Weight:=dat.ij$`*Pack Wt. (lbs.)`]
    times.ij[, Week:=dat.ij$Workout]
    times.ij[, Footwear:=dat.ij[, 4]]
    times.ij[, Footwear:=fcase(
      grepl("running", Footwear, ignore.case=TRUE), 1,
      grepl("hiking", Footwear, ignore.case=TRUE), 2,
      grepl("mtn", Footwear, ignore.case=TRUE), 3)]
    times.ij[, StudentID:=paste0(metadat[i, EntryYear], "-", 100 + j)]
    times <- rbindlist(list(times, times.ij), fill=TRUE)
  }
}

# perform basic data validation/cleaning
clean <- times[between(Rep1, 0, 60) & between(Rep2, 0, 60)]
clean[, .(Entries=.N, Students=length(unique(StudentID))), by=EntryYear]
clean[Rep4 > 60, Rep4:=Rep4 - (12 * 60)]
# pairs(clean[, paste0("Rep", 1:7)])
# hist(clean$Weight)

# transform to long format
long <- melt(clean,
             id.vars=c("EntryYear", "StudentID", "Date",
                       "Week", "Weight", "Footwear"),
             variable.name="RepNum", value.name="TimeMins", na.rm=TRUE,
             variable.factor=FALSE)
long[, EntryYear:=paste0(EntryYear, "-", EntryYear + 1 - 2000)]
long[, RepNum:=as.integer(gsub("Rep", "", RepNum))]
setkey(long, StudentID, Week, Date, RepNum)

# store as CSV file; Rdata file
fwrite(long, file.path("Data", "MtnSchoolStairReps.csv"))
saveRDS(long, "stair_reps.rds")
