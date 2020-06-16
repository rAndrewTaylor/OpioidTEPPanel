library("readxl")

get.incidmat <- function(dat) {
  individuals <- sort(unique(dat[[1]]))
  nstatements <- get.nsorted.dat(dat)
  mat.list <- list()
  # Making a matrix for each individual and storing them in a list #
  for (i in 1:length(individuals)) {
    group.dat <- dat[dat[,1]==individuals[i],]
    incid.mat <- matrix(0, nrow=nstatements, ncol=nstatements)
    for (j in 1:dim(group.dat)[1]) {
      cards <- group.dat[j, which(!is.na(group.dat[j,]))]
      cols  <- which(substr(colnames(cards),1,4) == "Item")
      cards <- as.numeric(sort(unlist(cards[cols])))
      vct   <- rep(0,nstatements)
      vct[cards] = 1
      incid.mat <- incid.mat + vct%*%t(vct)
    }
    mat.list[[i]] <- incid.mat
  }
  return(mat.list)
}

CheckData <- function(inc.mat) {
  nindiv <- length(inc.mat)
  nstatements <- dim(inc.mat[[1]])[1]
  data.issues <- c()
  for (i in 1:nindiv) {
    if (sum(inc.mat[[i]]) == nstatements^2) {
      data.issues <- c(data.issues,
                       sprintf("Sorter %s: %s",i, "All statements in one pile!"))
    }
    if (sum(inc.mat[[i]]) == nstatements) {
      data.issues <- c(data.issues,
                       sprintf("Sorter %s: %s",i, "Each statements in its own pile!"))
    }
    for(j in 1:(nstatements-1)) {
      nz = which(inc.mat[[i]][j,] > 1)
      if (length(nz) > 0) {
        data.issues <- c(data.issues,
                         sprintf("Sorter %s: statements %s in multiple piles", i, j))
      }
    }
    nz = which(diag(inc.mat[[i]]) < 1)
    if (length(nz) > 0) {
      data.issues <- c(data.issues,
                       sprintf("Sorter %s: did not sort statements %s", i, paste(nz,collapse=",")))
    }
  }
  paste(data.issues, collapse="\n")
}

cat("\f\n Select the Concept Mapping data file\n")
excel.file.name <- file.choose()

cat("\nChecking", excel.file.name, "...\n")

sheetnames <- excel_sheets(excel.file.name)
expectedsheets <- c("Statements", "SortedCards", "Demographics", "Ratings", "RatingsScale")
foundsheets <- expectedsheets %in% sheetnames

# check the overall structure (look for the specific sheet names)
fatalErr <- FALSE
for (sheetNo in 1:5) {
  if (!foundsheets[sheetNo]) {
    warning("The input file is missing the", expectedsheets[sheetNo]," sheet ***")
    if (sheetNo %in% c(1,2)) { fatalErr <- TRUE }
  }
}
if (fatalErr) { stop("Fix the problems in the data file and try again\n\n") }

# statement.dat must have two columns: StatementID and Statement
if (file_ext(excel.file.name) == "xlsx") {
  statement.dat <- read_xlsx(excel.file.name,sheet="Statements")
  sort.dat <- read_xlsx(excel.file.name,col_names=FALSE,sheet="SortedCards")
}
if (file_ext(excel.file.name) == "xls") {
  statement.dat <- read_xls(excel.file.name,sheet="Statements")
  sort.dat <- read_xls(excel.file.name,col_names=FALSE,sheet="SortedCards")
}
if (any(colnames(statement.dat) != c("StatementID", "Statement"))) {
  stop("The Statements sheet must have two columns - StatementID and Statement")
}

# check the statement IDs:
if(any(sort(statement.dat$StatementID) != 1:length(statement.dat$StatementID))) {
  stop("The Statement IDs must be consecutive numbers.")
}

# check the format of the SortedCards sheet:
colnames(sort.dat) <- c("SorterID","SorterLabel",
                        paste("Item",(1:(dim(sort.dat)[2]-2))))
if ((mode(sort.dat$SorterID) != "numeric") | (mode(sort.dat$SorterLabel) != "character")) {
  stop("The first column in the SortedCards sheet must contain numeric sorter IDs and the second column must contain pile labels from each sorter (could be blank)")
}

# check the the sorter IDs are consecutive numbers
sorters <- sort(unique(sort(as.vector(sort.dat$SorterID))))
if(any(sorters != 1:length(sorters))) {
  warning("The Sorter IDs must be consecutive numbers.")
}

# in the SortedCards sheet, check if any statement ID is invalid:
sortedStatements <- unique(sort(unlist(sort.dat[,-c(1,2)])))
oorstatement <- which(!(sortedStatements %in% 1:length(statement.dat$StatementID)) )
if (length(oorstatement) > 0)  {
  stop(paste("Sorted statement number(s) in the SortedCards sheet are out of range:", oorstatement))
}

# more detailed test for SortedCards (see function above)
message(CheckData(get.incidmat(sort.dat)),"\n")

# check the demographic data - are there any invalid sorter ID?
# warn about missing data
demographics.dat <- c()
if ("Demographics"%in%sheetnames) {
  if (file_ext(excel.file.name) == "xlsx")
    demographics.dat <- read_xlsx(excel.file.name,sheet="Demographics")
  if (file_ext(excel.file.name) == "xls")
    demographics.dat <- read_xls(excel.file.name,sheet="Demographics")
  
  if (colnames(demographics.dat)[1] != "SorterID") {
    stop("The first column in the Demographics sheet must be called SorterID")
  }
  
  sortersDemographic <- demographics.dat$SorterID
  dups <- duplicated(sortersDemographic)
  if (any(dups)) {
    stop(paste("Duplicated sorter ID in Demographics:", demographics.dat$SorterID[which(dups)]))
  }
  oorsorter <- which(!(sortersDemographic %in% sorters))
  if (length(oorsorter) > 0)  {
    warning("Sorter number(s) in the Demographics sheet are out of range: ", paste(oorsorter, collapse=", "))
  }
  if (any(is.na(demographics.dat[,-1]))) {
    warning("Missing data in the Demographics sheet")
  }
}

# check the ratings data

ratings.scales.dat <- c()
mins <- list()
maxs <- list()  
if ("RatingsScale"%in%sheetnames) {
  if (file_ext(excel.file.name) == "xlsx")
    ratings.scales.dat <- read_xlsx(excel.file.name,sheet="RatingsScale")
  if (file_ext(excel.file.name) == "xls")
    ratings.scales.dat <- read_xls(excel.file.name,sheet="RatingsScale")
  
  if ((any(ratings.scales.dat[,1] != c("Min","Max"))) | 
      (colnames(ratings.scales.dat)[1] != "Variable")) {
    stop("The RatingsScale sheet is not formatted properly.")
  }
  ratingVarNames <- colnames(ratings.scales.dat[-1])
  for (coln in 2:length(colnames(ratings.scales.dat))) {
    mins[[colnames(ratings.scales.dat)[coln]]] <- ratings.scales.dat[1,coln]
    maxs[[colnames(ratings.scales.dat)[coln]]] <- ratings.scales.dat[2,coln]
  }
}

ratings.dat <- c()
if ("Ratings"%in%sheetnames) {
  if (file_ext(excel.file.name) == "xlsx")
    ratings.dat <- read_xlsx(excel.file.name,sheet="Ratings")
  if (file_ext(excel.file.name) == "xls")
    ratings.dat <- read_xls(excel.file.name,sheet="Ratings")
  if (any(colnames(ratings.dat)[1:2] != c("UserID", "StatementID"))){
    stop("The Ratings sheet is not formatted properly.")
  }
  if (any(!(colnames(ratings.dat)[-(1:2)] %in% ratingVarNames))) {
    stop("The variables in the Ratings sheet are not the same as in the RatingsScale sheet.")
  }
  ratedStatements <- unique(sort(ratings.dat$StatementID))
  oorstatement <- which(!(ratedStatements %in% 1:length(statement.dat$StatementID))) 
  if (length(oorstatement) > 0)  {
    stop(paste("Sorted statement number(s) in the Ratings sheet are out of range:", oorstatement))
  }
  
  sortersRatings <- ratings.dat$UserID
  oorsorter <- which(!(sortersRatings %in% sorters))
  if (length(oorsorter) > 0)  {
    warning("Sorter number(s) in the Ratings sheet are out of range. Row numbers: ", paste(oorsorter, collapse=", "))
  }
  if (any(is.na(ratings.dat[,-(1:2)]))) {
    warning("Missing data in the Ratings sheet")
  }
}

cat("Done\n")
