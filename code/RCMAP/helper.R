
#######
# find the directory of the current (open) project
getActiveProject <- function() {
  tryfile <- sprintf("%s/%s",getwd(),"activesession.RData")
  if (file.exists(tryfile)) {
    activeProject <- getwd()
  } else {
    activeProject <- "None"
  }
  activeProject
}

#######
# open the spreadsheet, or a saved (Rdata) file that contains the data for a project
OpenProject <- function(excel.file) {
  errs <- c()
  excel.file.name <- excel.file$name # the actual file name
  excel.file <- excel.file$datapath  # the file path (set by RStudio)
  # if the file has an RData extension, assume it's a concept mapping saved project
  # and load it. Otherwise, assume it's an excel file in the appropriate format
  fromExcel = TRUE
  if (length(grep(".RData",excel.file.name) ) > 0) {
    fromExcel = FALSE
    # comment out ".name" from the line below in order to enable uploading files from outside of the working directory
    file.copy(excel.file #.name
              ,"activesession.RData",overwrite = T)
    cmap = ReadSavedData()
    
    plotChoices <<- c("Point Map (MDS)", "Clusters", 
                      "Univariate Rating - by Statement",
                      "Univariate Rating - by Cluster",
                      "Bar Chart", "Pattern Matching", "GoZone", "Dendrogram")
    if (ncol(cmap$ratings.dat) == 3)
      plotChoices <<- c("Point Map (MDS)", "Clusters", 
                        "Univariate Rating - by Statement",
                        "Univariate Rating - by Cluster",
                        "Bar Chart", "Dendrogram")
    if (ncol(cmap$ratings.dat) < 3)
      plotChoices <<- c("Point Map (MDS)", "Clusters", "Dendrogram")
    if (!is.null(cmap$excel.file)) { # an .RData file, but does not contain a concept mapping file
      errs <- c(errs,"The selected file does not seem to contain concept mapping data")
    }
  } else { # assume the data is in an excel file (with the concept mapping format. See documentation)
    # we expect to find a few named sheets - "Statements" and "SortedCards" must be included in the file
    cat(excel.file.name,"\n")
    sheetnames <- excel_sheets(excel.file.name)
    # statement.dat must have two columns: StatementID and Statement
    if (file_ext(excel.file.name) == "xlsx")
      statement.dat <- read_xlsx(excel.file,sheet="Statements")
    if (file_ext(excel.file.name) == "xls")
      statement.dat <- read_xls(excel.file,sheet="Statements")
    # clustname will contain a list of cluster names
    clustname <- new.env()
    
    # create list of objects to be included in final report
    reportObjects <- list()
    # for loop creates the reportObjects categories
    for (panelName in c("Plots", "Reports", "Analysis")) {
      choices = getChoices(panelName)
      for (choice in choices) {
        reportObjects[[panelName]][[choice]] = list()
      }
    }
    
    # The sheet SortedCards has the following format: SorterID, Label (can be left blank), followed by
    # columns with the statement IDs that were put together in a pile by the sorter.
    if (file_ext(excel.file.name) == "xlsx")
      sort.dat <- read_xlsx(excel.file,col_names=FALSE,sheet="SortedCards")
    if (file_ext(excel.file.name) == "xls")
      sort.dat <- read_xls(excel.file,col_names=FALSE,sheet="SortedCards")
    colnames(sort.dat) <- c("SorterID","SorterLabel",
                            paste("Item",(1:(dim(sort.dat)[2]-2))))
    # use the sorted cards to generate the incidence matrix:
    inc.mat <- get.incidmat(sort.dat)
    errs <- c(errs,CheckData(inc.mat))
    # print the errors on the main screen
    cat(errs,"\n")
    #     if (length(errs) > 0) {
    #      DataProblems(errs)
    #    } else { 
    n.indiv <- length(inc.mat)
    M1 = Reduce("+",inc.mat)
    S1 = diag(M1)
    group.sim.mat <- S1-M1
    exc <- c()
    for (colA in 1:(dim(group.sim.mat)[1]-1)) {
      for (colB in (colA+1):dim(group.sim.mat)[1]) {
        if ( max((group.sim.mat[,colA]-group.sim.mat[,colB])^2) == 0) {
          group.sim.mat[colA,colB] <- group.sim.mat[colA,colB] + 0.001
          group.sim.mat[colB,colA] <- group.sim.mat[colA,colB]
        }
      }
    }
    
    # multidimensional scaling, and clustering:
    fit.MDS <- mds(group.sim.mat)
    gsm <- group.sim.mat
    d <- dist(fit.MDS$conf)
    
    fit.WardClust <- hclust(d)
    n.phrases <- dim(group.sim.mat)[1]
    x <- fit.MDS$conf[,1]
    y <- fit.MDS$conf[,2]
    x <- (x-min(x))
    y <- (y-min(y))
    stress <- fit.MDS$stress
    
    # read the ratings scale:
    ratings.scales.dat <- c()
    mins <- list()
    maxs <- list()  
    if ("RatingsScale"%in%sheetnames) {
      if (file_ext(excel.file.name) == "xlsx")
        ratings.scales.dat <- read_xlsx(excel.file,sheet="RatingsScale")
      if (file_ext(excel.file.name) == "xls")
        ratings.scales.dat <- read_xls(excel.file,sheet="RatingsScale")
      
      for (coln in 2:length(colnames(ratings.scales.dat))) {
        mins[[colnames(ratings.scales.dat)[coln]]] <- ratings.scales.dat[1,coln]
        maxs[[colnames(ratings.scales.dat)[coln]]] <- ratings.scales.dat[2,coln]
      }
    }
    
    ratings.dat <- c()
    if ("Ratings"%in%sheetnames) {
      if (file_ext(excel.file.name) == "xlsx")
        ratings.dat <- read_xlsx(excel.file,sheet="Ratings")
      if (file_ext(excel.file.name) == "xls")
        ratings.dat <- read_xls(excel.file,sheet="Ratings")
      
      if(length(exc) > 0) {
        ratings.dat <- ratings.dat[-which(ratings.dat$StatementID%in%exc),]
        statement.dat <- statement.dat[-which(statement.dat$StatementID%in%exc),]
      }
    }
    
    plotChoices <<- c("Point Map (MDS)", "Clusters", 
                      "Univariate Rating - by Statement",
                      "Univariate Rating - by Cluster",
                      "Bar Chart", "Pattern Matching", "GoZone", "Dendrogram")
    if (ncol(ratings.dat) == 3)
      plotChoices <<- c("Point Map (MDS)", "Clusters", 
                        "Univariate Rating - by Statement",
                        "Univariate Rating - by Cluster",
                        "Bar Chart", "Dendrogram")
    if (ncol(ratings.dat) < 3)
      plotChoices <<- c("Point Map (MDS)", "Clusters", "Dendrogram")
    
    demographics.dat <- c()
    if ("Demographics"%in%sheetnames) {
      if (file_ext(excel.file.name) == "xlsx")
        demographics.dat <- read_xlsx(excel.file,sheet="Demographics")
      if (file_ext(excel.file.name) == "xls")
        demographics.dat <- read_xls(excel.file,sheet="Demographics")
      
      for (j in 1:dim(demographics.dat)[2]) {
        if (class(demographics.dat[[j]]) == "character") {
          demographics.dat[[j]] <- as.factor(demographics.dat[[j]])
        }
      }
    }
    critvar1 <-  "all" 
    critvar2 <-  "all" 
    n.clust=3
    
    description = ""
    demographicGroups <- list("All" = list(searchCols = NULL, 
                                           demographics.dat$SorterID))
  }
  if(fromExcel) {
    save(description, excel.file, excel.file.name, statement.dat, clustname, n.clust, reportObjects, demographicGroups,
         sort.dat, inc.mat, group.sim.mat, gsm, fit.MDS, x, y, stress, n.indiv, 
         n.phrases, fit.WardClust, ratings.dat,
         ratings.scales.dat, mins, maxs, demographics.dat, critvar1, critvar2,
         file="activesession.RData")
  } else {
    cmap$excel.file.name <- excel.file.name
    with(cmap, {
      excel.file = excel.file.name
      save(description, excel.file, excel.file.name, statement.dat, clustname, n.clust, reportObjects, demographicGroups,
           sort.dat, inc.mat, group.sim.mat, gsm, fit.MDS, x, y, stress, n.indiv, 
           n.phrases, fit.WardClust, ratings.dat,
           ratings.scales.dat, mins, maxs, demographics.dat, critvar1, critvar2,
           file="activesession.RData")  
    })
  }
  
  cat("Saved data in activesession.RData\n")
}

ReloadProject <- function() {
  getActiveProject()
  cmap = ReadSavedData()
}

####
# find the number of statements that were sorted by users (should be the same as the number
# of items in the Statements sheet)
get.nsorted <- function() {
  cm = ReadSavedData()
  dat = cm$sort.dat
  numcols = dim(dat)[2]
  return (length(na.exclude(unique(stack(dat[,3:numcols])[,1]))))
}

get.nsorted.dat <- function(dat) {
  numcols = dim(dat)[2]
  return (length(na.exclude(unique(stack(dat[,3:numcols])[,1]))))
}
####
# create incidence matrix for each sorter
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

####
# check the sorted cards data
# For each individual, the following rules are checked:
# 1. All statements are not in a single pile
# 2. All statements are not in their own pile
# 3. Every statement can not be in more than one pile
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
  data.issues
}

####
# get the data from previous sessions or steps (stored in activesession.RData)
ReadSavedData <- function(){
  conceptmap <- new.env()
  conceptmap[["errs"]] = c()
  if (file.exists("activesession.RData")) {
    load("activesession.RData")
    required.sort.vars = 
      c("excel.file", 
        "excel.file.name",
        "sort.dat", 
        "inc.mat", 
        "group.sim.mat", 
        "fit.MDS",
        "x",
        "y",
        "stress",
        "n.indiv", 
        "n.phrases",
        "n.clust",
        "fit.WardClust")
    required.rating.vars = 
      c("ratings.dat")
    
    for (req.var in required.sort.vars) {
      if (!req.var%in%ls()) {
        if (!"Load the sorting file!"%in%conceptmap[["errs"]]) {
          conceptmap[["errs"]] = c(conceptmap[["errs"]],"Load the sorting file!")
        }
      } else {
        conceptmap[[req.var]] = get(req.var)
      }
    }
    for (req.var in required.rating.vars) {
      if (!req.var%in%ls()) {
        if (!"Load the ratings file!"%in%conceptmap[["errs"]]) {
          conceptmap[["errs"]] = c(conceptmap[["errs"]],"Load the ratings file!")
        }
      } else {
        conceptmap[[req.var]] = get(req.var)
      }
    }
    if ("n.clust"%in%ls()) { conceptmap[["n.clust"]] = get("n.clust") }
    else { conceptmap[["n.clust"]] = 3 }
    if ("clust.method"%in%ls()) { conceptmap[["clust.method"]] = get("clust.method") }
    else { conceptmap[["clust.method"]] = "Ward" }
    if ("colpoints"%in%ls()) { conceptmap[["colpoints"]] = get("colpoints") }
    else { conceptmap[["colpoints"]] = 1 }
    if ("coltext"%in%ls()) { conceptmap[["coltext"]] = get("coltext") }
    else { conceptmap[["coltext"]] = 1 }
    if ("showStatementsVal"%in%ls()) { conceptmap[["showStatementsVal"]] = get("showStatementsVal") }
    else { conceptmap[["showStatementsVal"]] = 1 }
    conceptmap[["clustname"]] = get("clustname")
    conceptmap[["statement.dat"]] = get("statement.dat")
    conceptmap[["ratings.scales.dat"]] = get("ratings.scales.dat")
    conceptmap[["maxs"]] = get("maxs")
    conceptmap[["mins"]] = get("mins")
    conceptmap[["demographics.dat"]] = get("demographics.dat")
    conceptmap[["critvar1"]] = get("critvar1")
    conceptmap[["critvar2"]] = get("critvar2")
    conceptmap[["description"]] = get("description")
    
    conceptmap[["reportObjects"]] = get("reportObjects")
    conceptmap[["demographicGroups"]] = get("demographicGroups")
    conceptmap[["plotChoices"]] = get("plotChoices")
    #
  }
  else {
    conceptmap[["errs"]] = 
      c("Read sorting and rating files to initialize the project!")
  }
  conceptmap
}


####
# get labels for each phrase
phrase.labels <- function(dat,nphrases) {
  dim1 <- dim(dat)[1]
  dim2 <- dim(dat)[2]
  cm  = ReadSavedData()
  phraselist <- vector("list",nphrases)
  for (i in 1:dim1) {
    x <- as.vector(as.matrix(dat[i,3:dim2]))
    x <- x[!is.na(x)]
    x <- x[which(x%in%cm$statement.dat$StatementID)]
    for (j in 1:length(x)) {
      index = which(cm$statement.dat$StatementID == x[j])
      phraselist[[index]] <- c(phraselist[[index]],as.character(dat[i,2]))
    }  
  }
  return(phraselist)  
}

####
# get a list of labels for each cluster
cluster.labels <- function(sortdat,nphrases,clustfit,nclust) {
  groups <- cutree(clustfit,k=nclust)  	# determines cluser label for each phrase
  phraselist <- phrase.labels(sortdat,nphrases)
  clustphrases <- vector("list",nclust)
  for (i in 1:nclust) {
    grouplength <- length(groups[groups==i])
    for (j in 1:grouplength) {
      clustphrases[[i]] <- c(clustphrases[[i]],phraselist[[which(groups==i)[j]]])
    }	
  }
  return(clustphrases)
}

####
# find top 10 labels for each cluster
cluster.labels.pop <- function(clustlabels,topn) {
  clusttop <- vector("list",length(clustlabels))
  for (i in 1:length(clustlabels)) {
    clusttable <- as.data.frame(table(clustlabels[[i]]))
    names(clusttable) <- c("Label","Count")
    clusttable <- clusttable[order(-clusttable$Count),]
    clusttable <- data.frame("Label"=clusttable$Label,"Count"=clusttable$Count)
    clusttop[[i]] <- clusttable[1:topn,]
  }
  return(clusttop)
}

# show sorting summary (only sorters will appear in this summary)
get.summary <- function(demographicGroup) {
  cmap <- ReadSavedData()
  dat <- cmap$sort.dat
  cmap$demographicGroup <- demographicGroup
  cmap$selectedGroup = cmap$demographicGroups[[demographicGroup]]
  if (cmap$demographicGroup == "All")
    cmap$selectedGroup$rows <- 1:length(cmap$demographics.dat$SorterID)
  selected <- cmap$demographics.dat$SorterID[cmap$selectedGroup$rows]
  dat <- dat[dat[[1]] %in% selected,]
  
  # Get Number of Individuals:
  indiv <- unique(dat[[1]])
  indiv <- indiv[indiv %in% selected] 
  if (length(indiv) == 0)
    return(data.frame())
  
  # Get Number of Groupings for each Individual
  num.groupings <- rep(0,length(indiv))
  for (i in 1:length(indiv)) {
    temp.dat <- dat[dat[[1]]==indiv[i],]
    num.groupings[i] <- dim(temp.dat)[1]
  }
  
  # Get Number of statements sorted for each individual
  num.statements <- rep(0,length(indiv))
  for (i in 1:length(indiv)) {
    temp.dat <- dat[dat[[1]]==indiv[i],]
    y <- c()
    for (j in 1:dim(temp.dat)[1]) {
      x <- as.vector(as.matrix((temp.dat[j,3:dim(temp.dat)[2]])))
      x <- x[!is.na(x)]
      y <- c(y,x)  
    }
    num.statements[i] <- length(unique(y))
  }
  
  data.frame("Sorter.ID"=as.integer(indiv),
             "Statements.Sorted"=as.integer(num.statements),
             "Groupings"=as.integer(num.groupings))
}

showInfo <- function(setting,par){
  if (setting == "Cluster Names") {
    return(cluster.statements(par))
  }
}

showMoreInfo <- function(setting,par){
  if (setting == "Cluster Names") {
    return(cluster.top10.statements(par))
  }
}

plotCluster <- function(newclustername,oldclustername) {
  if (is.null(oldclustername))
    return(NULL)
  tmp = clusterLabels()
  clustno = which(tmp == newclustername)
  if (length(clustno) != 1)
    clustno = which(tmp == oldclustername)
  if (length(clustno)== 0)
    return(NULL)
  cm = ReadSavedData()
  dat = cm$ratings.dat; clustfit = cm$fit.WardClust; nclust = cm$n.clust
  nphrases = cm$n.phrases; pIDs = cm$statement.dat[[1]]; phrases = cm$statement.dat[[2]]
  groups <- cutree(clustfit,k=nclust)
  cm$cllbl = tmp[clustno]
  cm$slct = which(groups == clustno)
  cm$groups = groups
  with(cm,{ 
    plot(x[slct],y[slct],pch=15,col="#008080",cex=0.8,main="",xlab="",ylab="",
         xlim=c(min(x)-0.5,max(x)+0.5), axes=F)
    x.cent <- mean(x[slct],na.rm = TRUE)
    y.cent <- mean(y[slct],na.rm = TRUE)  
    for (i in 1:length(slct)) {
      text(x[slct[i]],y[slct[i]],cex=0.7,groups[i]+2,labels=slct[i],pos=3)
    }
    for(i in 1:length(slct))
      lines(c(x.cent,x[slct[i]]),c(y.cent,y[slct[i]]),lty=2,col=2)
    text(x.cent,y.cent,cex=1.5,labels=cllbl, font=3,col="blue")
  })
}

demographicData <- function(slctRows, var) {
  cmap <- ReadSavedData()
  if(slctRows[1] == "all") {
    slctRows = 1:length(cmap$demographics.dat$SorterID)
  }
  data.frame(cmap$demographics.dat[slctRows,],row.names=1)
}

####
# statements in a cluster (can sort by distance from center)
cluster.statements <- function(cluster) {
  cm = ReadSavedData()
  dat = cm$ratings.dat; clustfit = cm$fit.WardClust; nclust = cm$n.clust
  nphrases = cm$n.phrases; pIDs = cm$statement.dat[[1]]; phrases = cm$statement.dat[[2]]
  groups <- cutree(clustfit,k=nclust)
  tmp = getClusters()
  clustno = which(tmp == cluster)
  slct = which(groups == clustno)
  phrases = gsub("\\\\","",phrases[slct])
  df = data.frame(as.integer(pIDs[slct]),phrases)
  colnames(df) = c("ID","Statement")
  df  
}

cluster.top10.statements <- function(cluster) {
  cm = ReadSavedData()
  dat = cm$ratings.dat; clustfit = cm$fit.WardClust; nclust = cm$n.clust
  nphrases = cm$n.phrases; pIDs = cm$statement.dat[[1]]; phrases = cm$statement.dat[[2]]
  clust.labels <- cluster.labels(cm$sort.dat,nphrases,clustfit,nclust)
  tmp = getClusters()
  clustno = which(tmp == cluster)
  clust.labels.top10 <- cluster.labels.pop(clust.labels,topn=10)[[as.integer(clustno)]]   
  top10labels <- clust.labels.top10$Label
  top10count <- as.numeric(clust.labels.top10$Count)
  colnames(clust.labels.top10) = c("Label","Count")
  clust.labels.top10
}


####
# summary of statements (includes ratings, groups)
# updated to include demographic features
rating.summary <- function(col, demographicGroup) {
  cmap <- ReadSavedData()
  cmap$demographicGroup <- demographicGroup
  cmap$selectedGroup = cmap$demographicGroups[[demographicGroup]]
  if (cmap$demographicGroup == "All")
    cmap$selectedGroup$rows <- 1:length(cmap$demographics.dat$SorterID)
  dat <- cmap$ratings.dat
  selected <- cmap$demographics.dat$SorterID[cmap$selectedGroup$rows]
  dat <- dat[dat[[1]] %in% selected,]
  
  lbl = clusterLabels()
  clustfit = cmap$fit.WardClust; nclust = cmap$n.clust
  nphrases = cmap$n.phrases
  pIDs = as.integer(cmap$statement.dat[[1]])
  statements = cmap$statement.dat[[2]]
  groups <- cutree(clustfit,k=nclust)
  phrases = gsub("\\\\","",statements)
  ratings <- tapply(dat[[col]],INDEX=dat[[2]],FUN=mean,na.rm=T)
  ratings.sd <- round(tapply(dat[[col]],INDEX=dat[[2]],FUN=sd,na.rm=T),2)
  final.ratings <- rep(NA,length(pIDs))
  final.ratings.sd <- rep(NA,length(pIDs))
  final.ratings[as.integer(names(ratings))] <- ratings
  final.ratings.sd[as.integer(names(ratings))] <- ratings.sd
  dframe = data.frame(pIDs,statements,final.ratings,final.ratings.sd,lbl[groups])
  colnames(dframe) = c("ID","Statement","Mean","SD","Cluster ID")
  dframe  
}

####
# obtain the mean/SD rating for each cluster
rating.summary.cluster <- function(cluster.id,col, demographicGroup) {
  cmap <- ReadSavedData()
  cmap$demographicGroup <- demographicGroup
  cmap$selectedGroup = cmap$demographicGroups[[demographicGroup]]
  if (cmap$demographicGroup == "All")
    cmap$selectedGroup$rows <- 1:length(cmap$demographics.dat$SorterID)
  dat <- cmap$ratings.dat
  selected <- cmap$demographics.dat$SorterID[cmap$selectedGroup$rows]
  dat <- dat[dat[[1]] %in% selected,]
  ratings <- tapply(dat[[col]],INDEX=dat[[2]],FUN=mean,na.rm=T)
  clustfit <- cmap$fit.WardClust
  nclust <- cmap$n.clust; x = cmap$x; y= cmap$y
  groups <- cutree(clustfit,k=nclust)
  # Get average rating per cluster:
  cluster.ratings.n <- rep(0,nclust)
  cluster.ratings <- rep(0,nclust)
  cluster.ratings.sd <- rep(0,nclust)
  cluster.ratings.SSq <- rep(0,nclust)
  x.cent <- rep(0,nclust)
  y.cent <- rep(0,nclust)
  for (i in 1:nclust) {
    cluster.ratings.n[i] <- length(ratings[groups==i])
    cluster.ratings[i] <- mean(ratings[groups==i],na.rm=TRUE)
    cluster.ratings.sd[i] <- sd(ratings[groups==i],na.rm=TRUE) # edited
    x.cent[i] <- mean(x[groups==i])
    y.cent[i] <- mean(y[groups==i])  
    x.groups <- x[groups==i]
    y.groups <- y[groups==i]
    # need to decide how to compute this. Should SSq / n be computed with n = the number of statements in the cluster rated by at least one user?
    for (j in 1:length(x.groups)) {
      cluster.ratings.SSq[i] = cluster.ratings.SSq[i] + (x.cent[i]-x.groups[j])^2+(y.cent[i]-y.groups[j])^2
    }
  }
  dframe = data.frame(cluster.id,as.integer(cluster.ratings.n),
                      cluster.ratings.SSq/cluster.ratings.n,
                      cluster.ratings,cluster.ratings.sd)
  colnames(dframe) = c("Cluster ID","N","SSq/N","Mean","S.D")
  dframe
}


rating.Ftest.cluster <- function(cluster.id,col,demographicGroup) {
  cmap <- ReadSavedData()
  cmap$demographicGroup <- demographicGroup
  cmap$selectedGroup = cmap$demographicGroups[[demographicGroup]]
  if (cmap$demographicGroup == "All")
    cmap$selectedGroup$rows <- 1:length(cmap$demographics.dat$SorterID)
  dat <- cmap$ratings.dat
  selected <- cmap$demographics.dat$SorterID[cmap$selectedGroup$rows]
  dat <- dat[dat[[1]] %in% selected,]
  clustfit = cmap$fit.WardClust; nclust = cmap$n.clust
  clusters <- as.factor(cutree(clustfit,k=nclust))
  ratings <- tapply(dat[[col]],INDEX=dat[[2]],FUN=mean,na.rm=T)
  clusters <- clusters[as.integer(names(ratings))]
  ftest = summary(aov(ratings~clusters))
  rnms = rownames(ftest[[1]])
  ftest = data.frame(ftest[[1]][,1],format(ftest[[1]][,2],digits=5),
                     format(ftest[[1]][,3],digits=5),
                     c(format(ftest[[1]][1,4],digits=5),""),
                     c(format(ftest[[1]][1,5],digits=5),""))
  colnames(ftest) = c("DF","Sum Sq.","Mean Sq.","F value","p-value")
  rownames(ftest) = rnms
  ftest
}

rating.TukeyHSDtest.cluster <- function(cluster.id,col,demographicGroup) {
  cmap <- ReadSavedData()
  cmap$demographicGroup <- demographicGroup
  cmap$selectedGroup = cmap$demographicGroups[[demographicGroup]]
  if (cmap$demographicGroup == "All")
    cmap$selectedGroup$rows <- 1:length(cmap$demographics.dat$SorterID)
  dat <- cmap$ratings.dat
  selected <- cmap$demographics.dat$SorterID[cmap$selectedGroup$rows]
  dat <- dat[dat[[1]] %in% selected,]
  
  clustfit = cmap$fit.WardClust; nclust = cmap$n.clust
  lbl = clusterLabels()
  clusters <- as.factor(lbl[cutree(clustfit,k=nclust)])
  ratings <- tapply(dat[[col]],INDEX=dat[[2]],FUN=mean,na.rm=T)
  clusters <- clusters[as.integer(names(ratings))]
  tukeytest = TukeyHSD(aov(ratings~clusters))
  rnms = rownames(tukeytest[[1]])
  tukeytest = data.frame(tukeytest[[1]])
  colnames(tukeytest) = c("Difference","Lower CI","Upper CI","Adjusted p-value")
  rownames(tukeytest) = rnms
  tukeytest
}

####
# save a session
saveSession <- function(filename,description = "") {
  cm  = ReadSavedData()
  cm$filename = filename
  cm$description = description
  with(cm,{ # comment out excel.file.name at end of first line of args in both save calls to revert code back to original state
    save(description,filename,filename,statement.dat,clustname, n.clust, excel.file.name,reportObjects,demographicGroups,
         sort.dat, inc.mat, group.sim.mat,gsm, plotChoices,
         fit.MDS, x, y, stress, n.indiv, n.phrases, fit.WardClust,
         ratings.dat,colpoints,coltext,showStatementsVal,
         ratings.scales.dat, mins, maxs, demographics.dat, critvar1, critvar2,
         file=sprintf("%s.RData",filename))
    save(description,filename,filename,statement.dat,clustname, n.clust, excel.file.name,reportObjects,demographicGroups,
         sort.dat, inc.mat, group.sim.mat,gsm, plotChoices,
         fit.MDS, x, y, stress, n.indiv, n.phrases, fit.WardClust,
         ratings.dat,colpoints,coltext,showStatementsVal,
         ratings.scales.dat, mins, maxs, demographics.dat, critvar1, critvar2,
         file="activesession.RData")})
}


##################################
# GUI and settings functions     #
##################################

####
# set colors for clusters
# colors (http://www.rapidtables.com/web/color/RGB_Color.htm)
cols = rep(c(
  "#64b5f6",
  "#FF82AB", # palevioletred 1
  "#00EE76", #springgreen 1  
  "#3CB371", #medium sea green  
  "#DA70D6", #orchid  
  "#FFC1C1", #rosybrown 1
  "#00CED1", #dark turquoise
  "#6495ED", #corn flower blue
  "#9370DB", #medium purple
  "#C0C0C0" #silver
),10)
fontcol <- "gray33"

showProjectInfo <- function() {
  cm=ReadSavedData()
  activeProject = getActiveProject()
  if (is.null(cm$ratings.scales.dat)) {
    pst = ""
  } else {
    pst = paste(colnames(cm$ratings.scales.dat)[2:dim(cm$ratings.scales.dat)[2]],
                collapse=", ")
  }
  
  ret = c(cm$excel.file.name, cm$n.indiv, cm$n.phrases, cm$description, pst)  
}

# Return the list of variables for which there is rating data
ratingVariables <- function() {
  cmap <- ReadSavedData()
  colnames(cmap$ratings.scales.dat)[2:dim(cmap$ratings.scales.dat)[2]]
}

getClusters <- function() {
  cm=ReadSavedData()
  lbls = 1:cm$n.clust
  for (i in 1:cm$n.clust) {
    key = sprintf("clust_%02d_%02d",cm$n.clust,i)
    if (length(cm$clustname[[key]]) > 0) {
      lbls[i] = cm$clustname[[key]]
    }
  }
  lbls
}

####
# show the top-level (drop-down) menu

####
# show a window with data-related error messages (if any exist)
DataProblems <- function(err) {
  # return err as a tble/text
}



####   PLOTS    ####
## show the multi-dimensional scaled map:
ShowMDSPlot <- function() {
  cmap  = ReadSavedData()
  with(cmap,{
    plot(x,y,pch=15,col="#008080",cex=0.8,main="",xlab="",ylab="",
         xlim=c(min(x)*1.2,max(x)*1.2),
         sub=paste("Stress=",format(stress,digits=3)),axes=F)
    text(x,y,pos=3,labels=seq(1,n.phrases,1),cex=0.8)  
  })
}

####
## Cluster Map (Rays)
ShowClusterPlot <- function() {
  cm  = ReadSavedData()
  with(cm,{ 
    color=TRUE
    plot(x,y,pch=15,col="#008080",cex=0.8,main="",xlab="",ylab="",
         xlim=c(min(x)*1.2,max(x)*1.2),
         sub=paste("Stress=",format(stress,digits=3)),axes=F)
    groups <- cutree(fit.WardClust,k=n.clust)
    for (i in 1:n.phrases) {
      if (color==TRUE) {
        text(x[i],y[i],cex=0.7,groups[i]+2,labels=i,pos=3)
      } else {
        text(x[i],y[i],cex=0.7,labels=i,pos=3)
      }
    }
    
    # Determine Cluster Centers
    x.cent <- rep(0,n.clust)
    y.cent <- rep(0,n.clust)
    for (i in 1:n.clust)  {
      x.cent[i] <- mean(x[groups==i])
      y.cent[i] <- mean(y[groups==i])  
    }
    
    # Make Lines Connecting Cluster Centers and points in cluster
    for (i in 1:n.clust) {
      x.groups <- x[groups==i]
      y.groups <- y[groups==i]
      for (j in 1:length(x.groups)) {
        lines(c(x.cent[i],x.groups[j]),c(y.cent[i],y.groups[j]),lty=2,col=2)
      }
    }    
    # Plot Cluster Centers
    lbl = clusterLabels()
    text(x.cent,y.cent,cex=1.5,labels=lbl, font=3,col="blue")
  })
}

####
## Cluster Map (Polygons)
ShowClusterPolygonsPlot <- function() {
  cm  = ReadSavedData()
  with(cm,{ 
    color=TRUE
    lbl = clusterLabels()
    
    dep <- (max(y)-min(y))/50
    plot(x,y,type="n",main="",xlab="",ylab="",xlim=c(min(x)*1.2,max(x)*1.2),
         sub=paste("Stress=",format(stress,digits=3)),axes=F,
         col=colpoints) # (x,y) from MDS output
    groups <- cutree(fit.WardClust,k=n.clust)
    for (i in 1:n.phrases) {
      text(x[i],y[i],cex=0.7,col=coltext,labels=i,pos=3)
    }
    
    # Determine Cluster Centers
    x.cent <- rep(0,n.clust)
    y.cent <- rep(0,n.clust)
    for (i in 1:n.clust)	{
      x.cent[i] <- mean(x[groups==i])
      y.cent[i] <- mean(y[groups==i])	
    }
    
    # Make Polygons to fill in the clusters
    for (i in 1:n.clust) {
      ptsx = x[which(groups==i)]; ptsy = y[which(groups==i)]
      hpts = chull(ptsx, ptsy)
      hpts <- c(hpts, hpts[1])
      extpts <- kronecker(cbind(ptsx[hpts],ptsy[hpts]),rep(1,8))+
        kronecker(rep(1,length(hpts)),eps)
      hpts = chull(extpts)
      hpts <- c(hpts, hpts[1])
      polygon(extpts[hpts,],col=cols[i],border="white")
      points(x,y,pch=15,col=colpoints,cex=0.8)
    }
    
    for (i in 1:n.phrases) {
      text(x[i],y[i],cex=0.7,labels=i,font=3,pos=3)
    }
    
    
    # Plot Cluster Centers
    lbl = clusterLabels()
    text(x.cent,y.cent,cex=1.5,labels=lbl, font=3,col=fontcol)
  })
}

####
## plot the dendrogram
ShowDendrogram <- function(size) {
  cm  = ReadSavedData()
  cm$size = size
  with(cm,{
    if(size == "half")
      par(mfrow=c(2,1))
    plot(fit.WardClust,main="",xlab="",ylab="",axes=F
    ) 
    groups <- cutree(fit.WardClust,k=n.clust)
    rect.hclust(fit.WardClust,k=n.clust,border="red")  
  })
}



####
## plot the scaled data, with average rating for each statement
PlotRatingsLadder <- function(uv, demographicGroup) {
  cmap <- ReadSavedData()
  cmap$uv <- uv
  cmap$demographicGroup <- demographicGroup
  with(cmap,{
    if (demographicGroup == "All") {
      selectedSorters <- demographics.dat$SorterID
    } else {
      group <- which(names(demographicGroups) %in% demographicGroup)
      selectedSorters <- demographics.dat$SorterID[demographicGroups[[group]]$rows]
    }
    ratingsdat <- ratings.dat[which(ratings.dat$UserID %in% selectedSorters),]
    groups <- cutree(fit.WardClust,k=n.clust) 
    ratings <- tapply(ratingsdat[[uv]],
                      INDEX=ratingsdat[[2]],FUN=mean,na.rm=T)
    colmin = as.numeric(mins[[which(colnames(ratings.scales.dat)==uv)-1]])
    colmax = as.numeric(maxs[[which(colnames(ratings.scales.dat)==uv)-1]])
    unitHeight = 10*(colmax-colmin+1)/(30*n.phrases/2)
    unitWidth = 2*(max(x)-min(x)+1)/1000
    plot(x[!is.na(ratings)],y[!is.na(ratings)],type="n",
         xlim=c(min(x)*1.2,max(x)*1.2),
         main=sprintf("Rating Map for Group %s - %s",
                      demographicGroup,uv),
         xlab="",ylab="",
         ylim=c(min(y)-unitHeight,max(y)+unitHeight*colmax),
         axes=F,col=colpoints)
    
    for (i in 1:n.phrases) {
      if (!is.na(ratings[i])) {
        rect(x[i]-2*unitWidth, y[i],
             x[i]+2*unitWidth, y[i]+ratings[i]*unitHeight,
             col=cols[groups[i]],border = cols[groups[i]])
        for (j in 1:floor(ratings[i])) {
          lines(c(x[i]-2*unitWidth,x[i]+2*unitWidth),
                c(y[i]+j*unitHeight,y[i]+j*unitHeight),
                lwd=1,col="white")
        }
      }
    }
    text(x[!is.na(ratings)],y[!is.na(ratings)],
         labels=seq(1,n.phrases)[!is.na(ratings)],cex=0.7,col=coltext)
  })
}

####
## Cluster Rating Map
ShowClusterRankingsPlotShades <- function(uv, demographicGroup) {
  cmap <- ReadSavedData()
  cmap$uniratingcol <- which(colnames(cmap$ratings.dat) == uv)
  dat <- cmap$sort.dat
  cmap$demographicGroup <- demographicGroup
  cmap$selectedGroup <- cmap$demographicGroups[[demographicGroup]]
  if (cmap$demographicGroup == "All")
    cmap$selectedGroup$rows <- 1:length(cmap$demographics.dat$SorterID)
  cmap$selected <- cmap$demographics.dat$SorterID[cmap$selectedGroup$rows]
  
  with(cmap,{ 
    x.title=""
    y.title = ""
    main.title = "Cluster Rating Map"
    color=TRUE
    groups <- cutree(fit.WardClust,k=n.clust)    # determines cluser label for each phrase
    ratings.dat <- ratings.dat[which(ratings.dat$UserID %in% selected), ]
    ratings <- tapply(ratings.dat[[uniratingcol]],INDEX=ratings.dat[[2]],FUN=mean,na.rm=TRUE)    # determines average ranking for each phrase    
    
    colname = colnames(ratings.dat)[uniratingcol]
    colmin = as.numeric(mins[[which(colnames(ratings.scales.dat)==colname)-1]])
    colmax = as.numeric(maxs[[which(colnames(ratings.scales.dat)==colname)-1]])
    rescale = (colmax-colmin+1)/5
    ratings = ratings/rescale
    colmax = colmax/rescale
    colmin = colmin/rescale
    dep <- (max(y)-min(y))/50	# determines depth of each level of the cluster
    # Get average rating per cluster:
    cluster.ratings <- rep(0,n.clust)
    for (i in 1:n.clust) {
      cluster.ratings[i] <- mean(ratings[groups==i], na.rm = TRUE)
    }
    
    # Determine cluster centers
    x.cent <- rep(0,n.clust)
    y.cent <- rep(0,n.clust)
    for (i in 1:n.clust)	{
      x.cent[i] <- mean(x[groups==i])
      y.cent[i] <- mean(y[groups==i])	
    }
    
    # Plot the clusters as polygons as usual:
    plot.cluster.polygon(x,y,fit.WardClust,n.clust,n.phrases,
                         subymin=dep*(colmax-colmin),main.title=
                           sprintf("Cluster Rating Map for Group %s - %s",demographicGroup,colnames(ratings.dat)[uniratingcol]))
    # Determine Cluster Centers
    x.cent <- rep(0,n.clust)
    y.cent <- rep(0,n.clust)
    for (i in 1:n.clust)  {
      x.cent[i] <- mean(x[groups==i])
      y.cent[i] <- mean(y[groups==i])	
      ptsx = x[which(groups==i)]; ptsy = y[which(groups==i)]
      hpts = chull(ptsx, ptsy)
      hpts <- c(hpts, hpts[1])
      extpts <- kronecker(cbind(ptsx[hpts],ptsy[hpts]),rep(1,8))+
        kronecker(rep(1,length(hpts)),eps)
      hpts = chull(extpts)
      hpts <- c(hpts, hpts[1])
      if (!is.na(cluster.ratings[i])) {
        for (j in round(cluster.ratings[i]):1) {
          polygon(extpts[hpts,1],extpts[hpts,2]-dep*(j-1),col=cols[i],
                  border="white")
          points(x,y,pch=15,col=colpoints,cex=0.8)
        }
      }
    }
    
    # Add text again
    for (i in 1:n.phrases) {
      if (color==TRUE) {
        text(x[i],y[i],cex=0.7,labels=i,pos=3,col=coltext)
      }
    }
    
    # Plot Cluster Centers again
    lbl = clusterLabels()
    text(x.cent,y.cent,cex=1.5,labels=lbl, font=3,col=fontcol)
    #    points(x,y,pch=15,col=colpoints,cex=0.8)
  })
}


BarChart <- function(vars, demographicGroup) {
  cm  = ReadSavedData()
  cm$vars = vars[vars!="None"]
  cm$uniratingcols = sapply(cm$vars,function(uv) which(colnames(cm$ratings.dat) == uv))
  cm$demographicGroup = demographicGroup
  cm$selectedGroup = cm$demographicGroups[[demographicGroup]]
  if (cm$demographicGroup == "All")
    cm$selectedGroup$rows <- 1:length(cm$demographics.dat$SorterID)
  with(cm,{ 
    slct <- demographics.dat$SorterID[selectedGroup$rows]
    ratings.dat <- ratings.dat[which(ratings.dat$UserID %in% slct), ] # only keep ratings information for users in selected group
    x.title=""
    y.title = ""
    color=TRUE
    groups <- cutree(fit.WardClust,k=n.clust)    # determines cluser label for each phrase
    k = l = 1
    if (length(uniratingcols) >= 2) {
      k = l = 2
    }
    par(mfrow=c(k,l), oma=c(0,0,2,0))
    for (uniratingcol in uniratingcols) {
      ratings <- tapply(ratings.dat[[uniratingcol]],INDEX=ratings.dat[[2]],FUN=mean,na.rm=T)    # determines average ranking for each phrase
      colname = colnames(ratings.dat)[uniratingcol]
      colmin = mins[[which(colnames(ratings.scales.dat)==colname)-1]]
      colmax = maxs[[which(colnames(ratings.scales.dat)==colname)-1]]
      rescale = as.numeric((colmax-colmin+1)/5)
      ratings = ratings/rescale
      colmax = colmax/rescale
      colmin = colmin/rescale
      dep <- (max(y)-min(y))/50  # determines depth of each level of the cluster
      # Get average rating per cluster:
      cluster.ratings <- rep(0,n.clust)
      sds <- rep(0,n.clust)
      
      for (i in 1:n.clust) {
        cluster.ratings[i] <- mean(ratings[groups==i], na.rm = TRUE)
        if (length(which(groups == i)) == 1 | is.na(cluster.ratings[i]))
          sds[i] <- 0
        else
          sds[i] <- sd(ratings[groups==i], na.rm = TRUE)
      }
      cluster.ratings[is.na(cluster.ratings)] <- 0
      
      Yl = ceiling(max(cluster.ratings+sds+0.5, na.rm = TRUE)) # edited: added na.rm = TRUE
      bp <- barplot(cluster.ratings, main=colnames(ratings.dat)[uniratingcol],col=cols,ylim=c(0,Yl),
                    names.arg=clusterLabels(),cex.names=1.5,las=3,
                    width=rep(0.8,n.clust),space=0.2)
      segments(bp, cluster.ratings - sds, bp, cluster.ratings + sds, lwd=2)
      segments(bp - 0.1, cluster.ratings - sds, bp + 0.1, cluster.ratings - sds, lwd=2)
      segments(bp - 0.1, cluster.ratings + sds, bp + 0.1, cluster.ratings + sds, lwd=2)
    }
    mtext(sprintf("Bar Chart for Group %s", demographicGroup), outer=TRUE, cex=1.5)
  })
}

# plot rating variables on a bar chart
BarChartInOne <- function(vars) { # 'vars' is a vector containing rating variable names
  cm  = ReadSavedData()
  cm$vars = vars[vars!="None"] # 'None' entries should be removed from the input vector
  cm$uniratingcols = sapply(cm$vars,function(uv) which(colnames(cm$ratings.dat) == uv))
  with(cm,{ 
    x.title=""
    y.title = ""
    main.title = "Cluster Rating Map"
    color=TRUE
    groups <- cutree(fit.WardClust,k=n.clust)    # determines cluser label for each phrase
    ratings.matrix = c()
    sd.matrix = c()
    Yl = 0
    for (uniratingcol in uniratingcols) {
      ratings <- tapply(ratings.dat[[uniratingcol]],INDEX=ratings.dat[[2]],FUN=mean,na.rm=T)    # determines average ranking for each phrase
      colname = colnames(ratings.dat)[uniratingcol]
      colmin = mins[[which(colnames(ratings.scales.dat)==colname)-1]]
      colmax = maxs[[which(colnames(ratings.scales.dat)==colname)-1]]
      rescale = (colmax-colmin+1)/5
      ratings = ratings/rescale
      colmax = colmax/rescale
      colmin = colmin/rescale
      dep <- (max(y)-min(y))/50  # determines depth of each level of the cluster
      # Get average rating per cluster:
      cluster.ratings <- rep(0,n.clust)
      sds <- rep(0,n.clust)
      for (i in 1:n.clust) {
        cluster.ratings[i] <- mean(ratings[groups==i])
        sds[i] <- sd(ratings[groups==i])
      }
      sds[is.na(sds)] = 0
      Yl = ceiling(max(cluster.ratings+sds+0.5,Yl))
      ratings.matrix = rbind(ratings.matrix,cluster.ratings)
      sd.matrix = rbind(sd.matrix,sds)
    }
    colnames(ratings.matrix) <- getClusters()
    bp <- barplot(ratings.matrix,beside=TRUE,main=createTitle(vars),
                  col=sapply(1:n.clust,function(x) rep(cols[x],length(vars))),
                  ylim=c(0,Yl),xlab="Clusters"
    )
    # add shading
    barplot(ratings.matrix,beside=TRUE,add=TRUE,col=("black"),
            density=c(0,10,10,20)[1:length(vars)],angle=c(0,45,-45,45)[1:length(vars)],
            legend=vars,args.legend=list(title="Rating Variable"))
    # store only positive entries to create sd bars
    bp_pos = bp[sd.matrix>0]
    ratings.matrix_pos = ratings.matrix[sd.matrix>0]
    sd.matrix_pos = sd.matrix[sd.matrix>0]
    # add standard deviation bars
    segments(bp_pos, ratings.matrix_pos - sd.matrix_pos, bp_pos, ratings.matrix_pos + sd.matrix_pos, lwd=2)
    segments(bp_pos - 0.1, ratings.matrix_pos - sd.matrix_pos, bp_pos + 0.1, ratings.matrix_pos - sd.matrix_pos, lwd=2)
    segments(bp_pos - 0.1, ratings.matrix_pos + sd.matrix_pos, bp_pos + 0.1, ratings.matrix_pos + sd.matrix_pos, lwd=2)
  })
}

####
# an pattern matching graph
ClusterLadderGraph <- function(bv1, bv2, demographicGroup) {
  cmap <- ReadSavedData()
  cmap$bivratingcol[[1]] = which(colnames(cmap$ratings.dat) == bv1)
  cmap$bivratingcol[[2]] = which(colnames(cmap$ratings.dat) == bv2)
  dat <- cmap$sort.dat
  cmap$demographicGroup <- demographicGroup
  cmap$selectedGroup <- cmap$demographicGroups[[demographicGroup]]
  if (cmap$demographicGroup == "All")
    cmap$selectedGroup$rows <- 1:length(cmap$demographics.dat$SorterID)
  cmap$selected <- cmap$demographics.dat$SorterID[cmap$selectedGroup$rows]
  
  with(cmap,{ 
    x.title=""
    y.title = ""
    main.title = "Pattern Matching"
    color=TRUE
    groups <- cutree(fit.WardClust,k=n.clust)
    
    ratings.dat <- ratings.dat[which(ratings.dat$UserID %in% selected), ]
    ratings1 <- tapply(ratings.dat[[bivratingcol[1]]],INDEX=ratings.dat[[2]],FUN=mean,na.rm=T)
    ratings2 <- tapply(ratings.dat[[bivratingcol[2]]],INDEX=ratings.dat[[2]],FUN=mean,na.rm=T)
    
    # Get average rating per cluster:
    cluster.ratings1 <- rep(0,n.clust)
    cluster.ratings2 <- rep(0,n.clust)
    for (i in 1:n.clust) {
      cluster.ratings1[i] <- mean(ratings1[groups==i], na.rm = TRUE)
      cluster.ratings2[i] <- mean(ratings2[groups==i], na.rm = TRUE)
    }
    m = min(c(cluster.ratings1,cluster.ratings2), na.rm = TRUE)
    M = max(c(cluster.ratings1,cluster.ratings2), na.rm = TRUE)
    m[m == Inf] = 0
    M[M == -Inf] = 1
    
    strs <- rep("", n.clust)
    for (i in 1:n.clust) {
      if (is.na(cluster.ratings1[i]) | is.na(cluster.ratings2[i])) { next }
      key = sprintf("clust_%02d_%02d",n.clust,i)
      if (length(clustname[[key]]) > 0) {
        strs[i] = clustname[[key]]
      } else {
        strs[i] = i
      }
    }
    hpx=700; vpx = 900
    cx = 1.5 
    minH <- ceiling(par()$cra[2])*cx/hpx # minimum vertical distance for text
    L <- unlist(lapply(strs,nchar))
    m <- min(cluster.ratings1, cluster.ratings2)
    M <- max(cluster.ratings1, cluster.ratings2)
    ordL <- order(cluster.ratings1)
    cr1 <- cluster.ratings1[ordL] # sorted ratings, first variable
    diffs <- c(0, cr1[2:n.clust] - cr1[1:(n.clust-1)])
    diffs[which(diffs <= minH)] <- minH - diffs[which(diffs <= minH)]
    diffs[which(diffs > minH)] <- 0
    diffs[1] <- 0
    rtg1offset <- cr1 + diffs
    maxtextW <- 2+ceiling(max(L)*par()$cra[1])*cx/vpx
    plot(c(0,5), c(m,M), axes=F, xlab="", ylab="", col=0, 
         ylim=c(m-0.05, M+0.1), xlim=c(0,5))
    text(maxtextW, rtg1offset,paste(strs[ordL],format(cluster.ratings1[ordL], digits=2)),
         pos=2, cex=cx, col=cols[ordL])
    for (i in 1:n.clust) {
      lines(c(maxtextW, 4), c(cluster.ratings1[i],cluster.ratings2[i]), 
            col=cols[i], lwd=3)
    }
    text(4, cluster.ratings2, format(cluster.ratings2, digits=2), pos=4,
         cex=cx, col=cols)
    
    text(maxtextW-0.3,M+0.07,colnames(ratings.dat)[bivratingcol[1]],cex=1.5)
    text(4.2,M+0.07,colnames(ratings.dat)[bivratingcol[2]],cex=1.5)
  })
}

# create gozone graph using a list of clusters
GoZoneGraphClusterList <- function(clustIndices,bv1,bv2,demographicGroup) {
  # arguments:
  #   clustIndices = numeric vector of integers denoting the indices of the clusters to be plotted
  #                  (i.e. clustIndices = c(1,3) will plot the first and third clusters listed in clusterLabels())
  #   bv1, bv2 = character strings denoting names of variables to be plotted on the x- and y-axes, resp.
  cmap <- ReadSavedData()
  cmap$bivratingcol[[1]] = which(colnames(cmap$ratings.dat) == bv1)
  cmap$bivratingcol[[2]] = which(colnames(cmap$ratings.dat) == bv2)
  cmap$clustIndices = clustIndices
  dat <- cmap$sort.dat
  cmap$demographicGroup <- demographicGroup
  cmap$selectedGroup <- cmap$demographicGroups[[demographicGroup]]
  if (cmap$demographicGroup == "All")
    cmap$selectedGroup$rows <- 1:length(cmap$demographics.dat$SorterID)
  cmap$selected <- cmap$demographics.dat$SorterID[cmap$selectedGroup$rows]
  
  with(cmap,{ 
    lbl <- clusterLabels()
    ratings.dat <- ratings.dat[which(ratings.dat$UserID %in% selected), ]
    ratings1 <- tapply(ratings.dat[[bivratingcol[1]]],
                       INDEX=ratings.dat[[2]],FUN=mean,na.rm=T)
    ratings2 <- tapply(ratings.dat[[bivratingcol[2]]],
                       INDEX=ratings.dat[[2]],FUN=mean,na.rm=T)    
    
    m = c(min(ratings1, na.rm = TRUE),min(ratings2, na.rm = TRUE))
    if (any(m == Inf)) { plot.new(); return() }
    M = c(max(ratings1, na.rm = TRUE),max(ratings2, na.rm = TRUE))
    # overall means
    m1 = mean(ratings1, na.rm = TRUE)
    m2 = mean(ratings2, na.rm = TRUE)
    # selected clusters:    
    groups <- cutree(fit.WardClust,k=n.clust)
    idxplot <- which(groups %in% clustIndices)
    m1s = mean(ratings1[idxplot], na.rm = TRUE)
    m2s = mean(ratings2[idxplot], na.rm = TRUE)
    s1 = sd(ratings1[idxplot], na.rm = TRUE)
    s2 = sd(ratings2[idxplot], na.rm = TRUE)
    x.title=""
    y.title = ""
    main.title = sprintf("GoZone for Group %s",demographicGroup) #lbl[cl]
    color=TRUE
    plot(ratings1[idxplot],ratings2[idxplot],col=0,xlim=c(m[1]-0.2,M[1]+0.2),
         main=main.title,
         ylim=c(m[2]-0.2,M[2]+0.2),axes=F,
         xlab=colnames(ratings.dat)[bivratingcol[1]],
         ylab=colnames(ratings.dat)[bivratingcol[2]])
    axis(1);axis(2)
    rect(m[1]-0.2, m[2]-0.2, m1s,m2s, col = "#FF4500", border = 0)
    rect(m1s, m[2]-0.2, M[1]+0.2,m2s, col = "#FFD700", border = 0)
    rect(m[1]-0.2, m2s, m1s,M[2]+0.2, col = "#FFD700", border = 0)
    rect(m1s,m2s,M[1]+0.2, M[2]+0.2,  col = "#7CFC00", border = 0)
    text(ratings1[idxplot],ratings2[idxplot],names(ratings1[idxplot]),cex=1,col="grey33",pos=3)
    text(ratings1[idxplot],ratings2[idxplot],toupper(letters[1:length(clustIndices)][match(groups[groups %in% clustIndices], clustIndices)]))
    lines(c(m[1]-0.2,M[1]+0.2),c(m2,m2),lty=2,lwd=3, col="white")
    lines(c(m[1]-0.2,M[1]+0.2),c(m2s,m2s),lty=3,lwd=2)
    lines(c(m[1]-0.2,M[1]+0.2),c(m2s+s2,m2s+s2),lty=3)
    lines(c(m[1]-0.2,M[1]+0.2),c(m2s-s2,m2s-s2),lty=3)
    lines(c(m1,m1),c(m[2]-0.2,M[2]+0.2),lty=2,lwd=3, col="white")
    lines(c(m1s,m1s),c(m[2]-0.2,M[2]+0.2),lty=3,lwd=2)
    lines(c(m1s+s1,m1s+s1),c(m[2]-0.2,M[2]+0.2),lty=3)
    lines(c(m1s-s1,m1s-s1),c(m[2]-0.2,M[2]+0.2),lty=3)
  })
}

####
# show clusters as polygons. used by ShowClusterRankingsPlotShades
# to plot the clusters with average ratings
plot.cluster.polygon <- function(x,y,clustfit,nclust,nphrases,color=TRUE,
                                 main.title="Clusters",x.title="",y.title="",subymin=0) {
  # Note:  may want to change this so that it doesn't require input of MDS output
  # clustfit is the output of the hclust function
  cm  = ReadSavedData()
  dep <- (max(y)-min(y))/50
  plot(x,y,type="n",main=main.title,xlab=x.title,ylab=y.title,axes=F,
       xlim=c(min(x)*1.2,max(x)*1.2),
       ylim=c(min(y)-subymin,max(y)),  col=cm$colpoints) # (x,y) from MDS output
  #  rect(par("usr")[1],par("usr")[3],par("usr")[2],par("usr")[4],col = "#e0e0e0")
  groups <- cutree(clustfit,k=nclust)
  
  # Determine Cluster Centers
  x.cent <- rep(0,nclust)
  y.cent <- rep(0,nclust)
  for (i in 1:nclust)	{
    x.cent[i] <- mean(x[groups==i])
    y.cent[i] <- mean(y[groups==i])	
  }
  
  # Make Polygons to fill in the clusters
  for (i in 1:cm$n.clust) {
    ptsx = x[which(groups==i)]; ptsy = y[which(groups==i)]
    hpts = chull(ptsx, ptsy)
    hpts <- c(hpts, hpts[1])
    extpts <- kronecker(cbind(ptsx[hpts],ptsy[hpts]),rep(1,8))+
      kronecker(rep(1,length(hpts)),eps)
    hpts = chull(extpts)
    hpts <- c(hpts, hpts[1])
    polygon(extpts[hpts,],col=cols[i],border="white")
  }
  
  # Plot Cluster Centers
  text(x.cent,y.cent,cex=1.1,labels=seq(1,nclust))
}

####
# get the boundary points of clusters (as polygons). used by ShowClusterRankingsPlotShades
# to plot the clusters with average ratings
get.polypoints <- function(x,y,clustfit,nclust,nphrases) {
  
  groups <- cutree(clustfit,k=nclust)
  
  # Create Matrix to store the x and y polygon points
  x.polypoints <- matrix(rep(NA,nclust*nphrases),nrow=nclust)
  y.polypoints <- matrix(rep(NA,nclust*nphrases),nrow=nclust)
  
  # Make Polygons to fill in the clusters
  
  for (i in 1:nclust) {
    x.groups <- x[groups==i]
    y.groups <- y[groups==i]
    
    x.top <- x.groups[which.max(y.groups)]
    y.top <- max(y.groups)	
    
    x.bot <- x.groups[which.min(y.groups)]
    y.bot <- min(y.groups)
    
    x.left <- min(x.groups)
    y.left <- y.groups[which.min(x.groups)]
    
    x.right <- max(x.groups)
    y.right <- y.groups[which.max(x.groups)]
    
    
    # Top-Left Points #
    tl.slope <- (y.top-y.left)/(x.top-x.left)
    tl.int <- y.top - tl.slope*x.top
    x.inplay <- x.groups[x.groups>x.left & x.groups<x.top & y.groups<y.top & y.groups>y.left]
    y.inplay <- y.groups[x.groups>x.left & x.groups<x.top & y.groups<y.top & y.groups>y.left]
    y.linevalue <- tl.slope*x.inplay + tl.int
    y.diff <- y.inplay - y.linevalue
    x.TBadd <- x.inplay[y.diff>0]
    y.TBadd <- y.inplay[y.diff>0]
    x.add <- sort(x.TBadd)
    y.add <- y.TBadd[order(x.TBadd)]
    x.poly <- c(x.left,x.add,x.top)
    y.poly <- c(y.left,y.add,y.top)
    
    # Top-Right Points #
    tr.slope <- (y.top-y.right)/(x.top-x.right)
    tr.int <- y.top - tr.slope*x.top	
    x.inplay <- x.groups[x.groups<x.right & x.groups>x.top & y.groups<y.top & y.groups>y.right]
    y.inplay <- y.groups[x.groups<x.right & x.groups>x.top & y.groups<y.top & y.groups>y.right]
    y.linevalue <- tr.slope*x.inplay + tr.int
    y.diff <- y.inplay - y.linevalue
    x.TBadd <- x.inplay[y.diff>0]
    y.TBadd <- y.inplay[y.diff>0]
    x.add <- sort(x.TBadd)
    y.add <- y.TBadd[order(x.TBadd)]
    x.poly <- c(x.poly,x.add,x.right)
    y.poly <- c(y.poly,y.add,y.right)
    
    # Bot-Right Points #
    br.slope <- (y.bot-y.right)/(x.bot-x.right)
    br.int <- y.bot - br.slope*x.bot	
    x.inplay <- x.groups[x.groups<x.right & x.groups>x.bot & y.groups<y.right & y.groups>y.bot]
    y.inplay <- y.groups[x.groups<x.right & x.groups>x.bot & y.groups<y.right & y.groups>y.bot]
    y.linevalue <- br.slope*x.inplay + br.int
    y.diff <- y.inplay - y.linevalue
    x.TBadd <- x.inplay[y.diff<0]
    y.TBadd <- y.inplay[y.diff<0]
    x.add <- -sort(-x.TBadd)
    y.add <- y.TBadd[order(-x.TBadd)]
    x.poly <- c(x.poly,x.add,x.bot)
    y.poly <- c(y.poly,y.add,y.bot)
    
    # Bot-Left Points #
    bl.slope <- (y.bot-y.left)/(x.bot-x.left)
    bl.int <- y.bot - bl.slope*x.bot	
    x.inplay <- x.groups[x.groups<x.bot & x.groups>x.left & y.groups<y.left & y.groups>y.bot]
    y.inplay <- y.groups[x.groups<x.bot & x.groups>x.left & y.groups<y.left & y.groups>y.bot]
    y.linevalue <- bl.slope*x.inplay + bl.int
    y.diff <- y.inplay - y.linevalue
    x.TBadd <- x.inplay[y.diff<0]
    y.TBadd <- y.inplay[y.diff<0]
    x.add <- -sort(-x.TBadd)
    y.add <- y.TBadd[order(-x.TBadd)]
    x.poly <- c(x.poly,x.add)
    y.poly <- c(y.poly,y.add)
    
    x.NA <- rep(NA,nphrases-length(x.poly))
    y.NA <- rep(NA,nphrases-length(y.poly))
    
    x.polypoints[i,] <- c(x.poly,x.NA)	
    y.polypoints[i,] <- c(y.poly,y.NA)	
  }
  return(list("xclust"=x.polypoints,"yclust"=y.polypoints))
}

showStatements <- function() {
  cm  = ReadSavedData()
  with(cm,{
    statmentsID = statement.dat$StatementID
    statments = statement.dat$Statement
    groups = cutree(fit.WardClust,k=n.clust)
    clustname = seq(1,n.clust)
    
    tt  <- tktoplevel()
    tkwm.title(tt,"Statements")
    tkwm.geometry(tt,'-40+40')
    yscr <- tkscrollbar(tt, repeatinterval=5,
                        command=function(...)tkyview(txt,...))
    txt <- tktext(tt,bg="white",font=fontTextSmall,
                  yscrollcommand=function(...)tkset(yscr,...),width=50,height=40,wrap="word")
    tkgrid(txt,yscr)
    tkgrid.configure(yscr,sticky="ns",column=1)
    lbl = clusterLabels()
    allstatements = ""
    for(cl in 1:n.clust) {
      allstatements = sprintf("%s\n\n%s\n",allstatements,lbl[cl])
      inclust = which(groups == cl)
      for (i in inclust) {
        stm = sprintf("%3d  %s",statmentsID[i],statments[i])
        allstatements = sprintf("%s\n%s",allstatements,stm)
      }
    }
    tkinsert(txt,"end",allstatements)
    tkconfigure(txt, state="disabled")
    tkfocus(txt)    
  })
}


clusterLabels <- function() {
  cm  = ReadSavedData()
  lbl = rep("",cm$n.clust)
  for(cl in 1:cm$n.clust) {
    key = sprintf("clust_%02d_%02d",cm$n.clust,cl)
    if (length(cm$clustname[[key]]) > 0) {
      lbl[cl] = cm$clustname[[key]]
    } else {
      lbl[cl] = cl
    }
  }
  lbl
}


####
# set cluster name
SetNumClust  <- function(n.clust) {
  cm  = ReadSavedData()
  cm$n.clust = as.integer(n.clust)
  with(cm,{
    save(description,excel.file,excel.file.name,statement.dat,clustname,n.clust,reportObjects,demographicGroups,
         sort.dat, inc.mat, group.sim.mat,gsm, 
         fit.MDS, x, y, stress, n.indiv, n.phrases, fit.WardClust,
         ratings.dat,
         n.clust, colpoints,coltext,showStatementsVal,
         ratings.scales.dat, mins, maxs, demographics.dat, critvar1, critvar2,
         file="activesession.RData") 
  })
  load("activesession.RData")
  cm  = ReadSavedData()
}


updateClusterName <- function(n.clust,incluster,name) {
  cm  = ReadSavedData()
  cm$n.clust = as.integer(n.clust)
  tmp = getClusters()
  cm$incluster = which(tmp == incluster)
  cm$name = name
  with(cm,{ 
    clustkey = sprintf("clust_%02d_%02d",n.clust,incluster)
    clustname[[clustkey]] <- name
    save(description,excel.file,excel.file.name,statement.dat,clustname,n.clust,reportObjects,demographicGroups,
         sort.dat, inc.mat, group.sim.mat,gsm, 
         fit.MDS, x, y, stress, n.indiv, n.phrases, fit.WardClust,
         ratings.dat,
         n.clust, colpoints,coltext,showStatementsVal,
         ratings.scales.dat, mins, maxs, demographics.dat, critvar1, critvar2,
         file="activesession.RData") 
  })
}

### When statements are sorted by cluster, creates arguments to be passed to renderTable that generate lines between clusters

add.to.row.args <- function(byCluster,reportType,ratingVar,demographicGroup) {
  if (is.null(byCluster) || !byCluster || is.null(reportType) || reportType!="Statements" || is.null(ratingVar) || is.null(demographicGroup)) # return NULL if statements are not sorted by cluster
    return(NULL)
  df = rating.summary(col = ratingVar, demographicGroup)
  clustCol = df[,5] # keep only column containing cluster names
  clustCol = clustCol[order(clustCol,decreasing=FALSE)]
  clustNames = unique(clustCol) 
  clustNames = clustNames[-length(clustNames)] # remove last cluster name (a line isn't needed after the last cluster)
  pos = lapply(clustNames,function(clName) max(which(clustCol==clName))) # return indices of the last statement in each cluster
  
  cm = ReadSavedData()
  command = rep(paste0("<TH COLSPAN=",ncol(df),"><hr>"),cm$n.clust-1)
  return(list(pos,command))
}


createTitle <- function(strs) {
  len = length(strs)
  if (len<=2)
    return(Reduce(function(str1,str2) paste(str1,"and",str2),strs))
  else {
    title = paste0(paste(sapply(strs[-len],function(x) paste0(x,", ")),collapse=""),"and ",strs[len])
    return(title)
  }
}

### Functions for saving and accessing plots/tables to be included in report
getSavedObject <- function(input) {
  # return element in reportObjects associated with the current settings given by input
  cm = ReadSavedData()
  vars = objectRetrievalVars(input,cm) 
  if (is.null(vars)) {
    return(NULL)
  }
  return(cm$reportObjects[[input$conditionedPanels]][[vars$radioButtonValue]][[vars$retrievalString]])
}

saveObject <- function(input) {
  # save current plot/table to reportObjects list
  cm = ReadSavedData()
  vars = objectRetrievalVars(input,cm)
  if (is.null(vars)) {
    return(NULL)
  }
  cm$reportObjects[[input$conditionedPanels]][[vars$radioButtonValue]][[vars$retrievalString]] = vars$variableList
  with(cm, {
    save(description, excel.file, excel.file.name, statement.dat, clustname, n.clust, reportObjects, demographicGroups,
         sort.dat, inc.mat, group.sim.mat, gsm, fit.MDS, x, y, stress, n.indiv, 
         n.phrases, fit.WardClust, ratings.dat,
         ratings.scales.dat, mins, maxs, demographics.dat, critvar1, critvar2,
         file="activesession.RData")
  })
}

removeSavedObject <- function(input) {
  # remove current plot/table from reportObjects list
  cm = ReadSavedData()
  vars = objectRetrievalVars(input,cm)
  if (is.null(vars)) {
    return(NULL)
  }
  cm$reportObjects[[input$conditionedPanels]][[vars$radioButtonValue]][[vars$retrievalString]] = NULL
  with(cm, {
    save(description, excel.file, excel.file.name, statement.dat, clustname, n.clust, reportObjects, demographicGroups,
         sort.dat, inc.mat, group.sim.mat, gsm, fit.MDS, x, y, stress, n.indiv, 
         n.phrases, fit.WardClust, ratings.dat,
         ratings.scales.dat, mins, maxs, demographics.dat, critvar1, critvar2,
         file="activesession.RData")
  })
}

objectRetrievalVars <- function(input,cm) {
  # returns a list containing keys used to access an object in reportObjects
  reportObjects = cm$reportObjects
  radioButtonValue =  switch(input$conditionedPanels,
                             "Plots" = input$plotType,
                             "Reports" = input$reportType,
                             "Analysis" = input$analysisType)
  if (is.null(radioButtonValue)) {
    return(NULL) # radioButtonValue may take value NULL when opening one of the conditioned panels for the first time
  }
  variableList = switch(input$conditionedPanels,
                        "Plots" = switch(radioButtonValue,
                                         "Point Map (MDS)" = list("pointmap" = "pointmap"),
                                         "Clusters" = list("n.clust" = cm$n.clust,"clusterplot" = input$clusterplot,"demographicSelection" = input$demographicSelection),
                                         "Bridging Index" = list("n.clust" = cm$n.clust),
                                         "Univariate Rating - by Statement" = list("univariateRating" = input$univariateRating,"demographicSelection" = input$demographicSelection),
                                         "Univariate Rating - by Cluster" = list("n.clust" = cm$n.clust,"univariateRating" = input$univariateRating,"demographicSelection" = input$demographicSelection),
                                         "Bar Chart" = list("n.clust" = cm$n.clust,"chooseRating1" = input$chooseRating1,"chooseRating2" = input$chooseRating2,
                                                            "chooseRating3" = input$chooseRating3,"chooseRating4" = input$chooseRating4,"demographicSelection" = input$demographicSelection),
                                         "Pattern Matching" = list("n.clust" = cm$n.clust,"bivariateRating1" = input$bivariateRating1,"bivariateRating2" = input$bivariateRating2,"demographicSelection" = input$demographicSelection),
                                         "GoZone" = list("n.clust" = cm$n.clust,"goZoneIndices" = as.integer(input$showGoZone),"bivariateRating1" = input$bivariateRating1,"bivariateRating2" = input$bivariateRating2,"demographicSelection" = input$demographicSelection),
                                         "Dendrogram" = list("n.clust" = cm$n.clust)),
                        "Reports" = switch(radioButtonValue,
                                           "Sorters" = list("sorters" = "sorters","demographicSelection" = input$demographicSelection),
                                           "Statements" = list("n.clust" = cm$n.clust,"univariateRating" = input$univariateRating,"sortBy" = input$sortBy,"byCluster" = input$byCluster,"demographicSelection" = input$demographicSelection),
                                           "Clusters" = list("n.clust" = cm$n.clust,"univariateRating" = input$univariateRating,"demographicSelection" = input$demographicSelection)),
                        "Analysis" = list("n.clust" = cm$n.clust,"univariateRating" = input$univariateRating,"demographicSelection" = input$demographicSelection))
  retrievalString = paste(paste(names(variableList),variableList,sep="="),collapse="_")
  return(list("radioButtonValue" = radioButtonValue,"variableList" = variableList,"retrievalString" = retrievalString))
}
###


createReport <- function(vg = FALSE, filename) {
  vg = vg & ifelse(.Platform$OS.type == "windows", TRUE, FALSE)
  # creates a file containing all of the objects for which the "include in report" 
  # checkbox was checked
  filewidth=8.5
  fileheight=11
  fontsize=12
  tablefontsize=9
  margins=c(1,1,1,1)
  rtf <- RTF(filename,width=filewidth,height=fileheight,
             font.size=fontsize,omi=margins)
  cmap = ReadSavedData()
  reportObjects = cmap$reportObjects
  
  gznum <- 1
  plotNums <- c()
  for (panelName in c("Plots", "Reports", "Analysis")) {
    emptyPanel = is.null(unlist(reportObjects[[panelName]]))
    if (!emptyPanel) {
      if (panelName %in% c("Reports", "Analysis")) {
        addPageBreak(rtf)
      }
      addHeader(rtf, panelName, font.size=14)
      choices = getChoices(panelName)
      if ((panelName == "Plots") & (ncol(cmap$ratings.dat) == 3))
        choices = choices[c(1:5,9)]
      if ((panelName == "Plots") & (ncol(cmap$ratings.dat) < 3))
        choices = choices[c(1,2,9)]
      for (choice in choices) {
        emptyChoice = is.null(unlist(reportObjects[[panelName]][[choice]]))
        if (!emptyChoice) {
          addHeader(rtf, choice, font.size=12, TOC.level=2)
          #          addTitle(doc,value=choice,font.size=10.0,level=2)
          for (arglist in reportObjects[[panelName]][[choice]]) {
            # 'output' will either be a plotting function or a table, depending on whether panel == "Plots"
            output = with(arglist, switch(choice,
                                          "Point Map (MDS)" = { ShowMDSPlot },
                                          "Clusters" = {
                                            # "Clusters" appears as an option under both the "Plots" and "Reports" panels,
                                            # so we need to check which panel we're in using is.null(clusterplot)
                                            if (exists("clusterplot")) {
                                              if(clusterplot == "Rays") {
                                                ShowClusterPlot
                                              } else {
                                                ShowClusterPolygonsPlot
                                              }
                                            } else {
                                              lbl = clusterLabels()
                                              if (is.null(univariateRating))
                                                return(NULL)
                                              df <- rating.summary.cluster(cluster.id=lbl,
                                                                           col=univariateRating,
                                                                           demographicGroup=demographicSelection)
                                              for (i in 3:5) {
                                                df[,i] = formatCol(df[,i],3)
                                              }
                                              colnames(df) <- gsub("\\."," ",colnames(df))
                                              df
                                            }
                                          },
                                          "Univariate Rating - by Statement" = {
                                            function() PlotRatingsLadder(univariateRating,demographicGroup=demographicSelection)
                                          },
                                          "Univariate Rating - by Cluster" = {
                                            function() ShowClusterRankingsPlotShades(univariateRating,demographicGroup=demographicSelection)
                                          },
                                          "Bar Chart" = {
                                            chooseRatingVars = c(chooseRating1,chooseRating2,
                                                                 chooseRating3,chooseRating4)
                                            if (is.null(chooseRatingVars))
                                              return(NULL)
                                            function() BarChart(chooseRatingVars,demographicGroup=demographicSelection)
                                          },
                                          "Pattern Matching" = {
                                            function() ClusterLadderGraph(bivariateRating1,bivariateRating2,demographicGroup=demographicSelection)
                                          },
                                          "GoZone" = {
                                            if(length(goZoneIndices) == 0)
                                              return(NULL)
                                            function() GoZoneGraphClusterList(goZoneIndices,bivariateRating1,bivariateRating2,demographicGroup=demographicSelection)
                                          },
                                          "Dendrogram" = {function() ShowDendrogram("full")},
                                          "Sorters" = { 
                                            df = get.summary(demographicGroup=demographicSelection)
                                            colnames(df) <- gsub("\\."," ",colnames(df))
                                            df
                                          },
                                          "Statements" = {
                                            if (is.null(univariateRating))
                                              return(NULL)
                                            df <- rating.summary(col=univariateRating,demographicGroup=demographicSelection)
                                            if (!is.null(sortBy)){
                                              if (sortBy == "Mean (descending)") {
                                                df = df[order(df[,3],decreasing =TRUE),]
                                              }
                                              
                                              if (byCluster) {
                                                df = df[order(df[,5],decreasing =FALSE),]
                                                clustNames = unique(df[,5])
                                                clustNames = clustNames[-length(clustNames)]
                                                pos = sapply(clustNames,function(clName) max(which(df[,5]==clName)))
                                                for (rowNum in rev(pos)) {
                                                  df = rbind(df[1:rowNum,] , df[(rowNum+1):nrow(df),])
                                                }
                                              } 
                                            }
                                            df[,3] = formatCol(df[,3],2)
                                            df[,4] = formatCol(df[,4],2)
                                            colnames(df) <- gsub("\\."," ",colnames(df))
                                            df
                                          },
                                          "ANOVA" = {
                                            lbl = clusterLabels()
                                            df <- rating.Ftest.cluster(cluster.id=lbl,col=univariateRating,demographicGroup=demographicSelection)
                                            rnms <- rownames(df)
                                            df <- data.frame(rnms,df)
                                            colnames(df)[1] <- ""
                                            df
                                          },
                                          "Tukey" = {
                                            lbl = clusterLabels()
                                            df <- rating.TukeyHSDtest.cluster(cluster.id=lbl,col=univariateRating,demographicGroup=demographicSelection)
                                            for (i in 1:4) {
                                              df[,i] = formatCol(df[,i],3)
                                            }
                                            rnms <- rownames(df)
                                            df <- data.frame(rnms,df)
                                            colnames(df)[1] <- ""
                                            #df[,4] = signif(df[,4],4)
                                            colnames(df) <- gsub("\\."," ",colnames(df))    
                                            df
                                          })
            )
            if (is.function(output)) {
              # create a plot
              hgt = 5; wdt = 5;
              if (choice == "Dendrogram") {
                hgt = 4; wdt = 6.5
              }
              addPlot(rtf, output, width=wdt, height=hgt)
              addPageBreak(rtf)
              if (vg) {
                plotFiles = sort(list.files(homedir, pattern="^RCMapPlot"), decreasing = TRUE)
                if (length(plotFiles) == 0)
                  plotNo = 1
                else
                  plotNo = 1 + as.numeric(gsub("RCMapPlot(\\d\\d\\d).wmf","\\1",plotFiles[1]))
                win.metafile(filename = sprintf("%s/RCMapPlot%03d.wmf", 
                                                homedir, plotNo))
                output()
                dev.off()
                plotNums <- c(plotNums, plotNo)
              }
            }
            else {
              # create a table
              if (!is.null(output)) {
                tableText = sprintf("Group: %s",arglist$demographicSelection)
                tableText = paste(tableText, sprintf(", Rating variable: %s",arglist$univariateRating), sep="")
                addParagraph(rtf, tableText)
                if (choice != "Sorters") {
                  # indicate the rating variable and demographic group above the table
                  tableText = paste(tableText, sprintf(", Rating variable: %s",arglist$univariateRating), sep="")
                }
                if (choice == "Sorters") {
                  addTable(rtf, output, col.widths = c(0.8,1.25,1.0),
                           col.justify = rep("C",3),header.col.justify = rep("C",3))
                }
                if (choice == "Statements") {
                  addTable(rtf, output, col.widths = c(0.5, 3.7, 0.6, 0.6, 1.3),
                           col.justify = c("C","L",rep("C",3)),header.col.justify = c("C","L",rep("C",3)))
                }
                if (choice == "Clusters") {
                  addTable(rtf, output, col.widths = c(2,0.7,0.7,0.7,0.7),
                           col.justify = c("L",rep("C",4)),
                           header.col.justify = c("L",rep("C",4)))
                }
                if (choice == "ANOVA") {
                  addTable(rtf, output[,-1], row.names=TRUE,
                           col.widths = c(1,0.4,0.85,0.85,0.85,1.1),
                           col.justify = c("L",rep("C",5)),
                           header.col.justify = c("L",rep("C",5)))
                }
                if (choice == "Tukey") {
                  addTable(rtf, output[,-1], row.names=TRUE, 
                           col.widths = c(1.0,1.0,1.0,1.0,1.5),
                           col.justify = c("L",rep("C",4)),
                           header.col.justify =c("L",rep("C",4)))
                }
                addParagraph(rtf,"")
                addNewLine(rtf, n=3)
              }
            }
            if (choice == "GoZone") {
              #cmap$clustname$clust_03_01
              nclust <- reportObjects[["Plots"]][["GoZone"]][[gznum]]$n.clust
              lbls = 1:nclust
              goZoneIndices <- reportObjects[["Plots"]][["GoZone"]][[gznum]]$goZoneIndices
              cnkeys <- sprintf("clust_%02d_%02d",nclust,lbls)
              for (i in goZoneIndices) {
                if (length(cmap$clustname[[cnkeys[i]]]) > 0) {
                  lbls[i] = cmap$clustname[[cnkeys[i]]]
                }
              }
              tab <- data.frame(toupper(letters[1:length(lbls[goZoneIndices])]),
                                lbls[goZoneIndices])
              colnames(tab) <- c("Label","Cluster Name")
              twidth<-c(1,4)
              addTable(rtf, tab)
              gznum <- gznum + 1
            }
          }
        }
      }
    }
  }
  if (length(plotNums) > 0) {
    addText(rtf, sprintf("The selected figures are available in Windows Metafile format in the %s folder as individual files (%d - %d)",
                         homedir, plotNums[1], plotNums[length(plotNums)]))
  }
  addParagraph(rtf,"\n\n")
  timestamp = format(Sys.time(), "This file was created on %B %d, %Y at %H:%M:%S.")
  addText(rtf, timestamp)
  done(rtf)
}

formatCol <- function(col,maxNumDec) {
  # format a column of a data table (return type = character)
  # maxNumDec = maximum number of digits displayed to the right of decimal
  col = round(col,maxNumDec)
  absCol = abs(col)
  decCol = absCol%%1
  charCol = format(decCol)
  nsmall = nchar(charCol) - 2  
  format(col,nsmall)
}
