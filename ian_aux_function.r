

month2season=function(x) {
  res=NA
  if ( (x==1) || (x==2) ||(x==12)) {res="inv"};
  if ( (x==3) || (x==4) || (x==5) ) {res="prim"};
  if ( (x==6) || (x==7) ||(x==8)) {res="est"};
  if ( (x==9) || (x==10) ||(x==11)) {res="aut"};
  return(res);
  
}

messeri_iannuccilli_Iclass=function(par,category,treshpar=0.2,missing=-999.9) {
  
  require(dplyr)  
  
  if ( length(par) != length(category)) {stop("Vector have not the same length")}
  
  par[which(par==missing)]=NA
  
  temp_merge_df=data.frame(param=par,wt=category)
  
  freq=temp_merge_df %>% group_by(wt) %>% summarise(freq=n())
  
  variance=temp_merge_df %>% group_by(wt) %>% summarise(var=var(param,na.rm=T))
  
  
  
  temppar=temp_merge_df$param[which(!is.na(temp_merge_df$param))]
  temp_merge_df$POCC=NA;
  temppar=ifelse(temppar >=treshpar,1,0)
  temp_merge_df$POCC[which(!is.na(temp_merge_df$param))]=temppar
  
  variance_POCC=temp_merge_df  %>% group_by(wt) %>% summarise(var=var(POCC,na.rm=T))

  temp_aov=aov(param~as.factor(wt),data=temp_merge_df)

  tempsplit=split(temp_merge_df,as.factor(temp_merge_df$wt))

  listwt=list()

  z=1

  for (i in levels(as.factor(temp_merge_df$wt)))           { 
    listwt[[z]]=temp_merge_df[which(temp_merge_df$wt!=i),]
    z=z+1
  }

  res_ks=list()

  len_ks=1

  for (i in 1:length(levels(as.factor(temp_merge_df$wt))))   { 
    res_ks[[i]]=try(suppressWarnings(as.numeric(ks.test(listwt[[i]]$param,tempsplit[[i]]$param)$p.value)))
    z=z+1
    len_ks=len_ks+1
  }

  res_ks=ifelse(unlist(res_ks)<0.05,1,0)
  
  SDOCC=sum(((as.numeric(unlist(lapply(tempsplit,function(x) sum(x$POCC,na.rm=T))))/as.numeric(unlist(lapply(tempsplit,function(x) length(which(!is.na(x$POCC)))))))
             -(sum(temp_merge_df$POCC,na.rm=T)/length(which(!is.na(temp_merge_df$POCC)))))^2,na.rm=T)/(nrow(freq)-1)
  
  df_variance=data.frame(var=variance$var,freq=freq$freq)
  df_variance=df_variance[which(df_variance$freq!=1),]
  res= data.frame(t(unlist(summary(temp_aov))),
                  WSD=sqrt(weighted.mean( df_variance[,1],df_variance[,2])),
                  SDOCC=sqrt(SDOCC),
                  KS_rejection=sum(res_ks),
                  missing_data=length(which(is.na(temp_merge_df$param))))


  res$F.value2=NULL
  res$"Pr..F.2"=NULL
  names(res)[1:8]=c("DegF_Cat","DegF_Param","BSS","WSS","Variance_Cat","Variance_Param","F","Pvalue_F")
  
  res$TSS=res$BSS+res$WSS
  res$EV=1-(res$WSS/res$TSS)
  res=res[,c("TSS","BSS","WSS","EV","F","Pvalue_F","WSD","SDOCC","KS_rejection",
             "Variance_Cat","Variance_Param","DegF_Cat","DegF_Param","missing_data")]
  
  res$KS_stats=(sum(res_ks)/length(res_ks))*100
                
  return(res)
                
}
