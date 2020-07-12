# library        -------
library(tidyverse)
library(rvest)
library(readxl)

# function       ----------------------------------------------------------------

#URLからエクセルファイルのリンク先を取得
read_mhlw<-function(x)(
  x %>% 
    read_html() %>% 
    html_nodes("a") %>% 
    html_attr("href") %>% 
    str_subset("\\.xlsx") %>% 
    str_c("https://www.mhlw.go.jp",.)
)

#性年齢データの最後の成型
age_gender_fix<-function(x){
  x  %>% 
    gather(性年齢,値,contains("歳"),総計) %>% 
    mutate(値=round(値)) %>% 
    distinct_all()
}

#都道府県データの最後の成型
area_fix<-function(x){
  x   %>% 
    gather(都道府県,値,c(総計,`01北海道`:`47沖縄県`)) %>% 
    mutate(値=round(値)) %>% 
    distinct_all()
}

# データ取得     ----------------------------------------------------------------

#取得先のURL
urls<-list(
  "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000139390.html",#NDB1
  "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000177221.html",#NDB2
  "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000177221_00002.html",#NDB3
  "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000177221_00003.html"#NDB4
)

#エクセルからデータを取得
for(j in 1:4){
  dir<-paste(getwd(),"/NDB_v",j,sep="")
  dir.create(dir)
  html_data <- read_mhlw(urls[[j]])
  
  for(i in 1:length(html_data)){
    download.file(url = html_data[i],
                  destfile = str_c(file.path(dir),"/data",i,".xlsx"),
                  method="wget")
    Sys.sleep(0.5) 
    
  }
}





# データ読み込み -----------------------------------------------------------------

rm(list=ls())
dat<-list()

#大きく都道府県と性年代のデータセットがあるが、ファイル名称などから判断つかないため、一旦強制的に取得した後にデータ整形する。
for(i in 1:4){
  # 上記で取得したフォルダごとに読み込みを行う
  year<-paste("data_v",i,sep="")
  excel_list<-list.files(year)
  
  for(j in 1:length(excel_list)){
    # 上記で取得したフォルダ内のexcelを一つずつ読み込む
    sheets <- excel_sheets(paste(year,"/",excel_list[j],sep=""))
    
    for(k in 1:length(sheets)){
      # 上記で取得したexcelをシートごとに読み込む
      suppressMessages(tmp0<-read_excel(paste(year,"/",excel_list[j],sep="")))
      
      # 特定健診のファイルは対象外:セルA1は特定健診という文言を含まない
      # クロス表のファイルは対象外:セルA1は診療行為コードという文言を含まない
      # 二次医療圏別のファイルは対象外:セルA1は医療機関数という文言を含まない
      # 特定保険医療材料のファイルは対象外:セルA1は特定保険医療材料という文言を含まない
      if(!str_detect(colnames(tmp0),"特定健診")&!str_detect(colnames(tmp0),"診療行為コード")&!str_detect(colnames(tmp0),"医療機関数")&!str_detect(colnames(tmp0),"特定保険医療材料")){
        suppressMessages(tmp1<-read_excel(paste(year,"/",excel_list[j],sep=""),skip=2,sheet = k))
        suppressMessages(tmp2<-read_excel(paste(year,"/",excel_list[j],sep=""),skip=3,sheet = k))
        
        # 2行目の列名称を取得する。それ以外の自動で名称付けられる名称[...xx]を除外
        tmp3<-data.frame(tmp = colnames(tmp1) %>% str_remove("\r\n") %>% str_remove_all("(\\...[0-9][0-9])") %>% str_remove_all("(\\...[0-9])")) %>% 
          mutate(tmp=if_else(tmp=="",NA_character_,as.character(tmp))) %>% 
          fill(tmp) 
        
        # 3行目の列名称を取得し、それ以外の自動で名称付けられる名称[...xx]を除外。
        # 年度によって性別の表記ゆれがあるので統一。
        # 2列目と結合し、列名称を作成する。
        col<-str_c(tmp3$tmp,colnames(tmp2) %>% str_remove("\r\n") %>% str_remove_all("(\\...[0-9][0-9])") %>% str_remove_all("(\\...[0-9])")) %>%
          str_remove_all("性") %>% 
          str_replace_all("～","_")%>% 
          str_replace_all("-","_")
          
        #作成した列名称をデータフレームに適用する。
        #薬価列に読み込みミスがでるため、ここのみtext指定
        suppressMessages(tmp2<-read_excel(paste(year,"/",excel_list[j],sep=""),skip=3,sheet = k,col_types=ifelse(str_detect(c("薬価|点数"), col),"text","guess")))
        colnames(tmp2)<-col
        
        #加算に関するシートは空白列があるため、列を除外する。
        if(str_detect(sheets[k],"加算")){
          tmp2<-tmp2 %>% select(-c(2:3))
          tmp<-tmp2 %>% mutate_all(funs(str_replace_all(.,"-","0"))) %>% fill(1:5) %>% mutate(sheet= paste(colnames(tmp0)[1],"_",sheets[k],sep="")) 
        }
        else{
          tmp<-tmp2 %>% mutate_all(funs(str_replace_all(.,"-","0"))) %>% fill(1:5) %>% mutate(sheet= paste(colnames(tmp0)[1],"_",sheets[k],sep=""))
        }
        
        #データを結合する。
        dat<-bind_rows(dat,tmp)
      }
      #対象外のデータは読み込まない
      else{}
      print(paste("year:",i,"excel:",j,"sheet:",k))
    }
  }
}

rm(list=ls()[!ls() %in% c("dat","age_gender_fix","are_fix")])

# データ整形     -------------------------------------------------------------------

dat<-dat%>% 
  mutate(
    年度=case_when(
      str_detect(sheet,"H26年04月") ~"H26",
      str_detect(sheet,"H27年04月") ~"H27",
      str_detect(sheet,"H28年04月") ~"H28",
      str_detect(sheet,"H29年04月") ~"H29"),
    外来入院区分=case_when(
      str_detect(sheet,"外来") ~"外来",
      str_detect(sheet,"入院") ~"入院",
      str_detect(sheet,"全体") ~"全体"),
    処方薬区分=case_when(
      str_detect(sheet,"内服")&str_detect(sheet,"院内") ~"内服薬_外来院内",
      str_detect(sheet,"内服")&str_detect(sheet,"院外") ~"内服薬_外来院外",
      str_detect(sheet,"内服")&str_detect(sheet,"入院") ~"内服薬_入院",
      str_detect(sheet,"外用")&str_detect(sheet,"院内") ~"外用薬_外来院内",
      str_detect(sheet,"外用")&str_detect(sheet,"院外") ~"外用薬_外来院外",
      str_detect(sheet,"外用")&str_detect(sheet,"入院") ~"外用薬_入院",
      str_detect(sheet,"注射薬")&str_detect(sheet,"院内") ~"注射薬_外来院内",
      str_detect(sheet,"注射薬")&str_detect(sheet,"院外") ~"注射薬_外来院外",
      str_detect(sheet,"注射薬")&str_detect(sheet,"入院") ~"注射薬_入院"),
    #加算・区分名称・分類名称は重複しない。款は手術分類名称の大分類
    分類_加算=case_when(
      !is.na(款) ~str_c(款,分類名称,sep="_"),
      !is.na(分類名称) ~分類名称,
      !is.na(加算) ~加算,
      !is.na(区分名称) ~区分名称,
      TRUE~NA_character_)
    ) %>% 
  mutate_at(vars(c(contains("歳"),点数,薬価,総計,`01北海道`:`47沖縄県`)),funs(as.numeric(.)))

GenderAge_Treatment <-dat %>% 
  select(年度,外来入院区分,分類_加算,contains("診療行為"),点数,総計,contains("歳")) %>%
  filter(!is.na(診療行為)) %>% 
  age_gender_fix()

GenderAge_Medicine  <-dat %>% 
  select(年度,処方薬区分,contains("医薬品コード"),医薬品名,後発品区分,薬価,総計,contains("歳")) %>% 
  filter(!is.na(医薬品名))  %>% 
  age_gender_fix()

GenderAge_Dentistry <-dat %>% 
  select(年度,contains("傷病"),総計,contains("歳")) %>%
  filter(!is.na(傷病名)) %>% 
  age_gender_fix()

Area_Treatment <-dat %>% 
  select(年度,外来入院区分,分類_加算,contains("診療行為"),点数,総計,c(`01北海道`:`47沖縄県`)) %>%
  filter(!is.na(診療行為)) %>% 
  area_fix()

Area_Medicine  <-dat %>% 
  select(年度,処方薬区分,contains("医薬品コード"),医薬品名,後発品区分,薬価,総計,c(`01北海道`:`47沖縄県`)) %>% 
  filter(!is.na(医薬品名)) %>% 
  area_fix()

Area_Dentistry <-dat %>%
  select(年度,contains("傷病"),総計,c(`01北海道`:`47沖縄県`)) %>% 
  filter(!is.na(傷病名)) %>% 
  area_fix()
  

# データ書き出し ---------------------------------------------------------------

write.csv(Area_Dentistry,"summary_Area_Dentistry.csv",row.names = FALSE)
write.csv(Area_Medicine,"summary_Area_Medicine.csv",row.names = FALSE)
write.csv(Area_Treatment,"summary_Area_Treatment.csv",row.names = FALSE)
write.csv(GenderAge_Dentistry,"summary_GenderAge_Dentistry.csv",row.names = FALSE)
write.csv(GenderAge_Medicine,"summary_GenderAge_Medicine.csv",row.names = FALSE)
write.csv(GenderAge_Treatment,"summary_GenderAge_Treatment.csv",row.names = FALSE)