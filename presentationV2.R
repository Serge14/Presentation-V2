# This code generates presentation automatically

library(data.table)
library(stringr)
library(officer)
library(flextable)
library(ggplot2)
library(ggrepel)
library(magrittr)
library(ggpubr)
library(forcats)

# setwd("d:/Temp/3/Presentation-V2-master")
setwd("/home/sergiy/Documents/Work/Nutricia/Scripts/Presentation-V2")

YTD.No = 1

CP = "JAN 19"
vsPP = "vs. PP"
YTD = "YTD 19"
difYTD = "vs YTD18"
MAT = "MAT 19"
difMAT = "vs MAT18"

tableColnames3 = c("MAT 18", "MAT 19", " .", "YTD 18", "YTD 19", " ..",
                   "JAN 18", "FEB 18", "MAR 18", "APR 18", "MAY 18", "JUN 18", "JUL 18",
                   "AUG 18", "SEP 18", "OCT 18", "NOV 18", "DEC 18", "JAN 19")


# Read all necessary files

df = fread("/home/sergiy/Documents/Work/Nutricia/1/Data/df_Y19M01.csv", 
           header = TRUE, stringsAsFactors = FALSE, data.table = TRUE)
ppt = read_pptx("sample3.pptx")
dictColors = fread("dictColor.csv")
dictColors = dictColors[Color != ""]
# dictContent = fread("dictContent.csv")
dictContent = read.csv("dictContent.csv")

df = df[, c("SubBrand", "Size", "Age", "Scent", "Pieces", "Value", "Volume", 
            "Channel", "EC", "AC", "Acidified",
            "Scent2", "ScentType", "GlobalPriceSegment") := NULL]

df = df[Form != "Liquid"]

df = df[, .(ITEMSC = sum(PiecesC), VALUEC = sum(ValueC), VOLUMEC = sum(VolumeC)),
        by = .(Ynb, Mnb, Brand, PS0, PS2, PS3, PS, Company, PriceSegment, Form,
               Additives, Region)]

dfName = deparse(substitute(df))

customColors = dictColors$Color
names(customColors) = dictColors$Name

# Functions
makeTable = function(df) {
  cols = names(df)[-1]
  # df[,(cols) := round(.SD, 1), .SDcols = cols]
  
  cols = c(names(df)[1], "col_1", names(df)[2:3], "col_2", 
           names(df)[4:5], "col_3", names(df)[6:7])
  
  cols2 = c("Name", "col_1", "CP", "vsPP", "col_2", "YTD", "difYTD", "col_3", "MAT", "difMAT")
  
  names(df) = cols2[c(-2, -5, -8)]
  
  ft = regulartable(df, col_keys = cols2)
  ft = theme_zebra(ft)
  ft = bg(ft, bg = "#2F75B5", part = "header") #  , #0D47A1 - dark blue
  ft <- color(ft, color = "white", part = "header")
  
  #ft = fontsize(ft, size = 9)
  
  ft <- color(ft, ~ vsPP < 0, ~ vsPP, color = "red")
  ft <- color(ft, ~ difYTD < 0, ~ difYTD, color = "red")
  ft <- color(ft, ~ difMAT < 0, ~ difMAT, color = "red")
  
  # ft <- bold(ft, ~ Name == "MILUPA", bold = TRUE)
  
  ft = colformat_num(ft, cols2, digits = 1)
  
  ft <- add_footer(ft, Name = "* The table presents market shares and differences vs previous period or the same period last year." )
  ft <- merge_at(ft, j = 1:10, part = "footer")
  
  ft = width(ft, j = ~ col_1, width = .05)
  ft = width(ft, j = ~ col_2, width = .05)
  ft = width(ft, j = ~ col_3, width = .05)
  
  ft = width(ft, j = ~ Name, width = 1.65)
  ft = width(ft, j = ~ CP, width = .70)
  ft = width(ft, j = ~ vsPP, width = .72)
  ft = width(ft, j = ~ YTD, width = .70)
  ft = width(ft, j = ~ difYTD, width = .72)
  ft = width(ft, j = ~ MAT, width = .70)
  ft = width(ft, j = ~ difMAT, width = .72)
  
  ft = set_header_labels(ft, Name = cols[1], 
                         CP = CP, 
                         vsPP = vsPP,
                         YTD = YTD,
                         difYTD = difYTD,
                         MAT = MAT,
                         difMAT = difMAT)
  ft = empty_blanks(ft)
  
  
  
  #ft = empty_blanks(autofit(ft))
  ft
}

ftPriceSegmentAbs = function(df) {

cols = names(df)
names(df) = c("Name", "CP", "L3M", "YTD")
cols2 = c(names(df)[1], "col_1", names(df)[2], names(df)[3], names(df)[4])

ft = regulartable(df, col_keys = cols2)
ft = theme_zebra(ft)
ft = bg(ft, bg = "#2F75B5", part = "header") #  , #0D47A1 - dark blue 
ft <- color(ft, color = "white", part = "header")

#ft = fontsize(ft, size = 9)

ft = colformat_num(ft, cols2, digits = 1)

ft = width(ft, j = ~ col_1, width = .05)

ft = width(ft, j = ~ Name, width = 1.65) 
ft = width(ft, j = ~ CP, width = .72) 
ft = width(ft, j = ~ L3M, width = .72) 
ft = width(ft, j = ~ YTD, width = .72)

ft = set_header_labels(ft, Name = cols[1],
                       CP = cols[2],
                       L3M = cols[3],
                       YTD = cols[4])
ft = empty_blanks(ft)
ft
}

makeChart = function(df){
  
  names(df)[-1] = tableColnames3
  
  levelName = strsplit(gsub('\"', "", as.character(fopt1), fixed = TRUE), ", ")[[1]][2]
  
  df1 = melt.data.table(df, id.vars = levelName)
  # df1 = melt.data.table(df, id.vars = "Company")
  # df1[variable == "Blank1", variable := "."]
  # df1[variable == "Blank2", variable := ".."]
  
  LegendRowNumber = ceiling(length(df[, unique(get(levelName))])/7)
  
  maxY = ceiling(max(df1$value, na.rm = TRUE)/10)*10
  
  # customColors = c("NESTLE" = "red", "NUTRICIA" = "blue", "KHOROLSKII MK" = "orange",
  #                  "FRIESLAND CAMPINA" = "brown", "ABBOTT LAB" = "black", "DMK HUMANA" = "pink")

  # names(customColors) = str_pad(dictColors$Name, 17)
  
  
  toShow = df[1:3, get(levelName)]
  
  df.plot = ggplot(df1, 
                   aes(x=variable, 
                       y=value, 
                       col = get(levelName), 
                       group = get(levelName))) + 
    geom_line() + 
    geom_point() +
    geom_text_repel(aes(label = ifelse(get(levelName) %in% toShow, round(value, 1), "")),
                    direction = "y", nudge_y = 1,
                    show.legend = FALSE, size = 3.5) +
    scale_color_manual(values = customColors) +
    #scale_x_discrete(breaks = NULL) +
    theme_minimal() +
    ylab(NULL) + xlab(NULL) +
    theme(legend.position="top", legend.title = element_blank(),
          legend.text = element_text(size = 8),
          panel.grid.minor = element_blank(),
          panel.grid.major.y = element_line(linetype = "dotted", colour = "darkgrey"),
          panel.grid.major.x = element_blank(),
          panel.background = element_blank(),
          axis.text.x = element_text(angle = 90, hjust = 1)) +
    guides(color = guide_legend(nrow = LegendRowNumber)) + 
    coord_cartesian(ylim = c(0, maxY))
  
  # df1[, c(eval(levelName)) := str_pad(get(levelName), 17)]
  
  df.table = ggplot(df1[get(levelName) %in% toShow],
                    aes(
                      x = variable,
                      y = factor(str_pad(get(levelName), 19)),
                      label = value,
                      col = get(levelName),
                      group = get(levelName)
                    )) +
    geom_text(size = 3, aes(label = ifelse(is.na(value), "", sprintf("%0.1f", value)))) +
    scale_color_manual(values = customColors, guide = FALSE) +
    xlab(NULL) + ylab(NULL) +
    theme_bw() +
    theme(
      panel.grid.major = element_blank(),
      panel.border = element_blank(),
      axis.text.x = element_blank(),
      axis.ticks = element_blank(),
      plot.margin = unit(c(-0.5, 1, 0, 0.5), "lines")
    )
  
   ggarrange(df.plot, df.table, heights = c(7, 1),
            ncol = 1, nrow = 2, align = "v")
  
}


buildBarChart = function(df, fopt) {
  
  names(df)[-1] = tableColnames3
  
  chartTitle = strsplit(gsub('\"', "", as.character(fopt), fixed = TRUE), ", ")[[1]][1]
  # levelName = strsplit(gsub('\"', "", as.character(fopt1), fixed = TRUE), ", ")[[1]][2]
  # levelName = "PS3"
  levelName = names(df)[1]
  
  df1 = melt.data.table(df, id.vars = levelName)
  df1[, value := 100*value/sum(value), by = variable]
  df1[, labelPosition := cumsum(value), by = .(variable)] 
  # df1[, textColor := "white"] 
  # df1[get(levelName) != "Economy", textColor := "black"] 
  # df1$textColor = as.factor(df1$textColor)
  
  
  ggplot(df1, aes(x=variable,
                  y=value,
                  fill = fct_rev(get(levelName)),
                  label = value),
         color = textColor) +
    geom_bar(stat="identity") +
    geom_text(
      aes(label = ifelse(is.na(value), "", sprintf("%0.1f",value)),
          y = labelPosition),
      # y = labelPosition, color = textColor),
      vjust = 2.15,
      size=3) +
    scale_fill_manual(values = customColors) +
    scale_colour_manual(values = levels(df1$textColor)) +
    theme_minimal() +
    ylab(NULL) + xlab(NULL) +
    labs(title = chartTitle) +
    theme(legend.position=c(1, 1.05),
          legend.title = element_blank(),
          legend.text = element_text(size = 8),
          legend.justification="right",
          panel.grid.minor = element_blank(),
          panel.grid.major.y = element_blank(),
          panel.grid.major.x = element_blank(),
          panel.background = element_blank(),
          axis.text.x = element_text(angle = 90, hjust = 1),
          axis.text.y = element_blank(),
          legend.spacing.x = unit(0.1, 'cm')) +
    guides(colour = FALSE,
          fill = guide_legend(direction = "horizontal"))
  
}

buildPriceSegmentAbsChart = function(df) {

  levelName = names(df)[1]
  
p1 = ggplot(df1[grepl("MAT", variable)], aes(x=variable,
                                             y=value,
                                             fill = fct_rev(levelName),
                                             label = value),
            color = textColor) +
  geom_bar(stat="identity") +
  geom_text(
    aes(label = ifelse(is.na(value), "", sprintf("%0.0f",value)),
        y = labelPosition, color = textColor),
    vjust = 2.5,
    size=3) +
  # scale_fill_brewer(palette="GrandBudapest2") +
  scale_colour_manual(values = levels(df1$textColor)) +
  scale_fill_manual(values = customColorsPS) +
  theme_minimal() +
  ylab(NULL) + xlab(NULL) +
  # labs(title="Volume") +
  # theme(legend.position=c(1, 1.05),
  theme(legend.position="none",
        legend.title = element_blank(),
        legend.text = element_text(size = 8),
        legend.justification="right",
        panel.grid.minor = element_blank(),
        panel.grid.major.y = element_blank(),
        panel.grid.major.x = element_blank(),
        panel.background = element_blank(),
        axis.text.x = element_text(angle = 90, hjust = 1),
        axis.text.y = element_blank(),
        legend.spacing.x = unit(0.1, 'cm')) +
  guides(colour = FALSE,
         fill = guide_legend(direction = "horizontal"))

p2 = ggplot(df1[grepl("YTD", variable)], aes(x=variable,
                                             y=value,
                                             fill = fct_rev(levelName),
                                             label = value),
            color = textColor) +
  geom_bar(stat="identity") +
  geom_text(
    aes(label = ifelse(is.na(value), "", sprintf("%0.0f",value)),
        y = labelPosition, color = textColor),
    vjust = 2.5,
    size=3) +
  # scale_fill_brewer(palette="GrandBudapest2") +
  scale_colour_manual(values = levels(df1$textColor)) +
  scale_fill_manual(values = customColorsPS) +
  theme_minimal() +
  ylab(NULL) + xlab(NULL) +
  # labs(title="Volume") +
  theme(legend.position="none",
        legend.title = element_blank(),
        legend.text = element_text(size = 8),
        legend.justification="right",
        panel.grid.minor = element_blank(),
        panel.grid.major.y = element_blank(),
        panel.grid.major.x = element_blank(),
        panel.background = element_blank(),
        axis.text.x = element_text(angle = 90, hjust = 1),
        axis.text.y = element_blank(),
        legend.spacing.x = unit(0.1, 'cm')) +
  guides(colour = FALSE,
         fill = guide_legend(direction = "horizontal"))

p3 = ggplot(df1[19:57], aes(x=variable,
                            y=value,
                            fill = fct_rev(levelName),
                            label = value),
            color = textColor) +
  geom_bar(stat="identity") +
  geom_text(
    aes(label = ifelse(is.na(value), "", sprintf("%0.0f",value)),
        y = labelPosition, color = textColor),
    vjust = 2.5,
    size=3) +
  # scale_fill_brewer(palette="GrandBudapest2") +
  scale_colour_manual(values = levels(df1$textColor)) +
  scale_fill_manual(values = customColorsPS) +
  theme_minimal() +
  ylab(NULL) + xlab(NULL) +
  labs(title=" ") +
  theme(legend.position=c(0.0, 1.1),
        legend.title = element_blank(),
        legend.text = element_text(size = 8),
        legend.justification="left",
        panel.grid.minor = element_blank(),
        panel.grid.major.y = element_blank(),
        panel.grid.major.x = element_blank(),
        panel.background = element_blank(),
        axis.text.x = element_text(angle = 90, hjust = 1),
        axis.text.y = element_blank(),
        legend.spacing.x = unit(0.15, 'cm')) +
  guides(colour = FALSE,
         fill = guide_legend(direction = "horizontal"))

# ggarrange(p1, p2, p3, align = "v", nrow = 2, ncol= 2, common.legend = TRUE)
p=ggarrange(
  grid.arrange(
    grobs = list(p1, p2, p3),
    widths = c(4, 1, 1),
    layout_matrix = rbind(c(NA, 1, 2),
                          c(3, 3, 3))
  )
)
p
}


dataTable = function(data, measure, level, linesToShow, filterSegments = NULL) {
  
  # Create subset which consists only segments we'll work with
  #df = data[eval(parse(text = filterSegments)), .(Items = sum(ITEMSC), Value = sum(VALUEC), Volume = sum(VOLUMEC)),
  #          by = .(Brand, PS2, PS3, PS, Company, Ynb, Mnb, Form, Additives, PriceSegment)]
  
  df = data[eval(parse(text = filterSegments)), .(Items = sum(ITEMSC), Value = sum(VALUEC), Volume = sum(VOLUMEC)),
            by = .(Brand, PS0, PS2, PS3, PS, Company, Ynb, Mnb, PriceSegment)]
  
  
  
  if (level == "Company")  {df = data.table::dcast(df, Company~Ynb+Mnb, fun = sum, value.var = measure)}
  else if (level == "Brand") {df = data.table::dcast(df, Brand~Ynb+Mnb, fun = sum, value.var = measure)}
  else if (level == "Segment") {df = data.table::dcast(df, PriceSegment~Ynb+Mnb, fun = sum, value.var = measure)}
  
  nc = length(df)
  
  df$YTDLY = df[, Reduce(`+`, .SD), .SDcols = (nc - YTD.No - 11):(nc - 12)]
  df$YTD = df[, Reduce(`+`, .SD), .SDcols = (nc - YTD.No + 1):nc]
  df$MATLY = df[, Reduce(`+`, .SD), .SDcols = (nc - 23):(nc - 12)]
  df$MAT = df[, Reduce(`+`, .SD), .SDcols = (nc - 11):nc]
  
  df = df[, .SD, .SDcols=c(1, nc-1, nc, (nc+1):(nc+4))]
  
  df = df[, paste0(names(df)[2:length(df)], "_MS") := lapply(.SD, function(x) 100*x/sum(x)), 
          .SDcols = 2:length(df)]          
  
  df = df[, .SD, .SDcols = -c(2:7)]
  
  df$DeltaCM = df[, Reduce(`-`, .SD), .SDcols = c(3, 2)]
  df$DeltaYTD = df[, Reduce(`-`, .SD), .SDcols = c(5, 4)]
  df$DeltaMAT = df[, Reduce(`-`, .SD), .SDcols = c(7, 6)]
  
  setcolorder(df, c(1:3, 8, 4:5, 9, 6:7, 10))
  df = df[, .SD, .SDcols = -c(2, 5, 8)]
  
  result = head(df[order(-df[,4])], linesToShow)
  return(result)
}

dataChart = function(data, measure, level, linesToShow, filterSegments) {
  
  # Create subset which consists only segments we'll work with
  #df = data[eval(parse(text = filterSegments)), .(Items = sum(ITEMSC), Value = sum(VALUEC), Volume = sum(VOLUMEC)),
  #          by = .(Brand, PS2, PS3, PS, Company, Ynb, Mnb, Form, Additives)]
  
  df = data[eval(parse(text = filterSegments)), .(Items = sum(ITEMSC), Value = sum(VALUEC), Volume = sum(VOLUMEC)),
            by = .(Brand, PS0, PS2, PS3, PS, Company, Ynb, Mnb, PriceSegment)]
  
  if (level == "Company")  {df = data.table::dcast(df, Company~Ynb+Mnb, fun = sum, value.var = measure)}
  else if (level == "Brand") {df = data.table::dcast(df, Brand~Ynb+Mnb, fun = sum, value.var = measure)}
  else if (level == "Segment") {df = data.table::dcast(df, PriceSegment~Ynb+Mnb, fun = sum, value.var = measure)}
  
  nc = length(df)
  
  df$YTDLY = df[, Reduce(`+`, .SD), .SDcols = (nc - YTD.No - 11):(nc - 12)]
  df$YTD = df[, Reduce(`+`, .SD), .SDcols = (nc - YTD.No + 1):nc]
  df$MATLY = df[, Reduce(`+`, .SD), .SDcols = (nc - 23):(nc - 12)]
  df$MAT = df[, Reduce(`+`, .SD), .SDcols = (nc - 11):nc]
  
  df = df[, .SD, .SDcols=c(1, (nc-12):nc, (nc+1):(nc+4))]
  
  df = df[, paste0(names(df)[2:length(df)], "_MS") := lapply(.SD, function(x) 100*x/sum(x)), 
          .SDcols = 2:length(df)]          
  
  df = df[, .SD, .SDcols = -c(4:18)]
  df[,2] = NA
  df[,3] = NA
  names(df)[2:3] = c("Blank1", "Blank2")
  
  setcolorder(df, c(1, 19:20, 2, 17:18, 3, 4:16))
  
  result = head(df[order(-df[,6])], linesToShow)
  return(result)
}

dataSegmentTable = function(measure, level, linesToShow, filterSegments) {
  
  # Create subset which consists only segments we'll work with
  #df = data[eval(parse(text = filterSegments)), .(Items = sum(ITEMSC), Value = sum(VALUEC), Volume = sum(VOLUMEC)),
  #          by = .(Brand, PS2, PS3, PS, Company, Ynb, Mnb, Form, Additives)]
  
  df = data[eval(parse(text = filterSegments)), .(Items = sum(ITEMSC), Value = sum(VALUEC), Volume = sum(VOLUMEC)),
            by = .(Brand, PS0, PS2, PS3, PS, Company, Ynb, Mnb, PriceSegment)]
  
  if (level == "Company")  {df = data.table::dcast(df, Company~Ynb+Mnb, fun = sum, value.var = measure)}
  else if (level == "Brand") {df = data.table::dcast(df, Brand~Ynb+Mnb, fun = sum, value.var = measure)}
  else if (level == "PriceSegment") {df = data.table::dcast(df, PriceSegment~Ynb+Mnb, fun = sum, value.var = measure)}
  
  nc = length(df)
  
  df$L3M = df[, Reduce(`+`, .SD), .SDcols = (nc - 2):nc]
  df$YTD = df[, Reduce(`+`, .SD), .SDcols = (nc - YTD.No + 1):nc]
  
  
  
  df = df[, paste0(names(df)[2:length(df)], "_MS") := lapply(.SD, function(x) 100*x/sum(x)), 
          .SDcols = 2:length(df)]          
  
  df = df[, .SD, .SDcols = -c(2:(2*nc))]
  
  
  #setcolorder(df, c(1, 17:18, 19, 15:16, 20, 2:14))
  
  #result = head(df[order(-df[,6])], linesToShow)
  result = head(df[order(df[,1])], linesToShow)
  return(result)
}

dataSegmentChart = function(data, measure, level, linesToShow, filterSegments) {
  
  # Create subset which consists only segments we'll work with
  #df = data[eval(parse(text = filterSegments)), .(Items = sum(ITEMSC), Value = sum(VALUEC), Volume = sum(VOLUMEC)),
  #          by = .(Brand, PS2, PS3, PS, Company, Ynb, Mnb, Form, Additives)]
  
  df = data[eval(parse(text = filterSegments)), .(Items = sum(ITEMSC), Value = sum(VALUEC), Volume = sum(VOLUMEC)),
            by = .(Brand, PS0, PS2, PS3, PS, Company, Ynb, Mnb, PriceSegment)]
  
  if (level == "Company")  {df = data.table::dcast(df, Company~Ynb+Mnb, fun = sum, value.var = measure)}
  else if (level == "Brand") {df = data.table::dcast(df, Brand~Ynb+Mnb, fun = sum, value.var = measure)}
  else if (level == "PriceSegment") {df = data.table::dcast(df, PriceSegment~Ynb+Mnb, fun = sum, value.var = measure)}
  else if (level == "ICSegment") {df = data.table::dcast(df, PS~Ynb+Mnb, fun = sum, value.var = measure)}
  # else if (level == "IMFSegment") {df = data.table::dcast(df, PS3+PS2~Ynb+Mnb, fun = sum, value.var = measure)
  # df = df[,PS3 := paste(PS3, PS2)][,PS2 := NULL]}
  else if (level == "IMFSegment") {
    df1 = data.table::dcast(df[PS0 == "IMF" & PS3 == "Base"], 
                            PS3+PS2~Ynb+Mnb, fun = sum, value.var = measure)
    df1[, PS3 := paste(PS3, PS2)]
    df1[, PS2 := NULL]
    df2 = data.table::dcast(df[PS0 == "IMF" & (PS3 == "Specials" | PS3 == "Plus")], 
                             PS3 ~Ynb+Mnb, fun = sum, value.var = measure)
    df = rbindlist(list(df1, df2))
  }
  
  
  
  nc = length(df)
  
  df$YTDLY = df[, Reduce(`+`, .SD), .SDcols = (nc - YTD.No - 11):(nc - 12)]
  df$YTD = df[, Reduce(`+`, .SD), .SDcols = (nc - YTD.No + 1):nc]
  df$MATLY = df[, Reduce(`+`, .SD), .SDcols = (nc - 23):(nc - 12)]
  df$MAT = df[, Reduce(`+`, .SD), .SDcols = (nc - 11):nc]
  
  df = df[, .SD, .SDcols=c(1, (nc-12):nc, (nc+1):(nc+4))]
  
  #df = df[, paste0(names(df)[2:length(df)], "_MS") := lapply(.SD, function(x) 100*x/sum(x)), 
  #        .SDcols = 2:length(df)]          
  
  #df = df[, .SD, .SDcols = -c(4:18)]
  df$Blank1 = NA
  df$Blank2 = NA
  
  setcolorder(df, c(1, 17:18, 19, 15:16, 20, 2:14))
  
  #result = head(df[order(-df[,6])], linesToShow)
  result = head(df[order(df[,1])], linesToShow)
  return(result)
}


getTable = function(dfName, fopt) {
  
  if(is.null(fopt)) return(alist())
  eval(parse(text = paste("dataTable(", dfName, ",", fopt, ")")))
  
}

getChart = function(dfName, fopt) {
  
  if(is.null(fopt)) return(alist())
  eval(parse(text = paste("dataChart(", dfName, ",", fopt, ")")))
  
}

getBarChart = function(dfName, fopt) {
  
  if(is.null(fopt)) return(alist())
  eval(parse(text = paste("dataSegmentChart(", dfName, ",", fopt, ")")))
  
}


dataTable(df, "Volume", "Company", 11, 'PS0 == "IMF" | PS2 == "DRY FOOD" | PS3 == "SAVOURY MEAL" | PS3 == "FRUITS"')
dataChart(df, "Volume", "Company", 11, 'PS0 == "IMF" | PS2 == "DRY FOOD" | PS3 == "SAVOURY MEAL" | PS3 == "FRUITS"')
dataChart(df, "Volume", "Company", 11, 'PS0 == "IMF" | PS2 == "DRY FOOD" | PS3 == "SAVOURY MEAL" | PS3 == "FRUITS"')



fopt = dictContent$Formula1[2]

a = eval(parse(text = paste("dataChart(df, ", fopt, ")")))

dfName = deparse(substitute(df))

getTable(dfName, fopt)

fopt = gsub("\"", "", fopt)


# mytmp = layout_summary(ppt)[[2]][1]
ppt = ppt %>%
  add_slide(layout = "Two Content", master = "Office Theme") %>%
  ph_with_text(type = "title", str = dictContent$Name[dictContent$No == i]) %>% 
  ph_with_text(type = "body", str = "index 3", index = 3) %>% 
  ph_with_text(type = "body", str = "index 7", index = 7) %>% 
  ph_with_text(type = "body", str = "index 8", index = 8) %>% 
  ph_with_gg(value = makeChart(getChart(dfName, fopt)), index = 2) %>% 
  ph_with_flextable(type="body", value=makeTable(getTable(dfName, fopt)), index=1)

print(ppt, target="sample2.pptx")


for (j in unique(dictContent$Type)) {
  for (i in dictContent$No[dictContent$Type == j]) {
    
    fopt = dictContent$Formula1[dictContent$No == i]
    if (fopt != "") {
    ppt = ppt %>%
      add_slide(layout = "Two Content", master = "Office Theme") %>%
      ph_with_text(type = "title", str = dictContent$Name[dictContent$No == i]) %>%
      ph_with_text(type = "body", str = dictContent$Region[dictContent$No == i], index = 3) %>%
      ph_with_gg(value = makeChart(getChart(dfName, fopt)), index = 2) %>%
      ph_with_flextable(type="body", value=makeTable(getTable(dfName, fopt)), index=1)
      print(i)
      
}
  }
}

print(ppt, target="sample2.pptx")

df[PriceSegment == "PREMIUM", PriceSegment := "Premium"]
df[PriceSegment == "MAINSTREAM", PriceSegment := "Mainstream"]
df[PriceSegment == "ECONOMY", PriceSegment := "Economy"]

text_prop <- fp_text(color = "white", font.size = 18)

for (i in dictContent$No) {
  
  print(i)
  
  fopt1 = dictContent$Formula1[i]
  fopt2 = dictContent$Formula2[i]
 
  if (dictContent$Type[i] == "Two Tables") {
    
    ppt = ppt %>%
      add_slide(layout = "1_Two Content", master = "Office Theme") %>%
      ph_with_text(type = "title", str = dictContent$Name[i]) %>%
      ph_with_ul(style = text_prop, 
                 type = "body", 
                 index = 3, 
                 str_list = c(str_pad(dictContent$Region[i], 16)), 
                 level_list = c(1)) %>% 
      ph_with_flextable(type = "body",
                        value = makeTable(getTable(dfName, fopt2)),
                        index = 2) %>%
      ph_with_flextable(type = "body",
                        value = makeTable(getTable(dfName, fopt1)),
                        index = 1)
    
  }
  
  if (fopt1 != "" & i > 1 & 
      (dictContent$Level[i] == "Company" | dictContent$Level[i] == "Brand")) {
    
    ppt = ppt %>%
      add_slide(layout = "Two Content", master = "Office Theme") %>%
      ph_with_text(type = "title", str = dictContent$Name[i]) %>%
      ph_with_ul(style = text_prop, 
                 type = "body", 
                 index = 3, 
                 str_list = c(str_pad(dictContent$Region[i], 16)), 
                 level_list = c(1)) %>% 
      ph_with_gg(value = makeChart(getChart(dfName, fopt2)), index = 2) %>%
      ph_with_flextable(type = "body",
                        value = makeTable(getTable(dfName, fopt1)),
                        index = 1)
  }
    
    if (fopt1 != "" & i > 1 & 
        (dictContent$Level[i] == "PriceSegment" | 
         dictContent$Level[i] == "ICSegment" |
         dictContent$Level[i] == "IMFSegment")) {
      ppt = ppt %>%
        add_slide(layout = "1_Two Content", master = "Office Theme") %>%
        ph_with_text(type = "title", str = dictContent$Name[i]) %>%
        ph_with_ul(style = text_prop, 
                   type = "body", 
                   index = 3, 
                   str_list = c(str_pad(dictContent$Region[i], 16)), 
                   level_list = c(1)) %>%
        ph_with_gg(value = buildBarChart(getBarChart(dfName, fopt2), fopt2), index = 2) %>%
        ph_with_gg(value = buildBarChart(getBarChart(dfName, fopt1), fopt1), index = 1)
        
    }
  
  if (fopt1 != "" & i > 1 & 
      dictContent$Level[i] == "PriceSegmentAbs") {
    ppt = ppt %>%
      add_slide(layout = "2_Two Content", master = "Office Theme") %>%
      ph_with_text(type = "title", str = dictContent$Name[i]) %>%
      ph_with_ul(style = text_prop, 
                 type = "body", 
                 index = 3, 
                 str_list = c(str_pad(dictContent$Region[i], 16)), 
                 level_list = c(1)) %>%
      ph_with_gg(value = buildPriceSegmentAbsChart(getBarChart(dfName, fopt2), fopt2), index = 2) %>%
      ph_with_flextable(value = ftPriceSegmentAbs(getBarChart(dfName, fopt1), fopt1), index = 1)
    
  }  
  
  print(i)
}


print(ppt, target="sample3_4.pptx")

