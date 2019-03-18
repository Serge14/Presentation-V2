# This code generates presentation automatically

library(data.table)
library(stringr)

setwd("/home/sergiy/Documents/Work/Nutricia/Global/Ver2")

YTD.No = 1

# Read all necessary files

df = fread("/home/sergiy/Documents/Work/Nutricia/1/Data/df.csv", 
           header = TRUE, stringsAsFactors = FALSE, data.table = TRUE)
ppt = read_pptx("sample2.pptx")
dictColor = fread("/home/sergiy/Documents/Work/Nutricia/Scripts/Presentation-V2/dictColor.csv")
dictContent = fread("/home/sergiy/Documents/Work/Nutricia/Scripts/Presentation-V2/dictContent.csv")


df = df[Form != "Liquid", 
        c("SubBrand", "Size", "Age", "Scent", "Pieces", "Value", "Volume", 
            "Channel", "EC", "AC", "Acidified",
            "Scent2", "ScentType", "GlobalPriceSegment") := NULL]

df = df[, .(ITEMSC = sum(PiecesC), VALUEC = sum(ValueC), VOLUMEC = sum(VolumeC)),
        by = .(Ynb, Mnb, Brand, PS0, PS2, PS3, PS, Company, PriceSegment, Form,
               Additives, Region)]

# or read BFprocessed

makeTable = function(df) {
  cols = names(df)[-1]
  df[,(cols) := round(.SD, 1), .SDcols = cols]
  
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
  
  ft <- bold(ft, ~ Name == "MILUPA", bold = TRUE)
  
  ft = colformat_num(ft, cols2, digits = 1)
  
  ft <- add_footer(ft, Name = "* The table presents market shares and differences vs previous period or the same period last year." )
  ft <- merge_at(ft, j = 1:10, part = "footer")
  
  ft = width(ft, j = ~ col_1, width = .05)
  ft = width(ft, j = ~ col_2, width = .05)
  ft = width(ft, j = ~ col_3, width = .05)
  
  ft = width(ft, j = ~ Name, width = 1.65)
  ft = width(ft, j = ~ CP, width = .72)
  ft = width(ft, j = ~ vsPP, width = .72)
  ft = width(ft, j = ~ YTD, width = .72)
  ft = width(ft, j = ~ difYTD, width = .72)
  ft = width(ft, j = ~ MAT, width = .72)
  ft = width(ft, j = ~ difMAT, width = .72)
  
  ft = set_header_labels(ft, Name = cols[1], 
                         CP = cols[3], 
                         vsPP = cols[4],
                         YTD = cols[6],
                         difYTD = cols[7],
                         MAT = cols[9],
                         difMAT = cols[10])
  ft = empty_blanks(ft)
  
  
  
  #ft = empty_blanks(autofit(ft))
  ft
}

makeChart = function(df){
  df1 = melt.data.table(df, id.vars = "Company")
  
  maxY = round(max(df1$value, na.rm = TRUE), -1)
  
  # customColors = c("NESTLE" = "red", "NUTRICIA" = "blue", "KHOROLSKII MK" = "orange",
  #                  "FRIESLAND CAMPINA" = "brown", "ABBOTT LAB" = "black", "DMK HUMANA" = "pink")
  customColors = dictColors$Color
  names(customColors) = dictColors$Name
  
  
  
  toShow = df[1:3, Company]
  
  df.plot = ggplot(df1, aes(x=variable, y=value, col = Company, group = Company)) + 
    geom_line() + 
    #geom_point() +
    # geom_text_repel(aes(label = ifelse(Company %in% toShow, round(value, 1), "")),
    #                 direction = "y", nudge_y = 1,
    #                 show.legend = FALSE, size = 3.5) +
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
    guides(color = guide_legend(nrow = 1)) + 
    coord_cartesian(ylim = c(0, maxY))
  
  df.table = ggplot(df1[Company %in% toShow], aes(x=variable, 
                                                  y=factor(Company),
                                                  label = value,
                                                  col = Company, 
                                                  group = Company)) + 
    geom_text(size = 3, aes(label = ifelse(is.na(value), "", sprintf("%0.1f",value)))) +
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

dataTable = function(measure, level, linesToShow, filterSegments = NULL) {
  
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

dataChart = function(measure, level, linesToShow, filterSegments) {
  
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
  else if (level == "Segment") {df = data.table::dcast(df, PriceSegment~Ynb+Mnb, fun = sum, value.var = measure)}
  
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

dataSegmentChart = function(measure, level, linesToShow, filterSegments) {
  
  # Create subset which consists only segments we'll work with
  #df = data[eval(parse(text = filterSegments)), .(Items = sum(ITEMSC), Value = sum(VALUEC), Volume = sum(VOLUMEC)),
  #          by = .(Brand, PS2, PS3, PS, Company, Ynb, Mnb, Form, Additives)]
  
  df = data[eval(parse(text = filterSegments)), .(Items = sum(ITEMSC), Value = sum(VALUEC), Volume = sum(VOLUMEC)),
            by = .(Brand, PS0, PS2, PS3, PS, Company, Ynb, Mnb, PriceSegment)]
  
  if (level == "Company")  {df = data.table::dcast(df, Company~Ynb+Mnb, fun = sum, value.var = measure)}
  else if (level == "Brand") {df = data.table::dcast(df, Brand~Ynb+Mnb, fun = sum, value.var = measure)}
  else if (level == "Segment") {df = data.table::dcast(df, PriceSegment~Ynb+Mnb, fun = sum, value.var = measure)}
  else if (level == "ISSegment") {df = data.table::dcast(df, PS~Ynb+Mnb, fun = sum, value.var = measure)}
  else if (level == "IMFSegment") {df = data.table::dcast(df, PS3+PS2~Ynb+Mnb, fun = sum, value.var = measure)
  df = df[,PS3 := paste(PS3, PS2)][,PS2 := NULL]}
  
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

dataSubset = df[]


mytmp = layout_summary(ppt)[[2]][1]
ppt = ppt %>%
  add_slide(layout = "Two Content", master = mytmp) %>%
  ph_with_gg(value=df.plot, index = 2) %>% 
  ph_with_flextable(type="body", value=t1, index=1)

print(ppt, target="sample2.pptx")