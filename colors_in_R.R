################################################################################
#                                 Colors in R                                  #
#       Find the RGB and HEX codes for all named colors in R and ggplot2       #
################################################################################
###############                  DIAS Jean-Yves                  ###############                        



# Create a data.frame with HEX names and codes
color_df <- data.frame(
  name = colors(),
  t(col2rgb(colors()))
)

# add the HEX column
color_df$hex <- apply(color_df[, 2:4], 1, function(rgb) {
  rgb(rgb[1], rgb[2], rgb[3], maxColorValue = 255)
})

head(color_df)

# Little manipulation under Excel not shown #

# Automatically colour the lines and column according to the HEX code
# Package needed
library(openxlsx)

# Import file
wb <- loadWorkbook("color_df.xlsx")
sheet <- 1  # indicate the sheet name
df <- read.xlsx(wb, sheet = sheet) #store it as a dataframe

# Create a new workbook for styling
new_wb <- createWorkbook()
addWorksheet(new_wb, "Sheet1")

# Write the data without colors
writeData(new_wb, "Sheet1", df)

# Function to colour each cell
apply_fill <- function(wb, sheet, row, col, color) {
  style <- createStyle(fgFill = color)
  addStyle(wb, sheet, style, rows = row, cols = col, gridExpand = TRUE, stack = TRUE)
}

# Starting row
start_row <- 2
n_rows <- nrow(df)

# Column positions for each block because the column's names repeated themselves
block_starts <- c(1, 6, 11, 16)  

# Apply the coloration by block of columns
for (i in 0:(n_rows - 1)) {
  for (start_col in block_starts) {
    hex_value <- df[i + 1, start_col + 4]  # HEX column
    if (!is.na(hex_value) && grepl("^#([A-Fa-f0-9]{6})$", hex_value)) {
      cells_to_color <- start_col + c(0, 1, 2, 3, 4)  # Color name + R,G,B + HEX
      for (col in cells_to_color) {
        apply_fill(new_wb, "Sheet1", row = start_row + i, col = col, color = hex_value)
      }
    }
  }
}

# Save it
# saveWorkbook(new_wb, "color_df_fully_colored.xlsx", overwrite = TRUE)
