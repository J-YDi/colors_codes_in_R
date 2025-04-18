################################################################################
#                                 Colors in R                                  #
#     Find the RGB and HEX codes for all named palettes in R and ggplot2       #
################################################################################
###############                  DIAS Jean-Yves                  ###############   

# Packages
library(RColorBrewer)
library(paletteer)
library(tibble)
library(dplyr)
library(purrr)
library(wesanderson)
library(openxlsx)

# Continuous palettes from paletteer library
# ---- PARAMETERS ----
# Define the names of the palettes to use

palettes <- paste0(palettes_c_names$package,"::",palettes_c_names$palette)
nb_couleurs <- 15

# ---- GENERATE PALETTES & EXTRACT INFORMATION ----
# For each palette, generate the colors and extract the HEX, R, G and B values
df_paletteer_c <- map_df(palettes, function(palette_nom) {
  couleurs <- paletteer_c(palette_nom, nb_couleurs)
  tibble(
    palette = palette_nom,
    R = col2rgb(couleurs)[1, ],
    G = col2rgb(couleurs)[2, ],
    B = col2rgb(couleurs)[3, ],
    HEX = substr(as.character(couleurs), 1, 7)
  )
})


# Discrete palettes from paleteer library --> not coloured in the Excel because 
# too much cells

# ---- PARAMETERS ----
# Define the names of the palettes to use
palettes <- paste0(palettes_d_names$package,"::",palettes_d_names$palette)

# ---- GENERATE PALETTES & EXTRACT INFORMATION ----
# For each palette, generate the colors and extract the HEX, R, G and B values
df_paletteer_d <- map_df(palettes, function(palette_nom) {
  couleurs <- paletteer_d(palette_nom,type = c("discrete"))
  tibble(
    palette = palette_nom,
    R = col2rgb(couleurs)[1, ],
    G = col2rgb(couleurs)[2, ],
    B = col2rgb(couleurs)[3, ],
    HEX = substr(as.character(couleurs), 1, 7)
  )
})

# Dynamic palettes from paleteer library

palettes <- paste0(palettes_dynamic_names$package, "::", palettes_dynamic_names$palette)
nombre_couleurs <- palettes_dynamic_names$length  # vector indicating the number of colors per palette

# ---- GENERATE PALETTES & EXTRACT INFORMATION ----
df_paletteer_dy <- map2_df(palettes, nombre_couleurs, function(palette_nom, n) {
  # Generate the palette with the desired number 'n' of colors
  couleurs <- paletteer_dynamic(palette_nom, n = n)
  tibble(
    palette = palette_nom,
    R = col2rgb(couleurs)[1, ],
    G = col2rgb(couleurs)[2, ],
    B = col2rgb(couleurs)[3, ],
    HEX = substr(as.character(couleurs), 1, 7)
  )
})


# Brewer palette already included in paletteer 

# Load necessary libraries
#library(RColorBrewer)
#library(tibble)
#library(dplyr)
#library(purrr)
#
## ---- PARAMETERS ----
## Specify the names of the Brewer palettes to use.
#palettes <- rownames(brewer.pal.info)
#nb_couleurs <- 15  # Desired number, but the palette will use its max number if it's lower
#
## ---- GENERATE PALETTES & EXTRACT INFORMATION ----
## For each palette, retrieve the maximum number of available colors and use it directly.
#df_couleurs <- map_df(palettes, function(palette_nom) {
#  
#  # Get the maximum number of available colors for the palette
#  max_couleurs <- brewer.pal.info[palette_nom, "maxcolors"]
#  
#  # Use the minimum between nb_couleurs and max_couleurs
#  n_colors <- min(nb_couleurs, max_couleurs)
#  
#  # Generate the palette with the available number of colors
#  couleurs <- brewer.pal(n_colors, palette_nom)
#  
#  # Extract HEX, R, G and B values
#  tibble(
#    palette = palette_nom,
#    R = col2rgb(couleurs)[1, ],
#    G = col2rgb(couleurs)[2, ],
#    B = col2rgb(couleurs)[3, ],
#    HEX = couleurs
#  )
#})
#
## ---- PREVIEW ----
#print(df_couleurs)
#
## ---- EXPORT CSV (optional) ----
#write.csv(df_couleurs, "palette_rgb_hex.csv", row.names = FALSE)


# ---- FUNCTION FOR THE DEFAULT PALETTE ----
gg_color_hue <- function(n) {
  hues = seq(15, 375, length = n + 1)
  hcl(h = hues, l = 65, c = 100)[1:n]
}

# ---- PARAMETERS ----
nb_couleurs <- 15  # Desired number of colors

# ---- GENERATE PALETTE ----
couleurs <- gg_color_hue(nb_couleurs)

# ---- EXTRACT HEX + RGB ----
df_defaut <- tibble(
  palette = "Default",
  R = col2rgb(couleurs)[1, ],
  G = col2rgb(couleurs)[2, ],
  B = col2rgb(couleurs)[3, ],
  HEX = couleurs
)

# Wesanderson palette

# ---- PARAMETERS ----
# Define the names of the palettes to use
palettes <- names(wes_palettes)
nb_couleurs <- 15

# ---- GENERATE PALETTES & EXTRACT INFORMATION ----
# For each palette, generate the colors and extract the HEX, R, G and B values
df_wesanderson <- map_df(palettes, function(palette_nom) {
  couleurs <- wes_palette(palette_nom, nb_couleurs,type=c("continuous"))
  tibble(
    palette = palette_nom,
    R = col2rgb(couleurs)[1, ],
    G = col2rgb(couleurs)[2, ],
    B = col2rgb(couleurs)[3, ],
    HEX = substr(as.character(couleurs), 1, 7)
  )
})

# ---- GENERATE PALETTE ----
couleurs <- rainbow(15)

# ---- EXTRACT HEX + RGB ----
df_rainbow <- tibble(
  palette = "rainbow",
  R = col2rgb(couleurs)[1, ],
  G = col2rgb(couleurs)[2, ],
  B = col2rgb(couleurs)[3, ],
  HEX = couleurs
)

# ---- GENERATE PALETTE ----
couleurs <- heat.colors(15)

# ---- EXTRACT HEX + RGB ----
df_heatcolors <- tibble(
  palette = "heat.colors",
  R = col2rgb(couleurs)[1, ],
  G = col2rgb(couleurs)[2, ],
  B = col2rgb(couleurs)[3, ],
  HEX = couleurs
)

# ---- GENERATE PALETTE ----
couleurs <- terrain.colors(15)

# ---- EXTRACT HEX + RGB ----
df_terrain <- tibble(
  palette = "terrain.colors",
  R = col2rgb(couleurs)[1, ],
  G = col2rgb(couleurs)[2, ],
  B = col2rgb(couleurs)[3, ],
  HEX = couleurs
)

# ---- GENERATE PALETTE ----
couleurs <- topo.colors(15)

# ---- EXTRACT HEX + RGB ----
df_topo <- tibble(
  palette = "topo.colors",
  R = col2rgb(couleurs)[1, ],
  G = col2rgb(couleurs)[2, ],
  B = col2rgb(couleurs)[3, ],
  HEX = couleurs
)

# ---- GENERATE PALETTE ----
couleurs <- cm.colors(15)

# ---- EXTRACT HEX + RGB ----
df_cm <- tibble(
  palette = "cm.colors",
  R = col2rgb(couleurs)[1, ],
  G = col2rgb(couleurs)[2, ],
  B = col2rgb(couleurs)[3, ],
  HEX = couleurs
)

# Merge all dataframes
df_colours <- rbind(df_cm,df_defaut,df_heatcolors,df_paletteer_c,df_paletteer_d,df_paletteer_dy,df_rainbow,df_terrain,df_topo,df_wesanderson)

write.xlsx(df_colours,file = "/palettes_codes_in_R.xlsx")


# Coloring the excel

# Import file
wb <- loadWorkbook("palettes_codes_in_R.xlsx")
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
block_starts <- c(1)  

# Apply the coloring by blocks of columns
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
saveWorkbook(new_wb, "palettes_df_fully_colored.xlsx", overwrite = TRUE)
