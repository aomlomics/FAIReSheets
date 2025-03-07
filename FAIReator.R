#' FAIR eDNA Template Generator (FAIRe-ator)
#' This function is to generate FAIR eDNA data templates, based on the sample types, assay type, and requirement levels of your choice. 
#'
#' Instructions 
#' 
#' Step 1: Save the input files in the working directory
#' FAIRe_checklist_v1.0.xlsx
#' FAIRe_checklist_v1.0_FULLtemplate.xlsx
#' 
#' STEP 2. Source the FAIReator.R script to load the function into your R environment.
#'
#' Step 3: Run the FAIReator function with the below arguments
#' 
#' Arguments
#' 
#' @req_lev Requirement level(s) of each fields to include in the data template. Select one or more from "M", "HR", "R", and "O" (Mandatory, Highly recommended, Recommended, and Optional, respectively). Default is c("M", "HR", "R", "O")
#' @sample_type A (list of) Sample type(s). Select one or more from "Water", "Soil", "Sediment", "Air", "HostAssociated", "MicrobialMatBiofilm", and "SymbiontAssociated", or "other". "other" will include all the sample-type-specific fields in the output. 
#' @assay_type An approach applied to detect taxon/taxa. Select one "targeted" (e.g., q/d PCR based detection) or "metabarcoding" (e.g., metabarcoding)
#' @project_id A brief, concise project identifier with no spaces or special characters. This ID will be used in data file names as 'project_id'.
#' @assay_name A brief, concise assay name(s) with no spaces or special characters. This will be used in data file names as 'assay_name'. 
#' @projectMetadata_user (optional) A user-defined field or list of fields that are not listed in the FAIR eDNA metadata checklist. These fields will be appended to the end of the projectMetadata.
#' @sampleMetadata_user (optional) A user-defined field or list of fields that are not listed in the FAIR eDNA metadata checklist. These fields will be appended to the end of the sampleMetadata.
#' 
#' @examples
#' FAIReator(sample_type = c('Water', 'Sediment'), 
#'                   assay_type='metabarcoding', 
#'                   project_id = 'gbr2022', 
#'                   assay_name = c('MiFish', 'crust16S')) 
#'
#'
#' The expected outputs are:
#' A new folder called "template_<project_id>" (e.g., template_gbr2022) in your working directory. 
#' An Excel file called "<project_id>.xlsx" (e.g., gbr2022.xlsx) in the new template folder, containing multiple worksheets in the file, including README, projectMetadata, sampleMetadata and other assay_type specific tables (e.g., OTU table if assay_type = "metabarcoding")

FAIReator = function(req_lev = c('M', 'HR', 'R', 'O'), 
                             sample_type, 
                             assay_type, 
                             project_id, 
                             assay_name, 
                             projectMetadata_user = NULL,
                             sampleMetadata_user = NULL) {
  
  # install packages --------------------------------------------------------
  
  packages <- c("readxl", "openxlsx", "RColorBrewer", "dplyr", "here") 
  for (i in packages) {
    if (!require(i, character.only = TRUE)) {
      install.packages(i, dependencies = TRUE)
      library(i, character.only = TRUE)
    }
  }
  
  
  # set input and output ----------------------------------------------------
  
  FAIRe_checklist_ver <- 'v1.0'
  input_file_name <- paste0('FAIRe_checklist_', FAIRe_checklist_ver, ".xlsx")
  
  sheet_name <- 'checklist'
  input <- suppressMessages(
    readxl::read_excel(input_file_name, sheet = sheet_name)
    )
  
  full_temp_file_name <- paste0("FAIRe_checklist_", FAIRe_checklist_ver, "_FULLtemplate.xlsx")
  
  
  # create a directory for output templates
  if(!dir.exists(paste('template', project_id, sep='_'))){dir.create(paste('template', project_id, sep='_'))} 
  
  wb <- createWorkbook() # create a excel workbook
  
  # Make README -------------------------------------------------------------
  iso_current_time <- sub("(\\d{2})(\\d{2})$", "\\1:\\2", format(Sys.time(), "%Y-%m-%dT%H:%M:%S%z"))
  readme1 <- c('The templates were generated using the eDNA checklist version of;', 
               input_file_name, 
               '',
               'Date/Time generated;', 
               iso_current_time)
  
  readme2 <- c('The templates were generated based on the below arguments;',
               paste('project_id =', project_id),
               paste('assay_name =', paste(assay_name, collapse = ' | ')),
               paste('assay_type =', assay_type),
               paste('req_lev =', paste(req_lev, collapse = ' | '))
  )
  
  if (any(sample_type %>% tolower() == 'other')) {
    readme2 <- c(readme2, 
                 paste('sample_type =', paste(sample_type, collapse = ' | '), 
                       '(Note: this option provides sample-type-specific fields for ALL sample types)')
    )
  } else {
    readme2 <- c(readme2, 
                 paste('sample_type =', paste(sample_type, collapse = ' | '), 
                       '(Note: this option provides sample-type-specific fields for the selected sample type(s))')
    )
  }
  
  if (!is.null(projectMetadata_user)) {
    readme2 <- c(readme2, paste('projectMetadata_user =', paste(projectMetadata_user, collapse = ' | ')))
  }
  if (!is.null(sampleMetadata_user)) {
    readme2 <- c(readme2, paste('sampleMetadata_user =', paste(sampleMetadata_user, collapse = ' | ')))
  }
  
  readme3 <- c('Requirement levels;', 'M = Mandatory', 'HR = Highly recommended', 'R = Recommended', 'O = Optional')
  
  readme4 <- c('List of files;',
               paste0('projectMetadata_', project_id),
               paste0('sampleMetadata_', project_id)
               )
  if (assay_type=='metabarcoding') {
    readme4=c(readme4, # MT: add exp meta for multi taxon. 
              paste('experimentRunMetadata', project_id, sep='_'),
              paste('otuRaw', project_id, assay_name, '<seq_run_id>', sep = '_'),
              paste('otuFinal', project_id, assay_name, '<seq_run_id>', sep = '_'),
              paste('taxaRaw', project_id, assay_name, '<seq_run_id>', sep = '_'),
              paste('taxaFinal', project_id, assay_name, '<seq_run_id>', sep = '_'),
              'Note: otuRaw, otuFinal, taxaRaw and taxaFinal should be produced for each assay_name and seq_run_id',
              'Note: <seq_run_id> in the file names should match with seq_run_id in your experimentRunMetadata'
    )
    
  } else if (assay_type=='targeted') {
    readme4=c(readme4, 
              paste('stdData', project_id, sep='_'),
              paste0('eLowQuantData_', project_id, ' (if applicable)'),
              paste('ampData', project_id, assay_name, sep = '_'),
              'Note: ampData should be produced for each assay_name'
    )
  }
  
  readme.df=data.frame(c(readme1, '', readme2, '', readme3, '', readme4))
  
  
  addWorksheet(wb = wb, sheetName='README')
  writeData(wb, 'README', readme.df, colNames = F) 
  
  # Fill colours for requirement levels
  req_col_df <- tibble(
    'requirement_level' = c(
      "M = Mandatory",
      "HR = Highly recommended",
      "R = Recommended",
      "O = Optional"
    ),
    'requirement_level_code' = c("M", "HR", "R", "O"),
    'col' = c("#E26B0A", "#FFCC00", "#FFFF99", "#CCFF99")
  )
  
  for (i in req_col_df$requirement_level_code) {
    addStyle(wb, sheet = 'README', 
             style = createStyle(fgFill = req_col_df$col[which(req_col_df$requirement_level_code == i)]), 
             rows = which(readme.df[,1] == req_col_df$requirement_level[which(req_col_df$requirement_level_code == i)]),
             cols = 1, 
             gridExpand = T, stack = T)
  }
  
  # projectMetadata sheet  ------------------------------------------------
  sheet_name <- 'projectMetadata'
  sheet_df <- suppressMessages(
    readxl::read_excel(full_temp_file_name, sheet=sheet_name)
  )
  
  # remove sections from projectMetadata based on detection type
  section2rm <- if(assay_type == 'metabarcoding') {
    'Targeted assay detection'
  } else if(assay_type == 'targeted') {
    c('Library preparation sequencing', 'Bioinformatics', 'OTU/ASV')
  } else if (assay_type %>% tolower() == 'other') {
    NA
  }
  
  for (i in section2rm) {
    if (i %in% sheet_df$section) {
      sheet_df <- sheet_df[-which(sheet_df$section == i),] 
    }
  }
  
  # req_lev 
  req_lev2rm <- setdiff(c('M', 'HR', 'R', 'O'), req_lev)
  for (i in req_lev2rm) {
    if (i %in% sheet_df$requirement_level_code) {
      sheet_df <- sheet_df[-which(sheet_df$requirement_level_code == i),] 
    }
  }
  
  # add projectMetadata_user
  if (!is.null(projectMetadata_user)) {
    temp <- data.frame(matrix(nrow = length(projectMetadata_user), ncol = ncol(sheet_df)))
    colnames(temp) <- colnames(sheet_df)
    temp$term_name <- projectMetadata_user
    temp$requirement_level_code <- 'O'
    temp$section <- 'User defined'
    sheet_df <- rbind(sheet_df, temp)
  }
  
  # pre-fill some values
  sheet_df[which(sheet_df$term_name == 'project_id'), 'project_level'] <- project_id
  sheet_df[which(sheet_df$term_name == 'assay_type'), 'project_level'] <- assay_type
  sheet_df[which(sheet_df$term_name == 'checkls_ver'), 'project_level'] <- input_file_name
  if (length(assay_name) == 1) {
    sheet_df[which(sheet_df$term_name == 'assay_name'), 'project_level'] <- assay_name
  } else if (length(assay_name) > 0) {
    for (i in 1:length(assay_name)) {
      sheet_df[,paste0('assay', i)] <- NA
      sheet_df[which(sheet_df$term_name == 'assay_name'),paste0('assay', i)] = assay_name[i]
    }
  }
  
  
  addWorksheet(wb, sheet_name)
  writeData(wb, sheet_name, sheet_df, colNames = T)
  
  # sampleMetadata sheet  ------------------------------------------------
  sheet_name <- 'sampleMetadata'
  sheet_df <- suppressMessages(
    readxl::read_excel(full_temp_file_name, sheet=sheet_name, col_names = F)
  )
  
  # remove terms from sampleMetadata based on sample type
  samp_type_row2keep <- if(any(sample_type %>% tolower() == 'other')) {
    NA
  } else {
    filtered_input <- input %>% filter(is.na(input$sample_type_specificity) | input$sample_type_specificity =='ALL' | grepl(paste(sample_type, collapse = '|'), input$sample_type_specificity, ignore.case = TRUE))
    colnames(sheet_df) <- sheet_df[which(sheet_df[,1] == 'samp_name'),]
    sheet_df <- sheet_df[,intersect(filtered_input$term_name, colnames(sheet_df))]
  }
  # remove 'Targeted assay detection' section if assay_type == 'metabarcoding'
  if (assay_type == 'metabarcoding') {
    section_row <- which(sheet_df[,1] == '# section')
    cols2rm <- which(sheet_df[section_row,] == 'Targeted assay detection')
    sheet_df <- sheet_df[, -cols2rm]
  }
  # add detected_notDetected_<assay_name> if assay_type == 'targeted' & length(assay_name)>1
  if (assay_type == 'targeted' & length(assay_name)>1) {
    new_columns <- paste0('detected_notDetected_', assay_name)
    original_col_index <- which(names(sheet_df) == "detected_notDetected")
    new_cols_data <- as.data.frame(replicate(length(new_columns), sheet_df[,"detected_notDetected"]))
    colnames(new_cols_data) <- new_columns
    new_cols_data[which(new_cols_data[,1]=="detected_notDetected"),]  <- new_columns
    
    sheet_df <- sheet_df[-original_col_index]  # Remove the original column
    sheet_df <- cbind(sheet_df[, 1:(original_col_index - 1)],  # Columns before the original
                      new_cols_data, 
                      sheet_df[, original_col_index:ncol(sheet_df)])  # Columns after the original
  }
  
  # req_lev
  req_lev_row <- which(sheet_df[1,] == '# requirement_level_code')
  for (i in req_lev2rm) {
    if (i %in% sheet_df[req_lev_row,]) {
      sheet_df <- sheet_df[,-which(sheet_df[req_lev_row,] == i)] 
    }
  }
  
  # add sampleMetadata_user
  if (!is.null(sampleMetadata_user)) {
    temp <- data.frame(matrix(nrow = nrow(sheet_df), ncol = length(sampleMetadata_user)))
    colnames(temp) <- sampleMetadata_user
    rownames(temp) <- rownames(sheet_df)
    temp[which(sheet_df[,1] == 'samp_name'),] <- sampleMetadata_user
    temp[which(sheet_df[,1] == '# section'),] <- 'User defined'
    temp[which(sheet_df[,1] == '# requirement_level_code'),] <- 'O'
    sheet_df <- cbind(sheet_df, temp)
  }
  
  # pre-fill assay_name
  if (any(colnames(sheet_df) == 'assay_name')) {
    sheet_df[nrow(sheet_df)+1,'assay_name'] = paste(assay_name, collapse = " | ")
  }
  
  addWorksheet(wb, sheet_name)
  writeData(wb, sheet_name, sheet_df, colNames = F)
   
  # add all other sheets ------------------------------------------------
  # select worksheets based on assay_type
  if (assay_type == 'metabarcoding') {
    sheet_ls <- c('experimentRunMetadata', 'taxaRaw', 'taxaFinal')
  } else if (assay_type == 'targeted') {
    sheet_ls <- c('stdData', 'eLowQuantData', 'ampData')
  }
  
  for (sheet_name in sheet_ls) {
    sheet_df <- suppressMessages(
      readxl::read_excel(full_temp_file_name, sheet=sheet_name, col_names = F)
    )
    # req_lev
    req_lev_row <- which(sheet_df[1,] == '# requirement_level_code')
    for (i in req_lev2rm) {
      if (i %in% sheet_df[req_lev_row,]) {
        sheet_df <- sheet_df[,-which(sheet_df[req_lev_row,] == i)] 
      }
    }
    addWorksheet(wb, sheet_name)
    writeData(wb, sheet_name, sheet_df, colNames = F) 
  }
  
  # add Drop-down values sheet  ------------------------------------------------
  sheet_name <- 'Drop-down values'
  addWorksheet(wb, sheet_name)
  vocab_df <- suppressMessages(
    readxl::read_excel(full_temp_file_name, sheet=sheet_name, col_names = T)
  )
  
  if (assay_type == 'targeted' & length(assay_name) > 1) {
    filtered_row <- vocab_df[which(vocab_df$term_name == 'detected_notDetected'), ]
    temp <- filtered_row[rep(seq_len(nrow(filtered_row)), each = length(assay_name)), ]
    temp$term_name <- paste0('detected_notDetected_', assay_name)
    vocab_df <- rbind(vocab_df, temp)
  }
  writeData(wb, sheet_name, vocab_df, colNames = T) 
  
  # Fill colours, comments, dropdown ------------------------------------------------
  # comment should include requirement level (and requirement level condition if any), description, example, field type, and field type option (for cont. vocab) or field format (field format i.e. date) 
  #text styles
  c1 <- createStyle(fontColour = "black")
  c_bold <- createStyle(fontColour = "black", textDecoration = "bold")
  c_red <- createStyle(fontColour = 'red')
  c_bold_red <- createStyle(fontColour = "red", textDecoration = "bold")
  
  #vocab_df <- readWorkbook(wb, sheet = 'Drop-down values')
  
  
  for (sheet_name in c('projectMetadata', 'sampleMetadata', sheet_ls)) {
    if (sheet_name == 'projectMetadata') {
      sheet_df <- readWorkbook(wb, sheet = sheet_name)
      
      # colours for req_lev
      for (i in req_lev) {
        addStyle(wb, sheet = sheet_name, 
                 style = createStyle(fgFill = req_col_df$col[which(req_col_df$requirement_level_code == i)]), 
                 rows = which(sheet_df$requirement_level_code == i) + 1,
                 cols = which(colnames(sheet_df)=='requirement_level_code'), gridExpand = T, stack = T)
      }
      
      # bold text for field names and headers
      addStyle(wb, sheet = sheet_name, style = c_bold,
               rows = 1:(nrow(sheet_df)+1), cols = which(colnames(sheet_df) == 'term_name'), gridExpand = T) #+1 to account for col headers
      addStyle(wb, sheet = sheet_name, style = c_bold,
               rows = 1, cols = 1:ncol(sheet_df), gridExpand = T)
      
      # drop-down
      for (i in 1:nrow(vocab_df)) {
        if (any(sheet_df$term_name == vocab_df$term_name[i])) {
          col_start=which(colnames(vocab_df)=='vocab1')
          col_end=which(colnames(vocab_df)==paste0('vocab', vocab_df$n_options[i]))
          col_start_alphabet=LETTERS[col_start] #this is always D
          col_end_alphabet = c(LETTERS, paste0('A', LETTERS))[col_end] #if it exceeds Z, it'll continue with AA, AB, AC etc. 
          val=paste0("'Drop-down values'!$", col_start_alphabet, '$', i+1, ":$", col_end_alphabet, '$', i+1) #+1 to account for the headers
          
          if (length(assay_name) == 1) {
            dataVal_col=which(colnames(sheet_df)=='project_level')
          } else if (length(assay_name) > 1) {
            dataVal_col=c(which(colnames(sheet_df)=='project_level'), grep('assay', colnames(sheet_df)))
          }
          dataValidation(wb, sheet = sheet_name, 
                         col= dataVal_col,
                         row = which(sheet_df$term_name==vocab_df$term_name[i])+1, #+1 to acount for the column header
                         type="list",
                         value=val,
                         showErrorMsg = FALSE,
                         allowBlank = TRUE)
          
        }
      }
      
      # comments
      for (i in sheet_df$term_name) {
        if (!i %in% projectMetadata_user) {
          req_lev_comm=input[which(input$term_name==i),]$requirement_level
          req_lev_cond=input[which(input$term_name==i),]$requirement_level_condition
          description=input[which(input$term_name==i),]$description
          example=input[which(input$term_name==i),]$example
          fieldtype=input[which(input$term_name==i),]$term_type
          #comment list
          #comm 1: req_lev_comm
          if (is.na(req_lev_cond)) {
            comm1=c('Requirement level : ', paste0(req_lev_comm, '\n'))
            if (req_lev_comm=='Mandatory') {
              style1=c(c_bold, c_bold_red)
            } else {
              style1=c(c_bold, c1)
            }
          } else {
            comm1=c('Requirement level : ', paste0(req_lev_comm, ' (', req_lev_cond, ')', '\n'))
            if (req_lev_comm=='Mandatory') {
              style1=c(c_bold, c_bold_red)
            } else {
              style1=c(c_bold, c1)
            }
          }
          
          #comm2
          comm2 = c('Description : ', paste0(description, '\n'),
                    'Example : ', paste0(example, '\n'))
          style2=c(c_bold, c1, c_bold, c1)
          
          #comm3 (field type)
          comm3=c('Field type : ', paste0(fieldtype, '\n'))
          style3=c(c_bold, c1)
          if (fieldtype =='controlled vocabulary') {
            txt1=input[which(input$term_name==i),]$controlled_vocabulary_options
            comm3=c('Field type : ', paste0(fieldtype, ' (', txt1, ')', '\n'))
            style3=c(c_bold, c_red)
          } else if (fieldtype == 'fixed format') {
            txt1=input[which(input$term_name==i),]$fixed_format
            comm3=c('Field type : ', paste0(fieldtype, ' (', txt1, ')', '\n'))
            style3=c(c_bold, c_red)
          } 
          comm=c(comm1, comm2, comm3)
          style=c(style1, style2, style3)
          writeComment(wb, 
                       sheet = sheet_name, 
                       col = which(colnames(sheet_df)=="term_name"), 
                       row = which(sheet_df$term_name==i)+1, #+1 to account for the col headers 
                       comment = createComment(comm,
                                               style = style,
                                               width = 7,
                                               height=10,
                                               visible = FALSE))
          
        }
        
      }
      
    } else if (!sheet_name == 'projectMetadata') {
      sheet_df <- readWorkbook(wb, sheet = sheet_name, colNames = F)
      if (sheet_name == 'sampleMetadata') {
        term_name_row <- which(sheet_df[,1] == 'samp_name')
        #term_name_row <- nrow(sheet_df)-1 #-1 because the last raw was added to pre-fill assay_name
      } else {
        term_name_row <- nrow(sheet_df)
      }
      colnames(sheet_df) <- sheet_df[term_name_row,]
      # colours for req_lev
      req_lev_row <- which(sheet_df[,1] == '# requirement_level_code')
      for (i in req_lev) {
        addStyle(wb, sheet = sheet_name, 
                 style = createStyle(fgFill = req_col_df$col[which(req_col_df$requirement_level_code == i)]), 
                 rows = req_lev_row,
                 cols = which(sheet_df[req_lev_row,] == i),
                 gridExpand = T, stack = T)
      }
      # bold text for field names
      addStyle(wb, sheet = sheet_name, style = c_bold, rows = term_name_row, cols = 1:ncol(sheet_df), gridExpand = T) 

      # drop-down
      for (i in 1:nrow(vocab_df)) {
        if (any(colnames(sheet_df) == vocab_df$term_name[i])) {
          col_start=which(colnames(vocab_df)=='vocab1')
          col_end=which(colnames(vocab_df)==paste0('vocab', vocab_df$n_options[i]))
          col_start_alphabet=LETTERS[col_start] #this is always D
          col_end_alphabet = c(LETTERS, paste0('A', LETTERS))[col_end] #if it exceeds Z, it'll continue with AA, AB, AC etc. 
          val=paste0("'Drop-down values'!$", col_start_alphabet, '$', i+1, ":$", col_end_alphabet, '$', i+1) #+1 to account for the headers
          
          dataValidation(wb, sheet = sheet_name, 
                         col=which(colnames(sheet_df)==vocab_df$term_name[i]),
                         row=term_name_row+1:100, #+1 because I want the dropdown to start one row below 'term_name'
                         type="list",
                         value=val,
                         showErrorMsg = FALSE,
                         allowBlank = TRUE)
        }
      }
      
      # comments
      for (i in colnames(sheet_df)) {
        if (!i %in% sampleMetadata_user) {
          req_lev_comm=input[which(input$term_name==i),]$requirement_level
          req_lev_cond=input[which(input$term_name==i),]$requirement_level_condition
          description=input[which(input$term_name==i),]$description
          example=input[which(input$term_name==i),]$example
          fieldtype=input[which(input$term_name==i),]$term_type
          if (grepl('detected_notDetected_', i)) {
            req_lev_comm=input[which(input$term_name=='detected_notDetected'),]$requirement_level
            req_lev_cond=input[which(input$term_name=='detected_notDetected'),]$requirement_level_condition
            description=input[which(input$term_name=='detected_notDetected'),]$description
            example=input[which(input$term_name=='detected_notDetected'),]$example
            fieldtype=input[which(input$term_name=='detected_notDetected'),]$term_type
          }
          #comment list
          #comm 1: req_lev_comm
          if (is.na(req_lev_cond)) {
            comm1=c('Requirement level : ', paste0(req_lev_comm, '\n'))
            if (req_lev_comm=='Mandatory') {
              style1=c(c_bold, c_bold_red)
            } else {
              style1=c(c_bold, c1)
            }
          } else {
            comm1=c('Requirement level : ', paste0(req_lev_comm, ' (', req_lev_cond, ')', '\n'))
            if (req_lev_comm=='Mandatory') {
              style1=c(c_bold, c_bold_red)
            } else {
              style1=c(c_bold, c1)
            }
          }
          
          #comm2
          comm2 = c('Description : ', paste0(description, '\n'),
                    'Example : ', paste0(example, '\n'))
          style2=c(c_bold, c1, c_bold, c1)
          
          #comm3 (field type)
          comm3=c('Field type : ', paste0(fieldtype, '\n'))
          style3=c(c_bold, c1)
          if (fieldtype =='controlled vocabulary') {
            txt1=input[which(input$term_name==i),]$controlled_vocabulary_options
            comm3=c('Field type : ', paste0(fieldtype, ' (', txt1, ')', '\n'))
            style3=c(c_bold, c_red)
          } else if (fieldtype == 'fixed format') {
            txt1=input[which(input$term_name==i),]$fixed_format
            comm3=c('Field type : ', paste0(fieldtype, ' (', txt1, ')', '\n'))
            style3=c(c_bold, c_red)
          } 
          comm=c(comm1, comm2, comm3)
          style=c(style1, style2, style3)
          
          writeComment(wb, 
                       sheet = sheet_name, 
                       col = which(colnames(sheet_df)==i), 
                       row = term_name_row, #+1 to account for the col headers 
                       comment = createComment(comm,
                                               style = style,
                                               width = 7,
                                               height=10,
                                               visible = FALSE))
          
        }
        
      }
      
      
      
    }
  }

  saveWorkbook(wb, 
               paste0('template_', project_id, '/', project_id, '.xlsx'), 
               overwrite = T)

  print(paste0('Your templates are in ', getwd(), '/template_', project_id))
}
