#' rocx
#'
#' This is the function that is called internally to generate the memos in different
#' forms.  This is not meant to be called by the user.
#' @import xml2
#' @export
rocx <- function(reference_docx, draft = TRUE, keep_old = FALSE,
                 use_bookdown = TRUE, ...) {

  # Use bookdown if defined, otherwise use rmarkdown
  if (('bookdown' %in% rownames(installed.packages())) & use_bookdown) {
    wd <- bookdown::word_document2
  } else {
    wd <- rmarkdown::word_document
  }

  # If there is no reference_docx supplied, run bookdown/rmarkdown
  if (missing(reference_docx)) {
    warning('No reference_docx specified in rocx.')
    return(wd())
  }

  if (!file.exists(reference_docx))
    stop('Specified reference_docx file not found.')

  config <- wd(reference_docx = reference_docx, ...)

  # Set on_exit to equal the rocx_exit function defined in this package
  config$on_exit <- function() {

    # Verify that the output file has been created.
    # If it hasn't, exit
    my_envir <- parent.frame()
    if (my_envir$output_dir == '.') {
      my_file <- my_envir$output_file
    } else {
      my_file <- sprintf('%s/%s', my_envir$output_dir, my_envir$output_file)
    }
    if (!file.exists(my_file))
      return(invisible(NULL))

    # Rename the output file
    new_file <- sprintf('%s_old.rocx', my_file)
    file.rename(from = my_file, to = new_file)

    if (file.exists('rocx_temp'))
      unlink('rocx_temp', recursive = TRUE)
    unzip(new_file, exdir = 'rocx_temp')

    # Get the template material to add.  This reflects the changes from
    # substituting the YAML variables
    new_header <- get_template_info(my_envir$yaml_front_matter, reference_docx)
    new_header <- xml_children(xml_child(new_header, 1))

    # Here I will try to strip out the attributes from new_header that
    # are not included in in_file.  I am hoping this will eliminate the
    # corrupt file errors that I am encountering when I try to open
    # the file in Word.
    #new_header <- clean_namespace(new_header, in_file)

    in_file <- read_xml('rocx_temp/word/document.xml')

    my_xml <- xml_find_all(in_file, '//w:p')
    for (i in seq_along(my_xml)) {
      if ('&HEADER&' %in% xml_text(my_xml[i])) {
        if (length(new_header) > 1) {
          for (j in seq(from = length(new_header), to = 2, by = -1)) {
            xml_add_sibling(my_xml[i], new_header[j])
          }
        }
        xml_replace(my_xml[i], new_header[i])
      }
    }
    write_xml(in_file, 'rocx_temp/word/document.xml')

    # If draft == FALSE then remove any DRAFT indicators from the header
    if (!draft)
      remove_draft()

    # Re-compress the files back into a docx
    setwd('rocx_temp')
    zip(my_file, files = list.files(), flags = '-r', zip = 'zip')
    file.rename(from = my_file,
                to = sprintf('../%s', my_file))
    setwd('..')
    unlink('rocx_temp', recursive = TRUE)

    # Unless directed in YAML option to keep original pandoc output, remove it
    if (!keep_old)
      unlink(new_file)
  }


  config
}

#' Get Template Information
#'
#' This function will read in the header reference file and replace any codes with
#' values supplied in the YAML header.
get_template_info <- function(yaml_front_matter, reference_docx) {

  # Convert any unexecuted r code in the YAML header
  yaml_front_matter <- evaluate_yaml(yaml_front_matter)

  template <- read_xml(unz(reference_docx, 'word/document.xml'))

  # Check to make sure that template has only one child; otherwise, the xml
  # is structured in a way I do not understand
  if (xml_length(template) > 1)
    stop('Error in rocx::get_template_infor:  XML file has more than one child from beginning')

  # Select the paragraphs and verify that there is at least one
  my_text <- xml_find_all(template, '//w:p')
  n <- length(my_text)
  if (n < 1) {
    warning('Template file appears to be empty.')
    return(invisible(NULL))
  }

  # Create list of YAML variables as they would appear in the template
  yaml_names <- sprintf('$%s$', tolower(names(yaml_front_matter)))

  # Cycle through the paragraphs in the template looking for YAML variables
  i <- 1
  while(i <= n) {  # Cycle through each paragraph
    is_present <- (yaml_names %in% xml_text(my_text[i]))

    if (any(is_present)) {   # At least one of the YAML variables are in the paragraph
      for (k in which(is_present)) {
        header <- yaml_names[[k]]

        num_yamls <- length(yaml_front_matter[[k]])
        if (num_yamls == 1) {
          replace_text(my_text[i], yaml_front_matter[[k]][[1]])
        } else if (num_yamls > 1) {
          new_vars <- yaml_front_matter[[k]]
          for (j in seq(from = num_yamls, to = 2, by = -1)) {
            # Replace the text in the current node
            replace_text(my_text[i], yaml_front_matter[[k]][j])

            # Then copy the node
            xml_add_sibling(my_text[i], my_text[i])
          }
          xml_text(my_text[i]) <- new_vars[1]
        }
      }
    }
    i <- i + 1
  }

  ret_val <- xml_children(xml_child(template, 1))

  # We don't want to include the 'sectPr' or other extraneous tags.  I will only
  # use those that are 'p' or 'tbl' tags
  for (this_node in ret_val) {
    if (!xml_name(this_node) %in% c('p', 'tbl')) {
      xml_remove(this_node)
    }
  }

  # Bookmark tags are unnecessary (I think).  Remove them.
  invisible(lapply(xml_find_all(template, '//w:bookmarkStart'), xml_remove))
  invisible(lapply(xml_find_all(template, '//w:bookmarkEnd'), xml_remove))

  template
}

#' @import xml2
replace_text <- function(node, new_text) {
  temp <- xml_find_all(node, './/text()')
  if (length(temp) > 1) {
    for (j in seq(2, length(temp))) {
      xml_text(temp[j]) <- ''
    }
  }
  xml_text(temp[1]) <- new_text
}

#' Evaluate YAML
#'
#' This function reads in the YAML variables and evaluates any included R code.
#' @param yfm YAML front matter (stored by knitr as yaml_front_matter)
#' @return The YAML front matter after R code evaluation
evaluate_yaml <- function(yfm) {
  for (i in seq_along(yfm)) {
    for (j in seq_along(yfm[[i]])) {
      this_val <- yfm[[i]][j]

      if (is.character(this_val)) {
        if (grepl("`r (.*)`", this_val)) {
          this_command <- gsub('`r (.*)`', '\\1', this_val)
          try({
            result <- as.character(eval(parse(text = this_command)))
            this_val <- result
          })
        }
      }
      yfm[[i]][j] <- this_val
    }
  }
  yfm
}

#' Clean Namespace
#'
#' This function cleans the namespace file to remove formats that I don't believe
#' need to be there.  I am not sure this file is necessary.  Consider deprecating it
#' if you can confirm that the function does nothing of value.
#'
#' This function is not meant to be called by the user.
clean_namespace <- function(add_file, skeleton_file) {
  # Find elements in the namespace that appear in new_file but not in skeleton_file
  bad_names <- setdiff(names(xml_ns(add_file)), names(xml_ns(skeleton_file)))

  # I couldn't figure out how to strip out attributes using xml2, so I will
  # convert to a character variable and strip out the attributes that way.
  temp <- as.character(add_file)

  for (i in bad_names) {
    # Remove node attributes
    my_pattern <- sprintf(' %s:\\w*=\\\"\\w*\\\"', i)
    temp <- gsub(my_pattern, '', temp)

    #remove attribute namespace
    my_pattern <- sprintf(' \\w*:%s=[^ >]*\\\"', i)
    temp <- gsub(my_pattern, '', temp)
  }

  temp <- gsub(' mc:Ignorable=\"w14 wp14\"', '', temp)
  temp <- gsub(' w:rsidR=\"\\w*\"', '', temp)
  temp <- gsub(' w:rsidRPr=\"\\w*\"', '', temp)
  temp <- gsub(' w:rsidRDefault=\"\\w*\"', '', temp)
  temp <- gsub(' w:rsidP=\"\\w*\"', '', temp)

  add_file <- read_xml(temp)
}

#' Remove Draft
#'
#' This function will check the header files to determine if the text in the
#' header files indicate that this is a draft.  If found, the draft text
#' will be removed.
#' @import xml2
remove_draft <- function() {
  # There may be two header files in the xml
  for (n in c(1,2)) {
    header_file <- sprintf('rocx_temp/word/header%i.xml', n)
    if (file.exists(header_file)) {
      resave <- FALSE
      in_file <- read_xml(header_file)
      nodes <- xml_find_all(in_file, '//w:p')
      for (i in seq_along(nodes)) {
        if (xml_text(xml_child(nodes, i)) == 'DRAFT') {
          replace_text(xml_child(nodes, i), '')
          resave <- TRUE
        }
      }
      if (resave)
        write_xml(in_file, header_file)
    }
  }
  invisible(NULL)
}

#' ROCX Exit
#'
#' This function will serve as the new on_exit function passed to rmarkdown or
#' bookdown.
rocx_exit <- function() {

  # Verify that the output file has been created.
  # If it hasn't, exit
  my_envir <- parent.frame()
  if (my_envir$output_dir == '.') {
    my_file <- my_envir$output_file
  } else {
    my_file <- sprintf('%s/%s', my_envir$output_dir, my_envir$output_file)
  }
  if (!file.exists(my_file))
    return(invisible(NULL))

  # Rename the output file
  new_file <- sprintf('%s_old.rocx', my_file)
  file.rename(from = my_file, to = new_file)

  if (file.exists('rocx_temp'))
    unlink('rocx_temp', recursive = TRUE)
  unzip(new_file, exdir = 'rocx_temp')

  # Get the template material to add.  This reflects the changes from
  # substituting the YAML variables
  new_header <- get_template_info(my_envir$yaml_front_matter, reference_docx)
  new_header <- xml_children(xml_child(new_header, 1))

  # Here I will try to strip out the attributes from new_header that
  # are not included in in_file.  I am hoping this will eliminate the
  # corrupt file errors that I am encountering when I try to open
  # the file in Word.
  #new_header <- clean_namespace(new_header, in_file)

  in_file <- read_xml('rocx_temp/word/document.xml')

  my_xml <- xml_find_all(in_file, '//w:p')
  for (i in seq_along(my_xml)) {
    if ('&HEADER&' %in% xml_text(my_xml[i])) {
      if (length(new_header) > 1) {
        for (j in seq(from = length(new_header), to = 2, by = -1)) {
          xml_add_sibling(my_xml[i], new_header[j])
        }
      }
      xml_replace(my_xml[i], new_header[i])
    }
  }
  write_xml(in_file, 'rocx_temp/word/document.xml')

  # If draft == FALSE then remove any DRAFT indicators from the header
  if (!draft)
    remove_draft()

  # Re-compress the files back into a docx
  setwd('rocx_temp')
  zip(my_file, files = list.files(), flags = '-r', zip = 'zip')
  file.rename(from = my_file,
              to = sprintf('../%s', my_file))
  setwd('..')
  unlink('rocx_temp', recursive = TRUE)

  # Unless directed in YAML option to keep original pandoc output, remove it
  if (!keep_old)
    unlink(new_file)
}
