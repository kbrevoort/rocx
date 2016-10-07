#' Information Memo for the Director
#'
#' Generate an information memo for the Director.
#' @export
information_memo <- function(...) {
  reference_docx <- sprintf('%s/rmarkdown/templates/information2director/resources/information2director_template.docx',
                            find.package(package = 'rocuments'))
  memo_engine(reference_docx = reference_docx, ...)
}

#' CCP Analysis Memo
#'
#' Generate a memo detailing the results of a CCP analysis.  If the results are to
#' be circulated to the Directory, then use information_memo instead.
#' @export
ccp_memo <- function(...) {
  reference_docx <- sprintf('%s/rmarkdown/templates/ccp_memo/resources/ccp_memo_template.docx',
                            find.package(package = 'rocuments'))
  memo_engine(reference_docx = reference_docx, ...)
}

#' Memo Engine
#'
#' This is the function that is called internally to generate the memos in different
#' forms.  This is not meant to be called by the user.
#' @import xml2
memo_engine <- function(reference_docx = reference_docx, fig_caption = TRUE,
                        md_extensions = NULL, pandoc_args = NULL,
                        draft = TRUE, keep_old = FALSE, ...) {

  #ref_docx <- sprintf('%s/rmarkdown/templates/information2director/resources/information2director_template.docx',
  #                    find.package('rocuments'))
  #ref_docx <- '/Users/ken/Documents/template_files/revised_information_template.docx'
  config <- bookdown::word_document2(fig_caption = fig_caption,
                                     md_extensions = md_extensions,
                                     pandoc_args = pandoc_args,
                                     reference_docx = reference_docx,
                                     ...)

  config$on_exit <- function() {
    # At this point, the output file will have been created.  Verify this to start.
    my_envir <- parent.frame()

    if (my_envir$output_dir == '.') {
      my_file <- my_envir$output_file
    } else {
      my_file <- sprintf('%s/%s', my_envir$output_dir, my_envir$output_file)
    }

    if (file.exists(my_file)) {
      new_file <- sprintf('%s_old.docx', my_file)
      file.rename(from = my_file, to = new_file)

      if (file.exists('rocument_temp'))
        unlink('rocument_temp', recursive = TRUE)
      unzip(new_file, exdir = 'rocument_temp')

      # Now get the header material to add
      #header_xml <- get_header_info(my_envir$yaml_front_matter)
      new_header <- get_header_info(my_envir$yaml_front_matter, reference_docx)

      # Remove any bookmarks from the XML
      invisible(lapply(xml_find_all(new_header, '//w:bookmarkStart'), xml_remove))
      invisible(lapply(xml_find_all(new_header, '//w:bookmarkEnd'), xml_remove))

      in_file <- read_xml('rocument_temp/word/document.xml')

      # Here I will try to strip out the attributes from new_header that
      # are not included in in_file.  I am hoping this will eliminate the
      # corrupt file errors that I am encountering when I try to open
      # the file in Word.
      #new_header <- clean_namespace(new_header, in_file)
      new_header <- xml_children(xml_child(new_header, 1))

      my_text <- xml_find_all(in_file, '//w:p')
      for (i in seq(1, length(my_text)) ) {
        if ('&HEADER&' %in% xml_text(my_text[i])) {
          if (length(new_header) > 1) {
            for (j in seq(from = length(new_header), to = 2, by = -1)) {
              xml_add_sibling(my_text[i], new_header[j])
            }
          }
          xml_replace(my_text[i], new_header[1])
        }
      }
      write_xml(in_file, 'rocument_temp/word/document.xml')

      # If draft == FALSE then remove the DRAFT indicator from the header
      if (!draft) {
        in_file <- read_xml('rocument_temp/word/header1.xml')
        for (i in seq_along(in_file)) {
          if (xml_text(xml_child(in_file, i)) == 'DRAFT') {
            replace_text(xml_child(in_file, i), '')
          }
        }
        write_xml(in_file, 'rocument_temp/word/header1.xml')
      }

      # Re-zip file
      setwd('rocument_temp')
      zip(my_file, files = list.files(), flags = '-r', zip = 'zip')
      file.rename(from = my_file,
                  to = sprintf('../%s', my_file))
      setwd('..')
      unlink('rocument_temp', recursive = TRUE)

      # Unless directed to keep original pandoc output, remove it
      if (!keep_old)
        unlink(new_file)
    } else {
      stop('rocument:  Knitr file output does not seem to exist.')
    }
  }
  config
}

#' Get Header Information
#'
#' This function will read in the header reference file and replace any codes with
#' values supplied in the YAML header.
get_header_info <- function(yaml_front_matter, reference_docx) {

  yaml_front_matter <- convert_yaml(yaml_front_matter)

  yaml_names <- sprintf('$%s$', tolower(names(yaml_front_matter)))

  # First, read in the header file.  Since I'm not writing the XML file out, I can
  # read directly from a docx file
  in_file <- read_xml(unz(reference_docx, 'word/document.xml'))

  # Check to make sure that in_file has only one child (otherwise, something is wrong)
  if (xml_length(in_file) > 1)
    stop('Error in rocument::get_header_infor:  XML file has more than one child from beginning')

  # I want to cycle through the paragraphs to choose text to replace
  my_text <- xml_find_all(in_file, '//w:p')
  n <- length(my_text)
  if (n < 1) stop('Invalid number of rows in xml_find_all')

  i <- 1
  while(i <= n) {  # This is cycling through the paragraphs of in_file
    is_present <- (yaml_names %in% xml_text(my_text[i]))

    if (any(is_present)) {
      for (k in which(is_present)) {
        header <- yaml_names[[k]]

        if (length(yaml_front_matter[[k]]) == 1) {
          replace_text(my_text[i], yaml_front_matter[[k]][[1]])
        } else {
          new_header <- yaml_front_matter[[k]]
          for (j in length(new_header):2) {
            # First, rename the node
            replace_text(my_text[i], yaml_front_matter[[k]][j])

            # Then copy the node
            xml_add_sibling(my_text[i], my_text[i])
          }
          xml_text(my_text[i]) <- new_header[1]
        }
      }
    }
    i <- i + 1
  }

  ret_val <- xml_children(xml_child(in_file, 1))

  # I think the xml tags named 'sectPr' are causing problems.  I will only use
  # those that are 'p" or 'tbl' tags
  for (this_node in ret_val) {
    if (!xml_name(this_node) %in% c('p', 'tbl')) {
      xml_remove(this_node)
    }
  }
  in_file
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

#' Convert YAML
#'
#' This function reads in the YAML variables and converts any R code that is included
#' into the values the code returns.
convert_yaml <- function(yfm) {
  for (i in seq_along(yfm)) {
    for (j in seq(1, length(yfm[[i]]))) {
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


