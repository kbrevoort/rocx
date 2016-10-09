# ROCX:  An R Package for (Limited) Use of docx Templates with R Markdown

Yihui Xie's R packages for reproducible research (including [knitr](http://yihui.name/knitr/), [rmarkdown](http://rmarkdown.rstudio.com/), and [bookdown](https://bookdown.org/)) make implementing a reproducible workflow relatively easy.
While users of these packages can create templates to customize HTML or LaTeX output, there is no similar capability for working with Microsoft Word's docx format.
The `rocx` package seeks to fill this gap by allowing users to specify `reference_docx` files that can serve as templates for the generated docx documents.

*How rocx works*:  `rocx` is a function that substitutes for the output type in documents generated by `knitr`.  
Instead of specifying the output type in the YAML header of an R Markdown document as `rmarkdown::word_document` or `bookdown::word_document2`, the user specifies `rocx::rocx`.
The `rocx` function generates the docx file that would result from `rmarkdown` or `bookdown`, using a template provided by the user as the `reference_docx`.  
The function then copies any text or tables in the template, after replacing optionally supplied variable values, into that output file by editing the xml code that makes up the docx.

This package may be useful in helping to produce documents (such as memos) that need to be produced as Word documents using a specific look and feel.  
This package offers limited template flexibility relative to what is available in `knitr` for HTML or LaTeX, but it will hopefully prove useful to some.
