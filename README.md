# DlrDocsProcessor
Convert [DLR documentation](https://github.com/IronLanguages/dlr/tree/master/Docs) from Word to Markdown for inclusion on the [DLR repo](https://github.com/IronLanguages/dlr) wiki.

Some background at https://github.com/IronLanguages/dlr/issues/246.

This project carries out the following (everything takes place in the `Filter` project, unless noted otherwise):

* Identify and mark code blocks in such a way that they can be recognized by Pandoc (`Driver` project)
* Generate the per-Word-document TOC into `_sidebar.md`
* Split each Word document into multiple Markdown files, based on the 1st-level header. (This involves multiple passes over the document, each time returning to Pandoc only the relevant section)
* Document headings are auto-numbered; Pandoc doesn't recreate this numbering in Markdown. Recreate them in the text, but not in the generated IDs. This entails using embedded HTML to explicitly write the IDs
* Replace HTML markup for images with Markdown, by removing positioning information
* Replace HTML markup for underlines with Markdown **bold**
* Trim excess spaces from beginning of lines in code blocks
* Reduce bullet list spacing from the Pandoc-generated 3 spaces to 1 space (`Driver` project)
