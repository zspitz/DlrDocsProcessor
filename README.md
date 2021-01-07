# DlrDocsProcessor
Convert [DLR documentation](https://github.com/IronLanguages/dlr/tree/master/Docs) from Word to Markdown for inclusion on the [DLR repo](https://github.com/IronLanguages/dlr) wiki.

Some background at https://github.com/IronLanguages/dlr/issues/246.

Goals:

* Split each Word document into multiple Markdown files, based on the 1st-level header
* Replace HTML markup for images with Markdown
* Replace HTML markup for underlines with Markdown **bold**
* Document headings are auto-numbered; Pandoc doesn't recreate this numbering in Markdown. Recreate them in the text, but not in the generated IDs.
* Generate TOC for each document; put into `_sidebar.md`
  * Each document's TOC will show all the headings for that document + links to the other documents
  * Alternatively, the other document's headers will be hidden in a details sction, if that works on a wiki

