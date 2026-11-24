{smcl}
{* *! version 1.0  24sep2025}{...}
{findalias asfradohelp}{...}
{vieweralsosee "" "--"}{...}
{vieweralsosee "[R] help" "help help"}{...}
{viewerjumpto "Syntax" "docx##syntax"}{...}
{viewerjumpto "Description" "docx##description"}{...}
{viewerjumpto "Example" "docx##example"}{...}
{title:Title}

{phang}
{bf:docx} {hline 2} Extended functions for putdocx


{marker syntax}{...}
{title:Syntax}

{p 8 17 2}
{cmd:docx}
[{it:docx_spec}]
[{cmd:,} {it:options}]

{synoptset 23 tabbed}{...}
{synopthdr}
{synoptline}
{syntab:docx_spec}
{synopt:{opt begin}}Initiates a new docx document with page numbers in the footer.{p_end}
{synopt:{opt start}}See {cmd:begin}{p_end}
{synopt:{opt open}}See {cmd:begin}{p_end}
{synopt:{opt maintitle}}Adds the specified text as main title to the current page.{p_end}
{synopt:{opt date}}Adds the current date.{p_end}
{synopt:{opt author}}Adds the author name.{p_end}
{synopt:{opt starttoc}}Adds table of contents at the current point, with within-document links to titles and subtitles. Table of contents needs to be "started" before calling {cmd:docx title} and {cmd:docx subtitle}.{p_end}
{synopt:{opt title}}Inserts the specified text as a title.{p_end}
{synopt:{opt subtitle}}Inserts the specified text as a subtitle.{p_end}
{synopt:{opt paragraph}}Start a new paragraph with default settings ({it:halign(left) indent(left,0.5) font("`Arial'", 11)}).{p_end}
{synopt:{opt emptylines}}Adds empty lines.{p_end}
{synopt:{opt tableshader}}Chades every other line in a table, to improve readability.{p_end}
{synopt:{opt checkpoint}}Stores the document as docx file, but without closing the putdocx session. Thus, you can continue to edit the docx document in memory.{p_end}
{synopt:{opt save}}Stores the document as docx file, and closes the in-memory docx document.{p_end}
{synopt:{it:everything else}}What is not defined above is passed forward to {cmd:putdocx} without change, thus {cmd:docx} acts as an abbreviation to {cmd:putdocx}.{p_end}
{synoptline}
{syntab:maintitle options}
{synopt:{opt text(string)}}The text to be added as main title. By default, this starts a new paragraph.{p_end}
{synopt:{opt newline}}Continue the last paragraph in a new line, without starting a new paragraph. Recommended if the main title needs to have multiple rows.{p_end}

{syntab:author options}
{synopt:{opt text(string)}}The text to be added as author name.{p_end}

{syntab:starttoc options}
{synopt:{opt nopage}}Do not fill the page column title in the table of contents. If you add titles and subtitles and specify page numbers, these will still appear in the table of contents.{p_end}

{syntab:title options}
{synopt:{opt text(string)}}The text to be added as section title.{p_end}
{synopt:{opt toc}}Add the current title to the table of contents.{p_end}
{synopt:{opt page:number(#)}}Add the specified number as page number to the table of contents. Page number needs to be specified manually (at least up to Stata version 19.0).{p_end}

{syntab:subtitle options}
{synopt:{opt text(string)}}The text to be added as section subtitle.{p_end}
{synopt:{opt toc}}Add the current title to the table of contents.{p_end}
{synopt:{opt page:number(#)}}Add the specified number as page number to the table of contents. Page number needs to be specified manually (at least up to Stata version 19.0).{p_end}
{synopt:{opt indentation(#)}}In the table of contents, use extra indentation. # must be a nonnegative integer.{p_end}

{syntab:emptylines options}
{synopt:{opt lines(#)}}The number of empty lines to be added.{p_end}

{syntab:tableshader options}
{synopt:{opt tbl(table_name)}}The name of the table to be colored.{p_end}
{synopt:{opt head:errow(#)}}The number of header rows. These are colored uniformly using color1. Specify it only if you want to override the automatic obtainment of header row count.{p_end}
{synopt:{opt title}}Specify if the table contains a title (taking up the first row, that should not be colored).{p_end}
{synopt:{opt notes}}Specify if the table contains notes (taking up the last row, that should not be colored).{p_end}
{synopt:{opt color1(sspec)}}The color used in the header (see the {opt shader} option in {help putdocx table} for color format details).{p_end}
{synopt:{opt color2(sspec)}}The color used in every other row of the body of the table (see the {opt shader} option in {help putdocx table} for color format details).{p_end}

{syntab:checkpoint options}
{synopt:{opt name(string)}}The path and name of the document to be created. It needs to be enclosed in quotes or compound double quotes.{p_end}
{synopt:{opt replace}}Allow to overwrite an existing file.{p_end}

{syntab:save options}
{synopt:{opt path(string)}}The path of the document to be created. The path needs to be enclosed in quotes or compound double quotes.{p_end}
{synopt:{opt filestub(string)}}The stub name for the output file. {cmd:docx save} automatically adds a rolling number to the file name and avoids overwriting old files.
The filestub needs to be enclosed in quotes or compound double quotes.{p_end}
{synopt:{opt sec:uritylabel(string)}}The path and name of a document that is saved with a security label, that needs to be cloned into the new docx.{p_end}
{synopt:{opt rc}}Display the return information from the security label copying subcommand.{p_end}
{synoptline}


{marker description}{...}
{title:Description}

{pstd}
{cmd:docx} simplifies the generation of output files in docx format. On top of simplified commands and standardized text format, 
it allows to add a table of contents that is updated automatically when calling {cmd:docx title} or {cmd:docx subtitle}. Currently (in Stata 19.0), 
page numbers need to be defined manually, but the TOC is created by default with internal links to each section of the 
document. {cmd:docx} also allows to copy a security label from an existing docx file to the one created by Stata.


{marker example}{...}
{title:Example}

{phang}{cmd:. docx open}{p_end}

{phang}{cmd:. docx paragraph}{p_end}

{phang}{cmd:. docx emptylines, lines(14)}{p_end}

{phang}{cmd:. docx maintitle, text(This is the main title)}{p_end}

{phang}{cmd:. docx date}{p_end}

{phang}{cmd:. docx author, text(This is the author's name)}{p_end}

{phang}{cmd:. docx pagebreak}{p_end}

{phang}{cmd:. docx starttoc}{p_end}

{phang}{cmd:. docx pagebreak}{p_end}

{phang}{cmd:. docx title, text(1 Caption number one))}{p_end}

{phang}{cmd:. docx subtitle, text(1.1 First subsection))}{p_end}

{phang}{cmd:. putdocx textblock begin, indent(left,0.5)}{p_end}

{phang}{cmd:.} {it:Text to be added to the paragraph...}{p_end}

{phang}{cmd:. putdocx textblock end}{p_end}

{phang}{cmd:. docx subtitle, text(1.2 Second subsection))}{p_end}

{phang}{cmd:. docx save, path("C:/Project1/Results") filestub("output")}{p_end}
