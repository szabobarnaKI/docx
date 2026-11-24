{smcl}
{* *! version 1.0.2  24nov2026}{...}
{title:Title}

{phang}
{bf:docxlabel} {hline 2} Copy or report a Microsoft Purview sensitivity label of a .docx file


{title:Syntax}

{pstd}Copy the label of {it:source} to {it:target}

{p 8 17 2}
{cmd:docxlabel} {it:target}{cmd:,} {cmdab:from(}{it:source}{cmd:)} [{cmdab:sav:ing(}{it:newfile}{cmd:)} {cmd:replace}]

{pstd}Report the label of a file

{p 8 17 2}
{cmd:docxlabel} {it:file}

{synoptset 20 tabbed}{...}
{synopthdr}
{synoptline}
{synopt:{opt from(source)}}labelled file whose sensitivity label is copied{p_end}
{synopt:{opt sav:ing(newfile)}}write the result to {it:newfile}; default is to change {it:target} in place{p_end}
{synopt:{opt replace}}allow {cmd:saving()} to overwrite an existing file{p_end}
{synoptline}

{pstd}Names with spaces must be in double quotes. Relative names are relative to the current working directory.


{title:Description}

{pstd}
{cmd:docxlabel} copies the sensitivity (MIP) label metadata of an existing labelled Office Open XML file
(.docx, and also .xlsx, .pptx) to another file, for example one created with {helpb putdocx}, so that Word
opens the new file as already labelled. It does not start Word, PowerShell, VBScript or any other external
program: the file is rewritten with Stata's built-in Java ({cmd:java.util.zip}).

{pstd}
Word stores a label in two places, and {cmd:docxlabel} copies both when present:

{phang2}1. the {cmd:MSIP_Label_}{it:guid}{cmd:_*} custom document properties in {cmd:docProps/custom.xml}
(Enabled, SetDate, Method, Name, SiteId, ActionId, ContentBits);{p_end}
{phang2}2. the part {cmd:docMetadata/LabelInfo.xml} with its relationship and content type (newer Office versions).{p_end}

{pstd}
Existing label metadata in {it:target} is removed first; other custom properties of {it:target} are kept and
renumbered. Nothing else in {it:target} is changed.

{pstd}
{cmd:docxlabel} is part of the {helpb docx} package and can also be called as {cmd:docx label} with the same syntax.

{pstd}
Without {cmd:from()}, {cmd:docxlabel} only reports the label of {it:file}.


{title:Remarks}

{pstd}
Only the label metadata is copied. Visual content markings that a label may add (header, footer,
watermark) are not created; Word may add them the next time it saves the file.

{pstd}
Files encrypted by a label with protection (or by a password) are not zip files and cannot be read or
changed; {cmd:docxlabel} stops with an error for them.

{pstd}
Close {it:target} in Word before running {cmd:docxlabel}.


{title:Examples}

{phang2}{cmd:. putdocx begin}{p_end}
{phang2}{cmd:. putdocx paragraph}{p_end}
{phang2}{cmd:. putdocx text ("Results")}{p_end}
{phang2}{cmd:. putdocx save "report.docx", replace}{p_end}
{phang2}{cmd:. docxlabel "report.docx", from("H:/templates/labelled.docx")}{p_end}
{phang2}{cmd:. docxlabel "report.docx"}{p_end}


{title:Stored results}

{synoptset 16 tabbed}{...}
{p2col 5 16 20 2: Scalars}{p_end}
{synopt:{cmd:r(found)}}1 if a label was found (in {it:source} when copying){p_end}
{synopt:{cmd:r(nprops)}}number of {cmd:MSIP_Label_*} properties{p_end}
{synopt:{cmd:r(labelinfo)}}1 if {cmd:LabelInfo.xml} is present{p_end}
{p2col 5 16 20 2: Macros}{p_end}
{synopt:{cmd:r(name)}}label name (if stored in the file){p_end}
{synopt:{cmd:r(labelid)}}label GUID{p_end}
{synopt:{cmd:r(siteid)}}tenant GUID{p_end}
{synopt:{cmd:r(file)}}target or queried file{p_end}
{synopt:{cmd:r(source)}}source file (copy only){p_end}
{synopt:{cmd:r(saved)}}file written (copy only){p_end}


{title:Also see}

{psee}{helpb docx}, {helpb putdocx}, {helpb java}{p_end}
