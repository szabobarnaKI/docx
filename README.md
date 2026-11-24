# docx

**Extended functions for `putdocx` in Stata** – build consistently formatted Word (`.docx`) reports with short, readable commands, and give them the right Microsoft Purview sensitivity label.

`docx` adds one-line commands for the parts of a report that otherwise take several `putdocx` calls: title page, section headings, a linked table of contents, zebra-shaded tables, checkpoints, numbered output files and sensitivity labels. Anything `docx` does not define is passed straight to `putdocx`, so you can mix `docx …` and plain `putdocx` commands freely.

## Features

- **One-line document setup** – `docx begin` opens a new document with 0.5-inch margins and centred page numbers in the footer.
- **Title page helpers** – main title, date and author with a standard layout.
- **Standardised headings** – `docx title` and `docx subtitle` give uniform section and subsection headings.
- **Table of contents with internal links** – `docx starttoc` inserts a TOC; each heading added with the `toc` option gets a TOC line that links to it.
- **Zebra-shaded tables** – `docx tableshader` shades the header rows and every other body row of a `putdocx` table, and detects the number of header rows itself.
- **Checkpoints** – `docx checkpoint` writes the document to disk without closing it, so you can look at intermediate output and keep editing.
- **Never overwrite old output** – `docx save` adds a running number to the file name (`report1.docx`, `report2.docx`, …).
- **Sensitivity labels** – `docx label` (or `docxlabel` on its own) copies a Microsoft Purview sensitivity label from an already-labelled file, so Word opens the Stata output as labelled and editable instead of read-only. `docx save, securitylabel()` does this as part of saving.
- **`putdocx` pass-through** – `docx pagebreak`, `docx table …`, `docx image …`, `docx text …` and so on go to `putdocx` unchanged.

## Requirements

- Stata **19** or later.
- The label functions use only Stata's built-in Java (`java.util.zip`). They do not start Word, PowerShell, VBScript or any other program, so they also work where these are blocked by IT policy.

## Installation

```stata
net install docx, from("https://raw.githubusercontent.com/szabobarnaKI/docx/main/") replace
help docx
```

This installs `docx` and `docxlabel` with their help files. To update later, run the same `net install` line again.

> `docxlabel.ado` contains a Java class and must be loaded by autoloading from the ado-path (which `net install` takes care of). Do not `run` or `do` the ado-file: an ado-file has its own Java instance, and the command would then fail with a "class not found" error.

## Quick start

```stata
* --- title page ------------------------------------------------------------
docx begin
docx paragraph
docx emptylines, lines(14)
docx maintitle, text(Statistical analysis report)
docx maintitle, text(Project 1) newline
docx date
docx author, text(Jane Doe)
docx pagebreak                        // passed through to putdocx

* --- table of contents -----------------------------------------------------
docx starttoc
docx pagebreak

* --- body ------------------------------------------------------------------
docx title, text(1 Baseline characteristics) toc page(3)
docx subtitle, text(1.1 Study population) toc page(3)

putdocx textblock begin, indent(left,0.5)
Text to be added to the paragraph...
putdocx textblock end

sysuse auto, clear
putdocx table t1 = data(make price mpg) in 1/10, varnames headerrow(1)
docx tableshader, tbl(t1)

docx subtitle, text(1.2 Outcomes) toc page(4)

* --- intermediate copy, then final save with a sensitivity label -----------
docx checkpoint, name("C:/Project1/Results/draft.docx") replace
docx save, path("C:/Project1/Results") filestub("report") ///
    securitylabel("H:/templates/labelled.docx")
display "Saved to `r(fname)'"
```

## Subcommands

Subcommand names are not case-sensitive.

| Subcommand | What it does |
|---|---|
| `begin` (also `open`, `start`) | Clears any document in memory and starts a new one with 0.5-inch margins and a centred page number in the footer. Other options are passed to `putdocx begin` (e.g. `landscape`, `pagesize()`). |
| `maintitle, text()` `[newline]` | Main title: centred, Arial 24, italic. `newline` continues the previous main-title paragraph on a new line (for titles of several lines). |
| `author, text()` | Author line: centred, Arial 16, italic. |
| `date` | Today's date as a centred paragraph, in `YYYY-MM-DD` format. |
| `starttoc` `[, nopage]` | Inserts a table of contents at the current position. `nopage` leaves out the "Page" column heading. Must come **before** any `title`/`subtitle` that uses `toc`. |
| `title, text()` `[toc page(#)]` | Section heading: left-aligned, Arial 13, bold. `toc` adds a linked line to the table of contents; `page(#)` writes that page number in the TOC. |
| `subtitle, text()` `[toc page(#) indentation(#)]` | Subsection heading: left-aligned, Arial 11, bold. `indentation(#)` sets how far its TOC line is indented (default 1). |
| `paragraph` | New paragraph in the body style: left-aligned, 0.5-inch left indent, Arial 11. |
| `emptylines, lines(#)` | Adds `#` line breaks to the current paragraph. |
| `tableshader, tbl(name)` `[headerrow(#) title notes color1() color2()]` | Shades the header rows of table `name` in `color1` and every other body row in `color2` (defaults `navajowhite` and `antiquewhite`; same format as `shading()` in `putdocx table`). The number of header rows is read from the table unless `headerrow(#)` is given. Use `title` / `notes` when the first / last row holds a table title / notes that should not be shaded. |
| `checkpoint, name()` `[replace]` | Writes the current document to `name` and keeps it open for further editing. |
| `save, path() filestub()` `[securitylabel() rc]` | Saves the document as `path/filestubN.docx`, where `N` is the first number not yet used, and closes it. `securitylabel(file)` copies the label of a labelled file into the new one; `rc` shows the output of that step. For a fixed, unnumbered file name use `putdocx save` instead. |
| `label` *target* `[, from() saving() replace]` | Copies or reports a sensitivity label – the same as `docxlabel`, see below. |
| *anything else* | Passed unchanged to `putdocx`. |

Paths and file names in `path()`, `filestub()` and `name()` should be in double quotes (or compound double quotes).

## Sensitivity labels: `docx label` / `docxlabel`

In organisations that require sensitivity labels, Word opens a `putdocx` file without a label as read-only. `docx label` copies the label metadata from any file that already has the right label (for example a template saved from Word), so that Word opens the new file as labelled and editable.

The same command is available on its own as **`docxlabel`**, for `.docx` files that were not made with `docx`, and for `.xlsx` and `.pptx` files. `docx label …` and `docxlabel …` take the same syntax and store the same results.

```stata
* copy the label of a labelled template into a report (changes the report in place)
docx label "report.docx", from("H:/templates/labelled.docx")

* write the labelled copy to a new file instead
docx label "report.docx", from("H:/templates/labelled.docx") saving("report_labelled.docx") replace

* report the label of a file
docxlabel "report.docx"
```

| Option | Meaning |
|---|---|
| `from(source)` | Labelled file whose label is copied. Without `from()`, the command only reports the label of *target*. |
| `saving(newfile)` | Write the result to `newfile`; by default *target* is changed in place. |
| `replace` | Allow `saving()` to overwrite an existing file. |

Office stores a label in up to two places, and both are copied when present: the `MSIP_Label_*` custom document properties (`docProps/custom.xml`) and the `docMetadata/LabelInfo.xml` part with its relationship and content type. Any existing label in the target is removed first; nothing else in the file changes.

Limitations: only the label metadata is copied – visual markings that a label may add (header, footer, watermark) are not created; Word may add them the next time it saves the file. Files encrypted by a protecting label or a password cannot be read or changed. Close the target in Word before running the command.

## Stored results

| Subcommand | Results in `r()` |
|---|---|
| `docx save` | `r(fname)` – full name of the saved file |
| `docx tableshader` | `r(tablename)`; `r(title)`, `r(notes)` (`yes`/`no`); `r(header_rows)` (scalar) |
| `docx label`, `docxlabel` | scalars `r(found)` (1 if a label was found), `r(nprops)` (number of `MSIP_Label_*` properties), `r(labelinfo)` (1 if `LabelInfo.xml` is present); macros `r(name)` (label name, if stored), `r(labelid)`, `r(siteid)` (tenant id), `r(file)`, and when copying `r(source)`, `r(saved)` |

## Notes

- **TOC page numbers are set by hand.** `putdocx` cannot tell on which page a heading ends up (as of Stata 19), so give `page(#)` yourself or leave it out. The links from the TOC to the headings work either way. In Word you can also replace the table with a native Word TOC after saving.
- **`docx begin` discards any unsaved document** that is open in `putdocx`.
- **Automatic header-row detection** in `tableshader` uses the repeating-header-row flags that `putdocx table …, headerrow(#)` writes. For tables made without `headerrow()`, give `docx tableshader, … headerrow(#)`.
- `checkpoint` and the header-row detection use undocumented internals of Stata's `putdocx`; they may need updating in future Stata versions.
- Heading and title text goes inside `text()`; avoid double quotes within the text.

## Repository contents

| File | Purpose |
|---|---|
| `docx.ado`, `docx.sthlp` | The `docx` command and its help file |
| `docxlabel.ado`, `docxlabel.sthlp` | The `docxlabel` command (also `docx label`) and its help file |
| `docx.pkg`, `stata.toc` | Package description used by `net install` |

## Help

After installation, type `help docx` and `help docxlabel` in Stata.
