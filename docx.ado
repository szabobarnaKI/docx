//capture program drop docx
program define docx, rclass
	version 19.0
	if "`1'"!="" {
		local subcmd=subinstr("`1'",",","",.)
	    if inlist(strlower("`subcmd'"),"open","begin","start") {
			syntax anything, [*]
		    capture putdocx clear
			putdocx begin, footer(footer1) margin(all,0.5) `options'
			putdocx paragraph, tofooter(footer1) halign(center)
			putdocx pagenumber
		}
		else if inlist(strlower("`1'"),"maintitle,","maintitle") {
		    main_title `0'
		}
		else if strlower("`1'")=="date" {
			putdocx paragraph, halign(center) font("`Arial'", 16)
			putdocx text ("`c(current_date)'"), nformat(%tcCCYY-NN-DD)
		}
	    else if inlist(strlower("`1'"),"author,","author") {
		    main_author `0'
		}
		else if inlist(strlower("`1'"),"title,","title") {
		    paragraph_title `0'
		}
		else if inlist(strlower("`1'"),"subtitle,","subtitle") {
		    subparagraph_title `0'
		}
	    else if strlower("`1'")=="paragraph" {
			putdocx paragraph, halign(left) indent(left,0.5) font("`Arial'", 11)
		}
		else if inlist(strlower("`1'"),"emptylines,","emptylines") {
		    empty_lines `0'
		}
	    else if inlist(strlower("`1'"),"save,","save") {
		    save_docx `0'
			return add
		}
	    else if inlist(strlower("`1'"),"checkpoint,","checkpoint") {
		    save_checkpoint_docx `0'
		}
		else if inlist(strlower("`1'"),"starttoc,","starttoc") {
			syntax [anything], [nopage *]
			putdocx table toc=(1,18), layout(fixed) border(all,nil) cellmargin(bottom,0.125)
			putdocx table toc(.,.), font(Arial, 11)
			if !("`page'"=="nopage") putdocx table toc(1,17)=("Page"), colspan(2) font(Arial, 13) bold
			putdocx table toc(1,1)=("Contents"), colspan(16) font(Arial, 13) bold
			addtocline, contents( ) page(-1) indentation(0)
		}
	    else if inlist(strlower("`1'"),"tableshader,","tableshader") {
		    tbl_shader `0'
			return add
		}
		else putdocx `0'
	}
	else display "{txt}Error. Type of action not defined.{break}Define {it:open}, {it:maintitle}, {it:author}, {it:date}, {it:title}, {it:subtitle}, {it:paragraph}, {it:emptylines}, {it:starttoc} or {it:save}. You can also use docx with putdocx options."
end

program define main_title
	syntax anything, text(string asis) [newline]
	if "`newline'"!="newline" {
		putdocx paragraph, halign(center) font("Arial", 24)
	}
	else {
		putdocx text (""), linebreak(1)
	}
	putdocx text ("`text'"), italic
end

program define save_docx, rclass
	syntax anything, path(string) filestub(string) [SECuritylabel(string) rc]
	// stop logging the docx output, save as a numbered log-file
	// first, find the next number to save the log to
	local x=1
	while fileexists("`path'/`filestub'`x'.docx") {
		local x=`x'+1
	}
	// save the file
	putdocx save "`path'/`filestub'`x'.docx"
	// copy a security label from an example file, if requested
	if "`securitylabel'"!="" {
		`=cond("`rc'"=="","quietly","")' docxlabel "`path'/`filestub'`x'.docx", from(`"`securitylabel'"')
	}
	// return information
	return local fname="`path'/`filestub'`x'.docx"
end

program define save_checkpoint_docx
	syntax [anything], name(string) replace
	if "`replace'"!="replace" {
		confirm new file `"`name'"'
	}
	
	mata: fname = st_local("name")
	mata: doc_id = strtoreal(st_global("ST__DOCX_ID"))
	mata: rc = _docx_save(doc_id, fname, 1)
end

program define main_author
	syntax anything, text(string asis)
	putdocx paragraph, halign(center) font("Arial", 16)
	putdocx text ("`text'"), italic
end

program define tbl_shader, rclass
	syntax anything, tbl(string asis) [HEADerrow(int -1) Title Notes color1(string asis) color2(string asis)]

	if "`color1'"=="" local color1="navajowhite,black,clear"
	if "`color2'"=="" local color2="antiquewhite,black,clear"
	quietly: capture putdocx describe "`tbl'"
	if _rc!=0 { // `anything' is not an existing table name
		display "Docx table `tbl' is not found."
		error 197
	}
	else {
		local rowmax=r(nrows)
		if `headerrow'<0 {
			quietly save_checkpoint_docx, name("`c(tmpdir)'chkpoint.docx") replace
			quietly find_docx_tbl_header, fname("`c(tmpdir)'chkpoint.docx") tbl(`tbl')
			local headerrow=r(headerrowcount)
			erase "`c(tmpdir)'chkpoint.docx"
		}
		if "`notes'"=="notes" local rowmax=`rowmax'-1
		if "`title'"=="title" local start=1
			else local start=0
		local x1=`start'+`headerrow'+2 // start with white row after header
		forvalues i=`x1'(2)`rowmax' {
			putdocx table `tbl'(`i',.), shading(`color2')
		}
		local x1=`start'+1
		local x2=`start'+`headerrow'
		forvalues i=`x1'(1)`x2' {
			putdocx table `tbl'(`i',.), shading(`color1')
		}
	}
	return local tablename="`tbl'"
	if "`title'"== "title" return local title="yes"
		else return local title="no"
	if "`notes'"== "notes" return local notes="yes"
		else return local notes="no"
	return scalar header_rows=`headerrow'
end

program define paragraph_title
	syntax anything, text(string asis) [toc PAGEnumber(integer -1)]
	if "`toc'"!="" {
	    //paragraph title needs to be added to table of contents
		//if pagenumber==-1, no pagenumber shall be entered
		addtocline, contents(`text') page(`pagenumber') indentation(0)
	}
	putdocx paragraph, halign(left) font("Arial", 13)
	putdocx text ("`text'"), bold bookmark(toc_bm`=r(toc_row)')
end

program define subparagraph_title
	syntax anything, text(string asis) [toc PAGEnumber(integer -1) indentation(integer 1)]
	if "`toc'"!="" {
	    //paragraph title needs to be added to table of contents
		//if pagenumber==-1, no pagenumber shall be entered
		addtocline, contents(`text') page(`pagenumber') indentation(`indentation')
	}
	putdocx paragraph, halign(left) font("Arial", 11)
	putdocx text ("`text'"), bold bookmark(toc_bm`=r(toc_row)')
end

program define empty_lines
	syntax anything, lines(integer)
	putdocx text (""), linebreak(`lines')
end

program define addtocline, rclass
	syntax, contents(string asis) page(integer) indentation(integer)
	//if page==-1, no pagenumber shall be entered
	quietly putdocx describe toc
	local y=r(nrows)
	local y1=r(nrows)+1
	local x0=r(ncols)-1
	local x=`indentation'+1
	local x1=r(ncols)-2-`indentation'
	putdocx table toc(`y',.), addrows(1)
	if "`page'"!="-1" {
		putdocx table toc(`y1',`x0')=("`page'"), colspan(2)
	}
	putdocx table toc(`y1',`x')=("`contents'"), colspan(`x1') bookmarklink(toc_bm`y1')
	return scalar toc_row=`y1'
end

program define find_table_index, rclass
	syntax anything(name=tblname)
	
	mata: tbl = st_local("tblname")
	mata: doc_id = strtoreal(st_global("ST__DOCX_ID"))
	mata: tinfo = J(1,3,.)
	mata: sret = _docx_table_get_info(doc_id, tbl)
	mata: stinfo = tokens(sret, "|")
	mata: st_local("tbl_index", stinfo[1])
	return scalar tbl_index=`tbl_index'
end

program define find_docx_tbl_header, rclass
	syntax, fname(string) TBLname(string asis)
	quietly find_table_index `tblname'
	local tblnum=`r(tbl_index)'
	
	local cpwd=c(pwd)
	// unpack everything in a temporary directory
	quietly {
		cd "`c(tmpdir)'"
		capture docx_removedir "`c(tmpdir)'docx_file_unzip"
		mkdir "`c(tmpdir)'docx_file_unzip"
		
		cd "`c(tmpdir)'docx_file_unzip"
		unzipfile `"`fname'"', replace
	}

	// open the document and count header rows
	file open docfile using "word/document.xml", read text

	local continue=1
	local tbl_i=0
	local headerrowcount=0
	while `continue'==1 {
		file read docfile line
		if r(eof)==1 local continue=0
		if strpos(`"`line'"',"</w:tbl>")!=0 local tbl_i=`tbl_i'+1
		if `tbl_i'>`tblnum' local continue=0
		if strpos(`"`line'"',`"<w:tblHeader w:val="true"/>"')!=0 & `tbl_i'==`tblnum' local headerrowcount=`headerrowcount'+1
	}
 
	file close docfile
	cd `"`cpwd'"'
	quietly docx_removedir "`c(tmpdir)'docx_file_unzip"
	
	return scalar headerrowcount=`headerrowcount'
end

program define docx_removedir
	syntax anything(name=path)
	capture rmdir `path'
	if _rc!=0 {
		// remove all files in the directory
		local filelist: dir `path' files "*.*"
		foreach i of local filelist {
			local fname=`path/' + "/`i'"
			erase `"`fname'"'
		}
		// try to remove all subdirectories in the directory
		local dirlist: dir `path' dirs "*"
		foreach i of local dirlist {
			local dname=`path/' + "/`i'"
			docx_removedir `"`dname'"'
		}
		capture rmdir `path'
	}
end
