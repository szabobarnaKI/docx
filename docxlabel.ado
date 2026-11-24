*! version 1.0.2  24nov2026
*! docxlabel: copy a Microsoft Purview (MIP) sensitivity label from one Office
*! Open XML file to another, or report the label of a file. Pure Stata + built-in
*! Java (java.util.zip); no shell, no scripts, no Word.
program docxlabel, rclass
    version 16
    syntax anything(name=file id="file name"), [FROM(string) SAVing(string) REPLACE]

    gettoken file rest : file
    if `"`rest'"' != "" {
        di as err `"only one file name allowed; put quotes around names with spaces"'
        exit 198
    }
    _dl_abspath `"`file'"'
    local dl_tgt `"`dl_path'"'
    confirm file `"`dl_tgt'"'

    if `"`from'"' == "" {
        if `"`saving'"' != "" | "`replace'" != "" {
            di as err "saving() and replace may only be used with from()"
            exit 198
        }
        local dl_mode query
    }
    else {
        local dl_mode copy
        _dl_abspath `"`from'"'
        local dl_src `"`dl_path'"'
        confirm file `"`dl_src'"'
        if `"`saving'"' != "" {
            _dl_abspath `"`saving'"'
            local dest `"`dl_path'"'
            capture confirm new file `"`dest'"'
            if _rc & "`replace'" == "" {
                di as err `"file `dest' already exists; specify replace"'
                exit 602
            }
        }
        else local dest `"`dl_tgt'"'
        tempfile dl_out
    }

    local dl_rc 692
    local dl_msg "Java call did not run"
    capture noisily java: DocxLabelJ.run();
    if _rc {
        di as err "docxlabel: the Java part could not be called (rc " _rc ")."
        di as err "If the class was not found: load docxlabel.ado by autoloading (adopath),"
        di as err "not with run/do -- an ado-file has its own Java instance."
        exit _rc
    }
    if "`dl_rc'" != "0" {
        di as err `"docxlabel: `dl_msg'"'
        exit `dl_rc'
    }

    if "`dl_mode'" == "copy" {
        capture copy `"`dl_out'"' `"`dest'"', replace
        if _rc {
            di as err `"could not write `dest' (rc `=_rc'); is the file open in Word?"'
            exit 603
        }
        capture erase `"`dl_out'"'
    }

    if `"`dl_name'"' != "" local what `""`dl_name'""'
    else local what "(name not stored in file)"
    if "`dl_mode'" == "query" {
        if "`dl_found'" == "1" {
            di as txt `"Sensitivity label of `dl_tgt':"'
            di as txt "  name:   " as res `"`what'"'
            di as txt "  id:     " as res "`dl_id'"
            di as txt "  tenant: " as res "`dl_site'"
            di as txt "  stored: " as res "`dl_nprops' MSIP_Label_* custom properties" ///
                cond("`dl_labelinfo'" == "1", " + LabelInfo.xml", "")
        }
        else di as txt `"No sensitivity label found in `dl_tgt'"'
    }
    else {
        di as txt `"Sensitivity label "' as res `"`what'"' as txt " (id `dl_id')"
        di as txt `"  copied from `dl_src'"'
        di as txt `"  to          `dest'"'
    }

    return local file     `"`dl_tgt'"'
    if "`dl_mode'" == "copy" {
        return local source `"`dl_src'"'
        return local saved  `"`dest'"'
    }
    return local name     `"`dl_name'"'
    return local labelid  "`dl_id'"
    return local siteid   "`dl_site'"
    return scalar labelinfo = `dl_labelinfo'
    return scalar nprops    = `dl_nprops'
    return scalar found     = `dl_found'
end

program _dl_abspath
    args path
    local p = subinstr(`"`path'"', "\", "/", .)
    if !(substr(`"`p'"', 1, 1) == "/" | regexm(`"`p'"', "^[A-Za-z]:/")) {
        local p = subinstr(c(pwd), "\", "/", .) + "/" + `"`p'"'
    }
    c_local dl_path `"`p'"'
end

java:
class DocxLabelJ {
    static final String RT_CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
    static final String RT_LABEL = "http://schemas.microsoft.com/office/2020/02/relationships/classificationlabels";
    static final String CT_CUSTOM = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
    static final String CT_LABEL = "application/vnd.ms-office.classificationlabels+xml";
    static final String NS_CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
    static final String NS_VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
    static final java.util.regex.Pattern P_PROP = java.util.regex.Pattern.compile(
        "<property\\b[^>]*\\bname=\"MSIP_Label_[^\"]*\"[^>]*>.*?</property>", java.util.regex.Pattern.DOTALL);
    static final java.util.regex.Pattern P_PTAG = java.util.regex.Pattern.compile("<property\\b[^>]*>");
    static final java.util.regex.Pattern P_REL = java.util.regex.Pattern.compile("<Relationship\\b[^>]*>");
    static final java.util.regex.Pattern P_OVR = java.util.regex.Pattern.compile("<Override\\b[^>]*>");
    static final java.util.regex.Pattern P_LBL = java.util.regex.Pattern.compile("<(?:\\w+:)?label\\s[^>]*>");

    static int failRc = 0;
    static String lCustomPath;
    static String lPartPath;
    static String lRelType;
    static String lCt;
    static byte[] lLabelXml;
    static java.util.List<String> lProps;
    static String lName;
    static String lId;
    static String lSite;

    static Exception fail(int rc, String msg) {
        failRc = rc;
        return new Exception(msg);
    }

    static boolean found() { return !lProps.isEmpty() || lPartPath != null; }

    static String attr(String tag, String name) {
        java.util.regex.Matcher m = java.util.regex.Pattern.compile(
            "\\s" + name + "\\s*=\\s*\"([^\"]*)\"").matcher(tag);
        return m.find() ? m.group(1) : null;
    }

    static String norm(String target) {
        String t = target.replace('\\', '/');
        while (t.startsWith("/")) t = t.substring(1);
        if (t.startsWith("./")) t = t.substring(2);
        return t;
    }

    static byte[] entry(java.util.zip.ZipFile z, String name) throws java.io.IOException {
        java.util.zip.ZipEntry e = z.getEntry(name);
        if (e == null) {
            java.util.Enumeration<? extends java.util.zip.ZipEntry> en = z.entries();
            while (en.hasMoreElements()) {
                java.util.zip.ZipEntry x = en.nextElement();
                if (x.getName().equalsIgnoreCase(name)) { e = x; break; }
            }
        }
        if (e == null) return null;
        try (java.io.InputStream in = z.getInputStream(e)) { return in.readAllBytes(); }
    }

    static String str(byte[] b) {
        return b == null ? null : new String(b, java.nio.charset.StandardCharsets.UTF_8);
    }

    static byte[] bytes(String s) {
        return s.getBytes(java.nio.charset.StandardCharsets.UTF_8);
    }

    static String unesc(String s) {
        return s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", "\"")
                .replace("&apos;", "'").replace("&amp;", "&");
    }

    static String overrideCT(String ctx, String part) {
        java.util.regex.Matcher m = P_OVR.matcher(ctx);
        while (m.find()) {
            String pn = attr(m.group(), "PartName");
            if (pn != null && norm(pn).equalsIgnoreCase(part)) return attr(m.group(), "ContentType");
        }
        return null;
    }

    static String removeOverride(String ctx, String part) {
        java.util.regex.Matcher m = P_OVR.matcher(ctx);
        StringBuffer sb = new StringBuffer();
        while (m.find()) {
            String pn = attr(m.group(), "PartName");
            boolean drop = pn != null && norm(pn).equalsIgnoreCase(part);
            m.appendReplacement(sb, drop ? "" : java.util.regex.Matcher.quoteReplacement(m.group()));
        }
        m.appendTail(sb);
        return sb.toString();
    }

    static String removeLabelRels(String rels) {
        java.util.regex.Matcher m = P_REL.matcher(rels);
        StringBuffer sb = new StringBuffer();
        while (m.find()) {
            String t = attr(m.group(), "Type");
            boolean drop = t != null && t.endsWith("/classificationlabels");
            m.appendReplacement(sb, drop ? "" : java.util.regex.Matcher.quoteReplacement(m.group()));
        }
        m.appendTail(sb);
        return sb.toString();
    }

    static String insertBefore(String xml, String closing, String add) throws Exception {
        int i = xml.lastIndexOf(closing);
        if (i < 0) throw fail(610, "malformed package part: " + closing + " not found");
        return xml.substring(0, i) + add + xml.substring(i);
    }

    static String newRelId(String rels) {
        int k = 1;
        while (rels.contains("Id=\"rIdDL" + k + "\"")) k++;
        return "rIdDL" + k;
    }

    static void checkZip(String path) throws Exception {
        java.io.File f = new java.io.File(path);
        if (!f.isFile()) throw fail(601, "file " + path + " not found");
        byte[] h = new byte[4];
        int n;
        try (java.io.InputStream in = new java.io.FileInputStream(f)) { n = in.read(h); }
        if (n == 4 && (h[0] & 0xFF) == 0xD0 && (h[1] & 0xFF) == 0xCF && (h[2] & 0xFF) == 0x11 && (h[3] & 0xFF) == 0xE0)
            throw fail(610, path + " is encrypted (password, or a sensitivity label with protection); only unencrypted files can be read or changed");
        if (n < 2 || h[0] != 'P' || h[1] != 'K')
            throw fail(610, path + " is not a docx/OOXML (zip) file");
    }

    static void readLabel(java.util.zip.ZipFile z) throws Exception {
        lCustomPath = "docProps/custom.xml";
        lPartPath = null;
        lRelType = RT_LABEL;
        lCt = CT_LABEL;
        lLabelXml = null;
        lProps = new java.util.ArrayList<String>();
        lName = "";
        lId = "";
        lSite = "";
        String rels = str(entry(z, "_rels/.rels"));
        String ctx = str(entry(z, "[Content_Types].xml"));
        if (rels == null || ctx == null) throw fail(610, z.getName() + " is not a valid OOXML package (no _rels/.rels or [Content_Types].xml)");
        java.util.regex.Matcher m = P_REL.matcher(rels);
        while (m.find()) {
            String t = attr(m.group(), "Type");
            String tg = attr(m.group(), "Target");
            if (t == null || tg == null) continue;
            if (t.endsWith("/custom-properties")) lCustomPath = norm(tg);
            else if (t.endsWith("/classificationlabels") && lPartPath == null) { lRelType = t; lPartPath = norm(tg); }
        }
        if (lPartPath != null) {
            lLabelXml = entry(z, lPartPath);
            if (lLabelXml == null) lPartPath = null;
            else {
                String c = overrideCT(ctx, lPartPath);
                if (c != null) lCt = c;
                java.util.regex.Matcher lm = P_LBL.matcher(str(lLabelXml));
                while (lm.find()) {
                    String rem = attr(lm.group(), "removed");
                    if (rem != null && (rem.equals("1") || rem.equalsIgnoreCase("true"))) continue;
                    String id = attr(lm.group(), "id");
                    String site = attr(lm.group(), "siteId");
                    if (id != null) lId = id;
                    if (site != null) lSite = site;
                    break;
                }
            }
        }
        String cx = str(entry(z, lCustomPath));
        if (cx != null) {
            java.util.regex.Matcher pm = P_PROP.matcher(cx);
            while (pm.find()) lProps.add(pm.group());
        }
        java.util.Map<String, java.util.Map<String, String>> byGuid = new java.util.LinkedHashMap<String, java.util.Map<String, String>>();
        java.util.regex.Pattern pv = java.util.regex.Pattern.compile(">([^<]*)</[^>]*>\\s*</property>");
        for (String p : lProps) {
            String nm = attr(p.substring(0, p.indexOf('>') + 1), "name");
            if (nm == null || nm.length() < 49) continue;
            String guid = nm.substring(11, 47);
            String suffix = nm.substring(48);
            java.util.regex.Matcher vm = pv.matcher(p);
            String val = vm.find() ? unesc(vm.group(1)) : "";
            java.util.Map<String, String> g = byGuid.get(guid);
            if (g == null) { g = new java.util.HashMap<String, String>(); byGuid.put(guid, g); }
            g.put(suffix, val);
        }
        String pick = null;
        for (java.util.Map.Entry<String, java.util.Map<String, String>> e : byGuid.entrySet()) {
            if ("true".equalsIgnoreCase(e.getValue().getOrDefault("Enabled", ""))) { pick = e.getKey(); break; }
        }
        if (pick == null && !byGuid.isEmpty()) pick = byGuid.keySet().iterator().next();
        if (pick != null) {
            java.util.Map<String, String> v = byGuid.get(pick);
            lName = v.getOrDefault("Name", "");
            if (lId.isEmpty()) lId = pick;
            if (lSite.isEmpty()) lSite = v.getOrDefault("SiteId", "");
        }
        lId = lId.replace("{", "").replace("}", "").toLowerCase();
        lSite = lSite.replace("{", "").replace("}", "").toLowerCase();
    }

    static String addProps(String xml, java.util.List<String> props) throws Exception {
        int i = xml.indexOf("<Properties");
        if (i < 0) throw fail(610, "malformed custom properties part");
        if (xml.lastIndexOf("</Properties>") < 0) {
            int k = xml.indexOf("/>", i);
            if (k < 0) throw fail(610, "malformed custom properties part");
            xml = xml.substring(0, k) + "></Properties>" + xml.substring(k + 2);
        }
        if (!props.isEmpty()) {
            int j = xml.indexOf('>', i);
            if (!xml.substring(i, j).contains("xmlns:vt=")) {
                xml = xml.substring(0, j) + " xmlns:vt=\"" + NS_VT + "\"" + xml.substring(j);
            }
            xml = insertBefore(xml, "</Properties>", String.join("", props));
        }
        java.util.regex.Matcher m = P_PTAG.matcher(xml);
        StringBuffer sb = new StringBuffer();
        int pid = 2;
        while (m.find()) {
            String tag = m.group().replaceFirst("\\bpid\\s*=\\s*\"[0-9]*\"", "pid=\"" + pid + "\"");
            pid++;
            m.appendReplacement(sb, java.util.regex.Matcher.quoteReplacement(tag));
        }
        m.appendTail(sb);
        return sb.toString();
    }

    static void put(java.util.zip.ZipOutputStream zo, String name, long time, byte[] data) throws java.io.IOException {
        java.util.zip.ZipEntry ne = new java.util.zip.ZipEntry(name);
        if (time > 0) ne.setTime(time);
        zo.putNextEntry(ne);
        zo.write(data);
        zo.closeEntry();
    }

    static void copy(String srcPath, String tgtPath, String outPath) throws Exception {
        checkZip(srcPath);
        checkZip(tgtPath);
        try (java.util.zip.ZipFile sz = new java.util.zip.ZipFile(srcPath)) { readLabel(sz); }
        if (!found()) throw fail(459, "no sensitivity label found in " + srcPath);
        try (java.util.zip.ZipFile t = new java.util.zip.ZipFile(tgtPath)) {
            String rels = str(entry(t, "_rels/.rels"));
            String ctx = str(entry(t, "[Content_Types].xml"));
            if (rels == null || ctx == null) throw fail(610, tgtPath + " is not a valid OOXML package (no _rels/.rels or [Content_Types].xml)");
            String tCustom = null;
            java.util.Set<String> oldLabel = new java.util.HashSet<String>();
            java.util.regex.Matcher m = P_REL.matcher(rels);
            while (m.find()) {
                String ty = attr(m.group(), "Type");
                String tg = attr(m.group(), "Target");
                if (ty == null || tg == null) continue;
                if (ty.endsWith("/custom-properties")) tCustom = norm(tg);
                else if (ty.endsWith("/classificationlabels")) oldLabel.add(norm(tg).toLowerCase());
            }
            rels = removeLabelRels(rels);
            for (String o : oldLabel) ctx = removeOverride(ctx, o);
            String newLabel = null;
            if (lPartPath != null) {
                newLabel = lPartPath;
                rels = insertBefore(rels, "</Relationships>", "<Relationship Id=\"" + newRelId(rels)
                    + "\" Type=\"" + lRelType + "\" Target=\"" + newLabel + "\"/>");
                ctx = removeOverride(ctx, newLabel);
                ctx = insertBefore(ctx, "</Types>", "<Override PartName=\"/" + newLabel + "\" ContentType=\"" + lCt + "\"/>");
            }
            String customPath = tCustom != null ? tCustom : "docProps/custom.xml";
            String tcx = str(entry(t, customPath));
            String newCustom = null;
            if (tcx != null) {
                newCustom = addProps(P_PROP.matcher(tcx).replaceAll(""), lProps);
            } else if (!lProps.isEmpty()) {
                newCustom = addProps("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Properties xmlns=\""
                    + NS_CUSTOM + "\" xmlns:vt=\"" + NS_VT + "\"></Properties>", lProps);
            }
            if (newCustom != null) {
                if (tCustom == null) {
                    rels = insertBefore(rels, "</Relationships>", "<Relationship Id=\"" + newRelId(rels)
                        + "\" Type=\"" + RT_CUSTOM + "\" Target=\"" + customPath + "\"/>");
                }
                if (overrideCT(ctx, customPath) == null) {
                    ctx = insertBefore(ctx, "</Types>", "<Override PartName=\"/" + customPath + "\" ContentType=\"" + CT_CUSTOM + "\"/>");
                }
            }
            java.util.Set<String> done = new java.util.HashSet<String>();
            try (java.util.zip.ZipOutputStream zo = new java.util.zip.ZipOutputStream(
                    new java.io.BufferedOutputStream(new java.io.FileOutputStream(outPath)))) {
                java.util.Enumeration<? extends java.util.zip.ZipEntry> en = t.entries();
                while (en.hasMoreElements()) {
                    java.util.zip.ZipEntry e = en.nextElement();
                    String n = e.getName();
                    String nl = norm(n).toLowerCase();
                    if (done.contains(nl)) continue;
                    byte[] data;
                    if (nl.equals("_rels/.rels")) data = bytes(rels);
                    else if (nl.equals("[content_types].xml")) data = bytes(ctx);
                    else if (newCustom != null && nl.equals(customPath.toLowerCase())) data = bytes(newCustom);
                    else if (newLabel != null && nl.equals(newLabel.toLowerCase())) data = lLabelXml;
                    else if (oldLabel.contains(nl)) continue;
                    else if (e.isDirectory()) continue;
                    else { try (java.io.InputStream in = t.getInputStream(e)) { data = in.readAllBytes(); } }
                    put(zo, n, e.getTime(), data);
                    done.add(nl);
                }
                if (newCustom != null && !done.contains(customPath.toLowerCase())) put(zo, customPath, -1, bytes(newCustom));
                if (newLabel != null && !done.contains(newLabel.toLowerCase())) put(zo, newLabel, -1, lLabelXml);
            }
        }
    }

    static void query(String path) throws Exception {
        checkZip(path);
        try (java.util.zip.ZipFile z = new java.util.zip.ZipFile(path)) { readLabel(z); }
    }

    public static void run() {
        String rc = "0";
        String msg = "";
        boolean ok = false;
        failRc = 0;
        try {
            String mode = com.stata.sfi.Macro.getLocal("dl_mode");
            if (mode.equals("query")) query(com.stata.sfi.Macro.getLocal("dl_tgt"));
            else copy(com.stata.sfi.Macro.getLocal("dl_src"), com.stata.sfi.Macro.getLocal("dl_tgt"),
                      com.stata.sfi.Macro.getLocal("dl_out"));
            ok = true;
        } catch (Throwable e) {
            rc = failRc != 0 ? String.valueOf(failRc) : "692";
            msg = failRc != 0 ? e.getMessage() : e.toString();
        }
        com.stata.sfi.Macro.setLocal("dl_rc", rc);
        com.stata.sfi.Macro.setLocal("dl_msg", msg);
        if (ok) {
            com.stata.sfi.Macro.setLocal("dl_found", found() ? "1" : "0");
            com.stata.sfi.Macro.setLocal("dl_name", lName);
            com.stata.sfi.Macro.setLocal("dl_id", lId);
            com.stata.sfi.Macro.setLocal("dl_site", lSite);
            com.stata.sfi.Macro.setLocal("dl_nprops", String.valueOf(lProps.size()));
            com.stata.sfi.Macro.setLocal("dl_labelinfo", lPartPath != null ? "1" : "0");
        }
    }
}
end
