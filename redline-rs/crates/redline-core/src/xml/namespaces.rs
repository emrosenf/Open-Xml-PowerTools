#![allow(non_snake_case)]

use super::xname::XName;

/// XML Namespace namespace (for xmlns declarations)
pub mod XMLNS {
    pub const NS: &str = "http://www.w3.org/2000/xmlns/";
}

pub mod W {
    use super::XName;
    pub const NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    
    pub fn p() -> XName { XName::new(NS, "p") }
    pub fn r() -> XName { XName::new(NS, "r") }
    pub fn t() -> XName { XName::new(NS, "t") }
    pub fn rPr() -> XName { XName::new(NS, "rPr") }
    pub fn pPr() -> XName { XName::new(NS, "pPr") }
    pub fn body() -> XName { XName::new(NS, "body") }
    pub fn document() -> XName { XName::new(NS, "document") }
    pub fn ins() -> XName { XName::new(NS, "ins") }
    pub fn del() -> XName { XName::new(NS, "del") }
    pub fn delText() -> XName { XName::new(NS, "delText") }
    pub fn delInstrText() -> XName { XName::new(NS, "delInstrText") }
    pub fn moveFrom() -> XName { XName::new(NS, "moveFrom") }
    pub fn moveTo() -> XName { XName::new(NS, "moveTo") }
    pub fn rPrChange() -> XName { XName::new(NS, "rPrChange") }
    pub fn pPrChange() -> XName { XName::new(NS, "pPrChange") }
    pub fn sectPrChange() -> XName { XName::new(NS, "sectPrChange") }
    pub fn tblPrChange() -> XName { XName::new(NS, "tblPrChange") }
    pub fn tblGridChange() -> XName { XName::new(NS, "tblGridChange") }
    pub fn tcPrChange() -> XName { XName::new(NS, "tcPrChange") }
    pub fn trPrChange() -> XName { XName::new(NS, "trPrChange") }
    pub fn tblPrExChange() -> XName { XName::new(NS, "tblPrExChange") }
    pub fn numberingChange() -> XName { XName::new(NS, "numberingChange") }
    pub fn cellIns() -> XName { XName::new(NS, "cellIns") }
    pub fn cellDel() -> XName { XName::new(NS, "cellDel") }
    pub fn cellMerge() -> XName { XName::new(NS, "cellMerge") }
    pub fn customXmlInsRangeStart() -> XName { XName::new(NS, "customXmlInsRangeStart") }
    pub fn customXmlInsRangeEnd() -> XName { XName::new(NS, "customXmlInsRangeEnd") }
    pub fn customXmlDelRangeStart() -> XName { XName::new(NS, "customXmlDelRangeStart") }
    pub fn customXmlDelRangeEnd() -> XName { XName::new(NS, "customXmlDelRangeEnd") }
    pub fn customXmlMoveFromRangeStart() -> XName { XName::new(NS, "customXmlMoveFromRangeStart") }
    pub fn customXmlMoveFromRangeEnd() -> XName { XName::new(NS, "customXmlMoveFromRangeEnd") }
    pub fn customXmlMoveToRangeStart() -> XName { XName::new(NS, "customXmlMoveToRangeStart") }
    pub fn customXmlMoveToRangeEnd() -> XName { XName::new(NS, "customXmlMoveToRangeEnd") }
    pub fn moveFromRangeStart() -> XName { XName::new(NS, "moveFromRangeStart") }
    pub fn moveFromRangeEnd() -> XName { XName::new(NS, "moveFromRangeEnd") }
    pub fn moveToRangeStart() -> XName { XName::new(NS, "moveToRangeStart") }
    pub fn moveToRangeEnd() -> XName { XName::new(NS, "moveToRangeEnd") }
    pub fn tbl() -> XName { XName::new(NS, "tbl") }
    pub fn tr() -> XName { XName::new(NS, "tr") }
    pub fn trPr() -> XName { XName::new(NS, "trPr") }
    pub fn tc() -> XName { XName::new(NS, "tc") }
    pub fn bookmarkStart() -> XName { XName::new(NS, "bookmarkStart") }
    pub fn bookmarkEnd() -> XName { XName::new(NS, "bookmarkEnd") }
    pub fn bookmark_start() -> XName { XName::new(NS, "bookmarkStart") }
    pub fn bookmark_end() -> XName { XName::new(NS, "bookmarkEnd") }
    pub fn name() -> XName { XName::new(NS, "name") }
    pub fn sectPr() -> XName { XName::new(NS, "sectPr") }
    pub fn sdt() -> XName { XName::new(NS, "sdt") }
    pub fn sdtContent() -> XName { XName::new(NS, "sdtContent") }
    pub fn hyperlink() -> XName { XName::new(NS, "hyperlink") }
    pub fn fld() -> XName { XName::new(NS, "fld") }
    pub fn footnoteReference() -> XName { XName::new(NS, "footnoteReference") }
    pub fn endnoteReference() -> XName { XName::new(NS, "endnoteReference") }
    pub fn footnotes() -> XName { XName::new(NS, "footnotes") }
    pub fn footnote() -> XName { XName::new(NS, "footnote") }
    pub fn endnotes() -> XName { XName::new(NS, "endnotes") }
    pub fn endnote() -> XName { XName::new(NS, "endnote") }
    pub fn txbxContent() -> XName { XName::new(NS, "txbxContent") }
    pub fn drawing() -> XName { XName::new(NS, "drawing") }
    pub fn pict() -> XName { XName::new(NS, "pict") }
    pub fn br() -> XName { XName::new(NS, "br") }
    pub fn tab() -> XName { XName::new(NS, "tab") }
    pub fn author() -> XName { XName::new(NS, "author") }
    pub fn id() -> XName { XName::new(NS, "id") }
    pub fn date() -> XName { XName::new(NS, "date") }
    pub fn tblPr() -> XName { XName::new(NS, "tblPr") }
    pub fn tblGrid() -> XName { XName::new(NS, "tblGrid") }
    pub fn tblPrEx() -> XName { XName::new(NS, "tblPrEx") }
    pub fn tcPr() -> XName { XName::new(NS, "tcPr") }
    pub fn cr() -> XName { XName::new(NS, "cr") }
    pub fn dayLong() -> XName { XName::new(NS, "dayLong") }
    pub fn dayShort() -> XName { XName::new(NS, "dayShort") }
    pub fn monthLong() -> XName { XName::new(NS, "monthLong") }
    pub fn monthShort() -> XName { XName::new(NS, "monthShort") }
    pub fn noBreakHyphen() -> XName { XName::new(NS, "noBreakHyphen") }
    pub fn pgNum() -> XName { XName::new(NS, "pgNum") }
    pub fn ptab() -> XName { XName::new(NS, "ptab") }
    pub fn softHyphen() -> XName { XName::new(NS, "softHyphen") }
    pub fn sym() -> XName { XName::new(NS, "sym") }
    pub fn yearLong() -> XName { XName::new(NS, "yearLong") }
    pub fn yearShort() -> XName { XName::new(NS, "yearShort") }
    pub fn fldChar() -> XName { XName::new(NS, "fldChar") }
    pub fn instrText() -> XName { XName::new(NS, "instrText") }
    pub fn fldSimple() -> XName { XName::new(NS, "fldSimple") }
    pub fn object() -> XName { XName::new(NS, "object") }
    pub fn commentRangeStart() -> XName { XName::new(NS, "commentRangeStart") }
    pub fn commentRangeEnd() -> XName { XName::new(NS, "commentRangeEnd") }
    pub fn lastRenderedPageBreak() -> XName { XName::new(NS, "lastRenderedPageBreak") }
    pub fn proofErr() -> XName { XName::new(NS, "proofErr") }
    pub fn permEnd() -> XName { XName::new(NS, "permEnd") }
    pub fn permStart() -> XName { XName::new(NS, "permStart") }
    pub fn perm_end() -> XName { XName::new(NS, "permEnd") }
    pub fn perm_start() -> XName { XName::new(NS, "permStart") }
    pub fn proof_err() -> XName { XName::new(NS, "proofErr") }
    pub fn no_proof() -> XName { XName::new(NS, "noProof") }
    pub fn soft_hyphen() -> XName { XName::new(NS, "softHyphen") }
    pub fn last_rendered_page_break() -> XName { XName::new(NS, "lastRenderedPageBreak") }
    pub fn comment_range_start() -> XName { XName::new(NS, "commentRangeStart") }
    pub fn comment_range_end() -> XName { XName::new(NS, "commentRangeEnd") }
    pub fn comment_reference() -> XName { XName::new(NS, "commentReference") }
    pub fn annotation_ref() -> XName { XName::new(NS, "annotationRef") }
    pub fn endnote_reference() -> XName { XName::new(NS, "endnoteReference") }
    pub fn footnote_reference() -> XName { XName::new(NS, "footnoteReference") }
    pub fn fld_data() -> XName { XName::new(NS, "fldData") }
    pub fn fld_char() -> XName { XName::new(NS, "fldChar") }
    pub fn instr_text() -> XName { XName::new(NS, "instrText") }
    pub fn fld_simple() -> XName { XName::new(NS, "fldSimple") }
    pub fn r_style() -> XName { XName::new(NS, "rStyle") }
    pub fn p_style() -> XName { XName::new(NS, "pStyle") }
    pub fn web_hidden() -> XName { XName::new(NS, "webHidden") }
    pub fn r_pr() -> XName { XName::new(NS, "rPr") }
    pub fn p_pr() -> XName { XName::new(NS, "pPr") }
    pub fn w() -> XName { XName::new("http://www.w3.org/2000/xmlns/", "w") }
    pub fn footnoteRef() -> XName { XName::new(NS, "footnoteRef") }
    pub fn endnoteRef() -> XName { XName::new(NS, "endnoteRef") }
    pub fn separator() -> XName { XName::new(NS, "separator") }
    pub fn continuationSeparator() -> XName { XName::new(NS, "continuationSeparator") }
    pub fn sdtPr() -> XName { XName::new(NS, "sdtPr") }
    pub fn sdtEndPr() -> XName { XName::new(NS, "sdtEndPr") }
    pub fn smartTag() -> XName { XName::new(NS, "smartTag") }
    pub fn smartTagPr() -> XName { XName::new(NS, "smartTagPr") }
    pub fn smart_tag() -> XName { XName::new(NS, "smartTag") }
    pub fn sdt_content() -> XName { XName::new(NS, "sdtContent") }
    pub fn ruby() -> XName { XName::new(NS, "ruby") }
    pub fn rubyPr() -> XName { XName::new(NS, "rubyPr") }
    pub fn gridSpan() -> XName { XName::new(NS, "gridSpan") }
    pub fn vMerge() -> XName { XName::new(NS, "vMerge") }
    pub fn val() -> XName { XName::new(NS, "val") }
    // rsid attributes (revision session IDs) - stripped during hashing
    pub fn rsid() -> XName { XName::new(NS, "rsid") }
    pub fn rsids() -> XName { XName::new(NS, "rsids") }
    pub fn rsidDel() -> XName { XName::new(NS, "rsidDel") }
    pub fn rsidP() -> XName { XName::new(NS, "rsidP") }
    pub fn rsidR() -> XName { XName::new(NS, "rsidR") }
    pub fn rsidRDefault() -> XName { XName::new(NS, "rsidRDefault") }
    pub fn rsidRPr() -> XName { XName::new(NS, "rsidRPr") }
    pub fn rsidSect() -> XName { XName::new(NS, "rsidSect") }
    pub fn rsidTr() -> XName { XName::new(NS, "rsidTr") }
    pub fn rsid_del() -> XName { XName::new(NS, "rsidDel") }
    pub fn rsid_p() -> XName { XName::new(NS, "rsidP") }
    pub fn rsid_r() -> XName { XName::new(NS, "rsidR") }
    pub fn rsid_r_default() -> XName { XName::new(NS, "rsidRDefault") }
    pub fn rsid_r_pr() -> XName { XName::new(NS, "rsidRPr") }
    pub fn rsid_sect() -> XName { XName::new(NS, "rsidSect") }
    pub fn rsid_tr() -> XName { XName::new(NS, "rsidTr") }
    pub fn bdo() -> XName { XName::new(NS, "bdo") }
    pub fn customXml() -> XName { XName::new(NS, "customXml") }
    pub fn dir() -> XName { XName::new(NS, "dir") }
    pub fn alignment() -> XName { XName::new(NS, "alignment") }
    pub fn relativeTo() -> XName { XName::new(NS, "relativeTo") }
    pub fn leader() -> XName { XName::new(NS, "leader") }
    pub fn type_() -> XName { XName::new(NS, "type") }

    // Visual formatting elements (for visual redline feature)
    pub fn color() -> XName { XName::new(NS, "color") }
    pub fn u() -> XName { XName::new(NS, "u") }
    pub fn strike() -> XName { XName::new(NS, "strike") }
    pub fn b() -> XName { XName::new(NS, "b") }
    pub fn sz() -> XName { XName::new(NS, "sz") }
    pub fn szCs() -> XName { XName::new(NS, "szCs") }
    pub fn rFonts() -> XName { XName::new(NS, "rFonts") }
    pub fn ascii() -> XName { XName::new(NS, "ascii") }
    pub fn hAnsi() -> XName { XName::new(NS, "hAnsi") }
    pub fn cs() -> XName { XName::new(NS, "cs") }

    // Table formatting elements (for summary table)
    pub fn tblW() -> XName { XName::new(NS, "tblW") }
    pub fn tblBorders() -> XName { XName::new(NS, "tblBorders") }
    pub fn top() -> XName { XName::new(NS, "top") }
    pub fn left() -> XName { XName::new(NS, "left") }
    pub fn bottom() -> XName { XName::new(NS, "bottom") }
    pub fn right() -> XName { XName::new(NS, "right") }
    pub fn insideH() -> XName { XName::new(NS, "insideH") }
    pub fn insideV() -> XName { XName::new(NS, "insideV") }
    pub fn gridCol() -> XName { XName::new(NS, "gridCol") }
    pub fn shd() -> XName { XName::new(NS, "shd") }
    pub fn jc() -> XName { XName::new(NS, "jc") }
    pub fn spacing() -> XName { XName::new(NS, "spacing") }
    pub fn tblCellMar() -> XName { XName::new(NS, "tblCellMar") }
    pub fn w_val() -> XName { XName::new(NS, "w") }
    pub fn fill() -> XName { XName::new(NS, "fill") }
    pub fn space() -> XName { XName::new(NS, "space") }
    pub fn before() -> XName { XName::new(NS, "before") }
    pub fn after() -> XName { XName::new(NS, "after") }
    pub fn line() -> XName { XName::new(NS, "line") }
    pub fn lineRule() -> XName { XName::new(NS, "lineRule") }
    pub fn vAlign() -> XName { XName::new(NS, "vAlign") }
}

pub mod S {
    use super::XName;
    pub const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    
    pub fn worksheet() -> XName { XName::new(NS, "worksheet") }
    pub fn sheetData() -> XName { XName::new(NS, "sheetData") }
    pub fn row() -> XName { XName::new(NS, "row") }
    pub fn c() -> XName { XName::new(NS, "c") }
    pub fn v() -> XName { XName::new(NS, "v") }
    pub fn f() -> XName { XName::new(NS, "f") }
    pub fn is() -> XName { XName::new(NS, "is") }
    pub fn t() -> XName { XName::new(NS, "t") }
    pub fn sheets() -> XName { XName::new(NS, "sheets") }
    pub fn sheet() -> XName { XName::new(NS, "sheet") }
    pub fn definedNames() -> XName { XName::new(NS, "definedNames") }
    pub fn definedName() -> XName { XName::new(NS, "definedName") }
    pub fn si() -> XName { XName::new(NS, "si") }
    pub fn numFmts() -> XName { XName::new(NS, "numFmts") }
    pub fn numFmt() -> XName { XName::new(NS, "numFmt") }
    pub fn fonts() -> XName { XName::new(NS, "fonts") }
    pub fn font() -> XName { XName::new(NS, "font") }
    pub fn fills() -> XName { XName::new(NS, "fills") }
    pub fn fill() -> XName { XName::new(NS, "fill") }
    pub fn borders() -> XName { XName::new(NS, "borders") }
    pub fn border() -> XName { XName::new(NS, "border") }
    pub fn cellXfs() -> XName { XName::new(NS, "cellXfs") }
    pub fn xf() -> XName { XName::new(NS, "xf") }
    pub fn authors() -> XName { XName::new(NS, "authors") }
    pub fn author() -> XName { XName::new(NS, "author") }
    pub fn commentList() -> XName { XName::new(NS, "commentList") }
    pub fn comment() -> XName { XName::new(NS, "comment") }
    pub fn text() -> XName { XName::new(NS, "text") }
    pub fn dataValidations() -> XName { XName::new(NS, "dataValidations") }
    pub fn dataValidation() -> XName { XName::new(NS, "dataValidation") }
    pub fn formula1() -> XName { XName::new(NS, "formula1") }
    pub fn formula2() -> XName { XName::new(NS, "formula2") }
    pub fn mergeCells() -> XName { XName::new(NS, "mergeCells") }
    pub fn mergeCell() -> XName { XName::new(NS, "mergeCell") }
    pub fn hyperlinks() -> XName { XName::new(NS, "hyperlinks") }
    pub fn hyperlink() -> XName { XName::new(NS, "hyperlink") }
    // Style elements for markup
    pub fn styleSheet() -> XName { XName::new(NS, "styleSheet") }
    pub fn patternFill() -> XName { XName::new(NS, "patternFill") }
    pub fn fgColor() -> XName { XName::new(NS, "fgColor") }
    pub fn bgColor() -> XName { XName::new(NS, "bgColor") }
    pub fn cellStyleXfs() -> XName { XName::new(NS, "cellStyleXfs") }
    // Comments elements
    pub fn comments() -> XName { XName::new(NS, "comments") }
    pub fn r() -> XName { XName::new(NS, "r") }
    pub fn legacyDrawing() -> XName { XName::new(NS, "legacyDrawing") }
    // Workbook elements
    pub fn workbook() -> XName { XName::new(NS, "workbook") }
}

pub mod P {
    use super::XName;
    pub const NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    
    pub fn presentation() -> XName { XName::new(NS, "presentation") }
    pub fn sld() -> XName { XName::new(NS, "sld") }
    pub fn sld_sz() -> XName { XName::new(NS, "sldSz") }
    pub fn sld_id_lst() -> XName { XName::new(NS, "sldIdLst") }
    pub fn sld_id() -> XName { XName::new(NS, "sldId") }
    pub fn c_sld() -> XName { XName::new(NS, "cSld") }
    pub fn cSld() -> XName { XName::new(NS, "cSld") }
    pub fn bg() -> XName { XName::new(NS, "bg") }
    pub fn sp_tree() -> XName { XName::new(NS, "spTree") }
    pub fn spTree() -> XName { XName::new(NS, "spTree") }
    pub fn sp() -> XName { XName::new(NS, "sp") }
    pub fn pic() -> XName { XName::new(NS, "pic") }
    pub fn graphic_frame() -> XName { XName::new(NS, "graphicFrame") }
    pub fn grp_sp() -> XName { XName::new(NS, "grpSp") }
    pub fn cxn_sp() -> XName { XName::new(NS, "cxnSp") }
    pub fn nv_sp_pr() -> XName { XName::new(NS, "nvSpPr") }
    pub fn nv_pic_pr() -> XName { XName::new(NS, "nvPicPr") }
    pub fn nv_graphic_frame_pr() -> XName { XName::new(NS, "nvGraphicFramePr") }
    pub fn nv_grp_sp_pr() -> XName { XName::new(NS, "nvGrpSpPr") }
    pub fn nv_cxn_sp_pr() -> XName { XName::new(NS, "nvCxnSpPr") }
    pub fn c_nv_pr() -> XName { XName::new(NS, "cNvPr") }
    pub fn nv_pr() -> XName { XName::new(NS, "nvPr") }
    pub fn ph() -> XName { XName::new(NS, "ph") }
    pub fn sp_pr() -> XName { XName::new(NS, "spPr") }
    pub fn grp_sp_pr() -> XName { XName::new(NS, "grpSpPr") }
    pub fn tx_body() -> XName { XName::new(NS, "txBody") }
    pub fn txBody() -> XName { XName::new(NS, "txBody") }
    pub fn blip_fill() -> XName { XName::new(NS, "blipFill") }
    // Non-visual properties for markup
    pub fn c_nv_sp_pr() -> XName { XName::new(NS, "cNvSpPr") }
    pub fn notes() -> XName { XName::new(NS, "notes") }
    pub fn sp_locks() -> XName { XName::new(NS, "spLocks") }
}

pub mod A {
    use super::XName;
    pub const NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    
    pub fn p() -> XName { XName::new(NS, "p") }
    pub fn r() -> XName { XName::new(NS, "r") }
    pub fn t() -> XName { XName::new(NS, "t") }
    pub fn off() -> XName { XName::new(NS, "off") }
    pub fn ext() -> XName { XName::new(NS, "ext") }
    pub fn xfrm() -> XName { XName::new(NS, "xfrm") }
    pub fn p_pr() -> XName { XName::new(NS, "pPr") }
    pub fn bu_char() -> XName { XName::new(NS, "buChar") }
    pub fn bu_auto_num() -> XName { XName::new(NS, "buAutoNum") }
    pub fn r_pr() -> XName { XName::new(NS, "rPr") }
    pub fn fld() -> XName { XName::new(NS, "fld") }
    pub fn latin() -> XName { XName::new(NS, "latin") }
    pub fn solid_fill() -> XName { XName::new(NS, "solidFill") }
    pub fn srgb_clr() -> XName { XName::new(NS, "srgbClr") }
    pub fn blip() -> XName { XName::new(NS, "blip") }
    pub fn prst_geom() -> XName { XName::new(NS, "prstGeom") }
    pub fn cust_geom() -> XName { XName::new(NS, "custGeom") }
    pub fn graphic() -> XName { XName::new(NS, "graphic") }
    pub fn graphic_data() -> XName { XName::new(NS, "graphicData") }
    pub fn tbl() -> XName { XName::new(NS, "tbl") }
    pub fn tr() -> XName { XName::new(NS, "tr") }
    pub fn tc() -> XName { XName::new(NS, "tc") }
    pub fn tx_body() -> XName { XName::new(NS, "txBody") }
    // Line/outline elements for markup
    pub fn ln() -> XName { XName::new(NS, "ln") }
    pub fn no_fill() -> XName { XName::new(NS, "noFill") }
    pub fn body_pr() -> XName { XName::new(NS, "bodyPr") }
    pub fn lst_style() -> XName { XName::new(NS, "lstStyle") }
    pub fn avLst() -> XName { XName::new(NS, "avLst") }
}

pub mod R {
    use super::XName;
    pub const NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    
    pub fn id() -> XName { XName::new(NS, "id") }
    pub fn embed() -> XName { XName::new(NS, "embed") }
    pub fn r() -> XName { XName::new("http://www.w3.org/2000/xmlns/", "r") }
}

pub mod MC {
    use super::XName;
    pub const NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    
    pub fn AlternateContent() -> XName { XName::new(NS, "AlternateContent") }
    pub fn Choice() -> XName { XName::new(NS, "Choice") }
    pub fn Fallback() -> XName { XName::new(NS, "Fallback") }
    pub fn mc() -> XName { XName::new("http://www.w3.org/2000/xmlns/", "mc") }
    pub fn ignorable() -> XName { XName::new(NS, "Ignorable") }
}

pub mod CP {
    use super::XName;
    pub const NS: &str = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    
    pub fn revision() -> XName { XName::new(NS, "revision") }
    pub fn lastModifiedBy() -> XName { XName::new(NS, "lastModifiedBy") }
}

pub mod DC {
    use super::XName;
    pub const NS: &str = "http://purl.org/dc/elements/1.1/";
    
    pub fn creator() -> XName { XName::new(NS, "creator") }
}

pub mod PT {
    use super::XName;
    pub const NS: &str = "http://powertools.codeplex.com/2011";
    
    pub fn Unid() -> XName { XName::new(NS, "Unid") }
    pub fn SHA1Hash() -> XName { XName::new(NS, "SHA1Hash") }
    pub fn CorrelatedSHA1Hash() -> XName { XName::new(NS, "CorrelatedSHA1Hash") }
    pub fn StructureSHA1Hash() -> XName { XName::new(NS, "StructureSHA1Hash") }
}

pub mod M {
    use super::XName;
    pub const NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    
    pub fn f() -> XName { XName::new(NS, "f") }
    pub fn m() -> XName { XName::new("http://www.w3.org/2000/xmlns/", "m") }
    pub fn fPr() -> XName { XName::new(NS, "fPr") }
    pub fn ctrlPr() -> XName { XName::new(NS, "ctrlPr") }
    pub fn oMath() -> XName { XName::new(NS, "oMath") }
    pub fn oMathPara() -> XName { XName::new(NS, "oMathPara") }
    pub fn t() -> XName { XName::new(NS, "t") }
}

pub mod V {
    use super::XName;
    pub const NS: &str = "urn:schemas-microsoft-com:vml";
    
    pub fn textbox() -> XName { XName::new(NS, "textbox") }
    pub fn imagedata() -> XName { XName::new(NS, "imagedata") }
    pub fn group() -> XName { XName::new(NS, "group") }
    pub fn shape() -> XName { XName::new(NS, "shape") }
    pub fn rect() -> XName { XName::new(NS, "rect") }
    pub fn shapetype() -> XName { XName::new(NS, "shapetype") }
    pub fn fill() -> XName { XName::new(NS, "fill") }
    pub fn stroke() -> XName { XName::new(NS, "stroke") }
    pub fn shadow() -> XName { XName::new(NS, "shadow") }
    pub fn path() -> XName { XName::new(NS, "path") }
    pub fn formulas() -> XName { XName::new(NS, "formulas") }
    pub fn handles() -> XName { XName::new(NS, "handles") }
    pub fn textpath() -> XName { XName::new(NS, "textpath") }
}

pub mod WP {
    use super::XName;
    pub const NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    
    pub fn extent() -> XName { XName::new(NS, "extent") }
    pub fn docPr() -> XName { XName::new(NS, "docPr") }
    pub fn wp() -> XName { XName::new("http://www.w3.org/2000/xmlns/", "wp") }
}

pub mod W14 {
    use super::XName;
    pub const NS: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
    
    pub fn w14() -> XName { XName::new("http://www.w3.org/2000/xmlns/", "w14") }
    
    pub fn paraId() -> XName { XName::new(NS, "paraId") }
    pub fn textId() -> XName { XName::new(NS, "textId") }
}

pub mod O {
    use super::XName;
    pub const NS: &str = "urn:schemas-microsoft-com:office:office";
    
    pub fn relid() -> XName { XName::new(NS, "relid") }
    pub fn lock() -> XName { XName::new(NS, "lock") }
    pub fn extrusion() -> XName { XName::new(NS, "extrusion") }
    pub fn o() -> XName { XName::new("http://www.w3.org/2000/xmlns/", "o") }
}

pub mod W10 {
    use super::XName;
    pub const NS: &str = "urn:schemas-microsoft-com:office:word";
    
    pub fn wrap() -> XName { XName::new(NS, "wrap") }
    pub fn w10() -> XName { XName::new("http://www.w3.org/2000/xmlns/", "w10") }
}

pub mod EP {
    use super::XName;
    pub const NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
    
    pub fn total_time() -> XName { XName::new(NS, "TotalTime") }
}

pub mod DCTERMS {
    use super::XName;
    pub const NS: &str = "http://purl.org/dc/terms/";
    
    pub fn created() -> XName { XName::new(NS, "created") }
    pub fn modified() -> XName { XName::new(NS, "modified") }
}

pub mod VML {
    use super::XName;
    pub const NS: &str = "urn:schemas-microsoft-com:vml";
    
    pub fn vml() -> XName { XName::new("http://www.w3.org/2000/xmlns/", "v") }
    pub fn shape() -> XName { XName::new(NS, "shape") }
    pub fn shapetype() -> XName { XName::new(NS, "shapetype") }
}

pub mod W15 {
    use super::XName;
    pub const NS: &str = "http://schemas.microsoft.com/office/word/2012/wordml";

    pub fn w15() -> XName { XName::new("http://www.w3.org/2000/xmlns/", "w15") }

    // Comments extended
    pub fn commentsEx() -> XName { XName::new(NS, "commentsEx") }
    pub fn commentEx() -> XName { XName::new(NS, "commentEx") }
    pub fn paraId() -> XName { XName::new(NS, "paraId") }
    pub fn paraIdParent() -> XName { XName::new(NS, "paraIdParent") }
    pub fn done() -> XName { XName::new(NS, "done") }

    // People
    pub fn people() -> XName { XName::new(NS, "people") }
    pub fn person() -> XName { XName::new(NS, "person") }
    pub fn author() -> XName { XName::new(NS, "author") }
    pub fn presenceInfo() -> XName { XName::new(NS, "presenceInfo") }
    pub fn providerId() -> XName { XName::new(NS, "providerId") }
    pub fn userId() -> XName { XName::new(NS, "userId") }
}

pub mod W16CID {
    use super::XName;
    pub const NS: &str = "http://schemas.microsoft.com/office/word/2016/wordml/cid";

    pub fn w16cid() -> XName { XName::new("http://www.w3.org/2000/xmlns/", "w16cid") }

    pub fn commentsIds() -> XName { XName::new(NS, "commentsIds") }
    pub fn commentId() -> XName { XName::new(NS, "commentId") }
    pub fn paraId() -> XName { XName::new(NS, "paraId") }
    pub fn durableId() -> XName { XName::new(NS, "durableId") }
}

pub mod W16CEX {
    use super::XName;
    pub const NS: &str = "http://schemas.microsoft.com/office/word/2018/wordml/cex";

    pub fn w16cex() -> XName { XName::new("http://www.w3.org/2000/xmlns/", "w16cex") }

    pub fn commentsExtensible() -> XName { XName::new(NS, "commentsExtensible") }
    pub fn commentExtensible() -> XName { XName::new(NS, "commentExtensible") }
    pub fn durableId() -> XName { XName::new(NS, "durableId") }
    pub fn dateUtc() -> XName { XName::new(NS, "dateUtc") }
}

/// Word 2023 Date UTC namespace (w16du)
/// Used for UTC timestamps on revision elements (w:ins, w:del)
pub mod W16DU {
    use super::XName;
    pub const NS: &str = "http://schemas.microsoft.com/office/word/2023/wordml/word16du";

    /// Returns the xmlns:w16du namespace declaration attribute name
    pub fn w16du() -> XName { XName::new("http://www.w3.org/2000/xmlns/", "w16du") }

    /// Returns the w16du:dateUtc attribute name for UTC timestamps on revision elements
    pub fn dateUtc() -> XName { XName::new(NS, "dateUtc") }
}

pub mod W16SE {
    use super::XName;
    pub const NS: &str = "http://schemas.microsoft.com/office/word/2015/wordml/symex";
    
    pub fn w16se() -> XName { XName::new("http://www.w3.org/2000/xmlns/", "w16se") }
}

pub mod WNE {
    use super::XName;
    pub const NS: &str = "http://schemas.microsoft.com/office/word/2006/wordml";
    
    pub fn wne() -> XName { XName::new("http://www.w3.org/2000/xmlns/", "wne") }
}

pub mod WP14 {
    use super::XName;
    pub const NS: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
    
    pub fn wp14() -> XName { XName::new("http://www.w3.org/2000/xmlns/", "wp14") }
}

pub mod WPC {
    use super::XName;
    pub const NS: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
    
    pub fn wpc() -> XName { XName::new("http://www.w3.org/2000/xmlns/", "wpc") }
}

pub mod WPG {
    use super::XName;
    pub const NS: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
    
    pub fn wpg() -> XName { XName::new("http://www.w3.org/2000/xmlns/", "wpg") }
}

pub mod WPI {
    use super::XName;
    pub const NS: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingInk";
    
    pub fn wpi() -> XName { XName::new("http://www.w3.org/2000/xmlns/", "wpi") }
}

pub mod WPS {
    use super::XName;
    pub const NS: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    
    pub fn wps() -> XName { XName::new("http://www.w3.org/2000/xmlns/", "wps") }
}

pub mod XSI {
    use super::XName;
    pub const NS: &str = "http://www.w3.org/2001/XMLSchema-instance";
    
    pub fn schema_location() -> XName { XName::new(NS, "schemaLocation") }
    pub fn no_namespace_schema_location() -> XName { XName::new(NS, "noNamespaceSchemaLocation") }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn word_namespace_creates_valid_xnames() {
        let p = W::p();
        assert_eq!(p.namespace, Some(W::NS.to_string()));
        assert_eq!(p.local_name, "p");
    }

    #[test]
    fn spreadsheet_namespace_creates_valid_xnames() {
        let row = S::row();
        assert_eq!(row.namespace, Some(S::NS.to_string()));
        assert_eq!(row.local_name, "row");
    }
}
