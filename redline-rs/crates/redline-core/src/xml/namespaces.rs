#![allow(non_snake_case)]

use super::xname::XName;

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
    pub fn footnoteRef() -> XName { XName::new(NS, "footnoteRef") }
    pub fn endnoteRef() -> XName { XName::new(NS, "endnoteRef") }
    pub fn separator() -> XName { XName::new(NS, "separator") }
    pub fn continuationSeparator() -> XName { XName::new(NS, "continuationSeparator") }
    pub fn sdtPr() -> XName { XName::new(NS, "sdtPr") }
    pub fn sdtEndPr() -> XName { XName::new(NS, "sdtEndPr") }
    pub fn smartTag() -> XName { XName::new(NS, "smartTag") }
    pub fn smartTagPr() -> XName { XName::new(NS, "smartTagPr") }
    pub fn ruby() -> XName { XName::new(NS, "ruby") }
    pub fn rubyPr() -> XName { XName::new(NS, "rubyPr") }
    pub fn gridSpan() -> XName { XName::new(NS, "gridSpan") }
    pub fn val() -> XName { XName::new(NS, "val") }
    // rsid attributes (revision session IDs) - stripped during hashing
    pub fn rsid() -> XName { XName::new(NS, "rsid") }
    pub fn rsidDel() -> XName { XName::new(NS, "rsidDel") }
    pub fn rsidP() -> XName { XName::new(NS, "rsidP") }
    pub fn rsidR() -> XName { XName::new(NS, "rsidR") }
    pub fn rsidRDefault() -> XName { XName::new(NS, "rsidRDefault") }
    pub fn rsidRPr() -> XName { XName::new(NS, "rsidRPr") }
    pub fn rsidSect() -> XName { XName::new(NS, "rsidSect") }
    pub fn rsidTr() -> XName { XName::new(NS, "rsidTr") }
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
}

pub mod P {
    use super::XName;
    pub const NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    
    pub fn presentation() -> XName { XName::new(NS, "presentation") }
    pub fn sld() -> XName { XName::new(NS, "sld") }
    pub fn cSld() -> XName { XName::new(NS, "cSld") }
    pub fn spTree() -> XName { XName::new(NS, "spTree") }
    pub fn sp() -> XName { XName::new(NS, "sp") }
    pub fn txBody() -> XName { XName::new(NS, "txBody") }
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
}

pub mod R {
    use super::XName;
    pub const NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    
    pub fn id() -> XName { XName::new(NS, "id") }
}

pub mod MC {
    use super::XName;
    pub const NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    
    pub fn AlternateContent() -> XName { XName::new(NS, "AlternateContent") }
    pub fn Choice() -> XName { XName::new(NS, "Choice") }
    pub fn Fallback() -> XName { XName::new(NS, "Fallback") }
}

pub mod CP {
    pub const NS: &str = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
}

pub mod DC {
    pub const NS: &str = "http://purl.org/dc/elements/1.1/";
}

pub mod PT {
    use super::XName;
    pub const NS: &str = "http://powertools.codeplex.com/2011";
    
    pub fn Unid() -> XName { XName::new(NS, "Unid") }
    pub fn SHA1Hash() -> XName { XName::new(NS, "SHA1Hash") }
    pub fn CorrelatedSHA1Hash() -> XName { XName::new(NS, "CorrelatedSHA1Hash") }
}

pub mod M {
    use super::XName;
    pub const NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    
    pub fn f() -> XName { XName::new(NS, "f") }
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
}

pub mod W14 {
    use super::XName;
    pub const NS: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
    
    pub fn paraId() -> XName { XName::new(NS, "paraId") }
    pub fn textId() -> XName { XName::new(NS, "textId") }
}

pub mod O {
    use super::XName;
    pub const NS: &str = "urn:schemas-microsoft-com:office:office";
    
    pub fn relid() -> XName { XName::new(NS, "relid") }
    pub fn lock() -> XName { XName::new(NS, "lock") }
    pub fn extrusion() -> XName { XName::new(NS, "extrusion") }
}

pub mod W10 {
    use super::XName;
    pub const NS: &str = "urn:schemas-microsoft-com:office:word";
    
    pub fn wrap() -> XName { XName::new(NS, "wrap") }
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
