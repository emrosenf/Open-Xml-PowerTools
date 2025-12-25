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
    pub fn tbl() -> XName { XName::new(NS, "tbl") }
    pub fn tr() -> XName { XName::new(NS, "tr") }
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
