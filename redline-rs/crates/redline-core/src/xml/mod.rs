pub mod arena;
pub mod builder;
pub mod namespaces;
pub mod node;
pub mod parser;
pub mod xname;

pub use arena::XmlDocument;
pub use node::{XmlNode, XmlNodeData};
pub use xname::{XName, XAttribute};
pub use namespaces::{W, S, P, A, R, MC, CP, DC, PT, W16DU};
