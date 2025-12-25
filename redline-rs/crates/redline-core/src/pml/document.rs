use crate::error::Result;
use crate::package::OoxmlPackage;

pub struct PmlDocument {
    package: OoxmlPackage,
}

impl PmlDocument {
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let package = OoxmlPackage::open(bytes)?;
        Ok(Self { package })
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.package.save()
    }

    pub fn package(&self) -> &OoxmlPackage {
        &self.package
    }

    pub fn package_mut(&mut self) -> &mut OoxmlPackage {
        &mut self.package
    }
}
