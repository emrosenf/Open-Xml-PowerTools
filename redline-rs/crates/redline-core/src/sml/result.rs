use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlComparisonResult {
    pub sheets: Vec<SheetComparisonResult>,
    pub total_changes: usize,
}

impl SmlComparisonResult {
    pub fn new() -> Self {
        Self {
            sheets: Vec::new(),
            total_changes: 0,
        }
    }

    pub fn to_json(&self) -> String {
        serde_json::to_string_pretty(self).unwrap_or_default()
    }
}

impl Default for SmlComparisonResult {
    fn default() -> Self {
        Self::new()
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SheetComparisonResult {
    pub name: String,
    pub status: SheetStatus,
    pub cell_changes: Vec<CellChange>,
    pub row_changes: Vec<RowChange>,
    pub column_changes: Vec<ColumnChange>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum SheetStatus {
    Added,
    Deleted,
    Modified,
    Unchanged,
    Renamed,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellChange {
    pub cell_address: String,
    pub change_type: CellChangeType,
    pub old_value: Option<String>,
    pub new_value: Option<String>,
    pub old_formula: Option<String>,
    pub new_formula: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum CellChangeType {
    Added,
    Deleted,
    ValueChanged,
    FormulaChanged,
    FormatChanged,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RowChange {
    pub row_number: u32,
    pub change_type: RowChangeType,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum RowChangeType {
    Added,
    Deleted,
    Modified,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ColumnChange {
    pub column_letter: String,
    pub change_type: ColumnChangeType,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum ColumnChangeType {
    Added,
    Deleted,
    Modified,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn result_serializes_to_json() {
        let result = SmlComparisonResult::new();
        let json = result.to_json();
        assert!(json.contains("\"sheets\""));
        assert!(json.contains("\"total_changes\""));
    }
}
