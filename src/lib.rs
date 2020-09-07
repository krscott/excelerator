use calamine::{open_workbook_auto, DataType, Range, Reader, Sheets};
use std::collections::HashMap;
use std::path::Path;
use std::str::FromStr;

#[derive(Debug, thiserror::Error)]
pub enum LoadError {
    #[error("No data found in '{}'", .filename)]
    Empty { filename: String },

    #[error("No data found in sheet '{}' in '{}'", .sheet_name, .filename)]
    EmptySheet {
        filename: String,
        sheet_name: String,
    },

    #[error(transparent)]
    CalamineError(#[from] calamine::Error),
}

#[derive(Debug, thiserror::Error)]
pub enum DataError {
    #[error("Key '{}' value could not be parsed: {}", .key, .value)]
    ParseError { key: String, value: String },

    #[error("No data found for key '{}'", .0)]
    NoValue(String),
}

pub struct WorkbookData {
    header: HashMap<String, u32>,
    range: Range<DataType>,
    pub first_row: u32,
    pub last_row: u32,
    pub first_col: u32,
    pub last_col: u32,
}

impl WorkbookData {
    fn from_workbook_sheet_name(
        workbook: &mut Sheets,
        sheet_name: &str,
    ) -> Option<Result<Self, LoadError>> {
        let range = match workbook.worksheet_range(sheet_name)? {
            Ok(range) => range,
            Err(err) => return Some(Err(err.into())),
        };

        let (mut first_row, first_col) = range.start()?;
        let (last_row, last_col) = range.end()?;

        let min_cols = last_col - first_col + 1;

        let mut rows = range.rows();

        loop {
            first_row += 1;

            let row = rows.next()?;

            let row: Vec<_> = row.iter().map(|h| h.to_string()).collect();
            let count_cols = row.iter().filter(|x| !x.is_empty()).count() as u32;

            if count_cols >= min_cols {
                let header = row
                    .into_iter()
                    .enumerate()
                    .map(|(i, s)| (s, i as u32))
                    .collect();

                return Some(Ok(Self {
                    header,
                    range,
                    first_row,
                    last_row,
                    first_col,
                    last_col,
                }));
            }
        }
    }

    pub fn from_path<P: AsRef<Path>>(path: P) -> Result<Self, LoadError> {
        // For error message only
        let filename = path.as_ref().to_string_lossy().to_string();

        let mut workbook = open_workbook_auto(path)?;

        for s in workbook.sheet_names().to_owned() {
            if let Some(Ok(data)) = Self::from_workbook_sheet_name(&mut workbook, &s) {
                return Ok(data);
            }
        }

        Err(LoadError::Empty { filename })
    }

    pub fn from_path_with_sheet_name<P: AsRef<Path>>(
        path: P,
        sheet_name: &str,
    ) -> Result<Self, LoadError> {
        // For error message only
        let filename = path.as_ref().to_string_lossy().to_string();

        let mut workbook = open_workbook_auto(path)?;

        match Self::from_workbook_sheet_name(&mut workbook, sheet_name) {
            Some(Ok(data)) => Ok(data),
            Some(Err(err)) => Err(err.into()),
            None => Err(LoadError::EmptySheet {
                filename,
                sheet_name: sheet_name.to_owned(),
            }),
        }
    }

    pub fn get(&self, row_number: u32, column_header: &str) -> Option<String> {
        if row_number < self.first_row || row_number > self.last_row {
            return None;
        }

        let col_number = self.header.get(column_header)?;

        let value = self.range.get_value((row_number, *col_number))?;

        Some(value.to_string())
    }

    pub fn is_row_empty(&self, row_number: u32) -> bool {
        0 == self
            .header
            .keys()
            .filter_map(|h| self.get(row_number, h))
            .filter(|v| !v.is_empty())
            .count()
    }

    pub fn iter_rows<'a>(&'a self) -> RowsIterator<'a> {
        RowsIterator {
            source: self,
            current_row: self.first_row,
            last_row: self.last_row,
            first_col: self.first_col,
            last_col: self.last_col,
        }
    }
}

pub struct RowsIterator<'a> {
    source: &'a WorkbookData,
    pub current_row: u32,
    pub last_row: u32,
    pub first_col: u32,
    pub last_col: u32,
}

impl<'a> Iterator for RowsIterator<'a> {
    type Item = RowData<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.current_row > self.last_row {
            return None;
        }

        self.current_row += 1;

        Some(RowData {
            source: self.source,
            row_number: self.current_row,
        })
    }
}

pub struct RowData<'a> {
    source: &'a WorkbookData,
    row_number: u32,
}

impl<'a> RowData<'a> {
    /// Get the row number of this data in the source workbook
    pub fn number(&self) -> u32 {
        self.row_number
    }

    /// Get the value in the cell of this row with the matching column header
    pub fn get(&self, column_header: &str) -> Result<String, DataError> {
        match self.source.get(self.row_number, column_header) {
            Some(value) => Ok(value),
            None => Err(DataError::NoValue(column_header.into())),
        }
    }

    pub fn parse<T: FromStr>(&self, column_header: &str) -> Result<T, DataError> {
        let value_str = self.get(column_header)?;

        value_str.parse().map_err(|_| DataError::ParseError {
            key: column_header.into(),
            value: value_str,
        })
    }

    pub fn is_empty(&self) -> bool {
        self.source.is_row_empty(self.row_number)
    }
}

pub fn from_path<P: AsRef<Path>>(path: P) -> Result<WorkbookData, LoadError> {
    WorkbookData::from_path(path)
}

pub fn from_path_with_sheet_name<P: AsRef<Path>>(
    path: P,
    sheet_name: &str,
) -> Result<WorkbookData, LoadError> {
    WorkbookData::from_path_with_sheet_name(path, sheet_name)
}
