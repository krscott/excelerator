use calamine::{open_workbook_auto, DataType, Range, Reader, Sheets};
use std::collections::HashMap;
use std::path::Path;

#[derive(Debug, thiserror::Error)]
pub enum Error {
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

pub type Result<T> = std::result::Result<T, Error>;

pub struct WorkbookData {
    header: HashMap<String, u32>,
    range: Range<DataType>,
    pub first_row: u32,
    pub last_row: u32,
    pub first_col: u32,
    pub last_col: u32,
}

impl WorkbookData {
    fn from_workbook_sheet_name(workbook: &mut Sheets, sheet_name: &str) -> Option<Result<Self>> {
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

    pub fn from_path<P: AsRef<Path>>(path: P) -> Result<Self> {
        // For error message only
        let filename = path.as_ref().to_string_lossy().to_string();

        let mut workbook = open_workbook_auto(path)?;

        for s in workbook.sheet_names().to_owned() {
            if let Some(Ok(data)) = Self::from_workbook_sheet_name(&mut workbook, &s) {
                return Ok(data);
            }
        }

        Err(Error::Empty { filename })
    }

    pub fn from_path_with_sheet_name<P: AsRef<Path>>(path: P, sheet_name: &str) -> Result<Self> {
        // For error message only
        let filename = path.as_ref().to_string_lossy().to_string();

        let mut workbook = open_workbook_auto(path)?;

        match Self::from_workbook_sheet_name(&mut workbook, sheet_name) {
            Some(Ok(data)) => Ok(data),
            Some(Err(err)) => Err(err.into()),
            None => Err(Error::EmptySheet {
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
    pub fn get(&self, column_header: &str) -> Option<String> {
        self.source.get(self.row_number, column_header)
    }
}
