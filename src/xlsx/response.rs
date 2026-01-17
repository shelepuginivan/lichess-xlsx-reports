use std::io::Cursor;

use axum::{
    http::{StatusCode, header},
    response::{IntoResponse, Response},
};
use umya_spreadsheet::Spreadsheet;

pub struct XlsxResponse {
    filename: String,
    spreadsheet: Spreadsheet,
}

impl XlsxResponse {
    pub fn new(filename: impl Into<String>, spreadsheet: Spreadsheet) -> Self {
        Self {
            filename: filename.into(),
            spreadsheet,
        }
    }
}

impl IntoResponse for XlsxResponse {
    fn into_response(self) -> Response {
        let mut buffer = Cursor::new(Vec::new());

        match umya_spreadsheet::writer::xlsx::write_writer(&self.spreadsheet, &mut buffer) {
            Ok(_) => {
                let bytes = buffer.into_inner();

                (
                    StatusCode::OK,
                    [
                        (
                            header::CONTENT_TYPE,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        ),
                        (
                            header::CONTENT_DISPOSITION,
                            &format!(r#"attachment; filename="{}""#, &self.filename),
                        ),
                    ],
                    bytes,
                )
                    .into_response()
            }
            Err(e) => (
                StatusCode::INTERNAL_SERVER_ERROR,
                format!("Failed to generate spreadsheet: {e}"),
            )
                .into_response(),
        }
    }
}
