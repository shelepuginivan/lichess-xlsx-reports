use axum::{
    Json, Router,
    routing::{get, post},
};
use reqwest::StatusCode;

use crate::{
    data::Data,
    xlsx::{Report, XlsxResponse},
};

mod data;
mod lichess;
mod xlsx;

macro_rules! serve_static {
    ($path:literal, $content_type:expr) => {
        || async {
            (
                StatusCode::OK,
                [(axum::http::header::CONTENT_TYPE, $content_type)],
                include_bytes!(concat!("../static/", $path)),
            )
        }
    };
}

#[tokio::main]
async fn main() {
    let app = Router::new()
        .route("/", get(serve_static!("index.html", "text/html")))
        .route("/api/v1/report", post(generate_report));

    let listener = tokio::net::TcpListener::bind("127.0.0.1:8000")
        .await
        .unwrap();

    axum::serve(listener, app).await.unwrap();
}

async fn generate_report(Json(data): Json<Data>) -> Result<XlsxResponse, StatusCode> {
    let report = Report::new(data);

    let spreadsheet = report
        .generate_spreadsheet()
        .await
        .map_err(|_| StatusCode::INTERNAL_SERVER_ERROR)?;

    Ok(XlsxResponse::new(report.filename(), spreadsheet))
}
