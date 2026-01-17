use axum::{
    Json, Router,
    routing::{get, post},
};
use reqwest::StatusCode;

mod data;
mod lichess;
mod xlsx;

use crate::data::Data;
use crate::xlsx::{Report, XlsxResponse};

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
        .route("/style.css", get(serve_static!("style.css", "text/css")))
        .route("/app.js", get(serve_static!("app.js", "text/javascript")))
        .route(
            "/favicon.png",
            get(serve_static!("favicon.png", "image/png")),
        )
        .route("/api/v1/report", post(generate_report));

    let listener = tokio::net::TcpListener::bind("127.0.0.1:8000")
        .await
        .unwrap();

    axum::serve(listener, app).await.unwrap();
}

async fn generate_report(Json(data): Json<Data>) -> Result<XlsxResponse, (StatusCode, String)> {
    data.validate()
        .map_err(|e| (StatusCode::BAD_REQUEST, e.to_string()))?;

    let report = Report::new(data);

    let spreadsheet = report
        .generate_spreadsheet()
        .await
        .map_err(|e| (StatusCode::INTERNAL_SERVER_ERROR, e.to_string()))?;

    Ok(XlsxResponse::new(report.filename(), spreadsheet))
}
