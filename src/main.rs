use axum::{Json, Router, routing::post};
use reqwest::StatusCode;

use crate::{data::Data, xlsx::XlsxResponse};

mod data;
mod lichess;
mod xlsx;

#[tokio::main]
async fn main() {
    let app = Router::new().route("/api/v1/report", post(generate_report));

    let listener = tokio::net::TcpListener::bind("127.0.0.1:8000")
        .await
        .unwrap();

    axum::serve(listener, app).await.unwrap();
}

async fn generate_report(Json(data): Json<Data>) -> Result<XlsxResponse, StatusCode> {
    Err(StatusCode::NOT_IMPLEMENTED)
}
