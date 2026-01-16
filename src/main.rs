use axum::{Json, Router, routing::post};
use chrono::Local;
use reqwest::StatusCode;

use crate::{
    data::Data,
    xlsx::{Formatter, XlsxResponse},
};

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
    let report = match Formatter::new().format_data(&data).await {
        Ok(r) => r,
        Err(_) => return Err(StatusCode::INTERNAL_SERVER_ERROR),
    };

    let filename = format!(
        "Отчет.ТК.ФВиС.{}.{}.xlsx",
        data.id,
        Local::now().format("%Y-%m-%d").to_string(),
    );

    Ok(XlsxResponse::new(filename, report))
}
