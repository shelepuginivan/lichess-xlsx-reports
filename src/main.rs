use std::env;
use std::time::Duration;

use axum::{
    Json, Router,
    routing::{get, post},
};
use reqwest::StatusCode;
use tokio::signal;
use tower_http::timeout::TimeoutLayer;

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
        .route("/api/v1/report", post(generate_report))
        .layer(TimeoutLayer::with_status_code(
            StatusCode::REQUEST_TIMEOUT,
            Duration::from_secs(15),
        ));

    let addr = env::var("LISTEN_ADDR").unwrap_or(String::from("0.0.0.0:8000"));
    let listener = tokio::net::TcpListener::bind(addr).await.unwrap();

    axum::serve(listener, app)
        .with_graceful_shutdown(shutdown_signal())
        .await
        .unwrap();
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

async fn shutdown_signal() {
    let ctrl_c = async {
        signal::ctrl_c()
            .await
            .expect("failed to install Ctrl+C handler");
    };

    #[cfg(unix)]
    let terminate = async {
        signal::unix::signal(signal::unix::SignalKind::terminate())
            .expect("failed to install signal handler")
            .recv()
            .await;
    };

    #[cfg(not(unix))]
    let terminate = std::future::pending::<()>();

    tokio::select! {
        _ = ctrl_c => {},
        _ = terminate => {},
    }
}
