FROM rust:1.91 as builder

WORKDIR /app

COPY src src
COPY static static
COPY Cargo.toml Cargo.toml
COPY Cargo.lock Cargo.lock

RUN cargo build --release


FROM debian:trixie-slim AS app

RUN apt-get update && apt-get install -y --no-install-recommends \
    ca-certificates \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY --from=builder /app/target/release/lichess-xlsx-reports .

EXPOSE 8000
ENV LISTEN_ADDR=0.0.0.0:8000
ENTRYPOINT ["./lichess-xlsx-reports"]
