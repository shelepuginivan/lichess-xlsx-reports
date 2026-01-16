use anyhow::bail;
use pgnparse::parser::{PgnInfo, parse_pgn_to_rust_struct};
use reqwest::Url;
use serde::Deserialize;

use crate::lichess;

#[derive(Deserialize)]
pub struct Data {
    pub name: String,
    pub id: String,
    pub group: String,
    pub tournament: String,
    pub opponent: String,
    pub game_white_url: String,
    pub game_black_url: String,
}

impl Data {
    pub fn short_name(&self) -> String {
        let mut parts: Vec<&str> = self.name.split(" ").collect();

        if parts.len() >= 3 {
            parts.pop();
        }

        parts.join(" ")
    }

    pub async fn load_game_as_white(&self) -> anyhow::Result<PgnInfo> {
        self.load_game(&self.game_white_url).await
    }

    pub async fn load_game_as_black(&self) -> anyhow::Result<PgnInfo> {
        self.load_game(&self.game_black_url).await
    }

    async fn load_game(&self, game_url: &str) -> anyhow::Result<PgnInfo> {
        let url = Url::parse(game_url)?;
        let export_url = match lichess::game_url_to_export_url(&url) {
            Some(url) => url,
            None => bail!("cannot extract game export URL"),
        };
        let pgn = reqwest::get(export_url).await?.text().await?;

        Ok(parse_pgn_to_rust_struct(pgn))
    }
}
