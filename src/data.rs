use anyhow::bail;
use pgnparse::parser::{PgnInfo, parse_pgn_to_rust_struct};
use reqwest::Url;
use serde::Deserialize;

use crate::lichess;

#[derive(Deserialize)]
pub struct StudentData {
    pub name: String,
    pub group: String,
    pub id: String,
}

impl StudentData {
    pub fn short_name(&self) -> String {
        let mut parts: Vec<&str> = self.name.split(" ").collect();

        if parts.len() >= 3 {
            parts.pop();
        }

        parts.join(" ")
    }
}

#[derive(Deserialize)]
pub struct SubjectData {
    pub teacher: String,
    pub tournament: String,
}

#[derive(Deserialize)]
pub struct GameData {
    pub opponent: String,
    pub white_url: String,
    pub black_url: String,
}

impl GameData {
    fn validate(&self) -> anyhow::Result<()> {
        let url = match Url::parse(&self.white_url) {
            Ok(url) => url,
            Err(_) => bail!("Ссылка на игру белыми невалидна"),
        };

        if lichess::game_url_to_export_url(&url).is_none() {
            bail!("Неверный формат ссылки на игру белыми")
        }

        let url = match Url::parse(&self.black_url) {
            Ok(url) => url,
            Err(_) => bail!("Ссылка на игру черными невалидна"),
        };

        if lichess::game_url_to_export_url(&url).is_none() {
            bail!("Неверный формат ссылки на игру черными")
        }

        Ok(())
    }
}

#[derive(Deserialize)]
pub struct Data {
    pub student: StudentData,
    pub subject: SubjectData,
    pub game: GameData,
}

impl Data {
    pub fn validate(&self) -> anyhow::Result<()> {
        self.game.validate()?;

        Ok(())
    }

    pub async fn load_game_as_white(&self) -> anyhow::Result<PgnInfo> {
        self.load_game(&self.game.white_url).await
    }

    pub async fn load_game_as_black(&self) -> anyhow::Result<PgnInfo> {
        self.load_game(&self.game.black_url).await
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
