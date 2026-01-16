use serde::Deserialize;

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
