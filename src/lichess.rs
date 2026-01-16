use reqwest::Url;

/// Converts Lichess game URL to PGN export URL.
pub fn game_url_to_export_url(url: &Url) -> Option<String> {
    let domain = url.domain()?;
    let game_id = url.path_segments()?.next()?;

    Some(format!(
        "{}://{}/game/export/{}?evals=0&clocks=0",
        url.scheme(),
        domain,
        game_id
    ))
}
