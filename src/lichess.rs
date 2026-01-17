use reqwest::Url;

/// Converts Lichess game URL to PGN export URL.
pub fn game_url_to_export_url(url: &Url) -> Option<String> {
    let domain = url.domain()?;
    let mut game_id = url.path_segments()?.next()?;

    // The long game ID can be provided. In this case, the first 8 characters is the actual ID, and
    // the last 4 is the presentation suffix. The suffix can be safely trimmed, which results in
    // the analysis board ID.
    if game_id.len() == 12 {
        game_id = &game_id[0..8];
    }

    Some(format!(
        "{}://{}/game/export/{}?evals=0&clocks=0",
        url.scheme(),
        domain,
        game_id
    ))
}
