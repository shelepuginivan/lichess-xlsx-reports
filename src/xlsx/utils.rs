pub fn calc_row_count(moves_white: usize, moves_black: usize) -> u32 {
    let x = (moves_white / 2) as u32;
    let y = (moves_black / 2) as u32;

    let moves = x.max(y).max(60);
    let v = moves % 10;
    let remaining = (10 - v) % 10;

    moves + remaining
}
