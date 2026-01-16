use anyhow::bail;
use chrono::{DateTime, Local};
use pgnparse::parser::PgnInfo;
use umya_spreadsheet::{Spreadsheet, Worksheet};

use crate::data::Data;
use crate::xlsx::styles::Styles;
use crate::xlsx::utils::calc_row_count;

pub struct Report {
    data: Data,
    generation_time: DateTime<Local>,
}

impl Report {
    pub fn new(data: Data) -> Self {
        Self {
            data,
            generation_time: Local::now(),
        }
    }

    pub fn filename(&self) -> String {
        format!(
            "Отчет.ТК.ФВиС.{}.{}.xlsx",
            self.data.id,
            self.generation_time.format("%Y-%m-%d"),
        )
    }
}

impl Report {
    pub async fn generate_spreadsheet(&self) -> anyhow::Result<Spreadsheet> {
        let mut book = umya_spreadsheet::new_file();
        let sheet = match book.get_sheet_by_name_mut("Sheet1") {
            Some(s) => s,
            None => bail!("cannot find default sheet"),
        };

        self.write_title(sheet);
        self.write_info(sheet);
        self.write_game_info(sheet);
        self.write_games(sheet).await?;

        Ok(book)
    }

    fn write_title(&self, sheet: &mut Worksheet) {
        let title = r#"Отчет о результатах самостоятельной работы обучающегося по дисциплинам "Физическая культура" или "Элективные курсы по физической культуре и спорту""#;

        sheet.add_merge_cells("B1:M1");
        sheet
            .get_cell_mut("B1")
            .set_value(title)
            .set_style(Styles::title());
    }

    fn write_info(&self, sheet: &mut Worksheet) {
        sheet
            .get_cell_mut("B3")
            .set_value("Студ. билет")
            .set_style(Styles::header());
        sheet
            .get_cell_mut("B4")
            .set_value(&self.data.id)
            .set_style(Styles::header());

        sheet.add_merge_cells("C3:F3");
        sheet.add_merge_cells("C4:F4");
        sheet
            .get_cell_mut("C3")
            .set_value("ФИО")
            .set_style(Styles::header());
        sheet
            .get_cell_mut("C4")
            .set_value(&self.data.name)
            .set_style(Styles::header());

        sheet
            .get_cell_mut("G3")
            .set_value("Группа")
            .set_style(Styles::header());
        sheet
            .get_cell_mut("G4")
            .set_value(&self.data.group)
            .set_style(Styles::header());

        sheet.add_merge_cells("H3:J3");
        sheet.add_merge_cells("H4:J4");
        sheet
            .get_cell_mut("H3")
            .set_value("Спортивное отделение")
            .set_style(Styles::header());
        sheet
            .get_cell_mut("H4")
            .set_value("Шахматы")
            .set_style(Styles::header());

        sheet.add_merge_cells("K3:M3");
        sheet.add_merge_cells("K4:M4");
        sheet
            .get_cell_mut("K3")
            .set_value("Преподаватель")
            .set_style(Styles::header());
        sheet
            .get_cell_mut("K4")
            .set_value("С.В. Иванов")
            .set_style(Styles::header());
    }

    fn write_game_info(&self, sheet: &mut Worksheet) {
        let event_info = format!(
            "Шахматный турнир №{} {}",
            self.data.tournament,
            self.generation_time.format("%d.%m.%Y"),
        );

        sheet.add_merge_cells("B7:G7");
        sheet.add_merge_cells("B8:G8");
        sheet
            .get_cell_mut("B7")
            .set_value("Шахматная партия №1")
            .set_style(Styles::header());
        sheet
            .get_cell_mut("B8")
            .set_value(&event_info)
            .set_style(Styles::header());

        sheet
            .get_cell_mut("B9")
            .set_value("Белые")
            .set_style(Styles::header());
        sheet
            .get_cell_mut("B10")
            .set_value("Черные")
            .set_style(Styles::header());

        sheet.add_merge_cells("C9:G9");
        sheet.add_merge_cells("C10:G10");
        sheet.get_cell_mut("C9").set_value(self.data.short_name());
        sheet.get_cell_mut("C10").set_value(&self.data.opponent);

        sheet.add_merge_cells("H7:M7");
        sheet.add_merge_cells("H8:M8");
        sheet
            .get_cell_mut("H7")
            .set_value("Шахматная партия №2")
            .set_style(Styles::header());
        sheet
            .get_cell_mut("H8")
            .set_value(&event_info)
            .set_style(Styles::header());

        sheet
            .get_cell_mut("H9")
            .set_value("Белые")
            .set_style(Styles::header());
        sheet
            .get_cell_mut("H10")
            .set_value("Черные")
            .set_style(Styles::header());

        sheet.add_merge_cells("I9:M9");
        sheet.add_merge_cells("I10:M10");
        sheet.get_cell_mut("I9").set_value(&self.data.opponent);
        sheet.get_cell_mut("I10").set_value(self.data.short_name());
    }

    async fn write_games(&self, sheet: &mut Worksheet) -> anyhow::Result<()> {
        let mut game_white = self.data.load_game_as_white().await?;
        let mut game_black = self.data.load_game_as_black().await?;

        let moves = calc_row_count(game_white.moves.len(), game_black.moves.len());

        self.write_game(sheet, &mut game_white, moves, 0);
        self.write_game(sheet, &mut game_black, moves, 1);

        Ok(())
    }

    fn write_game(&self, sheet: &mut Worksheet, pgn: &mut PgnInfo, moves: u32, index: u32) {
        let base_col = index * 6 + 2;
        let base_row = 14;
        let height = moves / 2 + 2;

        sheet.get_cell_mut((base_col, 13)).set_value("№");
        sheet.get_cell_mut((base_col + 1, 13)).set_value("Белые");
        sheet.get_cell_mut((base_col + 2, 13)).set_value("Черные");
        sheet.get_cell_mut((base_col + 3, 13)).set_value("№");
        sheet.get_cell_mut((base_col + 4, 13)).set_value("Белые");
        sheet.get_cell_mut((base_col + 5, 13)).set_value("Черные");

        for i in base_row..base_row + height {
            let move_index = i - base_row;

            sheet
                .get_cell_mut((base_col, i))
                .set_value_number(move_index + 1);

            let white_move_index = (2 * move_index) as usize;
            let white_move = if white_move_index < pgn.moves.len() {
                pgn.moves[white_move_index].san.as_str()
            } else {
                "/"
            };

            sheet
                .get_cell_mut((base_col + 1, i))
                .set_value(white_move)
                .set_style(Styles::game_move());

            let black_move_index = (2 * move_index + 1) as usize;
            let black_move = if black_move_index < pgn.moves.len() {
                pgn.moves[black_move_index].san.as_str()
            } else {
                "/"
            };

            sheet
                .get_cell_mut((base_col + 2, i))
                .set_value(black_move)
                .set_style(Styles::game_move());
        }

        let move_offset = height;
        let height = height - 4;

        for i in base_row..base_row + height {
            let move_index = i - base_row + move_offset;

            sheet
                .get_cell_mut((base_col + 3, i))
                .set_value_number(i - base_row + 1 + move_offset);

            let white_move_index = (2 * move_index) as usize;
            let white_move = if white_move_index < pgn.moves.len() {
                pgn.moves[white_move_index].san.as_str()
            } else {
                "/"
            };

            sheet
                .get_cell_mut((base_col + 4, i))
                .set_value(white_move)
                .set_style(Styles::game_move());

            let black_move_index = (2 * move_index + 1) as usize;
            let black_move = if black_move_index < pgn.moves.len() {
                pgn.moves[black_move_index].san.as_str()
            } else {
                "/"
            };

            sheet
                .get_cell_mut((base_col + 5, i))
                .set_value(black_move)
                .set_style(Styles::game_move());
        }

        let result = pgn.get_header("Result");
        let mut r = result.split('-');
        let result_white = r.next().unwrap();
        let result_black = r.next().unwrap();

        sheet
            .get_cell_mut((base_col + 3, base_row + height + 2))
            .set_value("Итог:");
        sheet
            .get_cell_mut((base_col + 4, base_row + height + 1))
            .set_value("Белые");
        sheet
            .get_cell_mut((base_col + 4, base_row + height + 2))
            .set_value(result_white);
        sheet
            .get_cell_mut((base_col + 5, base_row + height + 1))
            .set_value("Черные");
        sheet
            .get_cell_mut((base_col + 5, base_row + height + 2))
            .set_value(result_black);
    }
}
