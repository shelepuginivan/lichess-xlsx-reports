use anyhow::{anyhow, bail};
use chrono::{DateTime, Local};
use pgnparse::parser::PgnInfo;
use umya_spreadsheet::{Border, Spreadsheet, Worksheet};

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
            "Otchet_TK_FViS_{}_{}.xlsx",
            self.data.student.id,
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
            .set_style(Styles::info_table());
        sheet
            .get_cell_mut("B4")
            .set_value(&self.data.student.id)
            .set_style(Styles::info_table());

        sheet.add_merge_cells("C3:F3");
        sheet.add_merge_cells("C4:F4");
        sheet
            .get_cell_mut("C3")
            .set_value("ФИО")
            .set_style(Styles::info_table());
        sheet
            .get_cell_mut("C4")
            .set_value(&self.data.student.name)
            .set_style(Styles::info_table());

        sheet
            .get_cell_mut("G3")
            .set_value("Группа")
            .set_style(Styles::info_table());
        sheet
            .get_cell_mut("G4")
            .set_value(&self.data.student.group)
            .set_style(Styles::info_table());

        sheet.add_merge_cells("H3:J3");
        sheet.add_merge_cells("H4:J4");
        sheet
            .get_cell_mut("H3")
            .set_value("Спортивное отделение")
            .set_style(Styles::info_table());
        sheet
            .get_cell_mut("H4")
            .set_value("Шахматы")
            .set_style(Styles::info_table());

        sheet.add_merge_cells("K3:M3");
        sheet.add_merge_cells("K4:M4");
        sheet
            .get_cell_mut("K3")
            .set_value("Преподаватель")
            .set_style(Styles::info_table());
        sheet
            .get_cell_mut("K4")
            .set_value(&self.data.subject.teacher)
            .set_style(Styles::info_table());

        // Fix right borders of the subject info table due to K3:M3 and K4:M4 being merged.
        sheet
            .get_cell_mut("M3")
            .get_style_mut()
            .get_borders_mut()
            .get_right_mut()
            .set_border_style(Border::BORDER_MEDIUM);
        sheet
            .get_cell_mut("M4")
            .get_style_mut()
            .get_borders_mut()
            .get_right_mut()
            .set_border_style(Border::BORDER_MEDIUM);
    }

    fn write_game_info(&self, sheet: &mut Worksheet) {
        let event_info = format!(
            "{} {}",
            self.data.subject.tournament,
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
            .set_style(Styles::game_info_table());
        sheet
            .get_cell_mut("B10")
            .set_value("Черные")
            .set_style(Styles::game_info_table());

        sheet.add_merge_cells("C9:G9");
        sheet.add_merge_cells("C10:G10");
        sheet
            .get_cell_mut("C9")
            .set_value(self.data.student.short_name())
            .set_style(Styles::student_name());
        sheet
            .get_cell_mut("C10")
            .set_value(&self.data.game.opponent)
            .set_style(Styles::student_name());

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
            .set_style(Styles::game_info_table());
        sheet
            .get_cell_mut("H10")
            .set_value("Черные")
            .set_style(Styles::game_info_table());

        sheet.add_merge_cells("I9:M9");
        sheet.add_merge_cells("I10:M10");
        sheet
            .get_cell_mut("I9")
            .set_value(&self.data.game.opponent)
            .set_style(Styles::student_name());
        sheet
            .get_cell_mut("I10")
            .set_value(self.data.student.short_name())
            .set_style(Styles::student_name());

        // Fix right borders of the game info table due to I9:M9 and I10:M10 being merged.
        sheet
            .get_cell_mut("M9")
            .get_style_mut()
            .get_borders_mut()
            .get_right_mut()
            .set_border_style(Border::BORDER_THIN);
        sheet
            .get_cell_mut("M10")
            .get_style_mut()
            .get_borders_mut()
            .get_right_mut()
            .set_border_style(Border::BORDER_THIN);
    }

    async fn write_games(&self, sheet: &mut Worksheet) -> anyhow::Result<()> {
        let (mut game_white, mut game_black) = tokio::try_join!(
            self.data.load_game_as_white(),
            self.data.load_game_as_black()
        )?;

        let moves = calc_row_count(game_white.moves.len(), game_black.moves.len());

        self.write_game(sheet, &mut game_white, moves, 0)?;
        self.write_game(sheet, &mut game_black, moves, 1)?;

        Ok(())
    }

    fn write_game(
        &self,
        sheet: &mut Worksheet,
        pgn: &mut PgnInfo,
        moves: u32,
        index: u32,
    ) -> anyhow::Result<()> {
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
                .set_value_number(move_index + 1)
                .set_style(Styles::game_move_number());

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
                .set_value_number(i - base_row + 1 + move_offset)
                .set_style(Styles::game_move_number());

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
        let result_white = r.next().ok_or(anyhow!("failed to get result"))?;
        let result_black = r.next().ok_or(anyhow!("failed to get result"))?;

        sheet
            .get_cell_mut((base_col + 3, base_row + height + 2))
            .set_value("Итог:")
            .set_style(Styles::game_result());
        sheet
            .get_cell_mut((base_col + 4, base_row + height + 1))
            .set_value("Белые")
            .set_style(Styles::game_result());
        sheet
            .get_cell_mut((base_col + 4, base_row + height + 2))
            .set_value(result_white)
            .set_style(Styles::game_result());
        sheet
            .get_cell_mut((base_col + 5, base_row + height + 1))
            .set_value("Черные")
            .set_style(Styles::game_result());
        sheet
            .get_cell_mut((base_col + 5, base_row + height + 2))
            .set_value(result_black)
            .set_style(Styles::game_result());

        for i in 0..4 {
            sheet
                .get_cell_mut((base_col + 3, base_row + height + i))
                .set_style(Styles::game_move_number_filler());
        }

        // Draw borders around the game moves area. The left border is always present as it is
        // inherited from the move number style.

        // Top and bottom border of game moves area.
        for i in 0..=5 {
            sheet
                .get_cell_mut((base_col + i, base_row))
                .get_style_mut()
                .get_borders_mut()
                .get_top_mut()
                .set_border_style(Border::BORDER_MEDIUM);

            sheet
                .get_cell_mut((base_col + i, base_row + height + 4))
                .get_style_mut()
                .get_borders_mut()
                .get_top_mut()
                .set_border_style(Border::BORDER_MEDIUM);
        }

        // Right border of game moves area.
        for row in 0..height + 4 {
            sheet
                .get_cell_mut((base_col + 5, base_row + row))
                .get_style_mut()
                .get_borders_mut()
                .get_right_mut()
                .set_border_style(Border::BORDER_MEDIUM);
        }

        Ok(())
    }
}
