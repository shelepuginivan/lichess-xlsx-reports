use umya_spreadsheet::{
    Alignment, Border, Color, Font, HorizontalAlignmentValues, Style, VerticalAlignmentValues,
};

pub struct Styles {
    accent_color: String,
    font_name: String,
}

impl Default for Styles {
    fn default() -> Self {
        Self {
            accent_color: String::from("ff2e75b6"),
            font_name: String::from("Calibri"),
        }
    }
}

impl Styles {
    pub fn title(&self) -> Style {
        Style::default()
            .set_font(self.font_bold().set_size(14_f64).to_owned())
            .set_alignment(self.align_center())
            .to_owned()
    }

    pub fn header(&self) -> Style {
        Style::default()
            .set_font(self.font_bold())
            .set_alignment(self.align_center())
            .to_owned()
    }

    pub fn info_table(&self) -> Style {
        let mut style = Style::default();

        let borders = style.get_borders_mut();
        let border = Border::BORDER_MEDIUM;

        borders.get_top_mut().set_border_style(border);
        borders.get_left_mut().set_border_style(border);
        borders.get_right_mut().set_border_style(border);
        borders.get_bottom_mut().set_border_style(border);

        style
            .set_font(self.font_bold())
            .set_alignment(self.align_center())
            .to_owned()
    }

    pub fn game_info_table(&self) -> Style {
        let mut style = self.header();

        let borders = style.get_borders_mut();
        let border = Border::BORDER_THIN;

        borders.get_top_mut().set_border_style(border);
        borders.get_left_mut().set_border_style(border);
        borders.get_right_mut().set_border_style(border);
        borders.get_bottom_mut().set_border_style(border);

        style
    }

    pub fn student_name(&self) -> Style {
        let mut style = Style::default()
            .set_alignment(self.align_center())
            .set_font(self.font_normal().set_size(11_f64).to_owned())
            .to_owned();

        let borders = style.get_borders_mut();
        let border = Border::BORDER_THIN;

        borders.get_top_mut().set_border_style(border);
        borders.get_left_mut().set_border_style(border);
        borders.get_right_mut().set_border_style(border);
        borders.get_bottom_mut().set_border_style(border);

        style
    }

    pub fn game_move(&self) -> Style {
        let mut style = Style::default().to_owned();
        let borders = style.get_borders_mut();
        let border = Border::BORDER_THIN;

        borders.get_top_mut().set_border_style(border);
        borders.get_left_mut().set_border_style(border);
        borders.get_right_mut().set_border_style(border);
        borders.get_bottom_mut().set_border_style(border);

        style
    }

    pub fn game_move_number(&self) -> Style {
        let mut style = Style::default().to_owned();
        let borders = style.get_borders_mut();
        let border = Border::BORDER_MEDIUM;

        borders.get_left_mut().set_border_style(border);
        borders.get_right_mut().set_border_style(border);

        let font = self
            .font_normal()
            .set_bold(true)
            .set_color(self.accent_color())
            .to_owned();

        style
            .set_alignment(self.align_center())
            .set_font(font)
            .to_owned()
    }

    pub fn game_move_number_filler(&self) -> Style {
        let mut style = Style::default().to_owned();
        let borders = style.get_borders_mut();
        let border = Border::BORDER_MEDIUM;

        borders.get_left_mut().set_border_style(border);
        borders.get_right_mut().set_border_style(border);

        style.set_alignment(self.align_center()).to_owned()
    }

    pub fn game_result(&self) -> Style {
        Style::default()
            .set_alignment(self.align_center())
            .set_font(self.font_normal())
            .to_owned()
    }

    fn align_center(&self) -> Alignment {
        let mut center = Alignment::default();

        center.set_wrap_text(true);
        center.set_vertical(VerticalAlignmentValues::Center);
        center.set_horizontal(HorizontalAlignmentValues::Center);

        center
    }

    fn accent_color(&self) -> Color {
        Color::default().set_argb(&self.accent_color).to_owned()
    }

    fn font_normal(&self) -> Font {
        Font::default()
            .set_name(&self.font_name)
            .set_size(10_f64)
            .to_owned()
    }

    fn font_bold(&self) -> Font {
        self.font_normal()
            .set_size(11_f64)
            .set_bold(true)
            .to_owned()
    }
}
