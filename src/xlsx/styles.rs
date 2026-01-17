use umya_spreadsheet::{
    Alignment, Border, Color, Font, HorizontalAlignmentValues, Style, VerticalAlignmentValues,
};

pub struct Styles {}

impl Styles {
    pub fn title() -> Style {
        let font = Font::default()
            .set_name("Calibri")
            .set_size(11f64)
            .set_bold(true)
            .to_owned();

        Style::default()
            .set_font(font)
            .set_alignment(Self::align_center())
            .to_owned()
    }

    pub fn header() -> Style {
        let font = Font::default()
            .set_name("Calibri")
            .set_size(11f64)
            .set_bold(true)
            .to_owned();

        Style::default()
            .set_font(font)
            .set_alignment(Self::align_center())
            .to_owned()
    }

    pub fn info_table() -> Style {
        let font = Font::default()
            .set_name("Calibri")
            .set_size(11f64)
            .set_bold(true)
            .to_owned();

        let mut style = Style::default();

        let borders = style.get_borders_mut();
        let border = Border::BORDER_MEDIUM;

        borders.get_top_mut().set_border_style(border);
        borders.get_left_mut().set_border_style(border);
        borders.get_right_mut().set_border_style(border);
        borders.get_bottom_mut().set_border_style(border);

        style
            .set_font(font)
            .set_alignment(Self::align_center())
            .to_owned()
    }

    pub fn game_info_table() -> Style {
        let mut style = Self::header();

        let borders = style.get_borders_mut();
        let border = Border::BORDER_THIN;

        borders.get_top_mut().set_border_style(border);
        borders.get_left_mut().set_border_style(border);
        borders.get_right_mut().set_border_style(border);
        borders.get_bottom_mut().set_border_style(border);

        style
    }

    pub fn student_name() -> Style {
        let mut style = Style::default()
            .set_alignment(Self::align_center())
            .set_font(Font::default().set_name("Calibri").to_owned())
            .to_owned();

        let borders = style.get_borders_mut();
        let border = Border::BORDER_THIN;

        borders.get_top_mut().set_border_style(border);
        borders.get_left_mut().set_border_style(border);
        borders.get_right_mut().set_border_style(border);
        borders.get_bottom_mut().set_border_style(border);

        style
    }

    pub fn game_move() -> Style {
        let mut style = Style::default().to_owned();
        let borders = style.get_borders_mut();
        let border = Border::BORDER_THIN;

        borders.get_top_mut().set_border_style(border);
        borders.get_left_mut().set_border_style(border);
        borders.get_right_mut().set_border_style(border);
        borders.get_bottom_mut().set_border_style(border);

        style
    }

    pub fn game_move_number() -> Style {
        let mut style = Style::default().to_owned();
        let borders = style.get_borders_mut();
        let border = Border::BORDER_MEDIUM;

        borders.get_left_mut().set_border_style(border);
        borders.get_right_mut().set_border_style(border);

        let font = Font::default()
            .set_name("Calibri")
            .set_bold(true)
            .set_size(10_f64)
            .set_color(Self::accent_color())
            .to_owned();

        style
            .set_alignment(Self::align_center().to_owned())
            .set_font(font)
            .to_owned()
    }

    pub fn game_move_number_filler() -> Style {
        let mut style = Style::default().to_owned();
        let borders = style.get_borders_mut();
        let border = Border::BORDER_MEDIUM;

        borders.get_left_mut().set_border_style(border);
        borders.get_right_mut().set_border_style(border);

        style
            .set_alignment(Self::align_center().to_owned())
            .to_owned()
    }

    pub fn game_result() -> Style {
        let font = Font::default()
            .set_name("Calibri")
            .set_size(10_f64)
            .to_owned();

        Style::default()
            .set_alignment(Self::align_center().to_owned())
            .set_font(font)
            .to_owned()
    }

    fn align_center() -> Alignment {
        let mut center = Alignment::default();

        center.set_wrap_text(true);
        center.set_vertical(VerticalAlignmentValues::Center);
        center.set_horizontal(HorizontalAlignmentValues::Center);

        center
    }

    fn accent_color() -> Color {
        Color::default().set_argb("ff2e75b6").to_owned()
    }
}
