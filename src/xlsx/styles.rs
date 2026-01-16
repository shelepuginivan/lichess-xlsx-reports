
use umya_spreadsheet::{
    Alignment, Border, Color, Font, HorizontalAlignmentValues, Style, VerticalAlignmentValues,
};

pub struct Styles {}

impl Styles {
    pub fn title() -> Style {
        Style::default()
            .set_font(Font::default().set_size(14f64).set_bold(true).to_owned())
            .set_alignment(Self::align_center())
            .to_owned()
    }

    pub fn header() -> Style {
        Border::default();

        Style::default()
            .set_font(Font::default().set_size(11f64).set_bold(true).to_owned())
            .set_alignment(Self::align_center())
            .to_owned()
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

        style
            .set_alignment(Self::align_center().to_owned())
            .set_font(
                Font::default()
                    .set_color(Color::default().set_argb(Color::COLOR_BLUE).to_owned())
                    .to_owned(),
            )
            .to_owned()
    }

    fn align_center() -> Alignment {
        let mut center = Alignment::default();

        center.set_wrap_text(true);
        center.set_vertical(VerticalAlignmentValues::Center);
        center.set_horizontal(HorizontalAlignmentValues::Center);

        center
    }
}
