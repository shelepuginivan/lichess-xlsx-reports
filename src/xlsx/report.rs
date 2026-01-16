use chrono::{DateTime, Local};

use crate::data::Data;

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
