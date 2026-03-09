use crate::chart::{BarDirection, Chart, ChartAxis, ChartGrouping, ChartSeries, ChartType};
use crate::error::{Result, XlsxError};
use crate::pivot_table::{
    PivotDataField, PivotField, PivotFieldSort, PivotSourceReference, PivotTable,
};
use crate::style::Style;
use crate::workbook::Workbook;
use crate::worksheet::{PivotValueSpec, Worksheet};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FinFormat {
    Bps1,
    Bps2,
    UsdMillions1,
    UsdMillions2,
    UsdBillions2,
    Pct1Signed,
    Pct2Signed,
    Multiple2,
    IntegerThousands,
}

impl FinFormat {
    pub fn format_code(self) -> &'static str {
        match self {
            Self::Bps1 => "0.0\\ \"bps\"",
            Self::Bps2 => "0.00\\ \"bps\"",
            Self::UsdMillions1 => "[$$-409]#,##0.0,,\\ \"mm\"",
            Self::UsdMillions2 => "[$$-409]#,##0.00,,\\ \"mm\"",
            Self::UsdBillions2 => "[$$-409]#,##0.00,,,\\ \"bn\"",
            Self::Pct1Signed => "+0.0%;-0.0%;0.0%",
            Self::Pct2Signed => "+0.00%;-0.00%;0.00%",
            Self::Multiple2 => "0.00x",
            Self::IntegerThousands => "#,##0",
        }
    }
}

impl Style {
    pub fn set_finance_format(&mut self, format: FinFormat) -> &mut Self {
        self.set_custom_format(format.format_code())
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MeasureType {
    Currency,
    Percentage,
    Bps,
    Multiple,
    Number,
}

impl MeasureType {
    fn default_format(self) -> FinFormat {
        match self {
            Self::Currency => FinFormat::UsdMillions2,
            Self::Percentage => FinFormat::Pct2Signed,
            Self::Bps => FinFormat::Bps1,
            Self::Multiple => FinFormat::Multiple2,
            Self::Number => FinFormat::IntegerThousands,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct FinanceDimension {
    name: String,
    members: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct FinanceMeasure {
    name: String,
    measure_type: MeasureType,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct FinanceModel {
    name: String,
    dimensions: Vec<FinanceDimension>,
    measures: Vec<FinanceMeasure>,
    scenarios: Vec<String>,
}

pub struct FinanceModelBuilder<'a> {
    workbook: &'a mut Workbook,
    model: FinanceModel,
}

impl Workbook {
    pub fn finance_model(&mut self, name: impl Into<String>) -> FinanceModelBuilder<'_> {
        FinanceModelBuilder {
            workbook: self,
            model: FinanceModel {
                name: name.into(),
                dimensions: Vec::new(),
                measures: Vec::new(),
                scenarios: Vec::new(),
            },
        }
    }

    pub fn pivot_on(
        &mut self,
        target_sheet: impl Into<String>,
        name: impl Into<String>,
    ) -> WorkbookPivotBuilder<'_> {
        WorkbookPivotBuilder {
            workbook: self,
            target_sheet: target_sheet.into(),
            name: name.into(),
            source: None,
            rows: Vec::new(),
            cols: Vec::new(),
            filters: Vec::new(),
            values: Vec::new(),
        }
    }
}

impl<'a> FinanceModelBuilder<'a> {
    pub fn dimension<I, S>(mut self, name: impl Into<String>, members: I) -> Self
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>,
    {
        self.model.dimensions.push(FinanceDimension {
            name: name.into(),
            members: members
                .into_iter()
                .map(|m| m.as_ref().to_string())
                .collect(),
        });
        self
    }

    pub fn measure(mut self, name: impl Into<String>, measure_type: MeasureType) -> Self {
        self.model.measures.push(FinanceMeasure {
            name: name.into(),
            measure_type,
        });
        self
    }

    pub fn scenario(mut self, name: impl Into<String>) -> Self {
        self.model.scenarios.push(name.into());
        self
    }

    pub fn build(self) -> Result<&'a mut Workbook> {
        if self.model.dimensions.is_empty() {
            return Err(XlsxError::InvalidWorkbookState(
                "finance model requires at least one dimension".to_string(),
            ));
        }
        if self.model.measures.is_empty() {
            return Err(XlsxError::InvalidWorkbookState(
                "finance model requires at least one measure".to_string(),
            ));
        }

        let meta_sheet_name = format!("{} Model", self.model.name);
        let data_sheet_name = format!("{} Data", self.model.name);

        if self.workbook.contains_sheet(meta_sheet_name.as_str())
            || self.workbook.contains_sheet(data_sheet_name.as_str())
        {
            return Err(XlsxError::InvalidWorkbookState(format!(
                "finance model sheets already exist for '{}'",
                self.model.name
            )));
        }

        {
            let ws = self.workbook.add_sheet(meta_sheet_name.as_str());
            ws.cell_mut("A1")?.set_value("Finance Model");
            ws.cell_mut("B1")?.set_value(self.model.name.as_str());

            ws.cell_mut("A3")?.set_value("Dimensions");
            ws.cell_mut("B3")?.set_value("Members");
            for (i, dimension) in self.model.dimensions.iter().enumerate() {
                let row = 4_u32
                    + u32::try_from(i).map_err(|_| {
                        XlsxError::InvalidWorkbookState(
                            "too many dimensions for finance model".to_string(),
                        )
                    })?;
                ws.cell_mut(build_cell_reference(1, row)?.as_str())?
                    .set_value(dimension.name.as_str());
                ws.cell_mut(build_cell_reference(2, row)?.as_str())?
                    .set_value(dimension.members.join(", "));
            }

            ws.cell_mut("D3")?.set_value("Measures");
            ws.cell_mut("E3")?.set_value("Type");
            ws.cell_mut("F3")?.set_value("Default Format");
            for (i, measure) in self.model.measures.iter().enumerate() {
                let row = 4_u32
                    + u32::try_from(i).map_err(|_| {
                        XlsxError::InvalidWorkbookState(
                            "too many measures for finance model".to_string(),
                        )
                    })?;
                ws.cell_mut(build_cell_reference(4, row)?.as_str())?
                    .set_value(measure.name.as_str());
                ws.cell_mut(build_cell_reference(5, row)?.as_str())?
                    .set_value(format!("{:?}", measure.measure_type));
                ws.cell_mut(build_cell_reference(6, row)?.as_str())?
                    .set_value(measure.measure_type.default_format().format_code());
            }

            ws.cell_mut("H3")?.set_value("Scenarios");
            for (i, scenario) in self.model.scenarios.iter().enumerate() {
                let row = 4_u32
                    + u32::try_from(i).map_err(|_| {
                        XlsxError::InvalidWorkbookState(
                            "too many scenarios for finance model".to_string(),
                        )
                    })?;
                ws.cell_mut(build_cell_reference(8, row)?.as_str())?
                    .set_value(scenario.as_str());
            }
        }

        {
            let ws = self.workbook.add_sheet(data_sheet_name.as_str());
            let mut headers: Vec<String> = self
                .model
                .dimensions
                .iter()
                .map(|d| d.name.clone())
                .collect();
            headers.push("Scenario".to_string());
            headers.extend(self.model.measures.iter().map(|m| m.name.clone()));

            for (i, header) in headers.iter().enumerate() {
                let col = 1_u32
                    .checked_add(u32::try_from(i).map_err(|_| {
                        XlsxError::InvalidWorkbookState(
                            "too many finance model headers".to_string(),
                        )
                    })?)
                    .ok_or_else(|| {
                        XlsxError::InvalidWorkbookState(
                            "finance model header column overflow".to_string(),
                        )
                    })?;
                ws.cell_mut(build_cell_reference(col, 1)?.as_str())?
                    .set_value(header.as_str());
            }

            let scenarios = if self.model.scenarios.is_empty() {
                vec!["base".to_string()]
            } else {
                self.model.scenarios.clone()
            };

            let dimension_vectors: Vec<Vec<String>> = self
                .model
                .dimensions
                .iter()
                .map(|d| {
                    if d.members.is_empty() {
                        vec!["n/a".to_string()]
                    } else {
                        d.members.clone()
                    }
                })
                .collect();
            let points = cartesian_product(dimension_vectors.as_slice());

            let mut row = 2_u32;
            for point in points {
                for scenario in &scenarios {
                    for (i, value) in point.iter().enumerate() {
                        let col = 1_u32
                            .checked_add(u32::try_from(i).map_err(|_| {
                                XlsxError::InvalidWorkbookState(
                                    "finance model point column overflow".to_string(),
                                )
                            })?)
                            .ok_or_else(|| {
                                XlsxError::InvalidWorkbookState(
                                    "finance model point column overflow".to_string(),
                                )
                            })?;
                        ws.cell_mut(build_cell_reference(col, row)?.as_str())?
                            .set_value(value.as_str());
                    }

                    let scenario_col = 1_u32
                        .checked_add(u32::try_from(self.model.dimensions.len()).map_err(|_| {
                            XlsxError::InvalidWorkbookState(
                                "finance model scenario column overflow".to_string(),
                            )
                        })?)
                        .ok_or_else(|| {
                            XlsxError::InvalidWorkbookState(
                                "finance model scenario column overflow".to_string(),
                            )
                        })?;
                    ws.cell_mut(build_cell_reference(scenario_col, row)?.as_str())?
                        .set_value(scenario.as_str());
                    row = row.checked_add(1).ok_or_else(|| {
                        XlsxError::InvalidWorkbookState("finance model row overflow".to_string())
                    })?;
                }
            }
        }

        Ok(self.workbook)
    }
}

pub struct WorkbookPivotBuilder<'a> {
    workbook: &'a mut Workbook,
    target_sheet: String,
    name: String,
    source: Option<String>,
    rows: Vec<String>,
    cols: Vec<String>,
    filters: Vec<String>,
    values: Vec<PivotValueSpec>,
}

impl<'a> WorkbookPivotBuilder<'a> {
    pub fn source(mut self, source: impl Into<String>) -> Self {
        self.source = Some(source.into());
        self
    }

    pub fn rows<I, S>(mut self, fields: I) -> Self
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>,
    {
        self.rows = fields
            .into_iter()
            .map(|field| field.as_ref().to_string())
            .collect();
        self
    }

    pub fn cols<I, S>(mut self, fields: I) -> Self
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>,
    {
        self.cols = fields
            .into_iter()
            .map(|field| field.as_ref().to_string())
            .collect();
        self
    }

    pub fn filters<I, S>(mut self, fields: I) -> Self
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>,
    {
        self.filters = fields
            .into_iter()
            .map(|field| field.as_ref().to_string())
            .collect();
        self
    }

    pub fn values<I>(mut self, values: I) -> Self
    where
        I: IntoIterator<Item = PivotValueSpec>,
    {
        self.values = values.into_iter().collect();
        self
    }

    pub fn validate_fields(self) -> Result<Self> {
        self.validate_internal()?;
        Ok(self)
    }

    pub fn place(self, target_cell: &str) -> Result<&'a mut Workbook> {
        self.validate_internal()?;
        let (target_col, target_row) = parse_cell_reference(target_cell)?;

        let mut pivot = PivotTable::new(
            self.name,
            PivotSourceReference::from_range(self.source.unwrap_or_default()),
        );
        pivot.set_target(target_row.saturating_sub(1), target_col.saturating_sub(1));

        for row_field in self.rows {
            let mut field = PivotField::new(row_field);
            field.set_sort_type(PivotFieldSort::Ascending);
            pivot.add_row_field(field);
        }
        for col_field in self.cols {
            let mut field = PivotField::new(col_field);
            field.set_sort_type(PivotFieldSort::Ascending);
            pivot.add_column_field(field);
        }
        for filter_field in self.filters {
            pivot.add_page_field(PivotField::new(filter_field));
        }
        for value in self.values {
            let mut data_field = PivotDataField::new(value.field_name().to_string());
            data_field.set_subtotal(value.subtotal());
            if let Some(name) = value.custom_name() {
                data_field.set_custom_name(name.to_string());
            }
            pivot.add_data_field(data_field);
        }

        let target_sheet = self.target_sheet.clone();
        let ws = self
            .workbook
            .sheet_mut(target_sheet.as_str())
            .ok_or_else(|| {
                XlsxError::InvalidWorkbookState(format!(
                    "target sheet '{}' does not exist",
                    target_sheet
                ))
            })?;
        ws.add_pivot_table(pivot);
        Ok(self.workbook)
    }

    fn validate_internal(&self) -> Result<()> {
        if !self.workbook.contains_sheet(self.target_sheet.as_str()) {
            return Err(XlsxError::InvalidWorkbookState(format!(
                "target sheet '{}' does not exist",
                self.target_sheet
            )));
        }

        let source = self.source.as_deref().unwrap_or("").trim();
        if source.is_empty() {
            return Err(XlsxError::InvalidWorkbookState(
                "pivot source is required".to_string(),
            ));
        }
        if self.values.is_empty() {
            return Err(XlsxError::InvalidWorkbookState(
                "pivot must include at least one value field".to_string(),
            ));
        }

        let mut seen = std::collections::HashSet::new();
        for field in self
            .rows
            .iter()
            .chain(self.cols.iter())
            .chain(self.filters.iter())
        {
            if !seen.insert(field.as_str()) {
                return Err(XlsxError::InvalidWorkbookState(format!(
                    "pivot field '{field}' appears more than once across rows/cols/filters"
                )));
            }
        }

        let (source_sheet_opt, source_range) = split_source_reference(source);
        let source_sheet = source_sheet_opt
            .map(normalize_source_sheet_name)
            .unwrap_or_else(|| self.target_sheet.clone());

        let source_ws = self.workbook.sheet(source_sheet.as_str()).ok_or_else(|| {
            XlsxError::InvalidWorkbookState(format!(
                "pivot source sheet '{}' does not exist",
                source_sheet
            ))
        })?;
        let headers = source_ws.source_headers_from_range(source_range)?;

        for field in self
            .rows
            .iter()
            .map(|value| value.as_str())
            .chain(self.cols.iter().map(|value| value.as_str()))
            .chain(self.filters.iter().map(|value| value.as_str()))
            .chain(self.values.iter().map(PivotValueSpec::field_name))
        {
            if !headers.iter().any(|header| header == field) {
                return Err(XlsxError::InvalidWorkbookState(format!(
                    "pivot field '{field}' not found in source header row"
                )));
            }
        }

        Ok(())
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartTemplate {
    PnlCurve,
    DrawdownCurve,
    ExposureBars,
    FactorContribStacked,
    WaterfallBridge,
}

pub struct FinanceChartTemplateBuilder<'a> {
    worksheet: &'a mut Worksheet,
    template: ChartTemplate,
    title: Option<String>,
    x: Option<String>,
    y: Option<String>,
    secondary_y: Option<String>,
}

impl Worksheet {
    pub fn chart_template(&mut self, template: ChartTemplate) -> FinanceChartTemplateBuilder<'_> {
        FinanceChartTemplateBuilder {
            worksheet: self,
            template,
            title: None,
            x: None,
            y: None,
            secondary_y: None,
        }
    }
}

impl<'a> FinanceChartTemplateBuilder<'a> {
    pub fn title(mut self, value: impl Into<String>) -> Self {
        self.title = Some(value.into());
        self
    }

    pub fn x(mut self, formula: impl Into<String>) -> Self {
        self.x = Some(formula.into());
        self
    }

    pub fn y(mut self, formula: impl Into<String>) -> Self {
        self.y = Some(formula.into());
        self
    }

    pub fn secondary_y(mut self, formula: impl Into<String>) -> Self {
        self.secondary_y = Some(formula.into());
        self
    }

    pub fn place(self, anchor: &str) -> Result<&'a mut Worksheet> {
        let x_formula = self.x.ok_or_else(|| {
            XlsxError::InvalidWorkbookState("chart template requires x range".to_string())
        })?;
        let y_formula = self.y.ok_or_else(|| {
            XlsxError::InvalidWorkbookState("chart template requires y range".to_string())
        })?;
        let (start_ref, end_ref) = split_anchor(anchor)?;
        let (from_col, from_row) = parse_cell_reference(start_ref)?;
        let (to_col, to_row) = parse_cell_reference(end_ref)?;

        let mut chart = match self.template {
            ChartTemplate::PnlCurve => {
                let mut c = Chart::new(ChartType::Line);
                c.set_grouping(ChartGrouping::Standard);
                c
            }
            ChartTemplate::DrawdownCurve => {
                let mut c = Chart::new(ChartType::Area);
                c.set_grouping(ChartGrouping::Standard);
                c
            }
            ChartTemplate::ExposureBars => {
                let mut c = Chart::new(ChartType::Bar);
                c.set_bar_direction(BarDirection::Column);
                c.set_grouping(ChartGrouping::Clustered);
                c
            }
            ChartTemplate::FactorContribStacked => {
                let mut c = Chart::new(ChartType::Bar);
                c.set_bar_direction(BarDirection::Column);
                c.set_grouping(ChartGrouping::Stacked);
                c
            }
            ChartTemplate::WaterfallBridge => {
                let mut c = Chart::new(ChartType::Bar);
                c.set_bar_direction(BarDirection::Column);
                c.set_grouping(ChartGrouping::Clustered);
                c.set_vary_colors(true);
                c
            }
        };

        if let Some(title) = self.title {
            chart.set_title(title);
        }

        let mut primary_series = ChartSeries::new(0, 0);
        primary_series.set_name("Primary");
        primary_series.set_categories(crate::chart::ChartDataRef::from_formula(x_formula));
        primary_series.set_values(crate::chart::ChartDataRef::from_formula(y_formula));
        chart.add_series(primary_series);

        if let Some(secondary) = self.secondary_y {
            let mut secondary_series = ChartSeries::new(1, 1);
            secondary_series.set_name("Secondary");
            secondary_series.set_categories(crate::chart::ChartDataRef::from_formula(
                chart
                    .series()
                    .first()
                    .and_then(|s| s.categories())
                    .and_then(|c| c.formula())
                    .unwrap_or_default(),
            ));
            secondary_series.set_values(crate::chart::ChartDataRef::from_formula(secondary));
            secondary_series.set_series_type(ChartType::Line);
            chart.add_series(secondary_series);
        }

        chart.add_axis(ChartAxis::new_category());
        chart.add_axis(ChartAxis::new_value());
        chart.set_anchor(
            from_col.saturating_sub(1),
            from_row.saturating_sub(1),
            to_col.saturating_sub(1),
            to_row.saturating_sub(1),
        );

        self.worksheet.add_chart(chart);
        Ok(self.worksheet)
    }
}

fn split_source_reference(source: &str) -> (Option<&str>, &str) {
    if let Some(pos) = source.find('!') {
        (Some(&source[..pos]), &source[pos + 1..])
    } else {
        (None, source)
    }
}

fn normalize_source_sheet_name(name: &str) -> String {
    let trimmed = name.trim();
    if trimmed.len() >= 2 && trimmed.starts_with('\'') && trimmed.ends_with('\'') {
        trimmed[1..trimmed.len() - 1].replace("''", "'")
    } else {
        trimmed.to_string()
    }
}

fn split_anchor(anchor: &str) -> Result<(&str, &str)> {
    let mut parts = anchor.split(':');
    let start = parts
        .next()
        .ok_or_else(|| XlsxError::InvalidWorkbookState("chart anchor is invalid".to_string()))?;
    let end = parts
        .next()
        .ok_or_else(|| XlsxError::InvalidWorkbookState("chart anchor is invalid".to_string()))?;
    if parts.next().is_some() {
        return Err(XlsxError::InvalidWorkbookState(
            "chart anchor is invalid".to_string(),
        ));
    }
    Ok((start, end))
}

fn parse_cell_reference(reference: &str) -> Result<(u32, u32)> {
    let normalized = crate::cell::normalize_cell_reference(reference)?;
    let split_index = normalized
        .char_indices()
        .find_map(|(index, ch)| ch.is_ascii_digit().then_some(index))
        .ok_or_else(|| XlsxError::InvalidCellReference(reference.to_string()))?;
    let (column_name, row_text) = normalized.split_at(split_index);
    let column_index = column_name
        .bytes()
        .try_fold(0_u32, |acc, byte| {
            acc.checked_mul(26)
                .and_then(|value| value.checked_add(u32::from(byte - b'A' + 1)))
        })
        .ok_or_else(|| XlsxError::InvalidCellReference(reference.to_string()))?;
    let row_index = row_text
        .parse::<u32>()
        .map_err(|_| XlsxError::InvalidCellReference(reference.to_string()))?;
    Ok((column_index, row_index))
}

fn build_cell_reference(column: u32, row: u32) -> Result<String> {
    if column == 0 || row == 0 {
        return Err(XlsxError::InvalidCellReference(format!(
            "{}{}",
            column, row
        )));
    }
    let mut current = column;
    let mut label = String::new();
    while current > 0 {
        let remainder = (current - 1) % 26;
        label.push(char::from(b'A' + remainder as u8));
        current = (current - 1) / 26;
    }
    let column_label: String = label.chars().rev().collect();
    Ok(format!("{}{}", column_label, row))
}

fn cartesian_product(groups: &[Vec<String>]) -> Vec<Vec<String>> {
    if groups.is_empty() {
        return Vec::new();
    }

    let mut acc: Vec<Vec<String>> = vec![Vec::new()];
    for group in groups {
        let mut next: Vec<Vec<String>> = Vec::new();
        for prefix in &acc {
            for member in group {
                let mut combined = prefix.clone();
                combined.push(member.clone());
                next.push(combined);
            }
        }
        acc = next;
    }
    acc
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::CellValue;
    use crate::{avg, sum, Workbook};

    #[test]
    fn finance_format_mapping_is_stable() {
        assert_eq!(FinFormat::Bps1.format_code(), "0.0\\ \"bps\"");
        assert_eq!(
            FinFormat::UsdMillions2.format_code(),
            "[$$-409]#,##0.00,,\\ \"mm\""
        );
        assert_eq!(FinFormat::Pct2Signed.format_code(), "+0.00%;-0.00%;0.00%");
    }

    #[test]
    fn finance_model_builder_creates_metadata_and_data_sheets() {
        let mut wb = Workbook::new();
        wb.finance_model("Core")
            .dimension("Region", ["US", "EMEA"])
            .dimension("Quarter", ["2025-Q1", "2025-Q2"])
            .measure("Revenue", MeasureType::Currency)
            .measure("Margin", MeasureType::Percentage)
            .scenario("base")
            .scenario("stress")
            .build()
            .expect("build finance model");

        assert!(wb.contains_sheet("Core Model"));
        assert!(wb.contains_sheet("Core Data"));
        assert_eq!(
            wb.sheet("Core Data")
                .and_then(|ws| ws.cell("A2"))
                .and_then(|cell| cell.value()),
            Some(&CellValue::String("US".to_string()))
        );
    }

    #[test]
    fn workbook_pivot_builder_validates_cross_sheet_headers() {
        let mut wb = Workbook::new();
        let data = wb.add_sheet("Data");
        data.cell_mut("A1").unwrap().set_value("Region");
        data.cell_mut("B1").unwrap().set_value("Quarter");
        data.cell_mut("C1").unwrap().set_value("Revenue");

        wb.add_sheet("Pivot");

        wb.pivot_on("Pivot", "CrossSheet")
            .source("Data!A1:C10")
            .rows(["Region"])
            .cols(["Quarter"])
            .values([sum("Revenue").name("Rev"), avg("Revenue")])
            .validate_fields()
            .expect("validation should pass")
            .place("A4")
            .expect("pivot placement should pass");

        let err = wb
            .pivot_on("Pivot", "BadCrossSheet")
            .source("Data!A1:C10")
            .rows(["Desk"])
            .values([sum("Revenue")])
            .validate_fields()
            .err()
            .expect("unknown field should fail");

        assert!(err.to_string().contains("Desk"));
    }

    #[test]
    fn chart_template_places_chart() {
        let mut wb = Workbook::new();
        let ws = wb.add_sheet("Dashboard");
        ws.chart_template(ChartTemplate::PnlCurve)
            .title("NAV")
            .x("Summary!$A$2:$A$10")
            .y("Summary!$B$2:$B$10")
            .secondary_y("Summary!$C$2:$C$10")
            .place("D2:L20")
            .expect("chart template should place chart");

        let chart = ws.charts().first().expect("chart exists");
        assert_eq!(chart.chart_type(), ChartType::Line);
        assert_eq!(chart.series().len(), 2);
        assert_eq!(chart.from_col(), 3);
        assert_eq!(chart.from_row(), 1);
        assert_eq!(chart.to_col(), 11);
        assert_eq!(chart.to_row(), 19);
    }
}
