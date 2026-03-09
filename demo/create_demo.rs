//! Build a financial-grade earnings workbook showcasing high-fidelity XLSX features.

use offidized_xlsx::{
    Alignment, BarDirection, Border, BorderSide, CellValue, Chart, ChartAxis, ChartDataRef,
    ChartGrouping, ChartSeries, ChartType, ConditionalFormatting,
    ConditionalFormattingOperator, Fill, Font, HorizontalAlignment, PivotDataField, PivotField,
    PivotFieldSort, PivotSourceReference, PivotSubtotalFunction, PivotTable, Sparkline,
    SparklineColors, SparklineGroup, SparklineType, Style, TotalFunction, VerticalAlignment,
    Workbook, WorksheetTable,
};

struct StyleIds {
    title: u32,
    header: u32,
    currency: u32,
    percent: u32,
    note: u32,
    metric: u32,
}

#[derive(Debug, Clone, Copy)]
struct QuarterMetrics {
    revenue: f64,
    ebitda: f64,
    margin: f64,
    fcf: f64,
    arr: f64,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Building financial-grade earnings showcase workbook...");

    let mut workbook = Workbook::new();
    let styles = install_styles(&mut workbook)?;

    build_earnings_fact_sheet(&mut workbook, &styles)?;
    build_summary_sheet(&mut workbook, &styles)?;
    build_dashboard_sheet(&mut workbook, &styles)?;
    build_pivot_sheet(&mut workbook, &styles)?;

    let output_path = "artifacts/earnings_showcase.xlsx";
    workbook.save(output_path)?;

    println!("Saved {}", output_path);
    println!();
    println!("Power features in this workbook:");
    println!("1. Styled financial model with durable number formats and layout.");
    println!("2. Native worksheet table + conditional formatting + freeze panes.");
    println!("3. Two dashboard charts with explicit axis/series metadata.");
    println!("4. Real pivot table with row/column/page/data fields.");
    println!("5. Sparklines and formula-driven KPI surfaces.");
    println!();
    println!("Suggested verification commands:");
    println!("  ofx pivots artifacts/earnings_showcase.xlsx");
    println!("  ofx charts artifacts/earnings_showcase.xlsx");
    println!("  ofx derive artifacts/earnings_showcase.xlsx --mode full > /tmp/earnings.ir");

    Ok(())
}

fn compute_quarter_metrics() -> Vec<QuarterMetrics> {
    let mut out = Vec::with_capacity(8);

    for q_idx in 0..8 {
        let mut revenue_sum = 0.0;
        let mut ebitda_sum = 0.0;
        let mut fcf_sum = 0.0;
        let mut arr_sum = 0.0;

        for r_idx in 0..4 {
            for s_idx in 0..3 {
                let seasonality = [1.00_f64, 1.04, 1.09, 1.22][q_idx % 4];
                let base = 16_000_000.0
                    + (q_idx as f64 * 1_900_000.0)
                    + (r_idx as f64 * 1_500_000.0)
                    + (s_idx as f64 * 1_100_000.0);
                let revenue = base * seasonality;
                let cogs = revenue * (0.40 + s_idx as f64 * 0.03 + r_idx as f64 * 0.01);
                let opex = revenue * (0.19 + (q_idx % 4) as f64 * 0.01);
                let ebitda = revenue - cogs - opex;
                let fcf = ebitda * 0.67 - 220_000.0 + r_idx as f64 * 45_000.0;
                let arr = revenue * (1.55 + s_idx as f64 * 0.10);

                revenue_sum += revenue;
                ebitda_sum += ebitda;
                fcf_sum += fcf;
                arr_sum += arr;
            }
        }

        out.push(QuarterMetrics {
            revenue: revenue_sum,
            ebitda: ebitda_sum,
            margin: if revenue_sum == 0.0 {
                0.0
            } else {
                ebitda_sum / revenue_sum
            },
            fcf: fcf_sum,
            arr: arr_sum,
        });
    }

    out
}

fn install_styles(workbook: &mut Workbook) -> offidized_xlsx::Result<StyleIds> {
    let mut title_font = Font::new();
    title_font
        .set_name("Aptos Display")
        .set_size("16")
        .set_bold(true)
        .set_color("FF0F172A");

    let mut title_align = Alignment::new();
    title_align
        .set_horizontal(HorizontalAlignment::Left)
        .set_vertical(VerticalAlignment::Center);

    let mut title_style = Style::new();
    title_style.set_font(title_font).set_alignment(title_align);
    let title = workbook.add_style(title_style)?;

    let mut header_font = Font::new();
    header_font
        .set_name("Aptos")
        .set_size("11")
        .set_bold(true)
        .set_color("FFFFFFFF");

    let mut header_fill = Fill::new();
    header_fill
        .set_pattern("solid")
        .set_foreground_color("FF1E3A8A")
        .set_background_color("FF1E3A8A");

    let mut header_border_side = BorderSide::new();
    header_border_side
        .set_style("thin")
        .set_color("FF94A3B8");
    let mut header_border = Border::new();
    header_border
        .set_left(header_border_side.clone())
        .set_right(header_border_side.clone())
        .set_top(header_border_side.clone())
        .set_bottom(header_border_side);

    let mut header_align = Alignment::new();
    header_align
        .set_horizontal(HorizontalAlignment::Center)
        .set_vertical(VerticalAlignment::Center)
        .set_wrap_text(true);

    let mut header_style = Style::new();
    header_style
        .set_font(header_font)
        .set_fill(header_fill)
        .set_border(header_border)
        .set_alignment(header_align);
    let header = workbook.add_style(header_style)?;

    let mut currency_align = Alignment::new();
    currency_align.set_horizontal(HorizontalAlignment::Right);
    let mut currency_style = Style::new();
    currency_style
        .set_number_format("$#,##0;[Red]($#,##0)")
        .set_alignment(currency_align);
    let currency = workbook.add_style(currency_style)?;

    let mut percent_align = Alignment::new();
    percent_align.set_horizontal(HorizontalAlignment::Right);
    let mut percent_style = Style::new();
    percent_style
        .set_number_format("0.0%")
        .set_alignment(percent_align);
    let percent = workbook.add_style(percent_style)?;

    let mut note_font = Font::new();
    note_font
        .set_name("Aptos")
        .set_size("10")
        .set_color("FF475569")
        .set_italic(true);
    let mut note_style = Style::new();
    note_style.set_font(note_font);
    let note = workbook.add_style(note_style)?;

    let mut metric_font = Font::new();
    metric_font
        .set_name("Aptos")
        .set_size("12")
        .set_bold(true)
        .set_color("FF0F172A");

    let mut metric_fill = Fill::new();
    metric_fill
        .set_pattern("solid")
        .set_foreground_color("FFE2E8F0")
        .set_background_color("FFE2E8F0");

    let mut metric_style = Style::new();
    metric_style
        .set_font(metric_font)
        .set_fill(metric_fill)
        .set_alignment(Alignment::new().set_horizontal(HorizontalAlignment::Right).clone())
        .set_number_format("$#,##0;[Red]($#,##0)");
    let metric = workbook.add_style(metric_style)?;

    Ok(StyleIds {
        title,
        header,
        currency,
        percent,
        note,
        metric,
    })
}

fn build_earnings_fact_sheet(
    workbook: &mut Workbook,
    styles: &StyleIds,
) -> Result<(), Box<dyn std::error::Error>> {
    let ws = workbook.add_sheet("Earnings Fact");
    ws.set_tab_color("1E3A8A");
    ws.set_default_column_width(14.0);
    ws.set_freeze_panes(0, 1)?;

    ws.cell_mut("A1")?.set_value("Period").set_style_id(styles.header);
    ws.cell_mut("B1")?.set_value("Quarter").set_style_id(styles.header);
    ws.cell_mut("C1")?.set_value("Region").set_style_id(styles.header);
    ws.cell_mut("D1")?.set_value("Segment").set_style_id(styles.header);
    ws.cell_mut("E1")?.set_value("Product").set_style_id(styles.header);
    ws.cell_mut("F1")?.set_value("Revenue").set_style_id(styles.header);
    ws.cell_mut("G1")?.set_value("COGS").set_style_id(styles.header);
    ws.cell_mut("H1")?.set_value("Operating Expense").set_style_id(styles.header);
    ws.cell_mut("I1")?.set_value("EBITDA").set_style_id(styles.header);
    ws.cell_mut("J1")?.set_value("EBITDA Margin").set_style_id(styles.header);
    ws.cell_mut("K1")?.set_value("Free Cash Flow").set_style_id(styles.header);
    ws.cell_mut("L1")?.set_value("ARR").set_style_id(styles.header);

    let quarters = [
        "2024-Q1", "2024-Q2", "2024-Q3", "2024-Q4", "2025-Q1", "2025-Q2", "2025-Q3",
        "2025-Q4",
    ];
    let regions = ["North America", "EMEA", "APAC", "LATAM"];
    let segments = ["Enterprise", "Mid-Market", "SMB"];

    let mut row = 2_u32;
    for (q_idx, quarter) in quarters.iter().enumerate() {
        for (r_idx, region) in regions.iter().enumerate() {
            for (s_idx, segment) in segments.iter().enumerate() {
                let product = match *segment {
                    "Enterprise" => "Platform Suite",
                    "Mid-Market" => "Growth Cloud",
                    _ => "Starter Stack",
                };

                let seasonality = [1.00_f64, 1.04, 1.09, 1.22][q_idx % 4];
                let base = 16_000_000.0
                    + (q_idx as f64 * 1_900_000.0)
                    + (r_idx as f64 * 1_500_000.0)
                    + (s_idx as f64 * 1_100_000.0);
                let revenue = base * seasonality;
                let cogs = revenue * (0.40 + s_idx as f64 * 0.03 + r_idx as f64 * 0.01);
                let opex = revenue * (0.19 + (q_idx % 4) as f64 * 0.01);
                let ebitda = revenue - cogs - opex;
                let margin = ebitda / revenue;
                let fcf = ebitda * 0.67 - 220_000.0 + r_idx as f64 * 45_000.0;
                let arr = revenue * (1.55 + s_idx as f64 * 0.10);

                ws.cell_mut(&format!("A{}", row))?
                    .set_value(format!("FY{}", 2024 + q_idx / 4));
                ws.cell_mut(&format!("B{}", row))?.set_value(*quarter);
                ws.cell_mut(&format!("C{}", row))?.set_value(*region);
                ws.cell_mut(&format!("D{}", row))?.set_value(*segment);
                ws.cell_mut(&format!("E{}", row))?.set_value(product);

                ws.cell_mut(&format!("F{}", row))?
                    .set_value(revenue)
                    .set_style_id(styles.currency);
                ws.cell_mut(&format!("G{}", row))?
                    .set_value(cogs)
                    .set_style_id(styles.currency);
                ws.cell_mut(&format!("H{}", row))?
                    .set_value(opex)
                    .set_style_id(styles.currency);
                ws.cell_mut(&format!("I{}", row))?
                    .set_value(ebitda)
                    .set_style_id(styles.currency);
                ws.cell_mut(&format!("J{}", row))?
                    .set_value(margin)
                    .set_style_id(styles.percent);
                ws.cell_mut(&format!("K{}", row))?
                    .set_value(fcf)
                    .set_style_id(styles.currency);
                ws.cell_mut(&format!("L{}", row))?
                    .set_value(arr)
                    .set_style_id(styles.currency);

                row += 1;
            }
        }
    }

    let mut table = WorksheetTable::new("EarningsFactTable", "A1:L97")?;
    table
        .set_style_name("TableStyleMedium2")
        .set_show_row_stripes(true)
        .set_show_first_column(true);

    let columns = [
        "Period",
        "Quarter",
        "Region",
        "Segment",
        "Product",
        "Revenue",
        "COGS",
        "Operating Expense",
        "EBITDA",
        "EBITDA Margin",
        "Free Cash Flow",
        "ARR",
    ];
    for column in columns {
        let col = table.add_column(column);
        if matches!(
            column,
            "Revenue" | "COGS" | "Operating Expense" | "EBITDA" | "Free Cash Flow" | "ARR"
        ) {
            col.set_totals_row_function(TotalFunction::Sum);
        }
    }
    ws.add_table(table);

    let mut low_margin_cf = ConditionalFormatting::cell_is(["J2:J97"], ["0.18"])?;
    low_margin_cf.set_operator(ConditionalFormattingOperator::LessThan);
    ws.add_conditional_formatting(low_margin_cf);

    Ok(())
}

fn build_summary_sheet(
    workbook: &mut Workbook,
    styles: &StyleIds,
) -> Result<(), Box<dyn std::error::Error>> {
    let ws = workbook.add_sheet("Summary");
    ws.set_tab_color("0F766E");
    ws.set_default_column_width(16.0);
    ws.set_freeze_panes(0, 1)?;

    ws.cell_mut("A1")?
        .set_value("Quarterly Earnings Summary")
        .set_style_id(styles.title);
    ws.cell_mut("A2")?.set_value("Quarter").set_style_id(styles.header);
    ws.cell_mut("B2")?.set_value("Revenue").set_style_id(styles.header);
    ws.cell_mut("C2")?.set_value("EBITDA").set_style_id(styles.header);
    ws.cell_mut("D2")?
        .set_value("EBITDA Margin")
        .set_style_id(styles.header);
    ws.cell_mut("E2")?
        .set_value("Free Cash Flow")
        .set_style_id(styles.header);
    ws.cell_mut("F2")?.set_value("ARR").set_style_id(styles.header);

    let quarters = [
        "2024-Q1", "2024-Q2", "2024-Q3", "2024-Q4", "2025-Q1", "2025-Q2", "2025-Q3",
        "2025-Q4",
    ];
    let metrics = compute_quarter_metrics();

    for (idx, quarter) in quarters.iter().enumerate() {
        let row = idx + 3;
        let m = metrics[idx];
        ws.cell_mut(&format!("A{}", row))?.set_value(*quarter);

        ws.cell_mut(&format!("B{}", row))?
            .set_formula(format!(
                "SUMIFS('Earnings Fact'!$F$2:$F$97,'Earnings Fact'!$B$2:$B$97,A{})",
                row
            ))
            .set_cached_value(CellValue::Number(m.revenue))
            .set_style_id(styles.currency);

        ws.cell_mut(&format!("C{}", row))?
            .set_formula(format!(
                "SUMIFS('Earnings Fact'!$I$2:$I$97,'Earnings Fact'!$B$2:$B$97,A{})",
                row
            ))
            .set_cached_value(CellValue::Number(m.ebitda))
            .set_style_id(styles.currency);

        ws.cell_mut(&format!("D{}", row))?
            .set_formula(format!("IF(B{}=0,0,C{}/B{})", row, row, row))
            .set_cached_value(CellValue::Number(m.margin))
            .set_style_id(styles.percent);

        ws.cell_mut(&format!("E{}", row))?
            .set_formula(format!(
                "SUMIFS('Earnings Fact'!$K$2:$K$97,'Earnings Fact'!$B$2:$B$97,A{})",
                row
            ))
            .set_cached_value(CellValue::Number(m.fcf))
            .set_style_id(styles.currency);

        ws.cell_mut(&format!("F{}", row))?
            .set_formula(format!(
                "SUMIFS('Earnings Fact'!$L$2:$L$97,'Earnings Fact'!$B$2:$B$97,A{})",
                row
            ))
            .set_cached_value(CellValue::Number(m.arr))
            .set_style_id(styles.currency);
    }

    ws.cell_mut("H2")?.set_value("YoY KPI").set_style_id(styles.header);
    ws.cell_mut("I2")?.set_value("Value").set_style_id(styles.header);
    ws.cell_mut("H3")?.set_value("Revenue Growth");
    ws.cell_mut("H4")?.set_value("EBITDA Growth");
    ws.cell_mut("H5")?.set_value("FCF Growth");

    let rev_growth = if metrics[3].revenue == 0.0 {
        0.0
    } else {
        (metrics[7].revenue - metrics[3].revenue) / metrics[3].revenue
    };
    let ebitda_growth = if metrics[3].ebitda == 0.0 {
        0.0
    } else {
        (metrics[7].ebitda - metrics[3].ebitda) / metrics[3].ebitda
    };
    let fcf_growth = if metrics[3].fcf == 0.0 {
        0.0
    } else {
        (metrics[7].fcf - metrics[3].fcf) / metrics[3].fcf
    };

    ws.cell_mut("I3")?
        .set_formula("IF(B6=0,0,(B10-B6)/B6)")
        .set_cached_value(CellValue::Number(rev_growth))
        .set_style_id(styles.percent);
    ws.cell_mut("I4")?
        .set_formula("IF(C6=0,0,(C10-C6)/C6)")
        .set_cached_value(CellValue::Number(ebitda_growth))
        .set_style_id(styles.percent);
    ws.cell_mut("I5")?
        .set_formula("IF(E6=0,0,(E10-E6)/E6)")
        .set_cached_value(CellValue::Number(fcf_growth))
        .set_style_id(styles.percent);

    ws.cell_mut("A12")?
        .set_value("Formulas remain editable and auditable via IR/content workflows.")
        .set_style_id(styles.note);

    Ok(())
}

fn build_dashboard_sheet(
    workbook: &mut Workbook,
    styles: &StyleIds,
) -> Result<(), Box<dyn std::error::Error>> {
    let ws = workbook.add_sheet("Dashboard");
    ws.set_tab_color("7C2D12");
    ws.set_default_column_width(18.0);
    let metrics = compute_quarter_metrics();

    ws.cell_mut("A1")?
        .set_value("Executive Earnings Dashboard")
        .set_style_id(styles.title);

    ws.cell_mut("A3")?
        .set_value("Latest Quarter Revenue")
        .set_style_id(styles.header);
    ws.cell_mut("A4")?
        .set_value("Latest Quarter EBITDA")
        .set_style_id(styles.header);
    ws.cell_mut("A5")?
        .set_value("Latest EBITDA Margin")
        .set_style_id(styles.header);
    ws.cell_mut("A6")?
        .set_value("Latest Free Cash Flow")
        .set_style_id(styles.header);

    ws.cell_mut("B3")?
        .set_formula("Summary!B10")
        .set_cached_value(CellValue::Number(metrics[7].revenue))
        .set_style_id(styles.metric);
    ws.cell_mut("B4")?
        .set_formula("Summary!C10")
        .set_cached_value(CellValue::Number(metrics[7].ebitda))
        .set_style_id(styles.metric);
    ws.cell_mut("B5")?
        .set_formula("Summary!D10")
        .set_cached_value(CellValue::Number(metrics[7].margin))
        .set_style_id(styles.percent);
    ws.cell_mut("B6")?
        .set_formula("Summary!E10")
        .set_cached_value(CellValue::Number(metrics[7].fcf))
        .set_style_id(styles.metric);

    let mut mix_chart = Chart::new(ChartType::Bar);
    mix_chart.set_anchor(3, 2, 11, 18);
    mix_chart.set_title("Revenue vs EBITDA by Quarter");
    mix_chart.set_bar_direction(BarDirection::Column);
    mix_chart.set_grouping(ChartGrouping::Clustered);

    let quarter_labels: Vec<String> = vec![
        "2024-Q1".to_string(),
        "2024-Q2".to_string(),
        "2024-Q3".to_string(),
        "2024-Q4".to_string(),
        "2025-Q1".to_string(),
        "2025-Q2".to_string(),
        "2025-Q3".to_string(),
        "2025-Q4".to_string(),
    ];
    let revenue_vals: Vec<Option<f64>> = metrics.iter().map(|m| Some(m.revenue)).collect();
    let ebitda_vals: Vec<Option<f64>> = metrics.iter().map(|m| Some(m.ebitda)).collect();
    let margin_vals: Vec<Option<f64>> = metrics.iter().map(|m| Some(m.margin)).collect();

    let mut revenue_cat_ref = ChartDataRef::from_formula("'Summary'!$A$3:$A$10");
    revenue_cat_ref.set_str_values(quarter_labels.clone());
    let mut revenue_val_ref = ChartDataRef::from_formula("'Summary'!$B$3:$B$10");
    revenue_val_ref.set_num_values(revenue_vals);
    let mut revenue_series = ChartSeries::new(0, 0);
    revenue_series
        .set_name("Revenue")
        .set_categories(revenue_cat_ref)
        .set_values(revenue_val_ref)
        .set_fill_color("FF1D4ED8");
    mix_chart.add_series(revenue_series);

    let mut ebitda_cat_ref = ChartDataRef::from_formula("'Summary'!$A$3:$A$10");
    ebitda_cat_ref.set_str_values(quarter_labels.clone());
    let mut ebitda_val_ref = ChartDataRef::from_formula("'Summary'!$C$3:$C$10");
    ebitda_val_ref.set_num_values(ebitda_vals);
    let mut ebitda_series = ChartSeries::new(1, 1);
    ebitda_series
        .set_name("EBITDA")
        .set_categories(ebitda_cat_ref)
        .set_values(ebitda_val_ref)
        .set_fill_color("FF16A34A");
    mix_chart.add_series(ebitda_series);

    let mut rev_axis = ChartAxis::new_category();
    rev_axis.set_title("Quarter");
    mix_chart.add_axis(rev_axis);

    let mut value_axis = ChartAxis::new_value();
    value_axis.set_title("USD");
    value_axis.set_min(0.0);
    mix_chart.add_axis(value_axis);

    ws.add_chart(mix_chart);

    let mut margin_chart = Chart::new(ChartType::Line);
    margin_chart.set_anchor(12, 2, 20, 18);
    margin_chart.set_title("EBITDA Margin Trend");

    let mut margin_cat_ref = ChartDataRef::from_formula("'Summary'!$A$3:$A$10");
    margin_cat_ref.set_str_values(quarter_labels);
    let mut margin_val_ref = ChartDataRef::from_formula("'Summary'!$D$3:$D$10");
    margin_val_ref.set_num_values(margin_vals);
    let mut margin_series = ChartSeries::new(0, 0);
    margin_series
        .set_name("Margin")
        .set_categories(margin_cat_ref)
        .set_values(margin_val_ref)
        .set_line_color("FFEA580C");
    margin_chart.add_series(margin_series);
    margin_chart.add_axis(ChartAxis::new_category());

    let mut margin_axis = ChartAxis::new_value();
    margin_axis.set_title("Percent");
    margin_axis.set_min(0.0);
    margin_chart.add_axis(margin_axis);
    ws.add_chart(margin_chart);

    ws.cell_mut("A8")?
        .set_value("Sparkline strip")
        .set_style_id(styles.header);

    let mut sparkline_group = SparklineGroup::new();
    sparkline_group.set_sparkline_type(SparklineType::Line);
    let mut spark_colors = SparklineColors::new();
    spark_colors.series = Some("FF1D4ED8".to_string());
    spark_colors.high = Some("FF16A34A".to_string());
    spark_colors.low = Some("FFDC2626".to_string());
    sparkline_group.set_colors(spark_colors);

    sparkline_group.add_sparkline(Sparkline::new("B8", "Summary!B3:B10"));
    sparkline_group.add_sparkline(Sparkline::new("C8", "Summary!C3:C10"));
    sparkline_group.add_sparkline(Sparkline::new("D8", "Summary!D3:D10"));
    ws.add_sparkline_group(sparkline_group);

    Ok(())
}

fn build_pivot_sheet(
    workbook: &mut Workbook,
    styles: &StyleIds,
) -> Result<(), Box<dyn std::error::Error>> {
    let ws = workbook.add_sheet("Pivot Analytics");
    ws.set_tab_color("4338CA");
    ws.set_default_column_width(16.0);
    ws.set_freeze_panes(0, 3)?;

    ws.cell_mut("A1")?
        .set_value("Multi-dimensional Pivot (Region x Segment x Quarter)")
        .set_style_id(styles.title);
    ws.cell_mut("A2")?
        .set_value("Use this tab for slice-and-dice analysis with native pivot metadata.")
        .set_style_id(styles.note);

    let source = PivotSourceReference::from_range("Earnings Fact!$A$1:$L$97");
    let mut pivot = PivotTable::new("EarningsPivot", source);
    pivot.set_target(3, 0);
    pivot.set_show_row_grand_totals(true);
    pivot.set_show_column_grand_totals(true);

    let mut region = PivotField::new("Region");
    region.set_sort_type(PivotFieldSort::Ascending);
    pivot.add_row_field(region);

    let segment = PivotField::new("Segment");
    pivot.add_row_field(segment);

    let quarter = PivotField::new("Quarter");
    pivot.add_column_field(quarter);

    let product_filter = PivotField::new("Product");
    pivot.add_page_field(product_filter);

    let mut sum_revenue = PivotDataField::new("Revenue");
    sum_revenue
        .set_custom_name("Total Revenue")
        .set_subtotal(PivotSubtotalFunction::Sum);
    pivot.add_data_field(sum_revenue);

    let mut sum_ebitda = PivotDataField::new("EBITDA");
    sum_ebitda
        .set_custom_name("Total EBITDA")
        .set_subtotal(PivotSubtotalFunction::Sum);
    pivot.add_data_field(sum_ebitda);

    let mut avg_margin = PivotDataField::new("EBITDA Margin");
    avg_margin
        .set_custom_name("Avg EBITDA Margin")
        .set_subtotal(PivotSubtotalFunction::Average);
    pivot.add_data_field(avg_margin);

    ws.add_pivot_table(pivot);

    Ok(())
}
