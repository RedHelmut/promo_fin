mod missing_report;
mod pdf;
use std::fs::File;


fn main() {
    //let file = r#"F:\3M Promo Data\May1-July31 2020\ViewExport Customer Detail.xlsx"#;
    let file = r##"F:\3M Promo Data\May1-July31 2020\data.csv"##;
    let json = r##"F:\3M Promo Data\May1-July31 2020\promo_May1_2020-July31_2020.json"##;

    let mut file_r = File::create("Missing Report.pdf");
    if let Ok(ref mut fl) = file_r {
        missing_report::run_missing_reports(file, json, Some(fl), r#"F:\rustprojects\promo_fin\promo.zip"#).expect("Fail");
    }

}
