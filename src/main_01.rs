use calamine::{open_workbook, DataType, Reader, Xlsx};
use dialoguer::Input;
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    // Prompt the user to enter the filename
    let filename: String = Input::new()
        .with_prompt("Enter the Excel file path")
        .interact()?;
    // opens a new workbook
    let mut workbook: Xlsx<_> = open_workbook(&filename)?;

    // Read whole worksheet data and print each cell
    if let Ok(range) = workbook.worksheet_range("Sheet1") {
        for row in range.rows() {
            for cell in row {
                match cell {
                    calamine::Data::Empty => print!("(Empty)\t"),
                    calamine::Data::String(s) => print!("{}\t", s),
                    calamine::Data::Float(f) => print!("{}\t", f),
                    calamine::Data::Int(i) => print!("{}\t", i),
                    calamine::Data::Bool(b) => print!("{}\t", b),
                    calamine::Data::Error(e) => print!("(Error: {:?})\t", e),
                    _ => print!("啥也不是"),
                }
            }
            println!(); // New line after each row
        }
    }

    // Check if the workbook has a vba project
    if let Some(Ok(mut vba)) = workbook.vba_project() {
        let vba = vba.to_mut();
        let module1 = vba.get_module("Module 1").unwrap();
        println!("Module 1 code:");
        println!("{}", module1);
        for r in vba.get_references() {
            if r.is_missing() {
                println!("Reference {} is broken or not accessible", r.name);
            }
        }
    }

    // You can also get defined names definition (string representation only)
    for name in workbook.defined_names() {
        println!("name: {}, formula: {}", name.0, name.1);
    }

    // Now get all formula!
    let sheets = workbook.sheet_names().to_owned();
    for s in sheets {
        println!(
            "found {} formula in '{}'",
            workbook
                .worksheet_formula(&s)
                .expect("error while getting formula")
                .rows()
                .flat_map(|r| r.iter().filter(|f| !f.is_empty()))
                .count(),
            s
        );
    }

    Ok(())
}
