use calamine::{open_workbook, DataType, Reader, Xlsx};
use dialoguer::Input;
use std::error::Error;
use std::fs::File;
use std::io::BufReader;

#[derive(Debug)]
struct Student {
    student_id: String,
    name: String,
    class: String,
    total_score: i32,
}

fn process_student_data(workbook: &mut Xlsx<BufReader<File>>) -> Vec<Student> {
    let mut students = vec![];

    // Assuming the sheet name is "Sheet1", modify if needed
    if let Ok(range) = workbook.worksheet_range("Sheet1") {
        for row in range.rows() {
            let student_id = &row[0];
            let name = &row[1];
            let class = &row[2];
            let score = &row[3];
            let student = Student {
                student_id: student_id.to_string(),
                name: name.to_string(),
                class: class.to_string(),
                total_score: score.to_string().parse().expect("REASON"),
            };
            students.push(student);
        }
    }
    println!("{:#?}", students);
    students
}

fn main() -> Result<(), Box<dyn Error>> {
    // Prompt the user to enter the filename
    let filename: String = Input::new()
        .with_prompt("Enter the Excel file path")
        .interact()?;

    // Opens a new workbook
    let mut workbook: Xlsx<_> = open_workbook(&filename)?;

    // Process student data
    let student_data = process_student_data(&mut workbook);

    // Print student data for verification (remove this in production)
    for student in &student_data {
        println!(
            "Student ID: {}, Name: {}, Class: {}, Total Score: {}",
            student.student_id, student.name, student.class, student.total_score
        );
    }

    Ok(())
}
