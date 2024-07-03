use calamine::{open_workbook, Reader, Xlsx};
use dialoguer::Input;
use std::error::Error;
use std::fs::File;
use std::io::BufReader;
use xlsxwriter::*;

#[derive(Debug)]
struct Student {
    student_id: String,
    name: String,
    class: String,
    total_score: f32,
}

fn process_student_data(workbook: &mut Xlsx<BufReader<File>>) -> Vec<Student> {
    let mut students = vec![];

    // Assuming the sheet name is "Sheet1", modify if needed
    if let Ok(range) = workbook.worksheet_range("Sheet1") {
        let mut student_map = std::collections::HashMap::new();

        // Iterate over each row in the worksheet
        // for row in range.rows() {
        //     let student_id = &row[0].to_string();
        //     let name = &row[1].to_string();
        //     let class = &row[2].to_string();
        //     let score = &row[3].to_string().parse::<f32>().expect("msg"); // Assuming score is float

        //     // Calculate total score for each student
        //     let mut total_score = student_map
        //         .entry(student_id.clone())
        //         .or_insert((name.clone(), class.clone(), 0.0))
        //         .2;

        //     total_score = total_score + score;
        //     // Clamp total_score to a maximum of 100
        //     if total_score > 100.0 {
        //         student_map.get_mut(student_id).unwrap().2 = 100.0;
        //     }
        // }
        for row in range.rows() {
            let student_id = row[0].to_string();
            let name = row[1].to_string();
            let class = row[2].to_string();
            let score = row[3]
                .to_string()
                .parse::<f32>()
                .expect("Failed to parse score");

            // 更新学生总成绩
            let entry = student_map
                .entry(student_id.clone())
                .or_insert((name, class, 0.0));
            entry.2 += score;

            // 限制总成绩不超过100
            if entry.2 > 100.0 {
                entry.2 = 100.0;
            }

            println!("{:#?}", row);
        }

        // Convert student_map into Student struct and push to students Vec
        for (student_id, (name, class, total_score)) in student_map {
            let student = Student {
                student_id: student_id.to_string(),
                name: name.to_string(),
                class: class.to_string(),
                total_score: total_score as f32, // Convert to i32
            };
            students.push(student);
        }
    }

    // Print the processed student data (for verification)
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

    // Output the processed student data to a new Excel file
    let new_filename = "processed_data.xlsx";
    let excel = Workbook::new(new_filename)?;
    let mut sheet = excel.add_worksheet(None)?;

    // Write headers
    sheet.write_string(0, 0, "Student ID", None)?;
    sheet.write_string(0, 1, "Name", None)?;
    sheet.write_string(0, 2, "Class", None)?;
    sheet.write_string(0, 3, "Total Score", None)?;

    let mut i = 1;
    // Write student data
    for student in &student_data {
        sheet.write_string(i, 0, &student.student_id, None)?;
        sheet.write_string(i, 1, &student.name, None)?;
        sheet.write_string(i, 2, &student.class, None)?;
        sheet.write_string(i, 3, &student.total_score.to_string(), None)?;
        i += 1;
    }

    excel.close()?;

    println!("Processed data saved to {}", new_filename);

    Ok(())
}
