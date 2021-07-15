use std::env;
use std::fmt;
use std::fs::File;
use std::io::prelude::*;
use std::path::Path;
use office::{Excel, DataType};


#[derive(Default, Debug)]
struct WorkoutData {
    time: f64,
    intensity: f64,
}

#[derive(Default, Debug)]
struct Interval {
    duration: f64,
    watt: f64,
    intensity_factor: f64,
    tss: f64,
}

impl Interval {
    fn new(wd1: &WorkoutData, wd2: &WorkoutData, ftp: f64) -> Self {
        let duration = wd2.time - wd1.time;
        let watt = (wd1.intensity + wd2.intensity) / 2.0 * ftp;
        let intensity_factor = watt / ftp;
        let tss = (duration/60.0) * intensity_factor.powf(2.0) * 100.0;
        Self {
            duration,
            watt,
            intensity_factor,
            tss,
        }
    }
}

#[derive(Default, Debug)]
struct Workout {
    ftp: f64,
    file_name: String,
    description: String,
    workout_data: Vec<WorkoutData>,
    intervals: Vec<Interval>,
    tss: f64,
}

impl fmt::Display for Workout {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, r"{:24} | TSS: {:5} | {}", 
               self.file_name, self.tss as u64, self.description)
    }
}

fn write_erg_file(workout: Workout) {
        let path = Path::new(&workout.file_name);
        let mut file = File::create(&path).expect("Couldn't open file");
        let mut file_content = format!("[COURSE HEADER]
VERSION = 2
UNITS = ENGLISH
DESCRIPTION = {}
FILE NAME = {}
FTP = {}
MINUTES WATTS
[END COURSE HEADER]
[COURSE DATA]
", workout.description, workout.file_name, workout.ftp);

        for data in workout.workout_data {
            file_content.push_str(&format!("{:.2}\t{}\n", data.time, (data.intensity * workout.ftp) as u64));
        }

        file_content.push_str("[END COURSE DATA]\n");
        file.write_all(file_content.as_bytes()).expect("Couldn't write file");
}

fn parse_workout(workbook: &mut Excel, worksheet: &str) -> Workout {

    let mut workout = Workout{.. Default::default()};

    if let Ok(range) = workbook.worksheet_range(worksheet) {
        let rows = range.rows();
        if let DataType::Float(ftp) = range.get_value(0, 1) {
            workout.ftp = *ftp;
        }
        if let DataType::String(file_name) = range.get_value(1, 1) {
            workout.file_name = file_name.to_string();
        }
        if let DataType::String(description) = range.get_value(2, 1) {
            workout.description = description.to_string();
        }

        for row in rows.skip(4) {
            match row {
                [DataType::Float(time), DataType::Float(intensity)] => {
                    workout.workout_data.push(
                        WorkoutData {
                            time: *time,
                            intensity: *intensity,
                        }
                    );
                },
                [DataType::Empty, DataType::Empty] => {
                    println!("EMPTY");
                    break;
                },
                _ => println!("Error in dataset"),
            }
        }
        
        assert!(workout.workout_data.len() % 2 == 0);

        let mut tss = 0.0;
        for i in 0..workout.workout_data.len() {
            if i % 2 == 0 {
                let interval = 
                    Interval::new(&workout.workout_data[i], &workout.workout_data[i+1], workout.ftp);
                tss += interval.tss;
                workout.intervals.push(interval);
            }
        }
        workout.tss = tss;
    }
    workout
}

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        panic!("Usage: {} <file>", args[0]);
    }
    let mut workbook = Excel::open(args[1].to_owned()).except("Couldn't open Excel file");
    let mut worksheets = workbook.sheet_names().except("Couldn't get worksheets");
    worksheets.sort();
    for worksheet in worksheets {
        if worksheet == "Rider" { continue; }
        let workout = parse_workout(&mut workbook, &worksheet);
        println!("{}", workout);
        write_erg_file(workout);
    }
}
