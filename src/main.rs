use std::env;
use std::fmt;
use std::fs::File;
use std::io::prelude::*;
use std::path::Path;
use office::{Excel, DataType};


/// `WorkoutData` represents a single row in the xlsx worksheet.
#[derive(Default, Debug)]
struct WorkoutData {
    /// Timestamp in minutes of the data point
    time: f64,
    /// Relative intensity at `time` in percent of FTP
    intensity: f64,
}

/// `Interval` represents an interval which is created in the 
/// `erg` file
#[derive(Default, Debug)]
struct Interval {
    /// Time in minutes the interval takes
    duration: f64,
    /// Average watts of the interval
    watt: f64,
    /// Approximation of the intensity factor, not 100% accurate when the 
    /// interval ramps up or down, but close enough
    intensity_factor: f64,
    /// Training Stress Score of the interval
    tss: f64,
}

impl Interval {
    /// Creates a new `Interval`, requires to consecutive `WorkoutData` points 
    /// and the current `FTP` as parameters.
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

/// The `Workout` struct represents the complete workout and contains 
/// the current `FTP`, `file_name`, the `description` of the workout, 
/// `Vectors` of `WorkoutData` and `Interval`s, as well as the total `TSS`.
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
    /// Custom formatting so that it a quick summary of the workout can be 
    /// printed to console after it is converted
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, r"{:24} | TSS: {:5} | {}", 
               self.file_name, self.tss as u64, self.description)
    }
}

/// Writes the parsed `Workout` to an `erg` file.
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
            file_content.push_str(&format!("{:.2}\t{}\n", 
                data.time, (data.intensity * workout.ftp) as u64));
        }

        file_content.push_str("[END COURSE DATA]\n");
        file.write_all(file_content.as_bytes()).expect("Couldn't write file");
}

/// Parses the workbook and iterates over each worksheet. All worksheets 
/// are then converted to `erg` files except the `Overview` worksheet.
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
                    Interval::new(&workout.workout_data[i], 
                                  &workout.workout_data[i+1], workout.ftp);
                tss += interval.tss;
                workout.intervals.push(interval);
            }
        }
        workout.tss = tss;
    }
    workout
}

fn main() {
    // Check argument
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        panic!("Usage: {} <file>", args[0]);
    }

    // open workbook and get worksheets
    let mut workbook = Excel::open(args[1].to_owned())
        .expect("Couldn't open Excel file");
    let mut worksheets = workbook.sheet_names()
        .expect("Couldn't get worksheets");
    worksheets.sort();

    // loop over worksheets, parse content and write `erg` files
    for worksheet in worksheets {
        if worksheet == "Overview" { continue; }
        let workout = parse_workout(&mut workbook, &worksheet);
        println!("{}", workout);
        write_erg_file(workout);
    }
}
