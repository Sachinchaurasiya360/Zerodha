# Face Recognition Based Attendance System

This project is a face recognition-based attendance system that uses OpenCV and Python. The system uses a camera to capture images of individuals and then compares them with the images in the database to mark attendance.

## Features

- Face detection and recognition using dlib's deep learning models
- Graphical interface for registering new faces
- Real-time attendance marking through facial recognition
- Modern web interface to view and manage attendance records
- Export attendance data to Excel for reporting and analysis

## Installation

1. Clone the repository to your local machine:
   ```
   git clone https://github.com/Arijit1080/Face-Recognition-Based-Attendance-System
   ```
2. Install the required packages using:
   ```
   pip install -r requirements.txt
   ```
3. Download the dlib models from https://drive.google.com/drive/folders/12It2jeNQOxwStBxtagL1vvIJokoz-DL4?usp=sharing and place the data folder inside the repo

## Usage

1. Collect the Faces Dataset by running:
   ```
   python get_faces_from_camera_tkinter.py
   ```
2. Convert the dataset into features:
   ```
   python features_extraction_to_csv.py
   ```
3. To take the attendance run:
   ```
   python attendance_taker.py
   ```
4. Check and manage the attendance records:
   ```
   python app.py
   ```
   Then open your browser and navigate to `http://127.0.0.1:5000/`

## Attendance Management

The web interface allows you to:

- View attendance by date
- See a list of all people who were present on a specific day
- View comprehensive statistics (total records, total days, registered people)
- Export complete attendance report to Excel (all dates with daily summary)
- Export specific date attendance to Excel
- Two-sheet Excel export with detailed records and summary data

## Contributing

Contributions are welcome! Please feel free to submit a pull request or open an issue if you find any bugs or have any suggestions.
