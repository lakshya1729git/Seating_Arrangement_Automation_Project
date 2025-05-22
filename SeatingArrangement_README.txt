
README - Seating Arrangement Automation

--------------------------------------------------
Project Title:
Seating Arrangement Automation - Internship Project

--------------------------------------------------
Description:

This project is a Streamlit-based web app that automates the generation of seating arrangements for exams. It allows users to upload a well-structured Excel file, configure preferences (like buffer seats and density), and then automatically allocates students to rooms based on course enrollments and room capacities.

It outputs:
- A complete seating plan
- A list of unallocated students (if rooms are insufficient)
- Individual attendance sheets for each date/session/room, packaged into a ZIP file for download

--------------------------------------------------
Dependencies Used:

streamlit
pandas
openpyxl
math
io
zipfile

Install with:
pip install streamlit pandas openpyxl

--------------------------------------------------
Input Excel Structure:

The uploaded Excel must contain the following sheets:

- in_timetable: Exam schedule with Morning and Evening sessions
- in_course_roll_mapping: Mapping of course codes to student roll numbers
- in_roll_name_mapping: Mapping of roll numbers to student names
- in_room_capacity: Room number, block, and seating capacity

--------------------------------------------------
User Inputs via UI:

- Buffer Seats: Number of empty seats to leave in each room
- Seating Mode:
  - "Sparse": Only 50% of effective capacity used
  - "Dense": Full effective capacity used

--------------------------------------------------
Core Logic and Loop Explanation:

1. Mappings and Preparation:
   - Normalize all inputs (strip spaces, uppercase roll numbers and courses).
   - Build name_lookup and course_to_students dictionaries for fast access.

2. Room Allocation Logic (assign_students):
   - Calculates usable seats after subtracting buffer.
   - Prefers rooms from the same block (B1 or B2) for a session.
   - Falls back to using all rooms (sorted by usable capacity) if needed.
   - Returns a list of (room, student list) tuples if allocation is possible.

3. Main Allocation Loop:
   - For each row in in_timetable:
     - Parse Morning and Evening courses.
     - Sort courses by number of students (large to small) to allocate bigger groups first.
     - Run assign_students() and update:
       - summary: For final seating data
       - left_out: For unallocated students
       - datewise_data: For generating attendance sheets

4. Excel Styling (format_worksheet):
   - Applies border and center alignment
   - Auto-sizes columns for better readability

5. ZIP File Creation:
   - Stores:
     - overall_seating.xlsx
     - seats_left.xlsx
     - Foldered attendance sheets: date/session/course_room.xlsx

--------------------------------------------------
Output Files Generated:

All generated inside a downloadable ZIP:

- overall_seating.xlsx: Master sheet with date-wise seating plans
- seats_left.xlsx: Lists courses that couldn't be allocated seats
- Attendance Sheets:
  - One Excel per course-room-session (under folder: date/morning/ or date/evening/)
  - Contains:
    - Roll no., student name, signature fields
    - Space for TAs and Invigilators to sign

--------------------------------------------------
Highlights:

- Intelligently distributes students to rooms by capacity and block
- Handles sparse/dense configurations
- Gracefully manages overflows and shows unallocated students
- Generates ready-to-print attendance sheets per session

--------------------------------------------------
Author:
This project was developed as part of an internship and can be easily adapted for institutions needing automated seating arrangements.
