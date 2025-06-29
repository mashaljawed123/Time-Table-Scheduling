# 📅 Weekly Timetable Generator (Semester-wise)

This Python project automatically generates a **weekly timetable** for university-level courses across multiple semesters using **graph theory** concepts such as **bipartite graphs** and **maximum matching**. The output is saved in an Excel workbook with **color-coded schedules** per semester and includes **graph visualizations** for better understanding.


## 🧠 How It Works

- Each **section-course pair** is treated as one part of a bipartite graph.
- Each **room-timeslot pair** forms the other part.
- The program attempts to match section-course pairs to room-timeslot pairs using **maximum bipartite matching**.
- Valid matchings are filtered to avoid time conflicts and duplicate course assignments.
- Final timetables are exported as **Excel sheets**, one per weekday.
- Each cell in the Excel sheet is **color-coded** based on the semester.


## 🎓 Supported Semesters

- **1st Year (Semester 1)**
- **2nd Year (Semester 3)**
- **3rd Year (Semester 5)**
- **4th Year (Semester 7)**  
  - Includes standard and SE-specific courses (`7AS`, `PPIT`, etc.)


## 🧩 Technologies & Libraries Used

- [`networkx`](https://networkx.org/) – Graph creation and maximum matching
- [`openpyxl`](https://openpyxl.readthedocs.io/) – Excel generation and styling
- [`matplotlib`](https://matplotlib.org/) – Bipartite graph visualization
- [`itertools`](https://docs.python.org/3/library/itertools.html) – Cartesian product for combinations


## 🗂️ Project Structure

- `sections_courses`: All combinations of sections and courses per semester
- `rooms_slots`: All room and timeslot combinations
- `G_section_course`: Bipartite graph for section-course ↔ room-timeslot
- `G_room_timeslot`: Graph for visualizing room-time usage
- `Weekly_Timetable.xlsx`: Final output Excel workbook (one sheet per weekday)
- Visual bipartite graphs shown for:
  - Overall section-to-course mapping
  - Day-wise room/time-to-course allocations


## 🎨 Semester Colors (in Excel)

| Semester | Color     | Hex     |
|----------|-----------|---------|
| 7        | Yellow    | `FFFF00`|
| 5        | Magenta   | `FF00FF`|
| 3        | Pink      | `FFCCCB`|
| 1        | Blue      | `CCE5FF`|


## 📊 Output Example (Excel Sheet)

```

| Timeslot | A1       | A2      | ... |
| -------- | -------- | ------- | --- |
| 8-9      | PF 1A    | ALGO 5C | ... |
| 9-10     | FYP-1 7B | DS 3D   | ... |
| ...      |          |         |     |

````

Each cell: `[Course Name] [Section]`


## 🧪 How to Run

1. **Install dependencies**
   ```bash
   pip install networkx openpyxl matplotlib
````

2. **Run the script**

   ```bash
   python timetable_generator.py
   ```

3. **Check output**

   * Graphs will be shown (optional).
   * A file named `Weekly_Timetable.xlsx` will be saved in your directory.


## 👩‍💻 Authors

Made with ♥ by:

* Mashal Jawed Ali
* Rida
* Mehdi


## 📜 License

This project is open-source and free to use for academic or educational purposes.
