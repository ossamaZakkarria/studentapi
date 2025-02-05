const express = require("express");
const cors = require("cors");
const fs = require("fs");
const XLSX = require("xlsx");

const app = express();
const PORT = process.env.PORT || 5000;
const FILE_PATH = "students.xlsx";

app.use(cors());
app.use(express.json());

const readExcel = () => {
  if (!fs.existsSync(FILE_PATH)) return [];

  const workbook = XLSX.readFile(FILE_PATH);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  return XLSX.utils.sheet_to_json(sheet) || [];
};

const writeExcel = (data) => {
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(data);

  XLSX.utils.book_append_sheet(workbook, worksheet, "Students");
  XLSX.writeFile(workbook, FILE_PATH);
};

// read student endpoint
app.get("/students", (req, res) => {
  const students = readExcel();
  res.json(students);
});

// add student endpoint
app.post("/students", (req, res) => {
  const { name, gender } = req.body;
  const students = readExcel();

  const maxId =
    students.length > 0 ? Math.max(...students.map((s) => s.id)) : 0;
  const newStudent = { id: maxId + 1, name, gender };

  students.push(newStudent);
  writeExcel(students);

  res.json(newStudent);
});

// delete student endoint
app.delete("/students/:id", (req, res) => {
  const students = readExcel();
  const newStudents = students.filter(
    (student) => student.id !== parseInt(req.params.id)
  );

  writeExcel(newStudents);
  res.json({ message: "Student deleted successfully!" });
});

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
