import express from "express";
import cors from "cors";
import mongoose from "mongoose";
import "dotenv/config";
import ExcelJS from "exceljs";
import {eachDayOfInterval } from 'date-fns';
import pdf from "pdfkit";
import {format} from 'date-fns'






const port = 3000;
const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(cors({ origin: "*" }));
const username = process.env.MONGO_USERNAME;
const password = encodeURIComponent(process.env.MONGO_PASSWORD);

const dbConnection = await mongoose.connect(
    "mongodb+srv://"
     + username
    + ":" + password +
     "@cluster0.r2vrqr2.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0/attendance"

);
if (dbConnection)
  app.listen(port, () => console.log("Server Started on port " + port));

const timeZone = "Asia/Kolkata";
const date = new Date();
const formattedDate = format(date, "dd-MM-yyyy HH:mm:ss", { timeZone });
let d = new Date().toLocaleString("en-US", { timeZone: "Asia/Kolkata" });

const attendanceSchema = new mongoose.Schema({
  date: {
    type: Date,
    default: d,
  },
  faculty: {
    type: String,
    required: true,
  },
  attendance: {
    type: Object,
    required: true,
  },
});


const facultySchema = new mongoose.Schema({
  facultyList: [{
    id: {
      type: String,
      required: true,
      default: () => new mongoose.Types.ObjectId().toString(),
    },
    name: {
      type: String,
      required: true,
    }
  }]
});


const studentSchema = new mongoose.Schema({
  id: {
    type: Number,
    default: Date.now(),
  },
  faculty: {
    type: String,
    required: true,
  },
  students: {
    type: Array,
    required: true,
  }
});


const attendanceModel = mongoose.model("record", attendanceSchema, "record");
const facultyModel = mongoose.model("faculty", facultySchema, "faculty");
const studentModel = mongoose.model("student", studentSchema, "student");



// Save attendance data
app.post("/saveAttendance", (req, res) => {
  const dataToSave = new attendanceModel(req.body);
  dataToSave.save()
    .then(response => res.send(response))
    .catch(error => {
      console.log(error);
      res.status(500).send("Error saving attendance data");
    });
});




app.get('/getAttendance', async (req, res) => {
  const { start, end } = req.query;
  if (!start || !end) {
    return res.status(400).send("Start date and end date are required");
  }

  try {
    const attendanceData = await attendanceModel.find({
      date: {
        $gte: new Date(start),
        $lte: new Date(end),
      },
    });
    res.send(attendanceData);
  } catch (error) {
    console.log(error);
    res.status(500).send("Error fetching attendance data");
  }
});



// Get faculty list
app.get("/getFaculty", async (req, res) => {
  try {
    const facultyList = await facultyModel.find({}).select("-__v");
    res.json(facultyList);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Error fetching faculty list" });
  }
});

// Save faculty data
app.post("/saveFaculty", async (req, res) => {
  const dataToSave = new facultyModel(req.body);
  dataToSave.save()
    .then(response => res.json(response))
    .catch(error => {
      console.log("Error saving faculty data:", error);
      res.status(500).send("Error saving faculty data");
    });
});

// Get student list based on faculty
app.get("/getStudent", async (req, res) => {
  const { faculty } = req.query;
  try {
    let studentList;
    if (faculty) {
      studentList = await studentModel.find({ faculty }).select("-__v");
    } else {
      studentList = await studentModel.find({}).select("-__v");
    }
    res.json(studentList);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Error fetching student list" });
  }
});

// Save student data
app.post("/saveStudent", async (req, res) => {
  const dataToSave = new studentModel(req.body);
  dataToSave.save()
    .then(response => res.json(response))
    .catch(error => {
      console.log("Error saving student data:", error);
      res.status(500).send("Error saving student data");
    });
});



//Get Faculty by Attendance
app.get('/getAttendanceFaculties', async (req, res) => {
  try {
      const faculties = await attendanceModel.distinct('faculty');
      res.json(faculties);
  } catch (error) {
      console.error('Error fetching faculties:', error);
      res.status(500).send('Error fetching faculties');
    }
});

// Delete a single student
app.delete("/deleteStudent/:id", async (req, res) => {
  const idToDelete = req.params.id;
  try {
      const result = await studentModel.updateOne(
          { "students.id": parseInt(idToDelete) },
          { $pull: { students: { id: parseInt(idToDelete) } } }
      );
      if (result.modifiedCount > 0) {
          res.json("Student Deleted");
      } else {
          res.status(404).json("Student not found");
      }
  } catch (error) {
      console.error(error);
      res.status(500).send("Error deleting student");
  }
});

//  delete single faculty

app.delete("/deleteFaculty/:id", async (req, res) => {
  const idToDelete = req.params.id;
  console.log("Received ID:", idToDelete); 

  try {
      const result = await facultyModel.updateOne(
          { "facultyList.id": idToDelete },
          { $pull: { facultyList: { id: idToDelete } } }
      );
      console.log("Update Result:", result);  

      if (result.modifiedCount > 0) {
          res.json("Faculty Deleted");
      } else {
          res.status(404).json("Faculty not found");
      }
  } catch (error) {
      console.error(error);
      res.status(500).send("Error deleting faculty");
  }
});


// // Delete multiple faculties

app.delete("/deleteFacultyMany", async (req, res) => {
  const facultyIdsToDelete = req.body.ids;
  try {
    const result = await facultyModel.updateMany(
      {},
      { $pull: { facultyList: { id: { $in: facultyIdsToDelete } } } }
    );
    if (result.modifiedCount > 0) {
      res.json("Faculties Deleted");
    } else {
      res.status(404).json("No faculties found");
    }
  } catch (error) {
    console.error(error);
    res.status(500).send("Error deleting faculties");
  }
});


app.get("/exportAttendance/:startDate/:endDate", async (req, res) => {
  const { startDate, endDate } = req.params;
  const { faculties } = req.query; 

  try {
      const endOfDay = new Date(endDate);
      endOfDay.setHours(23, 59, 59, 999);

    
      let filter = {
          date: {
              $gte: new Date(startDate),
              $lte: endOfDay
          }
      };

      if (faculties) {
          const facultyList = faculties.split(",");
          filter.faculty = { $in: facultyList };
      }

      const attendanceData = await attendanceModel.find(filter);

      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Attendance');

      const formattedStartDate = format(new Date(startDate), 'dd-MM-yyyy', { timeZone: 'Asia/Kolkata' });
      const formattedEndDate = format(new Date(endDate), 'dd-MM-yyyy', { timeZone: 'Asia/Kolkata' });

      const dateRange = eachDayOfInterval({ start: new Date(startDate), end: new Date(endOfDay) }).map(date => format(date, 'dd-MM-yyyy'));
      const headerRow = ['Faculty', 'Student Name', 'Date', ...dateRange];
      const headerRowCell = worksheet.addRow(headerRow);
      headerRowCell.height = 28;
      headerRowCell.eachCell((cell) => {
          cell.font = { bold: true };
          cell.alignment = { vertical: 'center', horizontal: 'center' };
      });

     
      worksheet.getColumn(1).width = 16;
      worksheet.getColumn(2).width = 16;
      worksheet.getColumn(3).width = 12;

      worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
          row.eachCell({ includeEmpty: true }, function (cell, colNumber) {
              if (colNumber === 1 || colNumber === 2 || colNumber === 3) {
                  cell.alignment = { horizontal: 'center' };
              }
          });
      });

      for (let i = 4; i < dateRange.length + 4; i++) {
          const column = worksheet.getColumn(i);
          column.width = 12;
          column.alignment = { horizontal: 'center' };
      }

   
      const attendanceMap = {};
      attendanceData.forEach(record => {
          const { date, faculty, attendance } = record;
          const formattedDate = format(date, 'dd-MM-yyyy');

          if (!attendanceMap[faculty]) {
              attendanceMap[faculty] = {};
          }

          for (const studentId in attendance) {
              const studentData = attendance[studentId];
              if (!attendanceMap[faculty][studentData.name]) {
                  attendanceMap[faculty][studentData.name] = {};
              }
              attendanceMap[faculty][studentData.name][formattedDate] = studentData.status;
          }
      });

      Object.keys(attendanceMap).forEach(faculty => {
          const students = attendanceMap[faculty];
          let isFirstStudent = true;

          Object.keys(students).forEach(studentName => {
              const attendanceByDate = students[studentName];
              const row = [isFirstStudent ? faculty : '', studentName, ''];

              dateRange.forEach(date => {
                  row.push(attendanceByDate[date] || 'A');
              });

              const newRow = worksheet.addRow(row);
              newRow.eachCell((cell) => {
                  cell.alignment = { horizontal: 'center' };
              });

              dateRange.forEach((date, index) => {
                  const cell = newRow.getCell(index + 4);
                  if (cell.value === 'P') {
                      cell.font = {
                          color: { argb: 'FF005500' },
                          bold: true
                      };
                  } else if (cell.value === 'A') {
                      cell.font = { color: { argb: 'FFFF0000' } };
                  }
              });

              isFirstStudent = false;
          });
      });

     
      const filename = `attendance_${startDate}_${endDate}.xlsx`;
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
      await workbook.xlsx.write(res);

      res.end();
  } catch (error) {
      console.error('Error exporting attendance:', error);
      res.status(500).send('Internal Server Error');
  }
});

// exportAttendancePDF
app.get('/exportAttendancePDF/:startDate/:endDate', async (req, res) => {
  const { startDate, endDate } = req.params;
  const { faculties } = req.query;

  try {
      const endOfDay = new Date(endDate);
      endOfDay.setHours(23, 59, 59, 999);

      const attendanceData = await attendanceModel.find({
          date: {
              $gte: new Date(startDate),
              $lte: endOfDay
          },
          faculty: { $in: faculties.split(',') } 
      });

      const doc = new pdf();
      const filename = `attendance_${startDate}_${endDate}.pdf`;
      res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
      res.setHeader("Content-Type", "application/pdf");
      doc.pipe(res);

      const startDateFormatted = formatDate(new Date(startDate));
      const endDateFormatted = formatDate(new Date(endDate));
      const title = `Attendance Report: ${startDateFormatted} - ${endDateFormatted}`;

      doc.font('Helvetica-Bold').fontSize(14);
      doc.text(title, { align: 'center' });
      doc.moveDown(0.5);

      const dateRange = eachDayOfInterval({ start: new Date(startDate), end: endOfDay }).map(date => format(date, 'dd'));

      const facultyMap = {};

      attendanceData.forEach(record => {
          const { date, faculty, attendance } = record;
          const formattedDate = format(date, 'dd');

          if (!facultyMap[faculty]) {
              facultyMap[faculty] = {
                  students: {}
              };
          }

          for (const studentId in attendance) {
              const studentData = attendance[studentId];
              if (!facultyMap[faculty].students[studentData.name]) {
                  facultyMap[faculty].students[studentData.name] = {};
              }
              facultyMap[faculty].students[studentData.name][formattedDate] = studentData.status;
          }
      });

      let y = 120;
      const leftMargin = 35;
      const lMargin = 110;
      const columnWidth = 16;

      doc.font('Helvetica').fontSize(10);

      Object.keys(facultyMap).forEach((faculty, index) => {
          const { students } = facultyMap[faculty];

          y += 10;
          doc.font('Helvetica-Bold').fontSize(12).text(`Faculty: ${faculty}`, leftMargin, y, { width: 650, text: 'center' });
          y += 20;

          doc.lineWidth(0.5);
          doc.rect(leftMargin, y, dateRange.length * columnWidth, 20).stroke();
          doc.font('Helvetica-Bold').fontSize(10).text('Student', leftMargin + 2, y + 3, { width: lMargin - leftMargin - 4, text: 'center' });
          dateRange.forEach((date, dayIndex) => {
              const x = lMargin + dayIndex * columnWidth;
              doc.font('Helvetica-Bold').fontSize(10).text(date, x, y + 3, { width: columnWidth, align: 'center' });
              doc.rect(x, y, columnWidth, 20).stroke();
          });

          y += 20;

          Object.keys(students).forEach((studentName, studentIndex) => {
              doc.font('Helvetica').fontSize(10).text(studentName, leftMargin, y,{text:"center"});
              doc.rect(leftMargin, y, lMargin - leftMargin + dateRange.length * columnWidth, 15).stroke();

              dateRange.forEach((date, dayIndex) => {
                  const formattedDate = `${date}`;
                  const status = students[studentName][formattedDate] || 'A';
                  const color = status === 'P' ? 'green' : 'red';
                  const x = lMargin + dayIndex * columnWidth + (columnWidth / 2) - (doc.widthOfString(status) / 2);
                  doc.fillColor(color).text(status, x, y + 3, { width: columnWidth, text: 'center' });
                  doc.fillColor('black');
                  doc.rect(lMargin + dayIndex * columnWidth, y, columnWidth, 15).stroke();
              });

              y += 15;
          });

          y += 10;
      });

      doc.end();
  } catch (error) {
      console.error("Error exporting attendance as PDF:", error);
      res.status(500).send("Internal Server Error");
  }
});

function formatDate(date) {
  const options = { year: 'numeric', month: 'long', day: 'numeric' };
  return date.toLocaleDateString('en-GB', options);
}

// exportAttendanceJSON
app.get("/exportAttendanceJSON/:startDate/:endDate", async (req, res) => {
  try {
      const { startDate, endDate } = req.params;
      const { faculties } = req.query; 
      const endOfDay = new Date(endDate);
      endOfDay.setHours(23, 59, 59, 999);
      let filter = {
          date: {
              $gte: new Date(startDate),
              $lte: endOfDay
          }
      };
      if (faculties) {
          filter.faculty = { $in: faculties.split(",") };
      }
      const attendanceData = await attendanceModel.find(filter);
      const jsonAttendanceData = attendanceData.map(record => {
          const { date, faculty, attendance } = record;
          const formattedDate = format(date, 'dd-MM-yyyy');
          const formattedAttendance = {};
          for (const studentId in attendance) {
              const studentData = attendance[studentId];
              formattedAttendance[studentData.name] = {
                  date: formattedDate,
                  status: studentData.status
              };
          }
          return {
              faculty,
              attendance: formattedAttendance
          };
      });
      res.json(jsonAttendanceData);
  } catch (error) {
      console.error("Error exporting attendance:", error);
      res.status(500).send("Internal Server Error");
    }
});












