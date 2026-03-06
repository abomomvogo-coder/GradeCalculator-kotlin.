// ============================================================
// Milestone1_Abomo.kt
// Student Grade Calculator — Version 2
// Author: Abomo
// Description: OOP Grade Calculator with Excel input/output,
//              higher-order functions, lambdas, and PDF export
// ============================================================

import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileOutputStream

data class Student(
    val studentId: String,
    val name: String,
    val math: Double?,
    val science: Double?,
    val english: Double?,
    val history: Double?
) {
    fun calculateAverage(): Double? {
        val grades = listOf(math, science, english, history).filterNotNull()
        return if (grades.isEmpty()) null else grades.sum() / grades.size
    }

    fun getLetterGrade(): String {
        val avg = calculateAverage() ?: return "N/A"
        return when {
            avg >= 90 -> "A"
            avg >= 80 -> "B"
            avg >= 70 -> "C"
            avg >= 60 -> "D"
            else      -> "F"
        }
    }

    fun isValid(): Boolean {
        return listOf(math, science, english, history)
            .filterNotNull()
            .all { it in 0.0..100.0 }
    }

    fun isPassing(): Boolean = (calculateAverage() ?: 0.0) >= 60.0

    fun formatReport(): String {
        val avg = calculateAverage()
        return buildString {
            appendLine("Student: $name ($studentId)")
            appendLine("  Math: ${math ?: "N/A"} | Science: ${science ?: "N/A"} | English: ${english ?: "N/A"} | History: ${history ?: "N/A"}")
            appendLine("  Average: ${avg?.let { "%.2f".format(it) } ?: "N/A"} | Grade: ${getLetterGrade()} | Status: ${if (isPassing()) "PASS" else "FAIL"}")
        }
    }
}

class ExcelReader(private val filePath: String) {
    fun readStudents(): List<Student> {
        val students = mutableListOf<Student>()
        val file = File(filePath)
        if (!file.exists()) return students
        val workbook = XSSFWorkbook(file.inputStream())
        val sheet = workbook.getSheetAt(0)
        sheet.drop(1).forEach { row ->
            val studentId = row.getCell(0)?.toString()?.trim() ?: return@forEach
            val name      = row.getCell(1)?.toString()?.trim() ?: return@forEach
            val math      = row.getCell(2)?.let { if (it.cellType == CellType.NUMERIC) it.numericCellValue else null }
            val science   = row.getCell(3)?.let { if (it.cellType == CellType.NUMERIC) it.numericCellValue else null }
            val english   = row.getCell(4)?.let { if (it.cellType == CellType.NUMERIC) it.numericCellValue else null }
            val history   = row.getCell(5)?.let { if (it.cellType == CellType.NUMERIC) it.numericCellValue else null }
            students.add(Student(studentId, name, math, science, english, history))
        }
        workbook.close()
        return students
    }
}

class ExcelWriter(private val outputPath: String) {
    fun writeResults(students: List<Student>, classAverage: Double?) {
        val workbook = XSSFWorkbook()
        val sheet = workbook.createSheet("Results")

        val headerStyle = workbook.createCellStyle().apply {
            fillForegroundColor = IndexedColors.DARK_BLUE.index
            fillPattern = FillPatternType.SOLID_FOREGROUND
            setFont(workbook.createFont().apply {
                bold = true
                color = IndexedColors.WHITE.index
                fontHeightInPoints = 12
            })
            alignment = HorizontalAlignment.CENTER
        }
        val altStyle = workbook.createCellStyle().apply {
            fillForegroundColor = IndexedColors.LIGHT_CORNFLOWER_BLUE.index
            fillPattern = FillPatternType.SOLID_FOREGROUND
            alignment = HorizontalAlignment.CENTER
        }
        val normalStyle = workbook.createCellStyle().apply {
            alignment = HorizontalAlignment.CENTER
        }
        val passStyle = workbook.createCellStyle().apply {
            fillForegroundColor = IndexedColors.LIGHT_GREEN.index
            fillPattern = FillPatternType.SOLID_FOREGROUND
            alignment = HorizontalAlignment.CENTER
            setFont(workbook.createFont().apply { bold = true })
        }
        val failStyle = workbook.createCellStyle().apply {
            fillForegroundColor = IndexedColors.ROSE.index
            fillPattern = FillPatternType.SOLID_FOREGROUND
            alignment = HorizontalAlignment.CENTER
            setFont(workbook.createFont().apply { bold = true })
        }

        val headers = listOf("Student ID", "Name", "Math", "Science", "English", "History", "Average", "Grade", "Status")
        val headerRow = sheet.createRow(0)
        headers.forEachIndexed { i, h ->
            headerRow.createCell(i).apply { setCellValue(h); cellStyle = headerStyle }
        }

        students.forEachIndexed { idx, s ->
            val row = sheet.createRow(idx + 1)
            val style = if (idx % 2 == 0) altStyle else normalStyle
            val avg = s.calculateAverage()
            fun cell(col: Int, value: String, st: CellStyle = style) =
                row.createCell(col).apply { setCellValue(value); cellStyle = st }
            cell(0, s.studentId); cell(1, s.name)
            cell(2, s.math?.toString() ?: "N/A")
            cell(3, s.science?.toString() ?: "N/A")
            cell(4, s.english?.toString() ?: "N/A")
            cell(5, s.history?.toString() ?: "N/A")
            cell(6, avg?.let { "%.2f".format(it) } ?: "N/A")
            cell(7, s.getLetterGrade())
            cell(8, if (s.isPassing()) "PASS" else "FAIL", if (s.isPassing()) passStyle else failStyle)
        }

        val summaryRow = sheet.createRow(students.size + 2)
        val boldStyle = workbook.createCellStyle().apply { setFont(workbook.createFont().apply { bold = true }) }
        summaryRow.createCell(0).apply { setCellValue("Class Average:"); cellStyle = boldStyle }
        summaryRow.createCell(1).apply { setCellValue(classAverage?.let { "%.2f".format(it) } ?: "N/A"); cellStyle = boldStyle }

        listOf(12, 20, 10, 10, 10, 10, 10, 8, 10).forEachIndexed { i, w -> sheet.setColumnWidth(i, w * 256) }
        FileOutputStream(outputPath).use { workbook.write(it) }
        workbook.close()
    }
}

class GradeBook {
    private val students = mutableListOf<Student>()

    fun loadFromExcel(path: String) {
        val loaded = ExcelReader(path).readStudents().filter { it.isValid() }
        students.addAll(loaded)
    }

    fun saveToExcel(path: String) {
        ExcelWriter(path).writeResults(students, getClassAverage())
    }

    fun processStudents(action: (Student) -> Unit) = students.forEach(action)

    fun filterStudents(predicate: (Student) -> Boolean): List<Student> = students.filter(predicate)

    fun getAllAverages(): List<Pair<String, Double?>> = students.map { it.name to it.calculateAverage() }

    fun getTopStudent(): Student? = students.maxByOrNull { it.calculateAverage() ?: 0.0 }

    fun getClassAverage(): Double? {
        val avgs = students.mapNotNull { it.calculateAverage() }
        return if (avgs.isEmpty()) null else avgs.sum() / avgs.size
    }

    fun getStudents(): List<Student> = students.toList()

    fun printSummary() {
        println("===== CLASS SUMMARY =====")
        println("Total Students : ${students.size}")
        println("Class Average  : ${"%.2f".format(getClassAverage() ?: 0.0)}")
        println("Top Student    : ${getTopStudent()?.name ?: "N/A"}")
        println("Passing        : ${filterStudents { it.isPassing() }.size}")
        println("Failing        : ${filterStudents { !it.isPassing() }.size}")
    }
}

fun main() {
    println("===== STUDENT GRADE CALCULATOR =====")
    val gradeBook = GradeBook()
    gradeBook.loadFromExcel("students.xlsx")
    println("\n--- Rapport de chaque etudiant ---")
    gradeBook.processStudents { student -> println(student.formatReport()) }
    println("--- Etudiants en echec ---")
    gradeBook.filterStudents { !it.isPassing() }
        .forEach { println("  ${it.name} - ${it.getLetterGrade()}") }
    println("\n--- Toutes les moyennes ---")
    gradeBook.getAllAverages()
        .forEach { (name, avg) -> println("  $name : ${avg?.let { "%.2f".format(it) } ?: "N/A"}") }
    gradeBook.saveToExcel("results.xlsx")
    gradeBook.printSummary()
}
