import javax.swing.*
import javax.swing.table.DefaultTableModel
import javax.swing.table.DefaultTableCellRenderer
import java.awt.*
import java.io.File
import javax.swing.filechooser.FileNameExtensionFilter

class GradeCalculatorGUI : JFrame("Student Grade Calculator") {

    private val gradeBook = GradeBook()
    private val tableModel = DefaultTableModel(
        arrayOf("ID", "Name", "Math", "Science", "English", "History", "Average", "Grade", "Status"), 0
    )
    private val table = JTable(tableModel)
    private val statusLabel = JLabel("  Aucun fichier charge.")

    init {
        setupUI()
        defaultCloseOperation = EXIT_ON_CLOSE
        setSize(1000, 600)
        setLocationRelativeTo(null)
        isVisible = true
    }

    private fun setupUI() {
        layout = BorderLayout(10, 10)

        val titlePanel = JPanel(FlowLayout(FlowLayout.CENTER))
        titlePanel.background = Color(46, 64, 87)
        val titleLabel = JLabel("Student Grade Calculator")
        titleLabel.font = Font("Segoe UI", Font.BOLD, 22)
        titleLabel.foreground = Color.WHITE
        titlePanel.add(titleLabel)
        add(titlePanel, BorderLayout.NORTH)

        val buttonPanel = JPanel(FlowLayout(FlowLayout.LEFT, 10, 10))
        val btnLoad = JButton("Charger Excel")
        val btnSave = JButton("Sauvegarder Excel")
        val btnPDF  = JButton("Exporter PDF")

        btnLoad.background = Color(52, 152, 219); btnLoad.foreground = Color.WHITE
        btnSave.background = Color(39, 174, 96);  btnSave.foreground = Color.WHITE
        btnPDF.background  = Color(231, 76, 60);  btnPDF.foreground  = Color.WHITE

        listOf(btnLoad, btnSave, btnPDF).forEach {
            it.font = Font("Segoe UI", Font.BOLD, 13)
            it.preferredSize = Dimension(180, 38)
            buttonPanel.add(it)
        }

        btnLoad.addActionListener { loadExcel() }
        btnSave.addActionListener { saveExcel() }
        btnPDF.addActionListener  { exportPDF() }

        add(buttonPanel, BorderLayout.NORTH)

        table.font = Font("Segoe UI", Font.PLAIN, 13)
        table.rowHeight = 28
        table.tableHeader.font = Font("Segoe UI", Font.BOLD, 13)
        table.tableHeader.background = Color(46, 64, 87)
        table.tableHeader.foreground = Color.WHITE

        table.columnModel.getColumn(8).cellRenderer = object : DefaultTableCellRenderer() {
            override fun getTableCellRendererComponent(
                t: JTable, value: Any?, selected: Boolean, focused: Boolean, row: Int, col: Int
            ): Component {
                val comp = super.getTableCellRendererComponent(t, value, selected, focused, row, col)
                horizontalAlignment = CENTER
                background = when (value?.toString()) {
                    "PASS" -> Color(212, 239, 223)
                    "FAIL" -> Color(253, 212, 212)
                    else   -> Color.WHITE
                }
                return comp
            }
        }

        add(JScrollPane(table), BorderLayout.CENTER)

        val bottomPanel = JPanel(BorderLayout())
        bottomPanel.background = Color(236, 240, 241)
        statusLabel.font = Font("Segoe UI", Font.ITALIC, 12)
        bottomPanel.add(statusLabel, BorderLayout.WEST)
        add(bottomPanel, BorderLayout.SOUTH)
    }

    private fun loadExcel() {
        val chooser = JFileChooser()
        chooser.fileFilter = FileNameExtensionFilter("Excel Files (*.xlsx)", "xlsx")
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            gradeBook.loadFromExcel(chooser.selectedFile.absolutePath)
            refreshTable()
            statusLabel.text = "  Fichier charge : ${chooser.selectedFile.name}"
        }
    }

    private fun saveExcel() {
        val chooser = JFileChooser()
        chooser.selectedFile = File("results.xlsx")
        if (chooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
            gradeBook.saveToExcel(chooser.selectedFile.absolutePath)
            statusLabel.text = "  Resultats sauvegardes !"
            JOptionPane.showMessageDialog(this, "Fichier Excel sauvegarde!", "Succes", JOptionPane.INFORMATION_MESSAGE)
        }
    }

    private fun exportPDF() {
        JOptionPane.showMessageDialog(this, "Fonction PDF disponible avec iText.", "PDF", JOptionPane.INFORMATION_MESSAGE)
    }

    private fun refreshTable() {
        tableModel.rowCount = 0
        gradeBook.getStudents().forEach { s ->
            val avg = s.calculateAverage()
            tableModel.addRow(arrayOf(
                s.studentId, s.name,
                s.math?.toString() ?: "N/A",
                s.science?.toString() ?: "N/A",
                s.english?.toString() ?: "N/A",
                s.history?.toString() ?: "N/A",
                avg?.let { "%.2f".format(it) } ?: "N/A",
                s.getLetterGrade(),
                if (s.isPassing()) "PASS" else "FAIL"
            ))
        }
    }
}

fun main() {
    SwingUtilities.invokeLater { GradeCalculatorGUI() }
}
