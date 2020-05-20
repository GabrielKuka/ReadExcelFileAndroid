package com.file.select

import android.os.Bundle
import android.os.Environment
import android.util.Log
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import com.obsez.android.lib.filechooser.ChooserDialog
import jxl.Workbook
import jxl.WorkbookSettings
import jxl.read.biff.BiffException
import kotlinx.android.synthetic.main.activity_main.*
import org.apache.poi.hssf.usermodel.HSSFDateUtil
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellValue
import org.apache.poi.ss.usermodel.FormulaEvaluator
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileNotFoundException
import java.io.IOException
import java.text.SimpleDateFormat


class MainActivity : AppCompatActivity() {


    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        button.setOnClickListener {
            chooseFile()

        }
    }

    private fun chooseFile() {
        ChooserDialog(this@MainActivity)
            .withStartFile(Environment.getExternalStorageState() + "/")
            .withChosenListener { path, pathFile ->
                Log.d("File: ", path)

                //readExcelFileWithJExcelApi(path)

                readExcelFileWithApachePOI(path)

            } // to handle the back key pressed or clicked outside the dialog:
            .withOnCancelListener { dialog ->
                Log.d("CANCEL", "CANCEL")
                dialog.cancel() // MUST have
            }
            .build()
            .show()
    }

    private fun readExcelFileWithJExcelApi(path: String) {
        val file = File(path)
        if (file.exists()) {
            val ws = WorkbookSettings()
            ws.gcDisabled = true
            try {

                val sheet = Workbook.getWorkbook(file).getSheet(0)
                Toast.makeText(this, sheet.getRow(0)[0].contents, Toast.LENGTH_SHORT).show()

            } catch (e: IOException) {
                e.printStackTrace()
            } catch (e: BiffException) {
                e.printStackTrace()
            }
        }
    }

    private fun readExcelFileWithApachePOI(path: String) {

        try {
            val excelFile = FileInputStream(File(path))
            val workBook = XSSFWorkbook(excelFile)
            val sheet = workBook.getSheetAt(0)
            val rowCount = sheet.physicalNumberOfRows
            val formulaEvaluator = workBook.creationHelper.createFormulaEvaluator()

            for (rowPos in 1 until rowCount) {
                val row = sheet.getRow(rowPos)
                val cellCount = row.physicalNumberOfCells

                for (cellPos in 0 until cellCount) {
                    val cellValue = getCellAsString(row, cellPos, formulaEvaluator)
                    Log.d("Cell: ", cellValue!!)
                }

            }

        } catch (e: FileNotFoundException) {
            e.printStackTrace()
        } catch (e: IOException) {
            e.printStackTrace()
        } catch (e: Exception) {
            e.printStackTrace()
        }

    }

    private fun getCellAsString(
        row: Row,
        c: Int,
        formulaEvaluator: FormulaEvaluator
    ): String? {
        var value = ""
        try {
            val cell: Cell = row.getCell(c)
            val cellValue: CellValue = formulaEvaluator.evaluate(cell)
            when (cellValue.cellType) {
                Cell.CELL_TYPE_BOOLEAN -> value = "" + cellValue.booleanValue
                Cell.CELL_TYPE_NUMERIC -> {
                    val numericValue: Int = cellValue.numberValue.toInt()
                    value = if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        val date: Double = cellValue.numberValue
                        val formatter = SimpleDateFormat("MM/dd/yy")
                        formatter.format(HSSFDateUtil.getJavaDate(date))
                    } else {
                        "0" + numericValue
                    }
                }
                Cell.CELL_TYPE_STRING -> value = "" + cellValue.stringValue
                else -> {
                }
            }
        } catch (e: NullPointerException) {
            e.printStackTrace()
        }
        return value
    }

}
