import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.io.InputStream

fun main() {

    val file = "src/myExcelRead.xlsx"
    val variable1: InputStream = FileInputStream(file)
    val workbook = XSSFWorkbook(variable1)
    //println(workbook)
    val sheet = workbook.getSheet("Sheet1")
    val totalRowNum: Int = sheet.physicalNumberOfRows
    val lastRowNum: Int = sheet.lastRowNum

    println("lastRowNum : $lastRowNum")

    println("totalRowNum $totalRowNum")
    println("-------------------------")

    // cell count
    //val lastCellNum = row.lastCellNum
    //println("lastCellNum $lastCellNum")

    for (i in 1..lastRowNum){
        // println("value of i : $i")
        val row: XSSFRow = sheet.getRow(i)
        //println("row :: $row")

        val excelData1 = row.getCell(0).rawValue
        val excelData2 = row.getCell(1).toString()
        println("""<string name="ram_${excelData1}">$excelData2</string>""")

        //<string name="ram_226">Please Provide Valid User Parameters.</string>

        // XSSFCell cell = row.getCell(j);
        //cell.setCellValue("data");

    }
}