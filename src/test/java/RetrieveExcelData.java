
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class RetrieveExcelData {

    protected ArrayList<String> getAllExcelData(String path, String sheetName,String columnName) throws IOException {
        /**
         *   Get excel work book and find all sheets
         */
        ArrayList<String> excelDataList = new ArrayList<String>();
        FileInputStream fis = new FileInputStream(path);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        Row firstRow;
        int sheets = workbook.getNumberOfSheets();

        for (int i = 0; i < sheets; i++) {
            if (workbook.getSheetName(i).equalsIgnoreCase(sheetName)) {
                XSSFSheet sheet = workbook.getSheetAt(i);
                /**
                 *   Find first row and value of first column
                 */
                Iterator<Row> rows = sheet.iterator();
                firstRow = rows.next();
                Iterator<Cell> ce = firstRow.cellIterator();
                int k = 0;
                int column = 0;
                while (ce.hasNext()) {
                    Cell cell = ce.next();
                    String value = null;
                    if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                        value = cell.getStringCellValue().trim();
                    } else {
                        value = NumberToTextConverter.toText(cell.getNumericCellValue()).trim();
                    }

                    if (value.equalsIgnoreCase(columnName)) {
                        column = k;
                    }

                    k++;
                }

                System.out.println("Start processing file from column : " + column);

                /**
                 *   Find first row and value of first column
                 */
                while (rows.hasNext()) {

                    Row row = rows.next();
                    row.getCell(column).setCellType(Cell.CELL_TYPE_STRING);
                    String firstColumnValue = row.getCell(column).getStringCellValue();

                    if (firstColumnValue.equalsIgnoreCase(columnName)) {

                        /**
                         *  Pull row column and save it in the list
                         */

                        Iterator<Cell> cv = row.cellIterator();
                        while (cv.hasNext()) {
                            Cell cell = cv.next();
                            String columnValue;
                            if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                                columnValue = cell.getStringCellValue().trim();
                                if (!columnValue.isEmpty()) {
                                    excelDataList.add(columnValue.trim());
                                }
                            } else {
                                columnValue = NumberToTextConverter.toText(cell.getNumericCellValue()).trim();
                                if (!columnValue.isEmpty()) {
                                    excelDataList.add(columnValue.trim());
                                }
                            }
                        }
                    }
                }
            }
        }
        return excelDataList;
    }


    protected ArrayList<String> getCISDExcelData(String path,String sheetName ,String columnName) throws IOException {
        /**
         *   Get excel work book and find all sheets
         */
        ArrayList<String> excelDataList = new ArrayList<String>();
        FileInputStream fis = new FileInputStream(path);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        Row firstRow;

        int sheets = workbook.getNumberOfSheets();

        for (int i = 0; i < sheets; i++) {

            if (workbook.getSheetName(i).equalsIgnoreCase(sheetName)) {
                XSSFSheet sheet = workbook.getSheetAt(i);
                /**
                 *   Find first row and value of first column
                 */
                Iterator<Row> rows = sheet.iterator();
                firstRow = rows.next();
                Iterator<Cell> ce = firstRow.cellIterator();
                int k = 0;
                int column = 0;

                while (ce.hasNext()) {
                    Cell cell = ce.next();
                    String value = null;
                    if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                        value = cell.getStringCellValue().trim();
                    } else {
                        value = NumberToTextConverter.toText(cell.getNumericCellValue()).trim();
                    }

                    if (value.equalsIgnoreCase(columnName)) {
                        column = k;
                    }
                    k++;
                }

                /**
                 *   Find first row and value of first column
                 */
                firstRow.getCell(column).setCellType(Cell.CELL_TYPE_STRING);
                String firstColumnValue = firstRow.getCell(column).getStringCellValue();

                if (firstColumnValue.equalsIgnoreCase(columnName)) {

                    Iterator<Cell> rowCells = firstRow.cellIterator();

                    while (rowCells.hasNext()) {
                        Cell cell =rowCells.next();
                        String columnValue;

                        if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                            columnValue = cell.getStringCellValue().trim();

                            excelDataList.add(columnValue);

                        } else {
                            columnValue = NumberToTextConverter.toText(cell.getNumericCellValue()).trim();
                            excelDataList.add(columnValue);
                        }
                    }
                }
            }
        }

        return excelDataList;
    }
}
