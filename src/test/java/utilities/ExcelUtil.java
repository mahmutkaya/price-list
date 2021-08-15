package utilities;

import org.apache.poi.ss.usermodel.*;
import org.testng.Assert;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelUtil {

    private Workbook workBook;
    private Sheet workSheet;
    private String path;

    /**
     * This Constructor is to open and access the excel file
     * @param path the path of excel file
     * @param sheetName the name of sheet in excel file
     */
    public ExcelUtil(String path, String sheetName) {
        this.path = path;
        try {
            // Opening the Excel file
            FileInputStream fileInputStream = new FileInputStream(path);
            // accessing the workbook
            workBook = WorkbookFactory.create(fileInputStream);
            //getting the worksheet
            workSheet = workBook.getSheet(sheetName);
            //asserting if sheet has data or not
            Assert.assertNotNull(workSheet, "Worksheet: \"" + sheetName + "\" was not found\n");
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * This will get the list of the data in the excel file
     * This takes the data as string and will return the data as a Map of String
     * @return List<Map<String, String>> all_data
     */
    public List<Map<String, String>> getDataList() {
        // getting all columns
        List<String> columns = getColumnsNames();
        // method will return this
        List<Map<String, String>> data = new ArrayList<>();
        for (int i = 1; i < rowCount(); i++) {
            // get each row
            Row row = workSheet.getRow(i);
            // creating map of the row using the column and value
            // key=column, value=cell
            Map<String, String> rowMap = new HashMap<String, String>();
            for (Cell cell : row) {
                int columnIndex = cell.getColumnIndex();
                rowMap.put(columns.get(columnIndex), cell.toString());
            }
            data.add(rowMap);
        }
        return data;
    }

    /**
     * Getting the number of columns in a specific single row
     * @return int column_numbers
     */
    public int columnCount() {
        //getting how many numbers in row 1
        return workSheet.getRow(0).getLastCellNum();
    }

    /**
     * Get the number of rows
     * @return int row_numbers
     */
    public int rowCount() {
        return workSheet.getLastRowNum() + 1; }//adding 1 to get the actual count

    /**
     * When you enter row and column number, then you get the data
     * @param rowNum the number of row
     * @param colNum the numbe of column
     * @return String cellData
     */
    public String getCellData(int rowNum, int colNum) {
        Cell cell;
        try {
            cell = workSheet.getRow(rowNum).getCell(colNum);
            String cellData = cell.toString();
            return cellData;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * getting all data into two dimentional array and returning the data
     * @return String[][] data of a specific row
     */
    public String[][] getDataArray() {
        String[][] data = new String[rowCount()][columnCount()];
        for (int i = 0; i < rowCount(); i++) {
            for (int j = 0; j < columnCount(); j++) {
                String value = getCellData(i, j);
                data[i][j] = value;
            }
        }
        return data;
    }

    /**
     * going to the first row and reading each row one by one
     * @return List<String> columns
     */
    public List<String> getColumnsNames() {
        List<String> columns = new ArrayList<>();
        for (Cell cell : workSheet.getRow(0)) {
            columns.add(cell.toString());
        }
        return columns;
    }

    /**
     * When you enter the row and column number, returning the value
     * @param value the new data
     * @param rowNum the number of row
     * @param colNum the number of column
     */
    public void setCellData(String value, int rowNum, int colNum) {
        Cell cell;
        Row row;
        try {
            row = workSheet.getRow(rowNum);
            cell = row.getCell(colNum);
            if (cell == null) {//if there is no value, create a cell.
                cell = row.createCell(colNum);
                cell.setCellValue(value);
            } else {
                cell.setCellValue(value);
            }
            FileOutputStream fileOutputStream = new FileOutputStream(path);
            workBook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * When you enter the row number and column name, returning the value
     * @param value the new data
     * @param columnName the name of column
     * @param row the number of row
     */
    public void setCellData(String value, String columnName, int row) {
        int column = getColumnsNames().indexOf(columnName);
        setCellData(value, row, column);
    }

    /**
     * this method will return data table as 2d array
     * so we need this format because of data provider.
     * @return String[][] all_data except first row
     */
    public String[][] getDataArrayWithoutFirstRow() {
        String[][] data = new String[rowCount()-1][columnCount()];
        for (int i = 1; i < rowCount(); i++) {
            for (int j = 0; j < columnCount(); j++) {
                String value = getCellData(i, j);
                data[i-1][j] = value;
            }
        }
        return data;
    }
}
