package ExcelReadWrite;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {

	    public static void main(String[] args) {
	        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
	            XSSFSheet sheet = workbook.createSheet("Sheet1");

	            Object[][] data = {
	                {"Name", "Age", "Email"},
	                {"John Doe", 30, "john@test.com"},
	                {"Jane Doe", 28, "jane@test.com"},
	                {"Bob Smith", 35, "bob@example.com"},
	                {"Swapnil", 37, "swapnil@example.com"}
	            };

	            int rowNum = 0;
	            for (Object[] row : data) {
	                XSSFRow excelRow = sheet.createRow(rowNum++);
	                int colNum = 0;
	                for (Object cellData : row) {
	                    Cell cell = excelRow.createCell(colNum++);
	                    if (cellData instanceof String) {
	                        cell.setCellValue((String) cellData);
	                    } else if (cellData instanceof Integer) {
	                        cell.setCellValue((Integer) cellData);
	                    }
	                }
	            }

	            try (FileOutputStream outputStream = new FileOutputStream("data.xlsx")) {
	                workbook.write(outputStream);
	            }

	            System.out.println("Excel file created successfully.");
	        } catch (Exception e) {
	            e.printStackTrace();
	        }
	        
	    }

}
