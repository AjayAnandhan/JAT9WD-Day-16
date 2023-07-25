package ExcelReadWrite;

import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {

	    public static void main(String[] args) {
	        try (FileInputStream fis = new FileInputStream("data.xlsx");
	             Workbook workbook = new XSSFWorkbook(fis)) {
	            Sheet sheet = workbook.getSheet("Sheet1");

	            for (Row row : sheet) {
	                for (Cell cell : row) {
	                    switch (cell.getCellType()) {
	                        case STRING:
	                            System.out.print(cell.getStringCellValue() + "\t");
	                            break;
	                        case NUMERIC:
	                            System.out.print(cell.getNumericCellValue() + "\t");
	                            break;
	                        default:
	                            System.out.print("\t");
	                            break;
	                    }
	                }
	                System.out.println();
	            }
	        } catch (Exception e) {
	            e.printStackTrace();
	        }
	    }

}
