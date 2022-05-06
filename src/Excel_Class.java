import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.Reporter;
import org.testng.annotations.Test;


public class Excel_Class {
	@Test
 public void Excel() throws Exception {
	 FileInputStream FileInput=new FileInputStream("E:\\Excel_File\\WriteData.xlsx");
	 XSSFWorkbook workbook=new XSSFWorkbook(FileInput);
	 XSSFSheet Sheet=workbook.getSheet("test");
	 System.out.println(Sheet.getSheetName());
	 System.out.println(Sheet.getLastRowNum());
	 System.out.println("Before Updating Cell Data is  "+ Sheet.getRow(2).getCell(1));
	 Reporter.log("Row data is Fetched");
	 System.out.println("Post Updates");
	 System.out.println("Post update 2");
	 //Write Data to Excel File
	 XSSFCell cell=Sheet.getRow(2).getCell(1);
	 cell.setCellValue("Test123456");
	 FileInput.close();
	 FileOutputStream FileOut=new FileOutputStream("E:\\Excel_File\\WriteData.xlsx");
	 Reporter.log("row Data is fetched");
	 workbook.write(FileOut);
	 System.out.println("Updated file After Write is done"+ cell.getStringCellValue());
	 FileOut.close();
 }
}
