package ExcelOperations;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ReadExcelData
{
      public static void main(String[] args) throws IOException {
          // declare filepath of the excel to be read
            String filePath = "F:\\ApachePOI\\src\\main\\resources\\TestData\\worldcities.xlsx";
            System.out.println(filePath);

            FileInputStream fis = new FileInputStream(filePath);
            //declare the workbook object
            XSSFWorkbook workbook = new XSSFWorkbook(fis);

            //get the sheet for the above workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            //get the number of rows and columns
            int rows = sheet.getLastRowNum();
            int cols = sheet.getRow(1).getLastCellNum();

            //use for loop to iterate over the rows and columns
            for(int r=0 ; r<rows ; r++)
            {
                  XSSFRow row = sheet.getRow(r);

                  for(int c=0 ; c<cols ; c++)
                  {
                        XSSFCell cell = row.getCell(c);

                        //get the types of the cell and perform operations based on the same
                        switch(cell.getCellType())
                        {
                              case 1:
                                     System.out.print(cell.getStringCellValue());
                                     break;
                              case 0:
                                     System.out.print(cell.getNumericCellValue());
                                     break;
                              default:
                                     System.out.print("No Cell types matched");
                                     break;
                        }
                        System.out.print(" | ");
                  }
                  System.out.println(" ");
            }
      }
}
