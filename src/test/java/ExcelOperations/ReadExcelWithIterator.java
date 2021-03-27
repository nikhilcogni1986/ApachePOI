package ExcelOperations;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ReadExcelWithIterator
{
      public static void main(String[] args) throws IOException {
          String filePath = "F:\\ApachePOI\\src\\main\\resources\\TestData\\worldcities.xlsx";
          FileInputStream fis = new FileInputStream(filePath);

          XSSFWorkbook workbook = new XSSFWorkbook(fis);
          XSSFSheet sheet = workbook.getSheetAt(0);

          //Using iterator we fetch the data from the excel
          Iterator itr = sheet.iterator();

          while(itr.hasNext())
          {
              XSSFRow row = (XSSFRow) itr.next();
              Iterator cellIterator = row.cellIterator();

              while(cellIterator.hasNext())
              {
                 XSSFCell cell = (XSSFCell) cellIterator.next();
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