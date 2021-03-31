package ExcelOperations;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class ExcelToHashMap
{
      public static void main(String[] args) throws IOException
      {
            FileInputStream fis = new FileInputStream(
                    "F:\\ApachePOI\\src\\main\\resources\\TestData\\StudentData.xlsx");

          XSSFWorkbook workbook = new XSSFWorkbook(fis);
          XSSFSheet sheet = workbook.getSheet("Student Data");
          int numberOfRows = sheet.getLastRowNum();
          System.out.println(numberOfRows);

          //create Hashmap
          HashMap<String, String> data = new HashMap<String, String>();

          for(int r=0 ; r<numberOfRows; r++)
          {
              String key = sheet.getRow(r).getCell(0).getStringCellValue();
              String value = sheet.getRow(r).getCell(1).getStringCellValue();
              data.put(key,value);
          }
          System.out.println("Data retrieved from Excel is:");
          for(Map.Entry entry :data.entrySet())
          {
              System.out.println(entry.getKey() + "  "+entry.getValue());
          }
      }
}