package ExcelOperations;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class HashMapToExcel
{
      public static void main(String[] args) throws IOException {
          XSSFWorkbook workbook = new XSSFWorkbook();
          XSSFSheet sheet = workbook.createSheet("Student Data");

          //create a Hashmap and add values to it.
          HashMap<String,String> map1 = new HashMap<String, String>();
          map1.put("101","John");
          map1.put("201","Smith");
          map1.put("301","James");
          map1.put("401","Matthew");
          map1.put("501","Kim");

          System.out.println(map1);
          int rowNumber = 0;

          for(Map.Entry entry :map1.entrySet())
          {
                XSSFRow row = sheet.createRow(rowNumber++);
                row.createCell(0).setCellValue((String)entry.getKey());
                row.createCell(1).setCellValue((String)entry.getValue());
          }
          FileOutputStream fout = new FileOutputStream("F:\\ApachePOI\\src\\main\\resources\\TestData\\StudentData.xlsx");
          workbook.write(fout);
          fout.close();


      }
}
