package ExcelOperations;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class WriteExcelIterator
{
      public static void main(String[] args) throws IOException {
          //create workbook object
          XSSFWorkbook workbook = new XSSFWorkbook();

          //create the sheet using workbook object
          XSSFSheet sheet = workbook.createSheet("Emp Info");

          //create an array/arraylist/HashMap to store data
          Object empdata[][] = {
                  {"EmpID", "Country", "Population"},
                  {101, "David", "200000"},
                  {101, "Suresh", "244000"},
                  {101, "Mahesh", "203000"},
          };
          int rowCount = 0;

          for(Object emp[]:empdata)
          {
              XSSFRow row = sheet.createRow(rowCount++);
              int columnCount = 0;

              for(Object value :emp)
              {
                  XSSFCell cell = row.createCell(columnCount++);

                  //store the data based on type of data
                  if(value instanceof String)
                      cell.setCellValue((String)value);
                  if(value instanceof Boolean)
                      cell.setCellValue((Boolean) value);
                  if(value instanceof Integer)
                      cell.setCellValue((Integer)value);
              }
          }
          //here one row of data is written and this continues for as many rows present
          String filePath = "F:\\ApachePOI\\src\\main\\resources\\TestData\\cities1.xlsx";
          FileOutputStream fout = new FileOutputStream(filePath);
          workbook.write(fout);
          fout.close();
          System.out.println("File is written successfully!");
      }
}