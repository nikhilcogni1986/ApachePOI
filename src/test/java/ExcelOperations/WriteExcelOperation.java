package ExcelOperations;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class WriteExcelOperation
{
      public static void main(String[] args) throws IOException {
          //create workbook object
          XSSFWorkbook workbook = new XSSFWorkbook();

          //create the sheet using workbook object
          XSSFSheet sheet = workbook.createSheet("Emp Info");

          //create an array/arraylist/HashMap to store data
          Object data[][] = {
                  {"EmpID", "Country", "Population"},
                  {101, "David", "200000"},
                  {101, "Suresh", "244000"},
                  {101, "Mahesh", "203000"},
          };

          int rows = data.length;
          int cols = data[0].length;

          System.out.println("No of rows in the array:"+rows);
          System.out.println("No of columns in the array:"+cols);

          for(int r=0 ; r<rows ; r++)
          {
              //create row object
              XSSFRow row = sheet.createRow(r);

              for(int col=0; col<cols ;col++)
              {
                 XSSFCell cell = row.createCell(col);
                 Object value = data[r][col];

                 //store the data based on type of data
                  if(value instanceof String)
                      cell.setCellValue((String)value);
                  if(value instanceof Boolean)
                      cell.setCellValue((Boolean) value);
                  if(value instanceof Integer)
                      cell.setCellValue((Integer)value);
              }
          }//here one row of data is written and this continues for as many rows present
          String filePath = "F:\\ApachePOI\\src\\main\\resources\\TestData\\cities.xlsx";
          FileOutputStream fout = new FileOutputStream(filePath);
          workbook.write(fout);
          fout.close();
          System.out.println("File is written successfully!");
      }
}