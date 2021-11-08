package utils;

import org.apache.poi.xssf.usermodel.*;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.json.JSONException;
import org.json.JSONObject;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.LinkedHashMap;

public class ReadExcelData
{
    public static void main(String args[]) throws IOException, JSONException
    {
        ReadExcelData read = new ReadExcelData();
        read.getRowCount();
    }

    private void getRowCount() throws IOException, JSONException
    {

        String excelFilePath = "C:\\Users\\prasanna.prabakaran\\Downloads\\MarkSheet.xlsx";

        FileInputStream inputStream = new FileInputStream(excelFilePath);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);
       String name= String.valueOf(workbook.getSheetAt(0).getRow(0).getCell(0));
       String age= String.valueOf(workbook.getSheetAt(0).getRow(0).getCell(1));
       String marks= String.valueOf(workbook.getSheetAt(0).getRow(0).getCell(2));

        // Using FOR loop

        int lengthOfRows = sheet.getLastRowNum();
        int lengthOfCells = sheet.getRow(1).getLastCellNum();

       // LinkedHashMap<String, XSSFCell> data =new LinkedHashMap<>();
        JSONObject jsonObject = new JSONObject();
        for (int rowIndex = 1; rowIndex < lengthOfRows; rowIndex++)
        {
            XSSFRow rowObj = sheet.getRow(rowIndex);
            for (int cellIndex = 0; cellIndex < lengthOfCells; cellIndex++)
            {
                XSSFCell cellObj = rowObj.getCell(cellIndex);
                //System.out.println(cellObj);
                switch (cellIndex){
                    case 0:
                        jsonObject.put(name,cellObj);
                        break;
                    case 1:
                        jsonObject.put(age,cellObj);
                        break;
                    case 2:
                        jsonObject.put(marks,cellObj);
                        break;

                }

            }
            System.out.println(jsonObject);
           // System.out.println();


        }
    }
}

     /* Iterator iterator= sheet.iterator();
      while (iterator.hasNext()){
         XSSFRow row= (XSSFRow) iterator.next();
        Iterator cellIterator= row.cellIterator();
        while (cellIterator.hasNext()){
           XSSFCell cell= (XSSFCell) cellIterator.next();
            switch (cell.getCellType()){
                case STRING:
                    //System.out.println("1");
                    System.out.print(cell.getStringCellValue());
                    break;

                case NUMERIC:
                    System.out.print(cell.getNumericCellValue());
                    break;
            }
            System.out.print("   ");


        }
          System.out.println();

      }*/



