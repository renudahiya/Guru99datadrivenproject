package datadriventestingusingPOI;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ReadGuru99ExcelFile {
    public void readExcel(String filePath,String fileName,String sheetName) throws IOException {
       //1. create an object of FILE class to open xlsx file

        File file=new File(filePath+"\\"+fileName);

        //2. create an object of FileInputStream class to read excel file
        FileInputStream inputStream=new FileInputStream(file);

        Workbook guru99Workbook=null;

        //find the file extension by splitting fileName in substring and getting only extension
        String fileExtensionName=fileName.substring(fileName.indexOf("."));

        if(fileExtensionName.equals(".xlsx")){
            //3. create object of XSSFWorkbook class
            guru99Workbook=new XSSFWorkbook(inputStream);
        }
        else
            if(fileExtensionName.equals(".xls")){
                guru99Workbook=new HSSFWorkbook(inputStream);

            }
            //4. read the sheet inside the workbook
        Sheet guru99Sheet=guru99Workbook.getSheet(sheetName);

            //5. find the number of rows in excel file
        int rowCount=guru99Sheet.getLastRowNum()-guru99Sheet.getFirstRowNum();


        //6. create a loop over all rows to read data
        for(int i=0;i<rowCount+1;i++){
            Row row=guru99Sheet.getRow(i);
            //7. create a loop to print cell values in a row
            for(int j=0;j<row.getLastCellNum();j++){
                System.out.print(row.getCell(j).getStringCellValue()+"||");

            }
           System.out.println();
        }
    }
    public static void main(String[] args) throws IOException{
        //create an object of ReadGuru99ExcelFile class
        ReadGuru99ExcelFile objectExcelFile=new ReadGuru99ExcelFile();

        //Prepare the path of excel file
        String filePath=System.getProperty("user.dir")+"\\src\\main\\resources";

        objectExcelFile.readExcel(filePath,"Guru99Book.xlsx","Sheet1");
    }
}
