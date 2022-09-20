package datadriventestingusingPOI;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteGuru99ExcelFile {
    public void writeExcel(String filePath, String fileName, String sheetName, String[] valueToWrite) throws IOException {
        //create file object
        File file=new File(filePath+"\\"+fileName);

        //create FileInputStream object
        FileInputStream inputStream=new FileInputStream(file);

        Workbook guru99Workbook=null;

        //find the file extension by splitting fileName in substring and getting only extension
        String fileExtensionName=fileName.substring(fileName.indexOf("."));

        if(fileExtensionName.equals(".xlsx")){
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
        //get first row
        Row row=guru99Sheet.getRow(0);
        //create a new  row
        Row newRow=guru99Sheet.createRow(rowCount+1);

        //create a loop over the cell of newly created Row
        for(int j=0;j<row.getLastCellNum();j++){
            //fill data in the row
            Cell cell=newRow.createCell(j);
            cell.setCellValue(valueToWrite[j]);
        }
        //close input stream
        inputStream.close();

        //create an object of FileOutputStream
        FileOutputStream outputStream=new FileOutputStream(file);

        //write data in the excel file
        guru99Workbook.write(outputStream);
        //close outputstream
        outputStream.close();
    }
    public static void main(String[] args)throws IOException {
        //create an array with the data
        String[] valueToWrite = {"F", "Noida"};

        //create an object of current class
        WriteGuru99ExcelFile object = new WriteGuru99ExcelFile();

        //write the file
        String filePath=System.getProperty("user.dir")+"\\src\\main\\resources";
        object.writeExcel(filePath,"Guru99Book.xlsx",
                "sheet1", valueToWrite);
    }
    }

