package ExcelOperations;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

	public static void main(String[] args) {
		//provide filename from where you want to read data
		String filepath="data.xlsx";
		
		try {
			//read the data from file 
          FileInputStream inpstream=new FileInputStream(filepath);
		
          //create an object of WorkBook
          Workbook workbook=new XSSFWorkbook(inpstream);
          
           Sheet sheet=workbook.getSheetAt(0);    
           Row row=  sheet.getRow(1);//read row data
           //enter cell to read
           Cell cell=row.getCell(0);
           
           if(cell.getCellType()==CellType.STRING) {
        	   System.out.println(cell.getStringCellValue());
           }else {
        	   System.out.println(cell.getNumericCellValue());

           }
           workbook.close();
           inpstream.close();
           
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

}

