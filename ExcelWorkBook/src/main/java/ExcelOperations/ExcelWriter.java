package ExcelOperations;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter {

	public static void main(String[] args) {
		
		//create a new excel file->create object of XSSF Work book
		
		//handle exceptions
		try(XSSFWorkbook workbook=new XSSFWorkbook()) {
			
			Sheet sheet=workbook.createSheet("Sheet1");
			
			
			//Define data data which we want to insert to excel file
			//create object 2-D Array 
			
			
			Object[][] data= {
					{"Name","Age","Email"},
					{"John Doe",30,"john@test.com"},
					{"John Doe",28,"john@test.com"},
					{"Bob Smith",35,"bob@example.com"},
					{"Swapnil",37,"swap@example.com"}
			};
			
			//traverse across the data of the array and write into sheet
			int rowNum=0;
			for(Object[] rowData:data) {
				Row row=sheet.createRow(rowNum++);//new row in sheet1
				int colNum=0;
				for(Object field:rowData) {
					Cell cell=row.createCell(colNum++);
					//check type of data 
					if(field instanceof String) {
						cell.setCellValue((String)field);
					}else if(field instanceof Integer) {
						cell.setCellValue((Integer) field);
					}
				}
				
			}
			//write wrokbook data content into excel file
			try(FileOutputStream outputstream=new FileOutputStream("data.xlsx"))
			{
				workbook.write(outputstream);
			}
			
			System.out.println("Excel file created successfully");
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}

}

