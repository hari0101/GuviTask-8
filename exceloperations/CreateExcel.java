package exceloperations;

//Java in-built Classes
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

//Import Apache POI dependency.
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/* FILE OPERATIONS

	Write a Java program to write data to an Excel file using Apache POI library.

*/

public class CreateExcel //Class name 

{

	public static void main(String[] args) throws IOException, Throwable
	{
		ArrayList<Object[]> data= new ArrayList<>(); //Declaring the Arraylist as name data. 
		
		//Passing the array type of values below. 
		data.add(new Object[] {"Name",     "Age",      "Email"});
		data.add(new Object[] {"John Doe",   30, "john@test.com"});
		data.add(new Object[] {"Jane Doe",   28, "john@test.com"});
		data.add(new Object[] {"Bob Smith",  35, "jacky@example.com"});
		data.add(new Object[] {"Swapnil  ",  37, "swapnil@example.com"});
		
		//Calling the method name write.
		CreateExcel.write(data);
		
	
	}
	
	public static void write(ArrayList<Object[]> data) throws Throwable, IOException //Declaring the method name write
	{
	// Creating the Workbook using the Apache POI Class XSSFWorkbook.	
	XSSFWorkbook workbook = new XSSFWorkbook();
	// Creating the sheet using the Class XSSFSheet.
	XSSFSheet sheet = workbook.createSheet("Sheet1");
	
	
	int rowCount = 0; //Initializing variable.
	
	// Using For Each loop to pass the ArrayList data.
	for(Object[] d : data) 
	{
		//Creating the required Rows by using Class XSSFRow.
		XSSFRow rows = sheet.createRow(rowCount++);
		
	int cellcount = 0; //Initializing variable.	
		//Outer foreach loop creating required columns based on cell.
		for(Object column : d)
		{
			//Creating the required Cells based on Arraylist we passed by using Class XSSFCell.
			XSSFCell cell = rows.createCell(cellcount++);
				
			//Using If-else 
						if(column instanceof String)
						{
							cell.setCellValue((String) column); //Set string value to cell.
						}else 
						if(column instanceof Integer) 
						{
							cell.setCellValue((Integer) column); // else set cell value as Integer.
						}
		
		}// innerLoop
		
		}//outerLoop 
	
	String filepath = "D:\\Utils\\Create_the_Excel.xlsx"; // Initializing the string to store the path of file
	// Creating object to get the excel file in filepath.
	File file = new File(filepath); 
	// Creating object of Fileoutputstream to write the data in excel file.
	try(FileOutputStream foutput = new FileOutputStream(file))
	{
		//Write or Save data in workbook
		workbook.write(foutput);
		//Close the function to avoid the data lose.
		workbook.close();
	}
	
	}
}
