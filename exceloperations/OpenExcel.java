package exceloperations;

//Java in-built Classes
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

//Imported Apache POI Classes.
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

/*FILE OPERATION
 	
 	Write a Java program to read data from an Excel file using Apache POI library and print it in the console.

*/

public class OpenExcel //Class name 
{

public static void main(String[] args) throws IOException 
{
	String path = "D:\\Excel read.xlsx"; //Intialized string to store the path.
	
	//Try-Catch the exception, if the data are not available in excel file.
	try {
	
		OpenExcel.ReadExcel(path); //Invoking the method by passing the path.
	
	}
	catch(FileNotFoundException ex)
	{
		System.out.println("ERROR");	
	}
}

	public static void ReadExcel(String path) throws IOException //New method name ReadExcel
	{
		// Fileinputstream creates the object to read the excel data.
		FileInputStream filein = new FileInputStream(path);
		// Using apache poi library XSSFWorkbook used open spreadsheet in the excel.
		XSSFWorkbook WB = new XSSFWorkbook(filein);
		// Using apache poi library XSSFSheet, we are passing the index number to begin reading in sheet. . 
		XSSFSheet sheet = WB.getSheetAt(0);
		// DataFormatter is Apache POI library helps to convert any datatype.
		DataFormatter convert = new DataFormatter();
		
		//Using the for-each pass the sheet with XSSFRow
		for(Row rows : sheet)
		{
			// Inner loop to check the end of row data in excel.
			for(Cell cell : rows)
			{
				//Invoking the convert object to store any value as String type : s.
				String s = convert.formatCellValue(cell);
				System.out.print(s + "\t");
			}
		System.out.println();
		}
		WB.close(); // To prevent the resource leaks.
		
	}	

}	