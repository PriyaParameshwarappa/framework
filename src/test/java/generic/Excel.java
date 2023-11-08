package generic;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Excel {

	public static String getData(String path,String sheet,int r,int c)
	{
		String v="";
		try 
		{
			Workbook wb = WorkbookFactory.create(new FileInputStream(path));
			v=wb.getSheet(sheet).getRow(r).getCell(c).toString();
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
		
		return v;
	}
	
	public static int getRowCount(String path,String sheet)
	{
		int rowCount=0;
		
		try
		{
			Workbook wb = WorkbookFactory.create(new FileInputStream(path));
			rowCount=wb.getSheet(sheet).getLastRowNum();
		}
		catch (Exception e) 
		{
			e.printStackTrace();
		}
		
		return rowCount;
	}
	
	//add a method to count columns
	
	public static int getColumnCount(String path,String sheet,int rowCount) {
		
		int rowCount1=0;
		int  columnCount=0;
				
		try {
			
			Workbook wb = WorkbookFactory.create(new FileInputStream(path));
			columnCount=wb.getSheet(sheet).getRow(rowCount1).getLastCellNum();
			
		} catch (Exception e) {
			e.printStackTrace();// TODO: handle exception
		}
		
		return columnCount;
		
	}
	
	//method to write the data
	
}
