package excelExample;


import java.io.FileInputStream;


import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;


public class ExcelReader {
																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																													
	
		public static HSSFWorkbook workbook;
		public static HSSFSheet worksheet;
		public static DataFormatter formatter = new DataFormatter();
		static String SheetName = "SheetName";
		public  static String file_location = System.getProperty("user.dir")+"/WorkSpace";
	
		
		
	
	public static void main(String[] args) throws Throwable  {
		// TODO Auto-generated method stub

	
		FileInputStream fis = new FileInputStream(file_location);
		workbook = new HSSFWorkbook(fis);
		worksheet =  workbook.getSheet(SheetName);
		HSSFRow  Row = worksheet.getRow(0);
		
		int rowNum = worksheet.getPhysicalNumberOfRows();
		int colNum = Row.getLastCellNum();
		
		Object Data[][] = new Object[rowNum-1][colNum];
		
		for(int i=0; i<rowNum-1;i++) {
			
			HSSFRow row = worksheet.getRow(i+1);
			
			for(int j=0;j<colNum;j++) {
				
				if (row == null)
					Data[i][j] = "";
				else
				{
					HSSFCell cell = row.getCell(j);
					if (cell==null)
						Data[i][j] = "";
						
				
				else
				{
					String value = formatter.formatCellValue(cell);
					Data[i][j] =  value ;
					System.out.println(value);
					
				}
				
					
			}
		}
		
			
					
			
		}
		}
		
		
	
}






	
	
	


	


