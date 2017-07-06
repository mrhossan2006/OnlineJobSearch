package utility;


import java.io.FileInputStream;

import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	private static FileInputStream ExcelFile;
	private static XSSFWorkbook ExcelWBook;	
	private static XSSFSheet ExcelWSheet;
	private static XSSFCell Cell;
	private static XSSFRow SheetRow;
	private static FileOutputStream fileOut;

	public static void SetExcelFile(String path , String SheetName ) throws Exception{
		try{
		
		ExcelFile = new FileInputStream(path);
		ExcelWBook = new XSSFWorkbook(ExcelFile);
		ExcelWSheet = ExcelWBook.getSheet(SheetName);
		}
		catch(Exception e){
			Log.error("excel File is not exist");
			throw(e);
		}


	}
	public static String GetCellData(int  RowNum, int ColNum) throws Exception{
		String cellData = "";
		try{
			Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
			cellData = Cell.getStringCellValue();
			ExcelWBook.close();
			ExcelFile.close();
		}
		catch(Exception e){
			Log.error("Cell Value is Empty");
			throw(e);
			
		}
		return cellData;
	}

	public static void SetCellData( int  RowNum, int ColNum, String Result ) throws Exception{
		
		try{
			//review the row and check for null
			SheetRow = ExcelWSheet.getRow(RowNum);
			if (SheetRow==null){
				SheetRow = ExcelWSheet.createRow(RowNum);
			}
			//update the value of the cell
			Cell = SheetRow.getCell(ColNum);
			if(Cell==null){
				Cell = SheetRow.createCell(ColNum);
			}
			Cell.setCellValue(Result);
			fileOut = new
		FileOutputStream(Constant.Path_TestData + Constant.File_TestData);
        ExcelWBook.write(fileOut);
		fileOut.flush();
		fileOut.close();
   }

	 catch(Exception e){
		 throw(e);
	 }
	}
}
