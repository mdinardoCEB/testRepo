import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataChecker {

	XSSFWorkbook templateWb;
	Sheet dataIdentSheet;

	String cellCol = "C";

	public DataChecker(File settingsFile) throws FileNotFoundException,
			IOException {

		templateWb = new XSSFWorkbook(new FileInputStream(settingsFile));
		dataIdentSheet = templateWb.getSheet("Data Identification");

	}

	public void getFormCells() {
		Cell c;
		for (Row r : dataIdentSheet) {
			
			c = r.getCell(CellReference.convertColStringToIndex(cellCol), Row.RETURN_BLANK_AS_NULL);
			
			if (c != null && c.getCellType() == 1) {
				
				System.out.println(c.getRichStringCellValue());
				
			}
			
			else if (c != null  && c.getCellType() == 2) {
				
				System.out.println(c.getCellFormula());
				
			}

		}

	}
	


	public static void main(String[] args) throws FileNotFoundException, IOException {
		//DataChecker test = new DataChecker(new File("CEB Wealth Dashboard Settings.xlsm"));
		//test.getFormCells();
		
		DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyy HH:mm");
		Date dat = new Date();
		System.out.println(dateFormat.format(dat));
		
	}
}
