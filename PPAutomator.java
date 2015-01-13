import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class PPAutomator {
	
	String mainBankName = "";
	String sourcePath = "C:\\Users\\mdinardo\\Documents\\DashboardAutomations\\AEDAutomation.xlsm";

	
	File templateFile = new File(
			"C:/Users/mdinardo/Documents/DashboardAutomations/testOutput/tester.txt");
	//File dataFile = new File("AEDDataTable.xlsx");
	//File peerFile = new File("AEDPeerGroups.xlsx");
	PeerList peerGroups;
	//XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(dataFile));
	//Sheet ws = wb.getSheet("Data");
	
	ArrayList<CopyObject> objectList = new ArrayList<CopyObject>();
	
	ArrayList<String> bankRows = new ArrayList<String>();
	ArrayList<String> peerRows = new ArrayList<String>();
	
	private int execRow = 0;
	ArrayList<String> managerRows = new ArrayList<String>();
	
	ArrayList<String> boxPositions = new ArrayList<String>();
	
	
	public PPAutomator() throws FileNotFoundException, IOException {
		
		
		
	}
	
	
	public void parseTemplateFile() throws FileNotFoundException {
		
		Scanner inputFile = new Scanner(templateFile);
		
		while (inputFile.hasNextLine()) {
			String line = inputFile.nextLine();
			
			if (line.length() > 1) {
				
				String[] dataSplit = line.split("_");
				
				String[] slideNumArray = dataSplit[0].split(":");
				String[] titleArray = dataSplit[1].split(":");
				String[] leftArray = dataSplit[2].split(":");
				String[] topArray = dataSplit[3].split(":");
				String[] widthArray = dataSplit[4].split(":");
				String[] heightArray = dataSplit[5].split(":");
				
				
				
				int slideNum = Integer.parseInt(slideNumArray[1]);
				String title = titleArray[1];
				String left = leftArray[1];
				String top = topArray[1];
				String width = widthArray[1];
				String height = heightArray[1];
				
				if (title.charAt(0) == '#') {
					
					String sheetName = titleArray[1];
					String tableName = titleArray[2];
					CopyObject tableObject = new CopyObject(sheetName, tableName, slideNum, left, top, width, height, mainBankName);
					
					objectList.add(tableObject);
					
				}
				
				else if (title.charAt(0) == '$') {
					
					String sheetName = titleArray[1];
					String range = titleArray[2];
					CopyObject tableObject = new CopyObject(sheetName, range, slideNum, left, top, width, height);
					
					objectList.add(tableObject);
					
				}
				
				else if ((title.length() > 5) && title.substring(0,5).equals("SLIDE")) {
					
					String positionInfo = "Slide number:" + slideNum + ",Title(Index):" + titleArray[2].substring(8, titleArray[2].length()) + ",Left:" + left + ",Top:" + top + ",Width:" + width + ",Height:" + height;
					System.out.println(positionInfo);
					boxPositions.add(positionInfo);
					
				}
				
				else {
				
				CopyObject chartObject = new CopyObject(slideNum, title, left, top, width, height, mainBankName);
				objectList.add(chartObject);
				
				}
			}

		}

	}
	
	public String replaceBankMetrics() {
		int bankMeanRow = 288;
		int peerMeanRow = 289;
		String code = "";
		
		// Change bank Means Title IMPORTANT
		
		code += "\nSheets(\"Data\").Select\nRange(\"A288\").Select\nActiveCell.FormulaR1C1 = \"" + mainBankName + "\"\n";
		
		//Change Means row
		code += "\nRange(\"B288\").Select\n" + "ActiveCell.FormulaR1C1 = \"=AVERAGE(";
		// Copy bank-mean into Bank Row
		for (int x = 0; x < bankRows.size(); x++) {
			
			code += "R[" + (Integer.parseInt(bankRows.get(x)) - bankMeanRow) + "]C";
			
			if (x < (bankRows.size() - 1)) {
				
				code += ",";
				
			}
		}
		
		code += ")\"\n" + "Range(\"B288\").Copy\nRange(\"B288:LK288\").Select\nActiveSheet.Paste\n";
		
		
		// Copy peer mean into peer row
		code += "\nRange(\"B289\").Select\n" + "ActiveCell.FormulaR1C1 = \"=AVERAGE(";
		
		for (int x = 0; x < peerRows.size(); x++) {
			
			code += "R[" + (Integer.parseInt(peerRows.get(x)) - peerMeanRow) + "]C";
			
			if (x < (peerRows.size() - 1)) {
				
				code += ",";
				
			}
			if (x % 40 == 0 && x > 0) {
				
				code += "\" & _\n\"";
				
				}
		}
		
		code += ")\"\n" + "Range(\"B289\").Copy\nRange(\"B289:LK289\").Select\nActiveSheet.Paste\n";
		
		
		return code;

	}

	public String replaceFrequency(String catNumber, int rowNumber,
			ArrayList<String> rowList) {

		String code = "Range(\"B" + rowNumber + "\").Select\n"
				+ "ActiveCell.FormulaR1C1 = \"=";

		for (int x = 0; x < rowList.size(); x++) {
			if (x != 0) {
				code += "+";
			}
			
			if (x % 40 == 0 && x > 0) {
				
				code += "\" & _\n\"";
				
			}
			code += "COUNTIF(R["
					+ (Integer.parseInt(rowList.get(x)) - rowNumber) + "]C,"
					+ catNumber + ")";
			

		}
		code += "\"\nRange(\"B" + rowNumber + "\").Copy\nRange(\"B" + rowNumber
				+ ":LK" + rowNumber + "\").Select\nActiveSheet.Paste\n";
		
		return code;
	}

	public void callAllObjects() throws FileNotFoundException {
		
		
		
		String output = "Sub callAllObjects(sourceString As String, saveString As String)\n\n"
				+ "Dim PowerPointApp As PowerPoint.Application\n"
				+ "Dim activeSlide As PowerPoint.Slide\n"
				+ "Set PowerPointApp = CreateObject(\"PowerPoint.Application\")\n"
				+ "PowerPointApp.Visible = True\n"
				+ "PowerPointApp.Presentations.Open (sourceString)\n\n";

		for (int i = 0; i < objectList.size(); i++) {

			output += objectList.get(i).toString();

		}

		output += "\nPowerPointApp.ActivePresentation.SaveAs (saveString)\nPowerPointApp.ActivePresentation.Close\nPowerPointApp.Quit\nEnd Sub";

		System.out.println(output);
		PrintWriter writer = new PrintWriter(".\\WMLC\\Resources\\Macros\\CallPPObjects.bas");
		writer.println(output);
		writer.close();

	}

	public String changeSplits() {

		// Executive Split
		String code = "";
		code += "Sheets(\"Executive Split\").Select\n"
				+ "Range(\"B290\").Select\nActiveCell.FormulaR1C1 = \"=R["
				+ (execRow - 290)
				+ "]C\"\nSelection.Copy\nRange(\"B290:LK290\").Select\nActiveSheet.Paste\n";
		
		
		// Manager Split

		if (managerRows.size() > 0) {

			code += "Sheets(\"Manager Split\").Select\n"
					+ "Range(\"B290\").Select\nActiveCell.FormulaR1C1 = \"=AVERAGE(";

			for (int x = 0; x < managerRows.size(); x++) {

				code += "R[" + (Integer.parseInt(managerRows.get(x)) - 290)
						+ "]C";

				if (x < (managerRows.size() - 1))
					code += ",";

			}

			code += ")\"\nRange(\"B290\").Copy\nRange(\"B290:LK290\").Select\nActiveSheet.Paste\n";
		}

		else if (managerRows.size() == 0) {
			code += "Sheets(\"Manager Split\").Select\n"
					+ "Range(\"B290\").Select\n"
					+ "ActiveCell.FormulaR1C1 = \"0\"\n"
					+ "Range(\"B290\").Copy\n"
					+ "Range(\"B290:LK290\").Select\nActiveSheet.Paste\n";
		}

		return code;

	}

	
	/*public void reportGeneratorScript() throws FileNotFoundException {

		String script = "Dim objExcel, objWorkbook\n"
				+ "Set objExcel = CreateObject(\"Excel.Application\")\n"
				+ "objExcel.Visible = True\n"
				+ "Set objWorkbook = objExcel.Workbooks.Open(\"" + sourcePath + "\")";
						
		for (int x = 0; x < peerGroups.getBankList().size(); x++) {
			
			script += "\nobjExcel.VBE.ActiveVBProject.VBComponents.Import \"C:\\Users\\mdinardo\\workspace\\NewPowerPointAuto\\AED Macros\\" + peerGroups.getBankList().get(x) + ".bas\"";
			
		}
		
		script += "\nobjExcel.VBE.ActiveVBProject.VBComponents.Import \"C:\\Users\\mdinardo\\workspace\\NewPowerPointAuto\\AED Macros\\runAll.bas\"" + "\nobjExcel.Application.Run \"runAll\"";
		
		System.out.println(script);
		PrintWriter writer = new PrintWriter("ReportGenerator.vbs");
		writer.println(script);
		writer.close();
		
		for (int x = 0; x < peerGroups.getBankList().size(); x++) {
			
			System.out.println(peerGroups.getBankList().get(x));
			
		}
		
	}*/
	
	

	

	public void setBankName(String bankName) {
		this.mainBankName = bankName;
	}
	
	

	public static void main(String[] args) throws IOException {

		PPAutomator test = new PPAutomator();
		test.parseTemplateFile();
		test.callAllObjects();
		//test.setBankName("First Midwest Bank");
		/*test.parseTemplateFile();
		test.callAllObjects();
		test.scanAEDDataFile();
		test.updateDataSheet();
		test.setAnswerPosition();*/
		//test.makeMacrosForAllBanks();
		//test.runAllMacro();
		//test.reportGeneratorScript();
		//test.callAllObjects();
		
	}

}
