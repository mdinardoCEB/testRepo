import java.awt.Color;
import java.awt.Component;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.util.ArrayList;

import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ParseIntake implements ActionListener {
	
	JFrame mainFrame = new JFrame("DashDesigner");
	JFrame chooseFrame = new JFrame("Select File");
	JLabel windowLab = new JLabel("<html><h1>Generating script, this may take some time</h1></html>");
	JLabel prelimDataLab = new JLabel("Writing preliminary table commands... ");
	JLabel peerDataLab = new JLabel("Parsing peer group data... ");
	JLabel dashDataLab = new JLabel("Writing Dashboard table commands... ");
	JLabel dashTabLab = new JLabel("Generating Dashboard... ");
	JLabel loadingLab = new JLabel("Script succesfully written. Executing Commands...");
	//PeerList peerList = new PeerList(new File("PeerGroupsFY2013.xlsx"));
	PeerList peerList;
	
	JFileChooser fc = new JFileChooser();
	
	int numberBanks = 0;
	int numberSheetsTable = 3;
	String startDataColumn = "E";
	String lastDataColumn;
	
	//File intakeForm = new File(
	//		"Corrected_Associated CEB Wealth Industry Benchmark Intake Form FY 2013_MASTER_3-6-13.xlsx");
	File toolForm;
	// File blankWorkbook = new File("blankForm.xlsx");
	Sheet mainSheet;
	Sheet dataSheet;
	XSSFWorkbook wb;
	XSSFWorkbook prelimTable;
	XSSFWorkbook toolWb;

	Sheet toolSheet;
	ArrayList<File> intakeFileList = new ArrayList<File>();
	ArrayList<String> cellList = new ArrayList<String>();
	ArrayList<String> bankNameList = new ArrayList<String>();
	
	String prelimScriptBody;
	
	String practiceName = "WMLC";
	
	boolean completeRunthrough = true;

	private boolean firstSheet = true;

	public ParseIntake(File toolForm) throws FileNotFoundException, IOException {
		
		this.toolForm = toolForm;
		System.out.println(toolForm == null);
		//wb = new XSSFWorkbook(new FileInputStream(intakeForm));
		//mainSheet = wb.getSheet("Data Intake Form");

		toolWb = new XSSFWorkbook(new FileInputStream(toolForm));
		toolSheet = toolWb.getSheet("Preview Table");

		
		
		prelimTable = new XSSFWorkbook(new FileInputStream("C:\\Users\\mdinardo\\Desktop\\CopyPasteTest\\PrelimtesterSAVED.xlsm"));
		dataSheet = prelimTable.getSheet("Data");
		
		boolean bankList = true;
		int x = 1;
		System.out.println("Here1");
		while (bankList == true) {
			if (toolSheet.getRow(x).getCell(0, Row.RETURN_BLANK_AS_NULL) != null) {
				bankNameList.add(toolSheet.getRow(x).getCell(0)
						.getRichStringCellValue().toString());
				numberBanks++;
			} else if (toolSheet.getRow(x).getCell(0, Row.RETURN_BLANK_AS_NULL) == null) {
				bankList = false;
			}
			x++;
		}
		System.out.println("Here2");
		boolean dataColumn = true;
		int cellCount = 0;
		while (dataColumn == true) {
			if (toolSheet.getRow(0).getCell(cellCount, Row.RETURN_BLANK_AS_NULL) != null) {
				lastDataColumn = CellReference.convertNumToColString(toolSheet.getRow(0).getCell(cellCount).getColumnIndex());
				cellCount++;
			}
			else if (toolSheet.getRow(0).getCell(cellCount, Row.RETURN_BLANK_AS_NULL) == null) {
				dataColumn = false;
			}
		}

		System.out.println(numberBanks);
		System.out.println(lastDataColumn);
		/*
		 * FormulaEvaluator evaluator =
		 * wb.getCreationHelper().createFormulaEvaluator();
		 * 
		 * for (int x = 30; x < 151; x++) { Row r = mainSheet.getRow(x); Cell
		 * cell = r.getCell(CellReference.convertColStringToIndex("K"));
		 * 
		 * if (cell!=null) { switch (evaluator.evaluateFormulaCell(cell)) { case
		 * Cell.CELL_TYPE_BOOLEAN:
		 * System.out.println(cell.getBooleanCellValue()); break; case
		 * Cell.CELL_TYPE_NUMERIC:
		 * System.out.println(cell.getNumericCellValue()); break; case
		 * Cell.CELL_TYPE_STRING:
		 * System.out.println(cell.getRichStringCellValue()); break; case
		 * Cell.CELL_TYPE_BLANK: break; case Cell.CELL_TYPE_ERROR:
		 * System.out.println(cell.getErrorCellValue()); break;
		 * 
		 * // CELL_TYPE_FORMULA will never occur case Cell.CELL_TYPE_FORMULA:
		 * break; } } }
		 */

	}
	
	public void createAndShowGUI() {
		
		fc.setFileSelectionMode(JFileChooser.FILES_ONLY);
		fc.addActionListener(this);
		
		mainFrame.setSize(600,200);
		mainFrame.getContentPane().setLayout(new GridLayout(0,1));
		mainFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		
		
		windowLab.setAlignmentX(Component.CENTER_ALIGNMENT);
		peerDataLab.setAlignmentX(Component.CENTER_ALIGNMENT);
		dashDataLab.setAlignmentX(Component.CENTER_ALIGNMENT);
		dashTabLab.setAlignmentX(Component.CENTER_ALIGNMENT);
		
		
		peerDataLab.setForeground(Color.GRAY);
		dashDataLab.setForeground(Color.GRAY);
		dashTabLab.setForeground(Color.GRAY);
		
		//mainFrame.add(windowLab);
		mainFrame.add(prelimDataLab);
		mainFrame.add(peerDataLab);
		mainFrame.add(dashDataLab);
		mainFrame.add(dashTabLab);
		
		mainFrame.setLocationRelativeTo(null);
		mainFrame.setVisible(true);
		
	}
	
	public static void chooseDatabase() throws FileNotFoundException, IOException {
		
		JFrame parent = new JFrame();
		JFileChooser fc2 = new JFileChooser();
		int returnVal = fc2.showOpenDialog(parent);
		File toolForm = fc2.getSelectedFile();
		
		ParseIntake program = new ParseIntake(toolForm);
		program.executeAll();
	}
	
	
	
	// Pastes all the rows from the preview table to the preliminary table
	public void prelimTableMacro() throws IOException {
		String macro = "Sub genPrelimTable()\n\nDim source as Workbook\nDim dest As Workbook\nSet source = Workbooks.Open(\"" + toolForm.getCanonicalPath() + "\")\nSet dest = ThisWorkbook";
		
		
		//macro += "\n"
		/*for (int x = 1; x <= (numberBanks + 1); x++) {
			macro += "\nsource.Activate\nSheets(\"Preview Table\").Select\nRows(\""
					+ x
					+ ":"
					+ x
				
					+ "\").Select\nSelection.Copy\ndest.Activate\nRows(\""
					+ x
					+ ":"
					+ x
					+ "\").Select\nSelection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False";
		}*/

		macro += "\nApplication.CutCopyMode = False\nsource.Close False\nEnd Sub";

		System.out.println(macro);

		PrintWriter writer = new PrintWriter(".\\WMLC\\Resources\\Macros\\GenPrelimTable.bas");
		writer.println(macro);
		writer.close();
		
		
	}
	
	// Copies the preliminary templates, pastes them, and inserts the relevant bank
	public void createPrelimTables() throws FileNotFoundException {
		String macro = "Sub createPrelimTables()\nApplication.CutCopyMode = False\n";
		
		int numSheetsBaseline = numberSheetsTable + 1;
		
		for (int x = 0; x < numberBanks; x++) {
			for (int y = 0; y < numberSheetsTable; y++) {
				macro += "\nSheets(\"Template." + (y + 1) + "\").Copy After:=Sheets(" + numSheetsBaseline + ")";
				macro += "\nSheets(\"Template." + (y + 1) + " (2)\").Name = \"" + bankNameList.get(x) + "." + (y + 1) + "\"";
				macro += "\nRange(\"A2\").Select\nActiveCell.FormulaR1C1 = \"=" + bankNameList.get(x) + "\"\n";
				numSheetsBaseline++;
			}
		}
		
		//macro+="\nApplication.CalculateFull\n\nEnd Sub";
		macro+="\nSheets(\"Data\").Select\nApplication.CalculateFull\n\nEnd Sub";
		
		PrintWriter writer = new PrintWriter(".\\" + practiceName + " Res\\Macros\\createPrelimTables.bas");
		writer.println(macro);
		writer.close();
		
	}
	
	public void createFinalTables() throws FileNotFoundException {
		
	String macro = "Sub createFinalTables()\nApplication.CutCopyMode = False\n";
		
		int numSheetsBaseline = numberSheetsTable + 1;
		
		for (int x = 0; x < numberBanks; x++) {
			for (int y = 0; y < numberSheetsTable; y++) {
				macro += "\nSheets(\"Template." + (y + 1) + "\").Copy After:=Sheets(" + numSheetsBaseline + ")";
				macro += "\nSheets(\"Template." + (y + 1) + " (2)\").Name = \"" + bankNameList.get(x) + "." + (y + 1) + "\"";
				macro += "\nRange(\"A2\").Select\nActiveCell.FormulaR1C1 = \"=" + bankNameList.get(x) + "\"\n";
				numSheetsBaseline++;
			}
		}
		
		//macro+="\nApplication.CalculateFull\n\nEnd Sub";
		macro+="\nSheets(\"Data\").Select\nApplication.CalculateFull\n\nEnd Sub";
		
		PrintWriter writer = new PrintWriter(".\\" + practiceName + " Res\\Macros\\createFinalTables.bas");
		writer.println(macro);
		writer.close();
		
	}
	
	// A script from going from preview table to preliminary reports
	public void generateScript() throws IOException {

		String script = "Dim objExcel, objWorkbook"
				+ System.getProperty("line.separator")
				+ "Set objExcel = CreateObject(\"Excel.Application\")\nobjExcel.Visible = True\n"
				+ System.getProperty("line.separator");
		
		prelimScriptBody = "\nSet objWorkbook = objExcel.Workbooks.Open(\"C:\\Users\\mdinardo\\workspace\\DashDesign\\" + practiceName + " Res\\" + practiceName + "PrelimTemplate.xlsm\")"
				+ System.getProperty("line.separator")
				+ "objExcel.VBE.ActiveVBProject.VBComponents.Import \"C:\\Users\\mdinardo\\workspace\\DashDesign\\" + practiceName + " Res\\Macros\\GenPrelimTable.bas\""
				+ System.getProperty("line.separator")
				+ "objExcel.VBE.ActiveVBProject.VBComponents.Import \"C:\\Users\\mdinardo\\workspace\\DashDesign\\" + practiceName + " Res\\Macros\\addLowsHighs.bas\""
				+ System.getProperty("line.separator")
				+ "objExcel.VBE.ActiveVBProject.VBComponents.Import \"C:\\Users\\mdinardo\\workspace\\DashDesign\\" + practiceName + " Res\\Macros\\labelLowsHighs.bas\""
				+ System.getProperty("line.separator")
				+ "objExcel.VBE.ActiveVBProject.VBComponents.Import \"C:\\Users\\mdinardo\\workspace\\DashDesign\\" + practiceName + " Res\\Macros\\PasteLow.bas\""
				+ System.getProperty("line.separator")
				+ "objExcel.VBE.ActiveVBProject.VBComponents.Import \"C:\\Users\\mdinardo\\workspace\\DashDesign\\" + practiceName + " Res\\Macros\\PasteHigh.bas\""
				+ System.getProperty("line.separator")
				+ "objExcel.VBE.ActiveVBProject.VBComponents.Import \"C:\\Users\\mdinardo\\workspace\\DashDesign\\" + practiceName + " Res\\Macros\\PrelimAll.bas\""
				+ System.getProperty("line.separator")
				+ "objExcel.VBE.ActiveVBProject.VBComponents.Import \"C:\\Users\\mdinardo\\workspace\\DashDesign\\" + practiceName + " Res\\Macros\\createPrelimTables.bas\""
				+ System.getProperty("line.separator")
				+ "objExcel.VBE.ActiveVBProject.VBComponents.Import \"C:\\Users\\mdinardo\\workspace\\DashDesign\\" + practiceName + " Res\\Macros\\ExportPrelimAsPDF.bas\""
				+ System.getProperty("line.separator");
		
		//if (completeRunthrough == true) script += "objExcel.VBE.ActiveVBProject.VBComponents.Import \"C:\\Users\\mdinardo\\workspace\\DashDesign\\WMLC Res\\Macros\\pastePrelimTabToFinal.bas\"\n";
				script += "\nobjExcel.Application.Run \"runAll\"\n";
		
		script += prelimScriptBody;
		/*
		 * + System.getProperty("line.separator") + "objWorkbook.Close False" +
		 * System.getProperty("line.separator");
		 * 
		 * script = script + "objExcel.Application.Quit" +
		 * System.getProperty("line.separator") + "objExcel.Quit" +
		 * System.getProperty("line.separator") + "Set objWorkbook = Nothing" +
		 * System.getProperty("line.separator") + "Set objExcel = Nothing";
		 */

		PrintWriter writer = new PrintWriter("PrelimReports.vbs");
		writer.println(script);
		writer.close();
		if (completeRunthrough == false) {
			Runtime.getRuntime().exec("cmd /c start PrelimReports.vbs");
		}
		
		prelimDataLab.setText("Writing preliminary table commands... DONE");
		prelimDataLab.setForeground(Color.green);
		peerDataLab.setForeground(Color.BLACK);
		
		mainFrame.getContentPane().validate();

	}

	public void pastePrelimTabToFinal() throws IOException {
		

		String macro = "Sub pastePrelimTabToFinal()\n\nDim source as Workbook\nDim dest As Workbook\nSet source = Workbooks.Open(\"C:\\Users\\mdinardo\\Documents\\DashboardAutomations\\" + practiceName + "\\" + practiceName + "PrelimReport.xlsm\")\nSet dest = ThisWorkbook\n";

		macro += "source.Activate\nSheets(\"Data\").Select\nRows(\"1:" + (numberBanks + 1) + "\").Copy\ndest.Activate\nSheets(\"Data\").Select\nRange(\"A1\").Select\nSelection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False\n";//ActiveWorkbook.SaveAs \"C:\\Users\\mdinardo\\Documents\\DashboardAutomations\\" + practiceName + "\\" + practiceName + "Dashboard.xlsm\", fileFormat:=52\ndest.Activate\n";
		macro += "\ndest.Activate\nEnd Sub";
		
		PrintWriter writer = new PrintWriter(".\\" + practiceName + " Res\\Macros\\pastePrelimTabToFinal.bas");
		writer.println(macro);
		writer.close();
		
	}

	// Creates a new folder for each bank in a specified directory (if it does
	// not already exist)
	public void generateNewDirectories(String rootDirectory) {
		for (int x = 0; x < bankNameList.size(); x++) {
			
			boolean success = new File(rootDirectory + bankNameList.get(x)).mkdirs();
			
		}
	}

	// Inserts High and Low columns into preliminary table, and copies/pastes the relevant formulas into the columns
	public void addLowsHighs() throws FileNotFoundException {
		String macro = "Sub lowsHighsPrelim()\nApplication.CutCopyMode = False\n";
		String labelNewColumns = "Sub labelNewColumns()\n";
		String pasteLowPercentiles = "Sub pasteLowPercentiles()";
		String pasteHighPercentiles = "Sub pasteHighPercentiles()";
		//String[] excludeList = { "B", "C", "D", "F", "H", "I", "CR", "CS", "CT" };
		String[] excludeList = { };

		ArrayList<Integer> excludedIndexList = this
				.parseExcludedList(excludeList);

		String colStringIndex = startDataColumn;

		for (int colIntIndex = CellReference.convertColStringToIndex(startDataColumn); colIntIndex <= CellReference
				.convertColStringToIndex(lastDataColumn); colIntIndex++) {
			boolean excluded = false;
			for (int x = 0; x < excludedIndexList.size(); x++) {
				if (colIntIndex == excludedIndexList.get(x))
					excluded = true;
			}

			if (excluded == false) {

				macro += "Columns(\""
						+ CellReference.convertNumToColString(CellReference
								.convertColStringToIndex(colStringIndex) + 1)
						+ ":"
						+ CellReference.convertNumToColString(CellReference
								.convertColStringToIndex(colStringIndex) + 1)
						+ "\").Select\nSelection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove\nSelection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove\n";
				//if (finReport == true) macro += "Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove\nSelection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove\nSelection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove\nSelection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove\n";
				labelNewColumns += "\nRange(\""
						+ CellReference.convertNumToColString(CellReference
								.convertColStringToIndex(colStringIndex) + 1)
						+ "1\").Select\nActiveCell.FormulaR1C1 = \"=RC[-1]&\"\"low\"\"\n"
						+ "Range(\""
						+ CellReference.convertNumToColString(CellReference
								.convertColStringToIndex(colStringIndex) + 2)
						+ "1\").Select\nActiveCell.FormulaR1C1 = \"=RC[-2]&\"\"high\"\"\n";

				pasteLowPercentiles += "\nRange(\""
						+ CellReference.convertNumToColString(CellReference
								.convertColStringToIndex(colStringIndex) + 1)
						+ "2\").Select\n ActiveCell.FormulaR1C1 = \"=PERCENTILE((R2C"
						+ (CellReference
								.convertColStringToIndex(colStringIndex) + 1)
						+ ":R" + (numberBanks + 1) + "C"
						+ (CellReference
								.convertColStringToIndex(colStringIndex) + 1)
						+ "),0.2)\"\nRange(\""
						+ CellReference.convertNumToColString(CellReference
								.convertColStringToIndex(colStringIndex) + 1)
						+ "2\").Copy\nRange(\""
						+ CellReference.convertNumToColString(CellReference
								.convertColStringToIndex(colStringIndex) + 1)
						+ "3:"
						+ CellReference.convertNumToColString(CellReference
								.convertColStringToIndex(colStringIndex) + 1)
						+ (numberBanks + 1) + "\").Select\nActiveSheet.Paste\n";

				pasteHighPercentiles += "\nRange(\""
						+ CellReference.convertNumToColString(CellReference
								.convertColStringToIndex(colStringIndex) + 2)
						+ "2\").Select\n ActiveCell.FormulaR1C1 = \"=PERCENTILE((R2C"
						+ (CellReference
								.convertColStringToIndex(colStringIndex) + 1)
						+ ":R" + (numberBanks + 1) + "C"
						+ (CellReference
								.convertColStringToIndex(colStringIndex) + 1)
						+ "),0.8)\"\nRange(\""
						+ CellReference.convertNumToColString(CellReference
								.convertColStringToIndex(colStringIndex) + 2)
						+ "2\").Copy\nRange(\""
						+ CellReference.convertNumToColString(CellReference
								.convertColStringToIndex(colStringIndex) + 2)
						+ "3:"
						+ CellReference.convertNumToColString(CellReference
								.convertColStringToIndex(colStringIndex) + 2)
						+ (numberBanks + 1) + "\").Select\nActiveSheet.Paste\n";

				colStringIndex = CellReference
						.convertNumToColString(CellReference
								.convertColStringToIndex(colStringIndex) + 3);
			}

			else if (excluded == true)
				colStringIndex = CellReference
						.convertNumToColString(CellReference
								.convertColStringToIndex(colStringIndex) + 1);
		}

		macro += "\nEnd Sub";
		labelNewColumns += "\nEnd Sub";
		pasteLowPercentiles += "\nEnd Sub";
		pasteHighPercentiles += "\nRange(\"A1\").Select\n\nEnd Sub";

		String macroAll = "Sub runAll()\n\nCall genPrelimTable\nCall lowsHighsPrelim\nCall labelNewColumns\nCall pasteLowPercentiles\nCall pasteHighPercentiles\nColumns(\"A:ZZ\").ColumnWidth = 15\nCall createPrelimTables\nCall exportPrelimPDF\nSheets(\"Data\").Select\n";
				
				
				macroAll += "\nActiveWorkbook.SaveAs \"C:\\Users\\mdinardo\\Documents\\DashboardAutomations\\" + practiceName + "\\" + practiceName + "PrelimReport.xlsm\", fileFormat:=52\nActiveWorkbook.Close True\nEnd Sub";

		System.out.println(macro);
		PrintWriter writer = new PrintWriter(".\\" + practiceName + " Res\\Macros\\addLowsHighs.bas");
		writer.println(macro);
		writer.close();

		PrintWriter writer2 = new PrintWriter(".\\" + practiceName + " Res\\Macros\\labelLowsHighs.bas");
		writer2.println(labelNewColumns);
		writer2.close();

		PrintWriter writer3 = new PrintWriter(".\\" + practiceName + " Res\\Macros\\PrelimAll.bas");
		writer3.println(macroAll);
		writer3.close();

		PrintWriter writer4 = new PrintWriter(".\\" + practiceName + " Res\\Macros\\PasteLow.bas");
		writer4.println(pasteLowPercentiles);
		writer4.close();

		PrintWriter writer5 = new PrintWriter(".\\" + practiceName + " Res\\Macros\\PasteHigh.bas");
		writer5.println(pasteHighPercentiles);
		writer5.close();

		// for (int x = CellReference.convertColStringToIndex(arg0))

		/*
		 * Sub Macro2() ' ' Macro2 Macro '
		 * 
		 * ' Columns("L:L").Select Selection.Insert Shift:=xlToRight,
		 * CopyOrigin:=xlFormatFromLeftOrAbove Selection.Insert
		 * Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove End Sub
		 */

	}
	
	// Creates an index list of columns that should be excluded from parsing
	public ArrayList<Integer> parseExcludedList(String[] excludeList) {

		ArrayList<Integer> excludedIndex = new ArrayList<Integer>();

		for (int x = 0; x < excludeList.length; x++) {
			excludedIndex.add(CellReference
					.convertColStringToIndex(excludeList[x]));
		}

		return excludedIndex;
	}
	
	// Not sure, does nothing right now
	public void getSelectedList() {

		/*
		 * String rangeSelection = "Range(\"";
		 * 
		 * int colKRangeStart = 31; int colKRangeEnd = 152;
		 * 
		 * int colLRangeStart = 44;
		 * 
		 * 
		 * String[] exclude = { "38", "43", "45", "51", "52", "54", "62", "69",
		 * "70", "72", "77", "83", "84", "95", "96", "102", "105", "106", "116",
		 * "119", "124", "125", "130", "137", "139", "140", "147", "148", "149"
		 * }; System.out.println("here??"); for (int x = colKRangeStart; x <
		 * colKRangeEnd; x++) { boolean excludedList = false; for (int y = 0; y
		 * < exclude.length; y++) { if (x == Integer.parseInt(exclude[y])) {
		 * excludedList = true; } }
		 * 
		 * if (excludedList == false) { rangeSelection+="K" + x + ","; } }
		 * rangeSelection+= "\").Select";
		 * 
		 * System.out.println(rangeSelection);
		 */
	}
	
	// Exports all prelim tables as PDFs into bank folder
	public void exportPrelimPDF() throws FileNotFoundException {
		String macro = "Sub exportPrelimPDF()\n\n";

		for (int x = 0; x < bankNameList.size(); x++) {
			macro += "Sheets(Array(";
			for (int y = 0; y < numberSheetsTable; y++) {
				macro += "\"" + bankNameList.get(x) + "." + (y + 1) + "\"";
				if (y < (numberSheetsTable - 1)) {
					macro += ",";
				}
			}
			macro += ")).Select";
			macro += "\nActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=\"C:\\Users\\mdinardo\\Documents\\DashboardAutomations\\" + practiceName + "\\"
					+ bankNameList.get(x)
					+ "\\PRELIM"
					+ bankNameList.get(x)
					+ ".pdf\", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False\n";
		}
		
		macro += "\nEnd Sub";
		
		PrintWriter writer = new PrintWriter(".\\" + practiceName + " Res\\Macros\\ExportPrelimAsPDF.bas");
		writer.println(macro);
		writer.close();
	}
	
	public void exportDashboardPDF() throws FileNotFoundException {
		String macro = "Sub exportDashboardPDF()\n\n";

		for (int x = 0; x < bankNameList.size(); x++) {
			macro += "Sheets(Array(";
			for (int y = 0; y < numberSheetsTable; y++) {
				macro += "\"" + bankNameList.get(x) + "." + (y + 1) + "\"";
				if (y < (numberSheetsTable - 1)) {
					macro += ",";
				}
			}
			macro += ")).Select";
			macro += "\nActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=\"C:\\Users\\mdinardo\\Documents\\DashboardAutomations\\" + practiceName + "\\"
					+ bankNameList.get(x)
					+ "\\DASHBOARD"
					+ bankNameList.get(x)
					+ ".pdf\", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False\n";
		}
		
		macro += "\nEnd Sub";
		
		PrintWriter writer = new PrintWriter(".\\" + practiceName + " Res\\Macros\\ExportDashboardAsPDF.bas");
		writer.println(macro);
		writer.close();
	}
	
	// I don't think this does anything properly right now
	/*public void generateMacro(String intakeName) {
		String macro = "Sub parseIntakeForms()\n\nDim source As Workbook\nDim dest As Workbook\nSet source = Workbooks.Open(\""
				+ intakeName
				+ "\")\nRange(\"K31:L31,K32:L32,K33:L33,K34:L34,K35:L35,K36:L36,K37:L37\").Copy\nSet dest = Workbooks.Open(\"C:\\Users\\mdinardo\\Desktop\\CopyPasteTest\\Prelimtester.xlsm\")\nRows(\"2:2\").Select\nSelection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=True\nsource.Activate\nRange(\"K44:L44,K46:L50,K53:L53,K55:L61,K63:L68,K71:L71:,K73:L76,K78:L82,K85:K89,K97:L101,K103:L104,K107:L115,K117:L118,K120:L123,K127:L129,K131:L136,K138:L138,K141:L146,K150:L152"
				+ "\").Copy\n";
		macro += "dest.Activate\nRange(\"H2\").Select\nSelection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=True\nEnd Sub";
		System.out.println(macro);
	}*/
	
	// Generates prelim table automatically, from a specified preview table to all exported PDFs (runs automatically if called).
	public void generatePrelim() throws IOException {
		
		this.generateNewDirectories("C:/Users/mdinardo/Documents/DashboardAutomations/WMLC/");
		this.prelimTableMacro();
		this.addLowsHighs();
		this.createPrelimTables();
		this.exportPrelimPDF();
		this.generateScript();
	}
	
	// Right now it takes the prelim table and makes a new Final table for each bank
	public void generateFinalReport() throws IOException {
		
		this.pastePrelimTabToFinal();
		this.insertNewColumns();
		this.addPeer33And66();
		this.createFinalTables();
		this.exportDashboardPDF();
		this.genFinalTableScript();
		
	}
	
	// Makes a copy of the final report template for each bank and pastes it into its folder
	/*public void copyFinalTemplate(File finalTemplate) throws IOException {
		
		File source = finalTemplate;
		
		for (int x = 0; x < bankNameList.size(); x++) {
			
			File dest = new File("C:\\Users\\mdinardo\\Desktop\\CopyPasteTest\\WMLC Preliminary reports\\" + bankNameList.get(x) + "\\" + bankNameList.get(x) + "FinReport.xlsm");
			FileUtils.copyFile(source, dest);
			
		}
		
	}*/
	
	// Copies the relevant row from the preliminary table and pastes it to the final table
	/*public void parsePrelimToFinal() throws FileNotFoundException, IOException {
		
		String macroAll = "Sub parseFinAll()\n";
		for (int x = 0; x < bankNameList.size(); x++) {
			
			

			System.out.println(bankNameList.get(x).replaceAll("\\s+", ""));
			// sourcePath:
			// Workbooks.Open(\"C:\\Users\\mdinardo\\Desktop\\PrelimTable.xlsm\")
			
			String fixedName = bankNameList.get(x).replaceAll("\\s+", "");
			fixedName = fixedName.replace("-", "");
			String macro = "Sub "
					+ fixedName + "finParse" 
					+ "()\n\nDim source as Workbook\nDim dest As Workbook\nSet source = ThisWorkbook\nSet dest = Workbooks.Open(\"C:\\Users\\mdinardo\\Desktop\\CopyPasteTest\\WMLC Preliminary reports\\"
					+ bankNameList.get(x) + "\\" + bankNameList.get(x)
					+ "FinReport.xlsm\")\n";
			macro += "source.Activate\n";
			macroAll += "Call " + fixedName + "finParse\n";
			int i = 1;
			String bName = "";

			Cell c = dataSheet.getRow(i).getCell(0);
			while (!c.getRichStringCellValue().toString()
					.equals(bankNameList.get(x))) {
				i++;
				c = dataSheet.getRow(i).getCell(0);
			}
			
			macro += "Rows(\""
					+ 1
					+ ":"
					+ 1
					+ "\").Select\nSelection.Copy\ndest.Activate\nRows(\"1:1\").Select\nActiveSheet.Paste\n";

			macro += "source.Activate\nRows(\""
					+ (c.getRowIndex() + 1)
					+ ":"
					+ (c.getRowIndex() + 1)
					+ "\").Select\nSelection.Copy\ndest.Activate\nRows(\"2:2\").Select\nSelection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False\nColumns(\"A:ZZ\").ColumnWidth = 15\ndest.Close True\n\nEnd Sub";
			
			
			
			PrintWriter writer = new PrintWriter(".\\" + practiceName + " Res\\Macros\\" + fixedName + "Parse.bas");
			writer.println(macro);
			writer.close();
			
			

		}
		
		macroAll += "\nEnd Sub";
		
		PrintWriter writer2 = new PrintWriter(".\\" + practiceName + " Res\\Macros\\ParseAll.bas");
		writer2.println(macroAll);
		writer2.close();
			
	}*/
	
	// A script for generating the final report
	public void finalReportScript() throws IOException {
		
		String script = "Dim objExcel, objWorkbook"
				+ System.getProperty("line.separator")
				+ "Set objExcel = CreateObject(\"Excel.Application\")"
				+ System.getProperty("line.separator");

		script += "objExcel.Visible = True\nSet objWorkbook = objExcel.Workbooks.Open(\"C:\\Users\\mdinardo\\workspace\\DashDesign\\PrelimTable.xlsm\")"
				+ System.getProperty("line.separator");
		
		for (int x = 0; x < bankNameList.size(); x++) {
			String fixedName = bankNameList.get(x).replaceAll("\\s+", "");
			fixedName = fixedName.replace("-", "");
				script += "objExcel.VBE.ActiveVBProject.VBComponents.Import \"C:\\Users\\mdinardo\\workspace\\DashDesign\\Macros\\" + fixedName + "Parse.bas\""
				+ System.getProperty("line.separator");
		}
		
		script += "objExcel.VBE.ActiveVBProject.VBComponents.Import \"C:\\Users\\mdinardo\\workspace\\DashDesign\\Macros\\" + "ParseAll.bas\""
				+ System.getProperty("line.separator") + "objExcel.Application.Run \"parseFinAll\"";
		
		
		PrintWriter writer = new PrintWriter("ParseFinalReport.vbs");
		writer.println(script);
		writer.close();
		
		Runtime.getRuntime().exec("cmd /c start ParseFinalReport.vbs");
	}
	
	public void insertNewColumns() throws FileNotFoundException {
		
		ArrayList<String> alreadyAddedCol = new ArrayList<String>();
		
		
		String macro = "Sub insertCols()\n";
		
		int x = 0;
		Cell c = dataSheet.getRow(0).getCell(x);
		while (c != null) {
			x++;
			c = dataSheet.getRow(0).getCell(x, Row.RETURN_BLANK_AS_NULL);
		}

		int lastIndex = x;

		while (lastIndex > 1) {

			c = dataSheet.getRow(0)
					.getCell(lastIndex, Row.RETURN_BLANK_AS_NULL);
			if (c != null && c.getRichStringCellValue().length() > 4
					&& c.getRichStringCellValue()
							.toString()
							.substring(
									c.getRichStringCellValue().toString()
											.length() - 4,
									c.getRichStringCellValue().toString()
											.length()).equals("high")) {
				
				boolean nextCol = true;
				
				for (int y = 0; y < alreadyAddedCol.size(); y++) {
					
					if (alreadyAddedCol.get(y).equals(c.getRichStringCellValue().toString())) {
						nextCol = false;
					}
					
				}
				
				if (nextCol == true) {
				macro += "Columns(\""
						+ CellReference.convertNumToColString(lastIndex + 1)
						+ ":"
						+ CellReference.convertNumToColString(lastIndex + 1)
						+ "\").Select\nSelection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove\nSelection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove\nSelection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove\nSelection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove\n";
				alreadyAddedCol.add(c.getRichStringCellValue().toString());
				}
			}
			
			lastIndex--;
		}
		
		macro += "\nEnd Sub";
		
		PrintWriter writer = new PrintWriter(".\\" + practiceName + " Res\\Macros\\insertCols.bas");
		writer.println(macro);
		writer.close();
	
		
		
	}
	
	// A method that adds 33 and 66th percentile columns to final report
	public void addPeer33And66() throws IOException {
		
		
		
		String macro3366 = "Sub addPeer33And66()\n";
		String macroIndPeerMed = "Sub addIndPeerMed()\n";
		
		ArrayList<String> macListHighLow = new ArrayList<String>();
		ArrayList<String> macList3366 = new ArrayList<String>();
		ArrayList<String> indPeerMed = new ArrayList<String>();
		String name;
		for (int z = 0; z < bankNameList.size(); z++) {
			name = bankNameList.get(z).replaceAll("\\s+", "");
			String mac3366 = "Sub " + name.replaceAll("-", "") + "3366()\nOn Error Resume Next\n";
			String macHighLow = "Sub " + name.replaceAll("-", "") + "HighLow()\nOn Error Resume Next\n";
			//String mac66 = "Sub " + name.replaceAll("-", "") + "66()\nOn Error Resume Next\n";
			macList3366.add(mac3366);
			macListHighLow.add(macHighLow);
			//macList66.add(mac66);
		}
		
		for (int z = 0; z < bankNameList.size(); z++) {
			name = bankNameList.get(z).replaceAll("\\s+", "");
			String medians = "Sub " + name.replaceAll("-", "") + "IndPeerMed()\nOn Error Resume Next\n";
			//String mac66 = "Sub " + name.replaceAll("-", "") + "66()\nOn Error Resume Next\n";
			indPeerMed.add(medians);
			//macList66.add(mac66);
		}
		
		

		int i = CellReference.convertColStringToIndex(startDataColumn);
		Cell c = dataSheet.getRow(0).getCell(i, Row.RETURN_BLANK_AS_NULL);
		int multiplier = 0;
		while (c != null) {

			if (c.getRichStringCellValue().toString().length() >= 4
					&& c.getRichStringCellValue()
							.toString()
							.substring(
									c.getRichStringCellValue().toString()
											.length() - 4,
									c.getRichStringCellValue().toString()
											.length()).equals("high")) {
				
				for (int z = 0; z < bankNameList.size(); z++) {
					String lowCmd = "Range(\"" + CellReference.convertNumToColString((i-1) + (4 * multiplier)) + (z + 2) + "\").Select\nActiveCell.FormulaR1C1= \"=PERCENTILE((";
					String highCmd = "Range(\"" + CellReference.convertNumToColString((i) + (4 * multiplier)) + (z + 2) + "\").Select\nActiveCell.FormulaR1C1= \"=PERCENTILE((";
					String percentileCmd33 = "Range(\"" + CellReference.convertNumToColString((i + 1) + (4 * multiplier)) + (z + 2) + "\").Select\nActiveCell.FormulaR1C1= \"=PERCENTILE((";
					String percentileCmd66 = "Range(\"" + CellReference.convertNumToColString((i + 2) + (4 * multiplier)) + (z + 2) + "\").Select\nActiveCell.FormulaR1C1= \"=PERCENTILE((";
					String indMedCmd = "Range(\"" + CellReference.convertNumToColString((i + 3) + (4 * multiplier)) + (z + 2) + "\").Select\nActiveCell.FormulaR1C1= \"=MEDIAN(R[" + (1 - (z + 2)) + "]" + "C[-5]:R[" + ((bankNameList.size() + 1) - (z + 2)) + "]C[-5])\"";
					String peerMedCmd = "Range(\"" + CellReference.convertNumToColString((i + 4) + (4 * multiplier)) + (z + 2) + "\").Select\nActiveCell.FormulaR1C1= \"=MEDIAN(";
					
					//String command = "\nRange(\"" + CellReference.convertNumToColString((i + 1) + (4 * multiplier)) + (z + 2) + "\").Select\nRange(\"";
					ArrayList<String> indexList = this.peerSelectionBody(bankNameList.get(z));
					for (int v = 0; v < indexList.size(); v++) {
						
						//command += CellReference.convertNumToColString((i-2) + (4 * multiplier)) + indexList.get(v);
						lowCmd += "R[" + (Integer.parseInt(indexList.get(v)) - (z + 2)) + "]C[-1]";
						highCmd += "R[" + (Integer.parseInt(indexList.get(v)) - (z + 2)) + "]C[-2]";
						percentileCmd33 += "R[" + (Integer.parseInt(indexList.get(v)) - (z + 2)) + "]C[-3]";
						percentileCmd66 += "R[" + (Integer.parseInt(indexList.get(v)) - (z + 2)) + "]C[-4]";
						//indMedCmd += "R[" + (Integer.parseInt(indexList.get(v)) - (z + 2)) + "]C[-5]";
						peerMedCmd += "R[" + (Integer.parseInt(indexList.get(v)) - (z + 2)) + "]C[-6]";
								
								
						if (v < (indexList.size() - 1)) {
							//command += ",";
							lowCmd += ",";
							highCmd += ",";
							percentileCmd33 += ",";
							percentileCmd66 += ",";
							peerMedCmd += ",";
						}
					}
					lowCmd += "),0.2)\"\n";
					highCmd += "),0.8)\"\n";
					percentileCmd33 += "),0.33)\"\n";
					percentileCmd66 += "),0.66)\"\n";
					indMedCmd += "\n";
					peerMedCmd += ")\n";
					//command += "\").Select\nActiveCell.FormulaR1C1 = Application.WorksheetFunction.Percentile(arr, 0.33)";
					//System.out.println(command);
					macListHighLow.set(z, macListHighLow.get(z) + lowCmd + highCmd);
					macList3366.set(z, macList3366.get(z) + percentileCmd33 + percentileCmd66);
					indPeerMed.set(z, indPeerMed.get(z) + indMedCmd + peerMedCmd);
				}

				macro3366 += "Range(\""
						+ CellReference.convertNumToColString((i + 1) + (4 * multiplier))
						+ "1\").Select\nActiveCell.FormulaR1C1 = \""
						+ c.getRichStringCellValue()
								.toString()
								.substring(
										0,
										(c.getRichStringCellValue().toString()
												.length() - 4)) + "33\n";
				
				
				
				macro3366 += "Range(\""
						+ CellReference.convertNumToColString((i + 2) + (4 * multiplier))
						 + "1\").Select\nActiveCell.FormulaR1C1 = \""
						+ c.getRichStringCellValue()
								.toString()
								.substring(
										0,
										(c.getRichStringCellValue().toString()
												.length() - 4)) + "66\"\n";
				
				macroIndPeerMed += "Range(\""
						+ CellReference.convertNumToColString((i + 3) + (4 * multiplier))
						 + "1\").Select\nActiveCell.FormulaR1C1 = \""
						+ c.getRichStringCellValue()
								.toString()
								.substring(
										0,
										(c.getRichStringCellValue().toString()
												.length() - 4)) + "indMedian\"\n";
				
				macroIndPeerMed += "Range(\""
						+ CellReference.convertNumToColString((i + 4) + (4 * multiplier))
						 + "1\").Select\nActiveCell.FormulaR1C1 = \""
						+ c.getRichStringCellValue()
								.toString()
								.substring(
										0,
										(c.getRichStringCellValue().toString()
												.length() - 4)) + "peerMedian\"\n";
				
				multiplier++;
			}
			
			i++;
			c = dataSheet.getRow(0).getCell(i, Row.RETURN_BLANK_AS_NULL);
		}
		
		peerDataLab.setText("Parsing peer group data... DONE");
		peerDataLab.setForeground(Color.green);
		dashDataLab.setForeground(Color.black);
		
		mainFrame.getContentPane().validate();
		
		
		for (int z = 0; z < macList3366.size(); z++) {
			
			macList3366.set(z, macList3366.get(z) + "\nEnd Sub");
			PrintWriter writerX = new PrintWriter(".\\" + practiceName + " Res\\Macros\\mac33Bank" + z + ".bas");
			writerX.println(macList3366.get(z));
			writerX.close();
			
			indPeerMed.set(z, indPeerMed.get(z) + "\nEnd Sub");
			PrintWriter writerY = new PrintWriter(".\\" + practiceName + " Res\\Macros\\indPeerMed" + z + ".bas");
			writerY.println(indPeerMed.get(z));
			writerY.close();
			
			macListHighLow.set(z, macListHighLow.get(z) + "\nEnd Sub");
			PrintWriter writerZ = new PrintWriter(".\\" + practiceName + " Res\\Macros\\macHL" + z + ".bas");
			writerZ.println(macListHighLow.get(z));
			writerZ.close();
		}

		macro3366 += "\nEnd Sub";
		macroIndPeerMed += "\nEnd Sub";

		PrintWriter writer = new PrintWriter(".\\" + practiceName + " Res\\Macros\\addPeer33And66.bas");
		writer.println(macro3366);
		writer.close();
		
		PrintWriter writer2 = new PrintWriter(".\\" + practiceName + " Res\\Macros\\addIndPeerMed.bas");
		writer2.println(macroIndPeerMed);
		writer2.close();
		
		dashDataLab.setText(dashDataLab.getText().toString() + "DONE");
		dashDataLab.setForeground(Color.green);
		dashTabLab.setForeground(Color.black);
		
		mainFrame.getContentPane().validate();
		
	}
	
	public void genFinalTableScript() throws IOException {
		
		String runAllMacro = "Sub runAllFinTable()\n";
		
		runAllMacro += "Call pastePrelimTabToFinal\nCall insertCols\nCall addPeer33And66\nCall addIndPeerMed\n";
		
		for (int x = 0; x < bankNameList.size(); x++) {
			
			String bName = bankNameList.get(x).replaceAll("\\s+", "");
			bName = bName.replaceAll("-", "");
			
			runAllMacro += "Call " + bName + "HighLow\n";
			runAllMacro += "Call " + bName + "3366\n";
			runAllMacro += "Call " + bName + "IndPeerMed\n";
			
		}
		runAllMacro += "Call createFinalTables\n";
		runAllMacro += "\nColumns(\"A:CCC\").ColumnWidth = 15\n\nApplication.CalculateFull\n\n Call exportDashboardPDF\nActiveWorkbook.SaveAs \"C:\\Users\\mdinardo\\Documents\\DashboardAutomations\\" + practiceName + "\\" + practiceName + "Dashboard.xlsm\"\nEnd Sub";

		String script = "Dim objExcel, objWorkbook, objExcel2, objWorkbook2"
				+ System.getProperty("line.separator")
				+ "Set objExcel = CreateObject(\"Excel.Application\")"
				+ System.getProperty("line.separator");

		script = script
				+ "objExcel.Visible = True\n";
		
		if (completeRunthrough == true) {
			
			script += prelimScriptBody;
			script += "\nobjExcel.Application.Run \"runAll\"\n";

		}

		// insertCols.bas
		script += "\n\nobjExcel.Application.Quit\nSet objExcel2 = CreateObject(\"Excel.Application\")\nobjExcel2.Visible = True\nSet objWorkbook2 = objExcel2.Workbooks.Open(\"C:\\Users\\mdinardo\\workspace\\DashDesign\\"
				+ practiceName
				+ " Res\\"
				+ practiceName
				+ "DashboardTemplate.xlsm\")\nobjExcel2.VBE.ActiveVBProject.VBComponents.Import \"C:\\Users\\mdinardo\\workspace\\DashDesign\\"
				+ practiceName + " Res\\Macros\\pastePrelimTabToFinal.bas\"\nobjExcel2.VBE.ActiveVBProject.VBComponents.Import \"C:\\Users\\mdinardo\\workspace\\DashDesign\\"
				+ practiceName + " Res\\Macros\\insertCols.bas\"\nobjExcel2.VBE.ActiveVBProject.VBComponents.Import \"C:\\Users\\mdinardo\\workspace\\DashDesign\\"
				+ practiceName + " Res\\Macros\\addPeer33And66.bas\"\nobjExcel2.VBE.ActiveVBProject.VBComponents.Import \"C:\\Users\\mdinardo\\workspace\\DashDesign\\"
				+ practiceName + " Res\\Macros\\addIndPeerMed.bas\"\n";

		for (int x = 0; x < bankNameList.size(); x++) {

			script += System.getProperty("line.separator")
					+ "objExcel2.VBE.ActiveVBProject.VBComponents.Import \"C:\\Users\\mdinardo\\workspace\\DashDesign\\" + practiceName + " Res\\Macros\\macHL"
					+ x
					+ ".bas\""
					+ System.getProperty("line.separator")
					+ "objExcel2.VBE.ActiveVBProject.VBComponents.Import \"C:\\Users\\mdinardo\\workspace\\DashDesign\\" + practiceName + " Res\\Macros\\mac33Bank"
					+ x
					+ ".bas\"\n"
					+ "objExcel2.VBE.ActiveVBProject.VBComponents.Import \"C:\\Users\\mdinardo\\workspace\\DashDesign\\" + practiceName + " Res\\Macros\\indPeerMed"
					+ x + ".bas\"\n";
		}
		
		script += "\nobjExcel2.VBE.ActiveVBProject.VBComponents.Import \"C:\\Users\\mdinardo\\workspace\\DashDesign\\" + practiceName + " Res\\Macros\\createFinalTables.bas\"\n";
		script += "objExcel2.VBE.ActiveVBProject.VBComponents.Import \"C:\\Users\\mdinardo\\workspace\\DashDesign\\" + practiceName + " Res\\Macros\\runAllFinTable.bas\"\n";
		script += "objExcel2.VBE.ActiveVBProject.VBComponents.Import \"C:\\Users\\mdinardo\\workspace\\DashDesign\\" + practiceName + " Res\\Macros\\ExportDashboardAsPDF.bas\"\n";
		script += System.getProperty("line.separator")
				+ "objExcel2.Application.Run \"runAllFinTable\"";
		
		dashTabLab.setText(dashTabLab.getText().toString() + "DONE");
		dashTabLab.setForeground(Color.green);
		mainFrame.add(loadingLab);
		
		mainFrame.getContentPane().validate();

		PrintWriter writer1 = new PrintWriter(".\\" + practiceName + " Res\\Macros\\runAllFinTable.bas");
		writer1.println(runAllMacro);
		writer1.close();
		
		PrintWriter writer2 = new PrintWriter("ExecuteFinalTable.vbs");
		writer2.println(script);
		writer2.close();

		System.out.println(script);
		
		
		Runtime.getRuntime().exec(
					"cmd /c start ExecuteFinalTable.vbs");

	}
	
	public ArrayList<String> peerSelectionBody(String bankName) throws FileNotFoundException, IOException {
		
		ArrayList<String> peerRowNumList = new ArrayList<String>();

		String peerBody = "";
		
		
		
		//System.out.println(peerList.toString());
		//System.out.println("here!");
		
		int index = 0;
		boolean found = false;
		
		for (int x = 0; x < peerList.getPeerGroupList().size(); x++) {
			
			if (peerList.getPeerGroupList().get(x).getUserList().contains(bankName)) {
				found = true;
				//System.out.println(peerList.getPeerGroupList().get(x).getName());

				index = x;
			}
		}
		
		if (found == false) {
			
			JFrame dialog = new JFrame();
			JOptionPane.showMessageDialog(dialog, "Peer group not found for " + bankName);
		}
		
		else {
		PeerGroup banksGroup = peerList.getPeerGroupList().get(index);
		
			for (Row r: dataSheet) {
				
				for (int y = 0; y < banksGroup.getBankList().size(); y++) {
					
					if (banksGroup.getBankList().get(y).equals(r.getCell(0).getRichStringCellValue().toString())) {
						found = true;
						peerBody += "A" + (r.getRowNum() + 1) + ",";
					
						peerRowNumList.add(Integer.toString((r.getRowNum() + 1)));
					}
				
			}
			}
		}
		
			//System.out.println(peerBody);
		return peerRowNumList;
		
	}
	
	public void executeAll() throws IOException {
		
		this.createAndShowGUI();
		this.generatePrelim();
		this.generateFinalReport();
		
	}

	public static void main(String[] args) throws FileNotFoundException,
			IOException {
		
		chooseDatabase();
		//ParseIntake intaker = new ParseIntake(new File("C:\\Users\\mdinardo\\Documents\\DashboardAutomations\\WMLC\\WMLCAggregateData.xlsx"));
		//intaker.generateNewDirectories();
		//intaker.copyFinalTemplate(new File("Finaltester.xlsm"));
		//intaker.parsePrelimToFinal();
		//intaker.finalReportScript();
		//intaker.addPeer33And66();
		//intaker.insertNewColumns(5);
		//intaker.generatePrelim();
		//intaker.genCellList("C:\\Users\\mdinardo\\Desktop\\Corrected_Associated CEB Wealth Industry Benchmark Intake Form FY 2013_MASTER_3-6-13.xlsx");
		// intaker.generateScript();
		// intaker.prelimTableMacro();
		
		//intaker.pastePrelimTabToFinal();
	    //intaker.insertNewColumns();
	    //intaker.generateScript();
		//intaker.addPeer33And66();
		
		//intaker.peerSelectionBody("Huntington");
		//intaker.addPeer33And66();
		//intaker.createFinalTables();
		//intaker.genFinalTableScript();
		//intaker.executeAll();
		//intaker.chooseDatabase();
		//intaker.createAndShowGUI();
		
		/*for (int x = 0; x < intaker.bankNameList.size(); x++) {
			
			intaker.peerSelectionBody(intaker.bankNameList.get(x));
			
		}*/
	}

	@Override
	public void actionPerformed(ActionEvent e) {
		String command = e.getActionCommand().toString();
		System.out.println(command);
		
		
		
	}
}
