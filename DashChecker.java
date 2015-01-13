
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.ArrayList;

import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;

import org.apache.pdfbox.exceptions.COSVisitorException;
import org.apache.pdfbox.util.PDFMergerUtility;
import org.apache.poi.ss.util.CellReference;


public class DashChecker {
	
	String root = System.getProperty("user.dir");
	private final String dirPath = "./Data";
	
	File dashTool = new File("./Settings/Dashboard Settings.xlsm");
	
	File aggMacro;
	File dir;
	File aggregateMacro;
	
	JFrame chooserFrame = new JFrame();
	JFileChooser fc = new JFileChooser();
	File selctedFile;
	File warningReports;
	String warningReportsString;
	String prelimTemplateString;
	String dashboardTemplateString;
	String scoreTemplateString;
	String outputDir;
	
	public DashChecker() throws IOException, InterruptedException {
		
		/*int returnVal = fc.showOpenDialog(chooserFrame);

		if (fc.getSelectedFile() == null) {
			System.exit(0); }
		warningReports = fc.getSelectedFile();
		warningReportsString = warningReports.getCanonicalPath();*/
		warningReportsString = root + "\\Data Warning Reports.xlsm";
		prelimTemplateString = root + "\\Settings\\PrelimTemplate.xlsm";
		dashboardTemplateString = root + "\\Settings\\DashboardTemplate.xlsm";
		scoreTemplateString = root + "\\Settings\\ScoreTemplate.xlsm";
		this.aggregateData();
		// this.aggregateScript();

	}

	public DashChecker(String outputDir) throws IOException,
			InterruptedException {

		// int returnVal = fc.showOpenDialog(chooserFrame);

		// if (fc.getSelectedFile() == null) {
		// System.exit(0); }
		// warningReports = fc.getSelectedFile();
		this.outputDir = outputDir;
		warningReportsString = root + "\\Data Warning Reports.xlsm";
		prelimTemplateString = root + "\\Settings\\PrelimTemplate.xlsm";
		dashboardTemplateString = root + "\\Settings\\DashboardTemplate.xlsm";
		scoreTemplateString = root + "\\Settings\\ScoreTemplate.xlsm";
		this.aggregateData();
		// this.aggregateScript();

	}

	public void aggregateData() throws IOException, InterruptedException {
		
		int i = 3;
		int prevNum = 2;

		String aggregateData = "Sub aggregateData()\nDim c As Range\nDim i As Integer\n";

		dir = new File(dirPath);
		File[] directoryListing = dir.listFiles();

		if (directoryListing != null) {
			for (File child : directoryListing) {
				aggregateData += "Call testParsing(\""
						+ child.getCanonicalPath() + "\",\""
						+ CellReference.convertNumToColString(i) + "\",\""
						+ CellReference.convertNumToColString(i + 1) + "\", \"" + prevNum + "\")\n";
				System.out.println(child.getName());
				i = i + 3;
				prevNum++;
			}

		}

		else {
			System.out.println("Not valid data directory");
		}

		aggregateData += "\nSheets(\"Data Report\").Select\ni = 1\n For Each c In ActiveSheet.UsedRange.Columns(\"B\").Cells" + "\nIf Not c.Text = \"\" Then\n i = i + 1\n End If \nNext c\nRange(\"B2:B\" & i).Copy\nSheets(\"Preview Table\").Select\nRange(\"A1\").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=True"
				+ "\nRange(\"A1\").Select\nActiveCell.FormulaR1C1 = \"Bank\"\nCall placeHighlightButton\n\nRange(\"A1\").Select\nEnd Sub";
		System.out.println(aggregateData);

		PrintWriter writer = new PrintWriter(new File("./Settings/Macros/aggregateData.bas"));
		writer.println(aggregateData);
		writer.close();
		
		aggMacro = new File("./Settings/Macros/aggregateData.bas");
		
		
	}
	
	/*public void aggregateScript() throws IOException {
		
	
		String script = "Dim objExcel, objWorkbook\n"
				+ "Set objExcel = CreateObject(\"Excel.Application\")\n"
				+ "objExcel.Visible = True\n"
				+ "Set objWorkbook = objExcel.Workbooks.Open(\"" + dashTool.getCanonicalPath() + "\")";
						
		
			
		script += "\nobjExcel.VBE.ActiveVBProject.VBComponents.Import \"" + aggMacro.getCanonicalPath() + "\"\n";

		
		script += "\nobjExcel.Application.Run \"aggregateData\"\nobjExcel.Visible = True\nobjExcel.Application.Run \"MakeWarningReport\"\nobjExcel.Application.Quit";
		PrintWriter writer = new PrintWriter("./Settings/Macros/dataScript.vbs");
		writer.println(script);
		writer.close();
		System.out.println(aggMacro.getCanonicalPath());
		
	}*/

	public void executeWarningScript() throws IOException, InterruptedException {

		String demoCode = "Dim objExcel, objWorkbook\n"
				+ "Set objExcel = CreateObject(\"Excel.Application\")\n"
				+ "objExcel.Visible = False\n"
				+ "Set objWorkbook = objExcel.Workbooks.Open(\""
				+ root
				+ "\\Settings\\Dashboard Settings.xlsm\")\n"
				+ "objExcel.VBE.ActiveVBProject.VBComponents.Import \""
				+ root
				+ "\\Settings\\Macros\\aggregateData.bas\""
				+ "\nobjExcel.Application.Run \"aggregateData\"\nOn Error Resume Next\nobjExcel.Visible = True\nobjExcel.Application.Run \"makeWarningReport\",\""
				+ warningReportsString + "\"\n";

		PrintWriter writer = new PrintWriter("./Settings/Macros/dataScript.vbs");
		writer.println(demoCode);
		writer.close();
		System.out.println(demoCode);

		Runtime.getRuntime().exec(
				"cmd /c start " + root + "\\Settings\\Macros\\dataScript.vbs");

		System.exit(0);

	}

	public void executePrelimScript() throws IOException, InterruptedException {

		String demoCode = "Dim objExcel, objWorkbook\n"
				+ "Set objExcel = CreateObject(\"Excel.Application\")\n"
				+ "objExcel.Visible = False\n"
				+ "Set objWorkbook = objExcel.Workbooks.Open(\""
				+ root
				+ "\\Settings\\Dashboard Settings.xlsm\")\n"
				+ "objExcel.VBE.ActiveVBProject.VBComponents.Import \""
				+ root
				+ "\\Settings\\Macros\\aggregateData.bas\""
				+ "\nobjExcel.Application.Run \"aggregateData\"\nOn Error Resume Next\nobjExcel.Visible = True\nobjExcel.Application.Run \"makePreliminaryReport\",\""
				+ prelimTemplateString + "\", \"" + outputDir + "\"\n";

		PrintWriter writer = new PrintWriter("./Settings/Macros/prelimReportScript.vbs");
		writer.println(demoCode);
		writer.close();
		System.out.println(demoCode);

		Runtime.getRuntime().exec(
				"cmd /c start " + root + "\\Settings\\Macros\\prelimReportScript.vbs");

		
		this.waitForFile(outputDir + "\\Dashboard Settings.xlsm", 1);
	}
	
	public void executeYourStories() throws IOException, InterruptedException {
		String demoCode = "Dim objExcel, objWorkbook\n"
				+ "Set objExcel = CreateObject(\"Excel.Application\")\n"
				+ "objExcel.Visible = False\n"
				+ "Set objWorkbook = objExcel.Workbooks.Open(\""
				+ root
				+ "\\Settings\\Dashboard Settings.xlsm\")\n"
				+ "objExcel.VBE.ActiveVBProject.VBComponents.Import \""
				+ root
				+ "\\Settings\\Macros\\aggregateData.bas\""
				+ "\nobjExcel.Application.Run \"aggregateData\"\nOn Error Resume Next\nobjExcel.Visible = True\nobjExcel.Application.Run \"makeYourStories\",\""
				+ scoreTemplateString + "\", \"" + outputDir + "\"\n";

		PrintWriter writer = new PrintWriter("./Settings/Macros/DashboardsScript.vbs");
		writer.println(demoCode);
		writer.close();
		System.out.println(demoCode);

		Runtime.getRuntime().exec(
				"cmd /c start " + root + "\\Settings\\Macros\\DashboardsScript.vbs");

		
		this.waitForFile(outputDir + "\\Dashboard Settings.xlsm", 1);
		
	}
	
	public void executeDashboardScript() throws IOException, InterruptedException {

		String demoCode = "Dim objExcel, objWorkbook\n"
				+ "Set objExcel = CreateObject(\"Excel.Application\")\n"
				+ "objExcel.Visible = False\n"
				+ "Set objWorkbook = objExcel.Workbooks.Open(\""
				+ root
				+ "\\Settings\\Dashboard Settings.xlsm\")\n"
				+ "objExcel.VBE.ActiveVBProject.VBComponents.Import \""
				+ root
				+ "\\Settings\\Macros\\aggregateData.bas\""
				+ "\nobjExcel.Application.Run \"aggregateData\"\nOn Error Resume Next\nobjExcel.Visible = True\nobjExcel.Application.Run \"makeDashboards\",\""
				+ dashboardTemplateString + "\", \"" + outputDir + "\"\n";

		PrintWriter writer = new PrintWriter("./Settings/Macros/DashboardsScript.vbs");
		writer.println(demoCode);
		writer.close();
		System.out.println(demoCode);

		Runtime.getRuntime().exec(
				"cmd /c start " + root + "\\Settings\\Macros\\DashboardsScript.vbs");

		
		this.waitForFile(outputDir + "\\Dashboard Settings.xlsm", 1);
	}
	
	public boolean isGlossary() {
		
		File testGloss = new File("./Settings/Glossary.pdf");
		
		return testGloss.exists();
		
	}
	
	public void mergePDFs(ArrayList<String> bankNames, String type)
			throws COSVisitorException, IOException {
		String destPath;
		String glossPath = root + "\\Settings\\Glossary.pdf";

		PDFMergerUtility ut;

		if (this.isGlossary() == true) {

			if (type.equals("Preliminary")) {
				for (int x = 0; x < bankNames.size(); x++) {
					destPath = outputDir + "\\" + bankNames.get(x)
							+ "\\Prelim Report " + bankNames.get(x) + ".pdf";
					ut = new PDFMergerUtility();
					ut.addSource(destPath);
					ut.addSource(glossPath);
					ut.setDestinationFileName(destPath);
					ut.mergeDocuments();

				}
			}
			
			if (type.equals("Dashboards")) {
				for (int x = 0; x < bankNames.size(); x++) {
					destPath = outputDir + "\\" + bankNames.get(x)
							+ "\\Dashboard " + bankNames.get(x) + ".pdf";
					ut = new PDFMergerUtility();
					ut.addSource(destPath);
					ut.addSource(glossPath);
					ut.setDestinationFileName(destPath);
					ut.mergeDocuments();

				}
			}
		}

	}
	
	public void waitForFile(String path, int continueCode) throws InterruptedException {
		// Preliminary Report: continueCode = 1 
		// Final Report: continueCode = 2
		// Score Files: continueCode = 3
		
		
		
		File createdFile = new File(path);
		int i = 0;
		while (createdFile.exists() == false && i < 900) {
			
			System.out.println("DNE");
			Thread.sleep(1000);
			i++;
			
		}
		
		if (i == 900) {
		JFrame dialog = new JFrame();
		JOptionPane.showMessageDialog(dialog, "Something went wrong, DashDesigner closed");
		System.exit(0);
		}
		
		
	}

	public void waitForSettings() throws InterruptedException {
		
		//File savedSettings = new File(".//WMLC//Output//Dashboard Settings.xlsm");
		
		//while (savedSettings.exists() == false) {
			
			System.out.println("DNE");
			Thread.sleep(30000);
			System.exit(0);
		//}
		
		// NextStep
			
	}
	
	public String getRootPath() throws IOException {
		return root;
	}
	
	public static void main(String[] args) throws IOException, InterruptedException, COSVisitorException {
		
		/*DateFormat dateFormat = new SimpleDateFormat("MM-dd-yyyy hh.mm");
		Calendar cal = Calendar.getInstance();
		System.out.println(dateFormat.format(cal.getTime()));*/
		//DashChecker test = new DashChecker();
		
		PDFMergerUtility ut = new PDFMergerUtility();
		
		ut.addSource("C:\\Users\\mdinardo\\Desktop\\WealthDataChecker\\Output\\10-20-2014 09.05\\Prelim Report Associated Bank.pdf");
		ut.addSource("C:\\Users\\mdinardo\\Desktop\\WealthDataChecker\\Output\\10-20-2014 09.05\\Glossary.pdf");
		
	
		ut.setDestinationFileName("C:\\Users\\mdinardo\\Desktop\\WealthDataChecker\\Output\\10-20-2014 09.05\\Prelim Report Associated Bank.pdf");
		ut.mergeDocuments();
		
		System.out.println("worked?)");
	}

}
