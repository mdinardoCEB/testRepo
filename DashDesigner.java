import java.awt.Color;
import java.awt.Component;
import java.awt.FlowLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintWriter;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;

import javax.swing.BoxLayout;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.UIManager;

import org.apache.pdfbox.exceptions.COSVisitorException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/* Demonstration
 * 
 * 
 */
public class DashDesigner implements ActionListener {
	JFrame mainFrame = new JFrame("DashDesigner");
	JLabel dataChecksLab = new JLabel(
			"1. Aggregating data and applying data checks...");
	JLabel prelimReportLab = new JLabel("2. Generating Preliminary Reports...");
	JLabel finalReportLab = new JLabel(
			"3. Calculating benchmarks and generating Dashboard...");
	JLabel scoreReportLab = new JLabel("4. Generating custom reports...");
	JButton nextButton = new JButton("Next");
	File rootDirectory = new File("Data Warning Reports.xlsm");
	String root = System.getProperty("user.dir");
	String outputDirectory;
	String productionID;
	JFrame loadingFrame = new JFrame("Loading");

	PeerList peers;
	File settings;
	XSSFWorkbook settingsWb;
	XSSFWorkbook dataWb;
	Sheet previewSheet;

	ArrayList<String> bankNameList = new ArrayList<String>();
	DashChecker checker;

	JFileChooser fc = new JFileChooser();

	JLabel mainMenLab = new JLabel("Choose an action");
	JButton warnReportBut = new JButton("Generate Warning Report");
	JButton preliminaryReportBut = new JButton("Generate Preliminary Report");
	JButton dashboardReportBut = new JButton("Generate Dashboard");

	public DashDesigner() throws FileNotFoundException, IOException {
		System.out.println(root);
		// settings = new File(root + "\\Settings\\Dashboard Settings.xlsm");

		// settingsWb = new XSSFWorkbook(new FileInputStream(settings));
		// peers = new PeerList(settingsWb);
		// this.loadBankNames();
		this.createAndShowGUI();

	}

	public void createAndShowGUI() {

		loadingFrame.setSize(300, 200);
		loadingFrame.getContentPane().setLayout(new FlowLayout());
		JLabel loadingLab = new JLabel("Loading...");
		loadingFrame.add(loadingLab);
		loadingFrame.setLocationRelativeTo(null);

		warnReportBut.setActionCommand("GenWarningReport");
		preliminaryReportBut.setActionCommand("GenPrelimReport");
		dashboardReportBut.setActionCommand("GenDashboard");

		warnReportBut.addActionListener(this);
		preliminaryReportBut.addActionListener(this);
		dashboardReportBut.addActionListener(this);

		mainFrame.getContentPane().setLayout(
				new BoxLayout(mainFrame.getContentPane(), BoxLayout.Y_AXIS));
		mainFrame.setSize(600, 100);

		JPanel topPane = new JPanel();
		// topPane.setAlignmentX(Component.LEFT_ALIGNMENT);
		mainMenLab.setAlignmentX(Component.LEFT_ALIGNMENT);
		topPane.add(mainMenLab);
		mainFrame.add(topPane);
		JPanel buttonPane = new JPanel();
		buttonPane.setLayout(new FlowLayout());
		buttonPane.add(warnReportBut);
		buttonPane.add(preliminaryReportBut);
		buttonPane.add(dashboardReportBut);
		mainFrame.add(buttonPane);

		mainFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		mainFrame.setLocationRelativeTo(null);
		mainFrame.setVisible(true);

	}

	public void showStartMenu() {

	}

	public void startDemonstration() throws IOException, InterruptedException {
		this.createAndShowGUI();
		/*
		 * this.loadBankNames(); this.makeNewDirectories(); checker = new
		 * DashChecker(); dataChecksLab
		 * .setText("1. Aggregating data and applying data checks... DONE");
		 * //For now System.exit(0); dataChecksLab.setForeground(Color.green);
		 * mainFrame.add(prelimReportLab);
		 * nextButton.setActionCommand("genPrelimReports");
		 * mainFrame.add(nextButton); mainFrame.validate();
		 */

		/*
		 * this.generatePrelimReports(); this.waitForPrelim();
		 * prelimReportLab.setText(prelimReportLab.getText() + "DONE");
		 * prelimReportLab.setForeground(Color.green);
		 * mainFrame.add(finalReportLab); mainFrame.validate();
		 * this.makeVBAPeerArrays(); this.executeDashboardScript();
		 * this.waitForDashboards(); this.executeScoreScript();
		 */

	}

	public void generatePrelimReports() throws IOException {

		String macro = "Sub genPrelimReports()\n"
				+ "Call pasteTableToPrelim(\"" + checker.getRootPath()
				+ "\\Output\\Dashboard Settings.xlsm\", \""
				+ checker.getRootPath()
				+ "\\Output\\Preliminary Report.xlsm\")\n" + "End Sub";

		PrintWriter writer = new PrintWriter(new File(
				".//WMLC//Resources//Macros//PrelimMacro.bas"));
		writer.println(macro);
		writer.close();

		String code = "Dim objExcel, objWorkbook\n"
				+ "Set objExcel = CreateObject(\"Excel.Application\")\n"
				+ "objExcel.Visible = True\n"
				+ "Set objWorkbook = objExcel.Workbooks.Open(\""
				+ checker.getRootPath()
				+ "\\Resources\\Templates\\Preliminary Report\\WMLCPrelimTemplate.xlsm\")\n"
				+ "objExcel.VBE.ActiveVBProject.VBComponents.Import \""
				+ checker.getRootPath()
				+ "\\Resources\\Macros\\PrelimMacro.bas\""
				+ "\nobjExcel.Application.Run \"genPrelimReports\"\n";

		PrintWriter writer2 = new PrintWriter(
				".//WMLC/Resources/Macros/PrelimScript.vbs");
		writer2.println(code);
		writer2.close();

		Runtime.getRuntime().exec(
				"cmd /c start " + checker.getRootPath()
						+ "\\Resources\\Macros\\PrelimScript.vbs");

	}

	public void makeVBAPeerArrays() throws FileNotFoundException {
		String masterString = "Function getPeerList(bankName As String) As Variant\n";

		String code = "";
		String code2 = "";
		String mstrUserList = "";
		mstrUserList = "Dim masterUsrList As Variant\n";
		mstrUserList += "masterUsrList = Array(";
		String mstrGroupList = "";
		mstrGroupList = "Dim masterGrpList As Variant\n";
		mstrGroupList += "masterGrpList = Array(";
		// For each peer group define an Array of bank names
		for (int x = 0; x < peers.getPeerGroupList().size(); x++) {

			code += "Dim peerGrp" + x + " As Variant\n";
			code += "peerGrp" + x + "= Array(";
			code2 += "Dim userGrp" + x + " As Variant\n";
			code2 += "userGrp" + x + "=Array(";

			mstrUserList += "userGrp" + x;
			mstrGroupList += "peerGrp" + x;

			if (x < (peers.getPeerGroupList().size() - 1)) {
				mstrUserList += ",";
				mstrGroupList += ",";
			}

			// For this array, add each bank name to the array
			for (int y = 0; y < peers.getPeerGroupList().get(x).getBankList()
					.size(); y++) {
				code += "\""
						+ peers.getPeerGroupList().get(x).getBankList().get(y)
						+ "\"";
				// If it isn't the last bank in the list, add a comma for
				// separation
				if (y < (peers.getPeerGroupList().get(x).getBankList().size() - 1)) {
					code += ",";
				}
			}
			code += ")\n";

			for (int y = 0; y < peers.getPeerGroupList().get(x).getUserList()
					.size(); y++) {
				code2 += "\""
						+ peers.getPeerGroupList().get(x).getUserList().get(y)
						+ "\"";
				// If it isn't the last bank in the list, add a comma for
				// separation
				if (y < (peers.getPeerGroupList().get(x).getUserList().size() - 1)) {
					code2 += ",";
				}
			}
			code2 += ")\n";

		}
		mstrUserList += ")\n";
		mstrGroupList += ")\n";

		System.out.println(code);
		System.out.println(mstrGroupList);
		System.out.println(code2);
		System.out.println(mstrUserList);

		masterString += code + "\n" + mstrGroupList + "\n" + code2 + "\n"
				+ mstrUserList;

		masterString += "\nDim x As Integer\nFor x = 0 To UBound(masterUsrList)\n"
				+ "For y = 0 To UBound(masterUsrList(x))\n"
				+ "If (masterUsrList(x)(y)) = bankName Then\n"
				+ "getPeerList = masterGrpList(x)\nEnd If\nNext y\nNext x\nEnd Function";

		PrintWriter writer = new PrintWriter(new File(
				".\\WMLC\\Resources\\Macros\\PeerGroupArrays.bas"));
		writer.println(masterString);
		writer.close();

		// this.loadBankNames();
		// this.makeNewDirectories();
	}

	public void executeDashboardScript() throws IOException {

		String pullMacro = "Sub runDashboard()\n" + "Call pullData(\""
				+ checker.getRootPath()
				+ "\\Output\\Dashboard Settings.xlsm\", \""
				+ checker.getRootPath()
				+ "\\Output\\Dashboards.xlsm\")\nEnd Sub";

		PrintWriter writer = new PrintWriter(
				".\\WMLC\\Resources\\Macros\\runDashboards.bas");
		writer.println(pullMacro);
		writer.close();

		String script = "Dim objExcel, objWorkbook\n"
				+ "Set objExcel = CreateObject(\"Excel.Application\")\n"
				+ "objExcel.Visible = True\n"
				+ "Set objWorkbook = objExcel.Workbooks.Open(\""
				+ checker.getRootPath()
				+ "\\Resources\\Templates\\Final Report\\WMLCDashboardTemplate.xlsm\")\n"
				+ "objExcel.VBE.ActiveVBProject.VBComponents.Import \""
				+ checker.getRootPath()
				+ "\\Resources\\Macros\\PeerGroupArrays.bas\"\n"
				+ "objExcel.VBE.ActiveVBProject.VBComponents.Import \""
				+ checker.getRootPath()
				+ "\\Resources\\Macros\\runDashboards.bas\"\n"
				+ "\nobjExcel.Application.Run \"runDashboard\"";

		PrintWriter writer2 = new PrintWriter(
				".\\WMLC\\Resources\\Macros\\DashboardScript.vbs");
		writer2.println(script);
		writer2.close();

		Runtime.getRuntime().exec(
				"cmd /c start " + checker.getRootPath()
						+ "\\Resources\\Macros\\DashboardScript.vbs");

	}

	public void executeScoreScript() throws IOException {

		String code = "Sub callAll()\n";
		for (int x = 0; x < bankNameList.size(); x++) {

			code += "Call changeBank(\""
					+ bankNameList.get(x)
					+ "\", \""
					+ checker.getRootPath()
					+ "\\Output\\"
					+ bankNameList.get(x)
					+ "\\"
					+ bankNameList.get(x)
					+ " Score File.xlsm\", \""
					+ checker.getRootPath()
					+ "\\Resources\\Templates\\Your Story Template\\Your Story Sample.pptx\",\""
					+ checker.getRootPath() + "\\Output\\"
					+ bankNameList.get(x) + "\\" + bankNameList.get(x)
					+ " Sample Deck.pptx\")\n";

		}

		code += "End Sub";

		PrintWriter writer = new PrintWriter(new File(
				".\\WMLC\\Resources\\Macros\\callAllScores.bas"));
		writer.println(code);
		writer.close();
		System.out.println(code);
		String script = "Dim objExcel, objWorkbook\n"
				+ "Set objExcel = CreateObject(\"Excel.Application\")\n"
				+ "objExcel.Visible = True\n"
				+ "Set objWorkbook = objExcel.Workbooks.Open(\""
				+ checker.getRootPath()
				+ "\\Resources\\Templates\\Your Story Template\\Wealth Score Template.xlsm\")\n"
				+ "objExcel.VBE.ActiveVBProject.VBComponents.Import \""
				+ checker.getRootPath()
				+ "\\Resources\\Macros\\callAllScores.bas\"\n"
				+ "\nobjExcel.Application.Run \"callAll\"\n";

		PrintWriter writer2 = new PrintWriter(new File(
				".\\WMLC\\Resources\\Macros\\makeScoreFiles.vbs"));
		writer2.println(script);
		writer2.close();

		Runtime.getRuntime().exec(
				"cmd /c start " + checker.getRootPath()
						+ "\\Resources\\Macros\\makeScoreFiles.vbs");

	}

	public void makeOutputFolder() throws IOException {

		// Each production run will have a new folder in .\Output\
		DateFormat dateFormat = new SimpleDateFormat("MM-dd-yyyy hh.mm");
		Calendar cal = Calendar.getInstance();

		String prodID = dateFormat.format(cal.getTime());
		productionID = prodID;
		File dir = new File("./Output/" + prodID);
		dir.mkdir();

		outputDirectory = dir.getCanonicalPath();

	}

	public void makeNewDirectories() {

		for (int x = 0; x < bankNameList.size(); x++) {
			System.out.println(x);
			File dir = new File("./Output/" + productionID + "/"
					+ bankNameList.get(x));
			dir.mkdir();

		}

	}

	public void loadBankNames() throws FileNotFoundException, IOException {

		dataWb = new XSSFWorkbook(new FileInputStream(new File(outputDirectory
				+ "\\Dashboard Settings.xlsm")));
		previewSheet = dataWb.getSheet("Preview Table");
		System.out.println("here");
		for (Row r : previewSheet) {

			if (r.getCell(0).getCellType() == 1
					&& !r.getCell(0, Row.RETURN_BLANK_AS_NULL)
							.getRichStringCellValue().equals(null)) {
				System.out.println(r.getCell(0).getRichStringCellValue()
						.toString());
				if (!r.getCell(0).getRichStringCellValue().toString()
						.equals("Bank"))
					bankNameList.add(r.getCell(0).getRichStringCellValue()
							.toString());

			}

		}

		/*
		 * for (int x = 0; x < peers.getPeerGroupList().size(); x++) {
		 * 
		 * for (int y = 0; y < peers.getPeerGroupList().get(x).getUserList()
		 * .size(); y++) {
		 * 
		 * bankNameList.add(peers.getPeerGroupList().get(x).getUserList()
		 * .get(y));
		 * 
		 * }
		 * 
		 * }
		 */

	}

	public void waitForDashboards() throws InterruptedException {

		File savedSettings = new File(".//WMLC//Output//Dashboards.xlsm");

		while (savedSettings.exists() == false) {

			System.out.println("DNE");
			Thread.sleep(1000);
		}

		// NextStep
	}

	public void waitForPrelim() throws InterruptedException {

		File savedSettings = new File(
				".//WMLC//Output//Preliminary Report.xlsm");

		while (savedSettings.exists() == false) {

			System.out.println("DNE");
			Thread.sleep(1000);
		}

		// NextStep
	}

	public static void main(String[] args) throws FileNotFoundException,
			IOException, InterruptedException {

		try {
			UIManager
					.setLookAndFeel("com.seaglasslookandfeel.SeaGlassLookAndFeel");
		} catch (Exception e) {
			System.out.println("error");
			e.printStackTrace();
		}

		DashDesigner test = new DashDesigner();

	}

	@Override
	public void actionPerformed(ActionEvent e) {
		String command = e.getActionCommand();

		if (command.equals("GenWarningReport")) {

			mainFrame.setVisible(false);

			try {
				checker = new DashChecker();
				checker.executeWarningScript();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			} catch (InterruptedException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}

		}

		if (command.equals("GenPrelimReport")) {
			loadingFrame.setVisible(true);
			mainFrame.setVisible(false);
			;

			loadingFrame.validate();
			try {
				this.makeOutputFolder();
				checker = new DashChecker(outputDirectory);
				checker.executePrelimScript();
				this.loadBankNames();
				this.makeNewDirectories();
				for (int x = 0; x < bankNameList.size(); x++) {

					System.out.println(bankNameList.get(x));
				}
				checker.waitForFile(outputDirectory
						+ "\\Preliminary Reports.xlsm", 1);
				checker.mergePDFs(bankNameList, "Preliminary");

				System.exit(0);
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			} catch (InterruptedException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			} catch (COSVisitorException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}

		}
		if (command.equals("GenDashboard")) {

			mainFrame.setVisible(false);
			loadingFrame.setVisible(true);
			loadingFrame.validate();
			try {
				this.makeOutputFolder();
				checker = new DashChecker(outputDirectory);
				checker.executeDashboardScript();
				this.loadBankNames();
				this.makeNewDirectories();
				for (int x = 0; x < bankNameList.size(); x++) {

					System.out.println(bankNameList.get(x));
				}
				checker.waitForFile(outputDirectory + "\\Dashboards.xlsm", 1);
				checker.mergePDFs(bankNameList, "Dashboards");

				System.exit(0);
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			} catch (InterruptedException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			} catch (COSVisitorException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
		}

		/*
		 * if (command.equals("genPrelimReports")) {
		 * 
		 * 
		 * try { this.generatePrelimReports(); this.waitForPrelim(); } catch
		 * (InterruptedException e1) { // TODO Auto-generated catch block
		 * e1.printStackTrace(); } catch (IOException e1) { // TODO
		 * Auto-generated catch block e1.printStackTrace(); }
		 * prelimReportLab.setText(prelimReportLab.getText() + "DONE");
		 * prelimReportLab.setForeground(Color.green);
		 * mainFrame.remove(nextButton);
		 * nextButton.setActionCommand("genDashboards");
		 * mainFrame.add(finalReportLab); mainFrame.add(nextButton);
		 * mainFrame.validate();
		 * 
		 * }
		 */

		if (command.equals("genDashboards")) {

			try {
				this.makeVBAPeerArrays();
			} catch (FileNotFoundException e2) {
				// TODO Auto-generated catch block
				e2.printStackTrace();
			}
			try {
				this.executeDashboardScript();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			try {
				this.waitForDashboards();
			} catch (InterruptedException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}

			finalReportLab.setText(finalReportLab.getText() + "DONE");
			finalReportLab.setForeground(Color.green);
			mainFrame.remove(nextButton);
			nextButton.setActionCommand("genScoresAndReports");
			mainFrame.add(scoreReportLab);
			mainFrame.add(nextButton);
			mainFrame.validate();

		}

		if (command.equals("genScoresAndReports")) {

			try {
				this.executeScoreScript();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}

			mainFrame.remove(nextButton);
			scoreReportLab.setText(scoreReportLab.getText() + "DONE");
			scoreReportLab.setForeground(Color.green);
			mainFrame.validate();

		}
	}
}
