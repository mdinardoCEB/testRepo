import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class PeerList {
	
	XSSFWorkbook wb;
	Sheet ws;
	
	private File listFile;
	private ArrayList<PeerGroup> peerGroupList = new ArrayList<PeerGroup>();
	
	public PeerList(XSSFWorkbook listFile) throws FileNotFoundException, IOException {
		
		this.wb = listFile;
		this.loadPeerGroups();
		System.out.println(this.toString());
		
	}
	
	public XSSFWorkbook getWb() {
		return wb;
	}

	public void setWb(XSSFWorkbook wb) {
		this.wb = wb;
	}

	public Sheet getWs() {
		return ws;
	}

	public void setWs(Sheet ws) {
		this.ws = ws;
	}

	public File getListFile() {
		return listFile;
	}

	public void setListFile(File listFile) {
		this.listFile = listFile;
	}

	public ArrayList<PeerGroup> getPeerGroupList() {
		return peerGroupList;
	}

	public void setPeerGroupList(ArrayList<PeerGroup> peerGroupList) {
		this.peerGroupList = peerGroupList;
	}

	public void loadPeerGroups() throws FileNotFoundException, IOException {
		
		String name = "";
		String description = "";
		
		ws = wb.getSheet("Peer Groups");
		
		int x = 1;
		Cell c = ws.getRow(0).getCell(x, Row.RETURN_BLANK_AS_NULL);
		
		
		
		// get Name and Description of peer group
		while (c != null) {
			ArrayList<String> bankList = new ArrayList<String>();
			ArrayList<String> userList = new ArrayList<String>();
			
			name = c.getRichStringCellValue().toString();
			description = ws.getRow(1).getCell(x).getRichStringCellValue().toString();
			
			
			int y = 2;
			
			Cell c2 = ws.getRow(y).getCell(x, Row.RETURN_BLANK_AS_NULL);
			
			
			// get list of banks in peer group
			while (c2 != null) {
				
				bankList.add(c2.getRichStringCellValue().toString());
				y++;
				if (ws.getRow(y) != null) {
					c2 = ws.getRow(y).getCell(x, Row.RETURN_BLANK_AS_NULL);
				}
				
				else if (ws.getRow(y) == null) {
					c2 = null;
				}
			}
			
			int z = 2;
			
			Cell c3 = ws.getRow(z).getCell(0, Row.RETURN_BLANK_AS_NULL);
			
			
			// find start of Bank List
			while (c3 == null) {
				
				z++;
				if (ws.getRow(z) != null) {
					c3 = ws.getRow(z).getCell(0,Row.RETURN_BLANK_AS_NULL);
				}
				
			}
			
			z++;
			
			while (ws.getRow(z) != null) {
				
				
				c3 = ws.getRow(z).getCell(1,Row.RETURN_BLANK_AS_NULL);
				
				
				if (c3 != null && c3.getRichStringCellValue().toString().equals(name)) {
					
					userList.add(ws.getRow(z).getCell(0).getRichStringCellValue().toString());
					
				}
				
				z++;
				//c3 = ws.getRow(z).getCell(1, Row.RETURN_BLANK_AS_NULL);
				
			}
			
			PeerGroup tempPeerGroup = new PeerGroup(name, description, bankList, userList);
			
			peerGroupList.add(tempPeerGroup);
			
			x++;
			
			c = ws.getRow(0).getCell(x, Row.RETURN_BLANK_AS_NULL);
			
		}
	}
	
	
	
	@Override
	public String toString() {
		String retString = "";
		System.out.println("Go");
		for (int x = 0; x < peerGroupList.size(); x++) {
			
			PeerGroup currentGroup = peerGroupList.get(x);
			retString += "NAME: " + currentGroup.getName() + "\nDESCRIPTION: " + currentGroup.getDescription() + "\nBANKS: ";
			
			for (int y = 0; y < currentGroup.getBankList().size(); y++) {
				
				retString += currentGroup.getBankList().get(y) + ", ";
				
			}
			
			retString += "\nUSERS: ";
			
			for (int z = 0; z < currentGroup.getUserList().size(); z++) {
				
				retString += currentGroup.getUserList().get(z) + ", ";
				
			}
			
			retString += "\n\n\n";
		}
		
		return retString;
	}
	
	public static void main(String[] args) throws FileNotFoundException, IOException {
		
		//File groups = new File("PeerGroupsFY2013.xlsx");
		
		//PeerList list = new PeerList(groups);
		
	//	list.loadPeerGroups();
		
		
	}

}
