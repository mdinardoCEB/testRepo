import java.util.ArrayList;


public class PeerGroup {
	
	private String name;
	private String description;
	private ArrayList<String> bankList;
	private ArrayList<String> userList;
	
	public PeerGroup(String name, String description, ArrayList<String> bankList, ArrayList<String> userList) {
		
		this.name = name;
		this.description = description;
		this.bankList = bankList;
		this.userList = userList;
		
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getDescription() {
		return description;
	}

	public void setDescription(String description) {
		this.description = description;
	}

	public ArrayList<String> getBankList() {
		return bankList;
	}

	public void setBankList(ArrayList<String> bankList) {
		this.bankList = bankList;
	}

	public ArrayList<String> getUserList() {
		return userList;
	}

	public void setUserList(ArrayList<String> userList) {
		this.userList = userList;
	}
	
	
	
}
