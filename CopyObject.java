
public class CopyObject {
	
	String bankName;
	int type;
	int slideNum;
	String sheetName;
	String tableName;
	String name;
	String left;
	String top;
	String width;
	String height;
	String range;
	
	public CopyObject(int slideNum, String name, String left, String top, String width, String height, String bankName) {
		
		this.type = 1;
		this.slideNum = slideNum;
		this.name = name;
		this.left = left;
		this.top = top;
		this.width = width;
		this.height = height;
		this.bankName = bankName;
		
	}
	
	public CopyObject(String sheetName, String tableName, int slideNum, String left, String top, String width, String height, String bankName) {
		
		this.type = 2;
		this.sheetName = sheetName;
		this.tableName = tableName;
		this.slideNum = slideNum;
		this.left = left;
		this.top = top;
		this.width = width;
		this.height = height;
		this.bankName = bankName;
		
	}
	
	public CopyObject(String sheetName, String range, int slideNum, String left, String top, String width, String height) {
		this.type = 3;
		this.sheetName = sheetName;
		this.range = range;
		this.slideNum = slideNum;
		this.left = left;
		this.top = top;
		this.width = width;
		this.height = height;
		
		
		
	}
	
	public String getBankName() {
// Changed
		return this.bankName;
	}
	
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}
	
	public String getSheetName() {
		return this.sheetName;
	}
	
	public void setTableName(String tableName) {
		this.tableName = tableName;
	}
	
	public String getTableName() {
		return this.tableName;
	}

	public int getSlideNum() {
		return slideNum;
	}

	public void setSlideNum(int slideNum) {
		this.slideNum = slideNum;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getLeft() {
		return left;
	}

	public void setLeft(String left) {
		this.left = left;
	}

	public String getTop() {
		return top;
	}

	public void setTop(String top) {
		this.top = top;
	}

	public String getWidth() {
		return width;
	}

	public void setWidth(String width) {
		this.width = width;
	}

	public String getHeight() {
		return height;
	}

	public void setHeight(String height) {
		this.height = height;
	}

	public String toString() {
		String output = "";

		if (this.type == 1) {

			output += "Call copyObject(\"" + this.getName() + "\","
					+ this.getSlideNum() + "," + this.getLeft() + ","
					+ this.getTop() + "," + this.getWidth() + ","
					+ this.getHeight() + ",PowerPointApp)\n";

		} else if (this.type == 2) {

			String sheetNameFixed = this.sheetName.substring(1,
					this.sheetName.length());

			output += "Call copyTable(\"" + sheetNameFixed + "\"," + "\""
					+ this.tableName + "\"," + this.getSlideNum() + ","
					+ this.getLeft() + "," + this.getTop() + ","
					+ this.getWidth() + "," + this.getHeight() + ",\""
					+ this.bankName + "\",PowerPointApp)\n";
		} else if (this.type == 3) {
			String sheetNameFixed = this.sheetName.substring(1,
					this.sheetName.length());
			String rangeFixed = this.range.replaceAll("-", ":");
			

			output += "Call copyPicture(\"" + sheetNameFixed + "\"," + "\""
					+ rangeFixed + "\"," + this.getSlideNum() + ","
					+ this.getLeft() + "," + this.getTop() + ","
					+ this.getWidth() + "," + this.getHeight() + ",PowerPointApp)\n";
		}

		return output;
	}

}
