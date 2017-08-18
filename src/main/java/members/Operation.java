package members;

public enum Operation {

	AVERAGE("Average"), SUM("Sum"), MIN("Min"), MAX("Max"), RANK("Rank");
	
	private String title;
	
	private Operation(String title) {
		this.title = title;
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

}
