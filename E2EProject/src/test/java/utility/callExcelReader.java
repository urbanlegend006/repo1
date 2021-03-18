package utility;

import java.util.List;

public class callExcelReader {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		excelReader xl = new excelReader();
		
		List<List<String>> dataXl = xl.getExcelDataAsLists("E:\\ImportExcel.xlsx");
		System.out.println(dataXl);
	
	}

}
