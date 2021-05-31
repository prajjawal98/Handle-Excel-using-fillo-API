package ExcelOpertaion;

import com.codoid.products.exception.FilloException;
import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import org.testng.annotations.Test;

public class UpdateExcelData {

	private String query;
	private String filePath;

	@Test
	public void test1() {
		query = "UPDATE \"Info Sheet2\" SET ProgrammingLanguage='Selenium' WHERE (StudentID=2 and \"First Name\"='Aditya' AND \"Last Name\"='Dubey')";
		filePath = "./files/SampleExcelFile1.xlsx";
		try {
			Fillo fillo = new Fillo();
			Connection connection = fillo.getConnection(filePath);
			connection.executeUpdate(query);
			connection.close();
		} catch (FilloException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
