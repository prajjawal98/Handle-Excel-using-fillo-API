package ExcelOpertaion;

import com.codoid.products.exception.FilloException;
import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import org.testng.annotations.Test;

public class InsertValuesInExcel {

	private String query;
	private String filePath;

	@Test
	public void test1() {
		query = "INSERT INTO \"Info - Sheet3\"(StudentID,\"First Name\",\"Last Name\",ProgrammingLanguage,\"Age\") VALUES(3,'popi90','kansal','x',67)";
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
