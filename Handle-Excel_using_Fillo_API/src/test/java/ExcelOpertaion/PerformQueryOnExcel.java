package ExcelOpertaion;

import com.codoid.products.exception.FilloException;
import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import com.codoid.products.fillo.Recordset;
import jdk.jfr.Description;
import org.testng.annotations.AfterTest;
import org.testng.annotations.Test;

public class PerformQueryOnExcel {

	private String query;
	private String filePath;

	@Description("display the complete data from the excel")
	@Test
	public void test1() {
		query = "SELECT * FROM InfoSheet1";
		filePath = "./files/SampleExcelFile1.xlsx";
		try {
			Fillo fillo = new Fillo();
			Connection connection = fillo.getConnection(filePath);
			Recordset recordset = connection.executeQuery(query);
			while (recordset.next()) {
				System.out.println(recordset.getField("StudentID") + " " + recordset.getField("First Name") + " "
						+ recordset.getField("Last Name") + " " + recordset.getField("ProgrammingLanguage") + " "
						+ recordset.getField("Age"));
			}
			recordset.close();
			connection.close();
		} catch (FilloException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}


	@Description("Total no of columns")
	@Test
	public void test2() {
		query = "SELECT * FROM InfoSheet1";
		filePath = "./files/SampleExcelFile1.xlsx";
		try {
			Fillo fillo = new Fillo();
			Connection connection = fillo.getConnection(filePath);
			Recordset recordset = connection.executeQuery(query);
			// get the number of columns
			int totalColumns = recordset.getFieldNames().size();
			System.out.println("Total number of columns is: " + totalColumns);
			// get the name of all the fields
			for (String eachField : recordset.getFieldNames()) {
				System.out.print(eachField + ", ");
			}
			System.out.println();
			recordset.close();
			connection.close();
		} catch (FilloException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

    @Description("Where condition use for particular value")
	@Test
	public void test3() {
		query = "SELECT * FROM InfoSheet1 WHERE ProgrammingLanguage='C'";
		filePath = "./files/SampleExcelFile1.xlsx";
		try {
			Fillo fillo = new Fillo();
			Connection connection = fillo.getConnection(filePath);
			Recordset recordset = connection.executeQuery(query);
			while (recordset.next()) {
				System.out.println(recordset.getField("StudentID") + " " + recordset.getField("First Name") + " "
						+ recordset.getField("Last Name") + " " + recordset.getField("ProgrammingLanguage") + " "
						+ recordset.getField("Age"));
			}
			recordset.close();
			connection.close();
		} catch (FilloException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}


	@Description("using WHERE clause with AND operator inside the query")
	@Test
	public void test4() {
		query = "SELECT * FROM InfoSheet1 WHERE (\"Last Name\"='Shyam' AND \"Age\"=32)";
		filePath = "./files/SampleExcelFile1.xlsx";
		try {
			Fillo fillo = new Fillo();
			Connection connection = fillo.getConnection(filePath);
			Recordset recordset = connection.executeQuery(query);
			while (recordset.next()) {
				System.out.println(recordset.getField("StudentID") + " " + recordset.getField("First Name") + " "
						+ recordset.getField("Last Name") + " " + recordset.getField("ProgrammingLanguage") + " "
						+ recordset.getField("Age"));
			}
			recordset.close();
			connection.close();
		} catch (FilloException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}




	@Description("get records where a particular column is empty")
	@Test
	public void test5() {
		query = "select * FROM InfoSheet1 WHERE StudentID=null";
		filePath = "./files/SampleExcelFile1.xlsx";
		try {
			Fillo fillo = new Fillo();
			Connection connection = fillo.getConnection(filePath);
			Recordset recordset = connection.executeQuery(query);
			while (recordset.next()) {
				System.out.println(recordset.getField("StudentID") + " " + recordset.getField("First Name") + " "
						+ recordset.getField("Last Name") + " " + recordset.getField("ProgrammingLanguage") + " "
						+ recordset.getField("Age"));
			}
			recordset.close();
			connection.close();
		} catch (FilloException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}


	@AfterTest
	public void afterMethod() {

		System.out.println("*****************************************************************");
	}

}
