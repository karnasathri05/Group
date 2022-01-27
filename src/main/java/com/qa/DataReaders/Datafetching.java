package com.qa.DataReaders;



import org.testng.annotations.Test;

import com.codoid.products.exception.FilloException;
import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import com.codoid.products.fillo.Recordset;

public class Datafetching {
	@Test
	public void movie() throws FilloException {
//public static void main(String[] args) throws FilloException {
	
		
		Fillo fillo = new Fillo();
		Connection connection =fillo.getConnection("C:\\Users\\admin\\Desktop\\MAVEN\\Excell\\Excelldata.xlsx");
		
		String strQuari ="Select * from Sheet1";
		Recordset rs =connection.executeQuery(strQuari);
		//print total Excell dsta
		while(rs.next()) {
			System.out.println(rs.getField("First name")+"----"+rs.getField("Last name")+"----"+rs.getField("Email")+"----"+rs.getField("cell no"));
			
		}
		//Total Rows in Excell Sheet
		System.out.println("Total Rows in excel" + rs.getCount());
		
		rs.moveLast();
		System.out.println(rs.getField("First name")+"----"+rs.getField("Last name")+"----"+rs.getField("Email")+"----"+rs.getField("cell no"));
         
		rs.movePrevious();
		System.out.println(rs.getField("First name")+"----"+rs.getField("Last name")+"----"+rs.getField("Email")+"----"+rs.getField("cell no"));

	
		
		rs.close();
		connection.close();
	}

}

