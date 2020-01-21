import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.Set;
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Database {
	private static Workbook wb;
	private static Sheet sheet;
	private static FileInputStream fis;
	public static void main(String args[]) throws Exception {
		Row row;
		try {
			Class.forName("com.mysql.cj.jdbc.Driver");
			Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306/db?serverTimezone=IST", "Sarina",
					"1234");
			Statement stat = con.createStatement();
			String query = "INSERT INTO newtable (`ResourceName`,`Signum`,`Jobstage`,`Project`,`ResourceStartDate`,`Location`,`Competency`) VALUES (?,?,?,?,?,?,?)";
			PreparedStatement ps = con.prepareStatement(query);
			fis = new FileInputStream("C:\\Users\\eracesh\\Downloads\\Current_Resource_List - Jan 02 2020.xlsx");
			wb = WorkbookFactory.create(fis);
			sheet = wb.getSheet("Sprint Account");
			int noOfRows = sheet.getLastRowNum();
			System.out.println("\n");
			Set<String> myset = new TreeSet<String>();
			// int noOfRows = sheet.getFirstRowNum();
			// System.out.println(noOfRows);
			for (int i = 1; i < noOfRows; i++) {
				row = sheet.getRow(i);
				if (null != row) {
					//ps.setString(1,row.getCell(0).toString());
					if(null != row.getCell(0)) {
					ps.setString(1,row.getCell(0).toString());
					}
					if(null != row.getCell(2)) {
					ps.setString(2,row.getCell(2).toString());
					}
					if(null != row.getCell(3)) {
					ps.setString(3,row.getCell(3).toString());
					}
					if(null != row.getCell(5)) {
					String sql = "select ProjectId from resource where Project ='"+row.getCell(5).toString()+"'";
					ResultSet rs=stat.executeQuery(sql); 
					//if(null != row.getCell(4)) {
					if(rs.next()) {
						ps.setInt(4, rs.getInt(1));
					}
					}
					if(null != row.getCell(6)) {
					ps.setString(5,row.getCell(6).toString());
					}
					if(null != row.getCell(10)) {
					ps.setString(6,row.getCell(10).toString());
					}
					if(null != row.getCell(12)) {
						String sql1 = "select ProjectId from comp where Competency='"+row.getCell(12).toString()+"'";
						ResultSet rs1=stat.executeQuery(sql1); 
						if(null != row.getCell(12)) {
							if(rs1.next()) {	
								ps.setInt(7, rs1.getInt(1));
							}
							}
					}
				
					/*if(null !=row.getCell(12)) {
						ps.setString(7, row.getCell(12).toString());
					}*/
				/*	String sql1 = "select ProjectId from comp where Competency='"+row.getCell(12).toString()+"'";
					ResultSet rs1=stat.executeQuery(sql1); 
					if(null != row.getCell(4)) {
						if(rs1.next()) {
                   
						}
					}*/
				try {
					ps.executeUpdate();
				//	System.out.println(myset);
				}	
		catch (Exception e) {
				e.printStackTrace();
			}
				}
	}
		}
		catch (Exception e) {
			e.printStackTrace();
		}

}
			
}	
					//ps.setString(7,row.getCell(12).toString());
					//ps.setString(8,row.getCell(12).toString());
			//	System.out.println("Values Inserted Successfully");
					/*if (null != sheet.getRow(i).getCell(0) && !sheet.getRow(i).getCell(0).toString().equals("")) {
						myset.add(sheet.getRow(i).getCell(0).toString());
						System.out.println(sheet.getRow(i).getCell(0));
						ps.executeUpdate();
					}
					if (null != sheet.getRow(i).getCell(2) && !sheet.getRow(i).getCell(2).toString().equals("")) {
						myset.add(sheet.getRow(i).getCell(2).toString());
						System.out.println(sheet.getRow(i).getCell(2));
						ps.executeUpdate();
					}
					if (null != sheet.getRow(i).getCell(3) && !sheet.getRow(i).getCell(3).toString().equals("")) {
						myset.add(sheet.getRow(i).getCell(3).toString());
						System.out.println(sheet.getRow(i).getCell(3));
						ps.executeUpdate();
					}
					if (null != sheet.getRow(i).getCell(5) && !sheet.getRow(i).getCell(5).toString().equals("")) {
						myset.add(sheet.getRow(i).getCell(5).toString());
						System.out.println(sheet.getRow(i).getCell(5));
						ps.executeUpdate();
					}
					if (null != sheet.getRow(i).getCell(6) && !sheet.getRow(i).getCell(6).toString().equals("")) {
						myset.add(sheet.getRow(i).getCell(6).toString());
						System.out.println(sheet.getRow(i).getCell(6));
						ps.executeUpdate();
					}
					if (null != sheet.getRow(i).getCell(10) && !sheet.getRow(i).getCell(10).toString().equals("")) {
						myset.add(sheet.getRow(i).getCell(10).toString());
						System.out.println(sheet.getRow(i).getCell(10));
						ps.executeUpdate();
					}
					if (null != sheet.getRow(i).getCell(12) && !sheet.getRow(i).getCell(12).toString().equals("")) {
						myset.add(sheet.getRow(i).getCell(12).toString());
						System.out.println(sheet.getRow(i).getCell(12));
						ps.executeUpdate();
					}
					//ps.executeUpdate();
				}
			}*/
		//	System.out.println(myset);
		//	System.out.println(ps.executeUpdate());
			// System.out.println(sheet.getRow(i).getCell(5));
			// Iterator<String> itr = myset.iterator();
			// int i=1;
			/*
			 * int i=1; while(itr.hasNext()) { ps.getString(1,itr.next());
			 * ps.getString(2,itr.next()); ps.getString(3,itr.next());
			 * ps.getString(5,itr.next()); ps.getString(6,itr.next());
			 * ps.getString(10,itr.next()); ps.getString(12,itr.next()); i++;
			 * 
			 * ps.executeUpdate(); }
			 */
		

