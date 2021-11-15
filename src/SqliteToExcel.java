import org.apache.poi.xssf.usermodel.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.PrintStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.*;
import java.util.ArrayList;
import org.sqlite.JDBC;
import org.sqlite.SQLiteConfig;

public class SqliteToExcel {
    
    String WB_NAME = "export.xlsx";
    
    public static void main(String[] args) throws SQLException {
        if(args.length>0 && args[0].endsWith(".db")) {
        	SqliteToExcel a = new SqliteToExcel();
	        if(args.length > 1 && args[1].endsWith(".xlsx"))
	            a.WB_NAME = args[1];
	        a.exportToExcel(args[0]);
        }else {
        	System.out.println("expected database file path (.db extension).");
        }
    }

    private Connection connect(String dbPath) throws SQLException {
    	if(Files.notExists(Paths.get(dbPath))) {
    		System.out.println("no database at \"" + dbPath + "\". check for case-sensitivity errors.");
    	}else {
	    	DriverManager.registerDriver(new JDBC());
	        String url = "jdbc:sqlite:" + dbPath;
	        SQLiteConfig config = new SQLiteConfig();
	        config.enforceForeignKeys(true);
	        try{
                return DriverManager.getConnection(url, config.toProperties());
	        }catch(SQLException e){
	            System.out.println(e.getMessage());
	        }
    	}
        return null;
    }

    public void exportToExcel(String dbPath) throws SQLException {
        OutputStream os;
        String[] tableType = {"TABLE"};
//        System.setErr(null);
        try {
			System.setErr(new PrintStream(new FileOutputStream("err.txt")));
		}
        catch (FileNotFoundException e1) {
			e1.printStackTrace();
		}
        
        try (Connection con = connect(dbPath) ;XSSFWorkbook wb = new XSSFWorkbook()){
            if(con == null)
                return;
            
            // get table names from connection
            ArrayList<String> tableNames = new ArrayList<>();
            DatabaseMetaData conMeta =  con.getMetaData();
            ResultSet tables = conMeta.getTables(null, null, null, tableType);
            while(tables.next()) {
            	tableNames.add(tables.getString("TABLE_NAME"));
            }
            
            
            for (String sheetName : tableNames) {
                XSSFSheet sheet = wb.createSheet(sheetName);
                try (Statement stat = con.createStatement()){
                    short rowNum = 0;
                    XSSFRow rowHead = sheet.createRow(rowNum++);
                    String sql = "select * from " + sheetName;
                    ResultSet res = stat.executeQuery(sql);
                    
                    //get column names
                    ResultSetMetaData meta = res.getMetaData();
                    int columnCount = meta.getColumnCount();
                    if(res.next()) {
                        int currColumn = 0;
                        while (currColumn < columnCount) {
                            XSSFCell cell = rowHead.createCell(currColumn);
                            cell.setCellValue(meta.getColumnName(++currColumn)); // write column-names at the top of the sheet
                        }
                    }else {
                    	XSSFCell cell = rowHead.createCell(0);
                        cell.setCellValue("no columns in this table.");
                    }
                    
                    //insert all data line by line
                    while(res.next()){
                        rowHead = sheet.createRow(rowNum++);
                        int currColumn = 0;
                        while (currColumn < columnCount) {
                            XSSFCell cell = rowHead.createCell(currColumn);
                            cell.setCellValue(res.getString(++currColumn));// column index begin at 1
                        }
                    }

                    Path f = new File(WB_NAME).toPath();
                    if(!Files.exists(f)){
                        os = new FileOutputStream(Files.createFile(Paths.get(WB_NAME)).toString());
                    }else{
                        os = new FileOutputStream(f.toString());
                    }              
                    wb.write(os);
                    os.close();
                }catch (FileNotFoundException e2) { // probably export.xlsx is open
					System.out.println("check for correct path. if so, \"" + WB_NAME + "\" might be open.");
					break;
                } catch (SQLException e) {
                    e.printStackTrace();
				}
            }
        } catch (IOException e) {
            System.out.println("unable to connect database. check for correct path or open DB file.");
        }

    }
}