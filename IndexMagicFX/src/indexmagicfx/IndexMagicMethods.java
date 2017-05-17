/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package indexmagicfx;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.SQLWarning;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 *
 * @author Glenn Dimaliwat
 */
public class IndexMagicMethods {
    
    //private String excelFile = "C:\\Users\\INDRA\\Documents\\BI_Phase2_Index.xls";
    private final String sapTableFile = "C:\\Users\\INDRA\\Documents\\BI_Phase2_SAP_TableData.xls";
    private static final String hostname = "172.18.1.74";
    private static final String port = "1433";
    private static final String databaseName = "RM_TARGET";
    
    private List<List<String>> sqlData = new ArrayList<>();
    private List<List<String>> indexData = new ArrayList<>();
    private List<List<String>> sapTableData = new ArrayList<>();
    
    protected void checkFields(String strDatabaseName, String strExcelFile) {
        try {
            FileInputStream fin = new FileInputStream(strExcelFile);
            POIFSFileSystem fs = new POIFSFileSystem(fin);
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            HSSFSheet sheet = wb.getSheetAt(0);
            HSSFRow row;
	    //HSSFCell cell;
            
            int rows; // No of rows
	    rows = sheet.getPhysicalNumberOfRows();
	
	    int cols = 0; // No of columns
	    int tmp = 0;
            
            // This trick ensures that we get the data properly even if it doesn't start from first few rows
	    for(int i = 0; i < 10 || i < rows; i++) {
	        row = sheet.getRow(i);
	        if(row != null) {
	            tmp = sheet.getRow(i).getPhysicalNumberOfCells();
	            if(tmp > cols) cols = tmp;
	        }
	    }
            
            // Connect to MS SQL
            String url = "jdbc:jtds:sqlserver://"+hostname+":"+port+"/"+strDatabaseName;
            String driver = "net.sourceforge.jtds.jdbc.Driver";
            String user = "Administrator";
            String pass = "Adm1n1strat012";
            
            Class.forName(driver);
            Connection conn;
            conn = DriverManager.getConnection(url, user, pass);
            for( SQLWarning warn = conn.getWarnings(); warn != null; warn = warn.getNextWarning() ) {
                    System.out.println( "SQL Warning:" ) ;
                    System.out.println( "State  : " + warn.getSQLState()  ) ;
                    System.out.println( "Message: " + warn.getMessage()   ) ;
                    System.out.println( "Error  : " + warn.getErrorCode() ) ;
            }
            String query = "";
            
            // Loop through Excel Rows
            for(int r = 0; r < rows; r++) {
	        row = sheet.getRow(r);
                if(row != null) {
                    if(row.getCell(0)!=null) {
                        if(row.getCell(0).toString().equalsIgnoreCase("JOB NAME")) {
                            // If Header - Do Nothing
                        }
                        if(!row.getCell(2).toString().equalsIgnoreCase(strDatabaseName) ) {
                            // If Not RM_TARGET - Do Nothing
                        }
                        else {
                            
                            //System.out.println(row.getCell(3));
                            String strTableName = "";
                            String strIndex1    = "";
                            String strIndex2    = "";
                            String strIndex3    = "";
                            String strIndex4    = "";
                            String strIndex5    = "";
                            String strIndex6    = "";
                            String strIndex7    = "";
                            String strIndex8    = "";
                            String strIndex9    = "";
                            String strIndex10   = "";
                            String strIndex11   = "";
                            
                            if(row.getCell(3)!=null) {
                                strTableName = row.getCell(3).toString();
                            }
                            if(row.getCell(4)!=null) {
                                strIndex1 = row.getCell(4).toString();
                            }
                            if(row.getCell(5)!=null) {
                                strIndex2 = row.getCell(5).toString();
                            }
                            if(row.getCell(6)!=null) {
                                strIndex3 = row.getCell(6).toString();
                            }
                            if(row.getCell(7)!=null) {
                                strIndex4 = row.getCell(7).toString();
                            }
                            if(row.getCell(8)!=null) {
                                strIndex5 = row.getCell(8).toString();
                            }
                            if(row.getCell(9)!=null) {
                                strIndex6 = row.getCell(9).toString();
                            }
                            if(row.getCell(10)!=null) {
                                strIndex7 = row.getCell(10).toString();
                            }
                            if(row.getCell(11)!=null) {
                                strIndex8 = row.getCell(11).toString();
                            }
                            if(row.getCell(12)!=null) {
                                strIndex9 = row.getCell(12).toString();
                            }
                            if(row.getCell(13)!=null) {
                                strIndex10 = row.getCell(13).toString();
                            }
                            if(row.getCell(14)!=null) {
                                strIndex11 = row.getCell(14).toString();
                            }
                            
                            //System.out.println("TABLE: "+strTableName);
                            String strInputIndex = strIndex1+" "+strIndex2+" "+strIndex3+" "+strIndex4+" "+strIndex5+" "+strIndex6+" "+strIndex7+" "+strIndex8+" "+strIndex9+" "+strIndex10+" "+strIndex11;
                            System.out.println("----------- BEGIN -----------");
                            System.out.println(strTableName +" INPUT INDEX: "+strInputIndex.trim());
                            
                            // Perform SQL Query
                            query =  "SELECT ORDINAL_POSITION, COLUMN_NAME";
                            query += " FROM INFORMATION_SCHEMA.COLUMNS ";
                            query += " WHERE TABLE_NAME = '"+strTableName+"'";
                            query += " ORDER BY ORDINAL_POSITION";
                            
                            sqlData.clear();
                            try (Statement stmt = conn.createStatement(); ResultSet rs = stmt.executeQuery( query )) {
                                // Loop through the result set
                                while( rs.next() ) {
                                    List<String> sqlDataRow = new ArrayList<>();
                                    sqlDataRow.add(rs.getString(1));
                                    sqlDataRow.add(rs.getString(2));
                                    sqlData.add(sqlDataRow);
                                }
                                // Required for MS SQL Server
                                stmt.cancel();
                                stmt.close();
                                rs.close();
                            }
                            
                            int fieldCount = 0;
                            if(!strIndex1.isEmpty()) {
                                fieldCount++;
                            }
                            if(!strIndex2.isEmpty()) {
                                fieldCount++;
                            }
                            if(!strIndex3.isEmpty()) {
                                fieldCount++;
                            }
                            if(!strIndex4.isEmpty()) {
                                fieldCount++;
                            }
                            if(!strIndex5.isEmpty()) {
                                fieldCount++;
                            }
                            if(!strIndex6.isEmpty()) {
                                fieldCount++;
                            }
                            if(!strIndex7.isEmpty()) {
                                fieldCount++;
                            }
                            if(!strIndex8.isEmpty()) {
                                fieldCount++;
                            }
                            if(!strIndex9.isEmpty()) {
                                fieldCount++;
                            }
                            if(!strIndex10.isEmpty()) {
                                fieldCount++;
                            }
                            if(!strIndex11.isEmpty()) {
                                fieldCount++;
                            }
                            
                            int matchCount = 0;
                            for(List<String> strData : sqlData) {
                                if(strData.get(1).equalsIgnoreCase(strIndex1)) {
                                    matchCount++;
                                }
                                if(strData.get(1).equalsIgnoreCase(strIndex2)) {
                                    matchCount++;
                                }
                                if(strData.get(1).equalsIgnoreCase(strIndex3)) {
                                    matchCount++;
                                }
                                if(strData.get(1).equalsIgnoreCase(strIndex4)) {
                                    matchCount++;
                                }
                                if(strData.get(1).equalsIgnoreCase(strIndex5)) {
                                    matchCount++;
                                }
                                if(strData.get(1).equalsIgnoreCase(strIndex6)) {
                                    matchCount++;
                                }
                                if(strData.get(1).equalsIgnoreCase(strIndex7)) {
                                    matchCount++;
                                }
                                if(strData.get(1).equalsIgnoreCase(strIndex8)) {
                                    matchCount++;
                                }
                                if(strData.get(1).equalsIgnoreCase(strIndex9)) {
                                    matchCount++;
                                }
                                if(strData.get(1).equalsIgnoreCase(strIndex10)) {
                                    matchCount++;
                                }
                                if(strData.get(1).equalsIgnoreCase(strIndex11)) {
                                    matchCount++;
                                }
                            }
                            if(matchCount==fieldCount) {
                                System.out.println(strTableName+" FIELDS ACCURATE");
                            }
                            else {
                                System.out.println(strTableName+" FIELDS NOT ACCURATE");
                            }
                            System.out.println("------------ END ------------");
                        }
                    }
                }
            }
         
            sqlData.clear();
            System.gc();
        }catch(IOException | ClassNotFoundException | SQLException e) {
            e.printStackTrace();
        }
    }
    
    protected void checkIndex(String strExcelFile) {
        try {
            
            // Create 2D List Comparator
            final Comparator<List<String>> comparator = new Comparator<List<String>>() {
                @Override
                public int compare(List<String> pList1, List<String> pList2) {
                    return pList1.get(0).compareTo(pList2.get(0));
                }
            };
            
            
            // SAP Table File
            FileInputStream finSAP = new FileInputStream(sapTableFile);
            POIFSFileSystem fsSAP = new POIFSFileSystem(finSAP);
            HSSFWorkbook wbSAP = new HSSFWorkbook(fsSAP);
            HSSFSheet sheetSAP = wbSAP.getSheetAt(0);
            HSSFRow rowSAP;
	    //HSSFCell cellSAP;
            
            int rowsSAP; // No of rows
	    rowsSAP = sheetSAP.getPhysicalNumberOfRows();
	
	    int colsSAP = 0; // No of columns
	    int tmpSAP = 0;
            
            // This trick ensures that we get the data properly even if it doesn't start from first few rows
	    for(int i = 0; i < 10 || i < rowsSAP; i++) {
	        rowSAP = sheetSAP.getRow(i);
	        if(rowSAP != null) {
	            tmpSAP = sheetSAP.getRow(i).getPhysicalNumberOfCells();
	            if(tmpSAP > colsSAP) colsSAP = tmpSAP;
	        }
	    }
            
            // Clear SAP Table Data
            sapTableData.clear();
            
            
            // Loop through Excel Rows
            for(int r = 0; r < rowsSAP; r++) {
	        rowSAP = sheetSAP.getRow(r);
                if(rowSAP != null) {
                    if(rowSAP.getCell(0)!=null) {
                        if(rowSAP.getCell(0).toString().equalsIgnoreCase("Table Name")) {
                        }
                        else {
                            List<String> sapDataRow = new ArrayList<>();
                            sapDataRow.add(rowSAP.getCell(0).toString());
                            sapDataRow.add(rowSAP.getCell(1).toString());
                            sapDataRow.add(rowSAP.getCell(2).toString());
                            sapTableData.add(sapDataRow);
                        }
                    }
                }
            }
            Collections.sort(sapTableData, comparator);
            
            // Excel File
            FileInputStream fin = new FileInputStream(strExcelFile);
            POIFSFileSystem fs = new POIFSFileSystem(fin);
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            HSSFSheet sheet = wb.getSheetAt(0);
            HSSFRow row;
	    //HSSFCell cell;
            
            int rows; // No of rows
	    rows = sheet.getPhysicalNumberOfRows();
	
	    int cols = 0; // No of columns
	    int tmp = 0;
            
            // This trick ensures that we get the data properly even if it doesn't start from first few rows
	    for(int i = 0; i < 10 || i < rows; i++) {
	        row = sheet.getRow(i);
	        if(row != null) {
	            tmp = sheet.getRow(i).getPhysicalNumberOfCells();
	            if(tmp > cols) cols = tmp;
	        }
	    }
            
            // Connect to MS SQL
            String url = "jdbc:jtds:sqlserver://"+hostname+":"+port+"/"+databaseName;
            String driver = "net.sourceforge.jtds.jdbc.Driver";
            String user = "Administrator";
            String pass = "Adm1n1strat012";
            
            Class.forName(driver);
            Connection conn;
            conn = DriverManager.getConnection(url, user, pass);
            for( SQLWarning warn = conn.getWarnings(); warn != null; warn = warn.getNextWarning() ) {
                    System.out.println( "SQL Warning:" ) ;
                    System.out.println( "State  : " + warn.getSQLState()  ) ;
                    System.out.println( "Message: " + warn.getMessage()   ) ;
                    System.out.println( "Error  : " + warn.getErrorCode() ) ;
            }
            String query = "";
            
            // Clear Index Data
            indexData.clear();
            
            // Loop through Excel Rows
            for(int r = 0; r < rows; r++) {
	        row = sheet.getRow(r);
                if(row != null) {
                    if(row.getCell(0)!=null) {
                        if(row.getCell(0).toString().equalsIgnoreCase("JOB NAME")) {
                            // If Header - Do Nothing
                        }
                        if(!row.getCell(2).toString().equalsIgnoreCase("RM_TARGET") &&
                                !row.getCell(2).toString().equalsIgnoreCase("DWH_DEF_TARGET_PHASE_2") &&
                                !row.getCell(2).toString().equalsIgnoreCase("SAP")) {
                            // If Not RM_TARGET - Do Nothing
                        }
                        else {
                            
                            //System.out.println(row.getCell(3));
                            String strTableName = "";
                            String strIndex1    = "";
                            String strIndex2    = "";
                            String strIndex3    = "";
                            String strIndex4    = "";
                            String strIndex5    = "";
                            String strIndex6    = "";
                            String strIndex7    = "";
                            String strIndex8    = "";
                            String strIndex9    = "";
                            String strIndex10   = "";
                            String strIndex11   = "";
                            
                            if(row.getCell(3)!=null) {
                                strTableName = row.getCell(3).toString();
                            }
                            if(row.getCell(4)!=null) {
                                strIndex1 = row.getCell(4).toString();
                            }
                            if(row.getCell(5)!=null) {
                                strIndex2 = row.getCell(5).toString();
                            }
                            if(row.getCell(6)!=null) {
                                strIndex3 = row.getCell(6).toString();
                            }
                            if(row.getCell(7)!=null) {
                                strIndex4 = row.getCell(7).toString();
                            }
                            if(row.getCell(8)!=null) {
                                strIndex5 = row.getCell(8).toString();
                            }
                            if(row.getCell(9)!=null) {
                                strIndex6 = row.getCell(9).toString();
                            }
                            if(row.getCell(10)!=null) {
                                strIndex7 = row.getCell(10).toString();
                            }
                            if(row.getCell(11)!=null) {
                                strIndex8 = row.getCell(11).toString();
                            }
                            if(row.getCell(12)!=null) {
                                strIndex9 = row.getCell(12).toString();
                            }
                            if(row.getCell(13)!=null) {
                                strIndex10 = row.getCell(13).toString();
                            }
                            if(row.getCell(14)!=null) {
                                strIndex11 = row.getCell(14).toString();
                            }
                            
                            //System.out.println("TABLE: "+strTableName);
                            String strRepository = row.getCell(2).toString();
                            String strInputIndex = strIndex1+" "+strIndex2+" "+strIndex3+" "+strIndex4+" "+strIndex5+" "+strIndex6+" "+strIndex7+" "+strIndex8+" "+strIndex9+" "+strIndex10+" "+strIndex11;
                            System.out.println("----------- BEGIN -----------");
                            System.out.println(strTableName +" INPUT INDEX: "+strInputIndex.trim());
                            
                            boolean matchFound = false;
                            boolean matchFoundForOptimization = false;
                                
                            if(!row.getCell(2).toString().equalsIgnoreCase("SAP")) { // Not SAP - For RM_TARGET and DWH
                                // Perform SQL Query
                                query =  "SELECT";
                                query += " ORDINAL_POSITION = ic.index_column_id,";
                                query += " COLUMN_NAME = col.name";
                                query += " FROM sys.indexes ind ";
                                query += " INNER JOIN sys.index_columns ic ON  ind.object_id = ic.object_id and ind.index_id = ic.index_id";
                                query += " INNER JOIN sys.columns col ON ic.object_id = col.object_id and ic.column_id = col.column_id";
                                query += " INNER JOIN sys.tables t ON ind.object_id = t.object_id";
                                query += " WHERE";
                                query += " ind.is_unique_constraint = 0";
                                query += " AND t.is_ms_shipped = 0";
                                query += " AND t.name = '"+strTableName+"'";
                                query += " ORDER BY ind.name, ind.index_id, ic.index_column_id";

                                sqlData.clear();
                                try (Statement stmt = conn.createStatement(); ResultSet rs = stmt.executeQuery( query )) {
                                    // Loop through the result set
                                    while( rs.next() ) {
                                        List<String> sqlDataRow = new ArrayList<>();
                                        sqlDataRow.add(rs.getString(1));
                                        sqlDataRow.add(rs.getString(2));
                                        sqlData.add(sqlDataRow);
                                    }
                                    // Required for MS SQL Server
                                    stmt.cancel();
                                    stmt.close();
                                    rs.close();
                                }
                                
                                int ctr = 0;
                                int indexID = 1;
                                String strIndex = "";
                                matchFound = false;
                                matchFoundForOptimization = false;
                                for(List<String> strData : sqlData) {
                                    if(strData.get(0).equalsIgnoreCase("1")) {
                                        // If not first row of sql data
                                        if(ctr!=0) {
                                            // Trim index string
                                            strIndex = strIndex.trim();
                                            System.out.println(strTableName+" INDEX "+indexID+": "+strIndex);
                                            //Assign New ID
                                            indexID++;
                                            //If retrieved index is not empty
                                            if(!strIndex.isEmpty()) {
                                                strInputIndex = strInputIndex.trim();
                                                //If retrieved index is same with the input index or it retrieved starts with input index
                                                if(strIndex.equalsIgnoreCase(strInputIndex) || strIndex.startsWith(strInputIndex)) {
                                                    matchFound = true;
                                                    System.out.println("---- MATCH FOUND ----");
                                                    break;
                                                }
                                                else {
                                                    // Check if condition is viable for optimization by removing CLIENT ID or MANDT
                                                    
                                                    if( (strIndex.contains("CLIENT_ID") && strInputIndex.contains("CLIENT_ID")) ||
                                                        (strIndex.contains("MANDT") && strInputIndex.contains("MANDT"))) {
                                                            
                                                        String strIndexNoClient = strIndex.replaceAll("(\\sCLIENT_ID$)|(\\sMANDT$)","");
                                                        strIndexNoClient = strIndexNoClient.trim();

                                                        if(strIndexNoClient.equalsIgnoreCase(strIndex)) {
                                                            strIndexNoClient = strIndex.replaceAll("(^CLIENT_ID\\s)|(^MANDT\\s)","");
                                                            strIndexNoClient = strIndexNoClient.trim();
                                                        }

                                                        String strInputIndexNoClient = strInputIndex.replaceAll("(\\sCLIENT_ID$)|(\\sMANDT$)","");
                                                        strInputIndexNoClient = strInputIndexNoClient.trim();

                                                        if(strInputIndexNoClient.equalsIgnoreCase(strInputIndex)) {
                                                            strInputIndexNoClient = strInputIndexNoClient.replaceAll("(^CLIENT_ID\\s)|(^MANDT\\s)","");
                                                            strInputIndexNoClient = strInputIndexNoClient.trim();
                                                        }

                                                        if(strIndexNoClient.equalsIgnoreCase(strInputIndexNoClient)) {
                                                            //matchFound = true;
                                                            matchFoundForOptimization = true;
                                                            System.out.println("---- MATCH FOUND FOR OPTIMIZATION ----");
                                                            break;
                                                        }
                                                    }
                                                }
                                            }

                                        }
                                        // Reset strIndex
                                        strIndex = strData.get(1);
                                        strIndex += " ";
                                        matchFound = false;
                                        matchFoundForOptimization = false;
                                    }
                                    else {
                                        strIndex += strData.get(1);
                                        strIndex += " ";
                                    }
                                    ctr++;
                                }
                                if(matchFound==false && matchFoundForOptimization==false && !strIndex.isEmpty()) {
                                    // Trim index string
                                    strIndex = strIndex.trim();
                                    System.out.println(strTableName+" INDEX "+indexID+": "+strIndex);
                                    //If retrieved index is not empty
                                    if(!strIndex.isEmpty()) {
                                        strInputIndex = strInputIndex.trim();
                                        //If retrieved index is same with the input index or it retrieved starts with input index
                                        if(strIndex.equalsIgnoreCase(strInputIndex) || strIndex.startsWith(strInputIndex)) {
                                            matchFound = true;
                                            System.out.println("---- MATCH FOUND ----");
                                        }
                                        else {
                                            // Check if condition is viable for optimization by removing CLIENT ID or MANDT

                                            if( (strIndex.contains("CLIENT_ID") && strInputIndex.contains("CLIENT_ID")) ||
                                                (strIndex.contains("MANDT") && strInputIndex.contains("MANDT"))) {

                                                String strIndexNoClient = strIndex.replaceAll("(\\sCLIENT_ID$)|(\\sMANDT$)","");
                                                strIndexNoClient = strIndexNoClient.trim();

                                                if(strIndexNoClient.equalsIgnoreCase(strIndex)) {
                                                    strIndexNoClient = strIndex.replaceAll("(^CLIENT_ID\\s)|(^MANDT\\s)","");
                                                    strIndexNoClient = strIndexNoClient.trim();
                                                }

                                                String strInputIndexNoClient = strInputIndex.replaceAll("(\\sCLIENT_ID$)|(\\sMANDT$)","");
                                                strInputIndexNoClient = strInputIndexNoClient.trim();

                                                if(strInputIndexNoClient.equalsIgnoreCase(strInputIndex)) {
                                                    strInputIndexNoClient = strInputIndexNoClient.replaceAll("(^CLIENT_ID\\s)|(^MANDT\\s)","");
                                                    strInputIndexNoClient = strInputIndexNoClient.trim();
                                                }

                                                if(strIndexNoClient.equalsIgnoreCase(strInputIndexNoClient)) {
                                                    //matchFound = true;
                                                    matchFoundForOptimization = true;
                                                    System.out.println("---- MATCH FOUND FOR OPTIMIZATION ----");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            else { // For SAP
                                
                                int ctr = 0;
                                int indexID = 1;
                                String strIndexSAP = "";
                                matchFound = false;
                                matchFoundForOptimization = false;
                                for(List<String> sapData : sapTableData) {
                                    if(strTableName.equalsIgnoreCase("Table Name")) {
                                    }
                                    else if(strTableName.equalsIgnoreCase(sapData.get(0).toString())) {
                                        if(sapData.get(1).equalsIgnoreCase("1")) {
                                            // If not first row of sql data
                                            if(ctr!=0) {
                                                // Trim index string
                                                strIndexSAP = strIndexSAP.trim();
                                                System.out.println(strTableName+" INDEX "+indexID+": "+strIndexSAP.trim());
                                                //Assign New ID
                                                indexID++;
                                                //If retrieved index is not empty
                                                if(!strIndexSAP.isEmpty()) {
                                                    strInputIndex = strInputIndex.trim();
                                                    //If retrieved index is same with the input index or it retrieved starts with input index
                                                    if(strIndexSAP.equalsIgnoreCase(strInputIndex) || strIndexSAP.startsWith(strInputIndex)) {
                                                        matchFound = true;
                                                        System.out.println("---- MATCH FOUND ----");
                                                        break;
                                                    }
                                                    else {
                                                        // Check if condition is viable for optimization by removing CLIENT ID or MANDT

                                                        if( (strIndexSAP.contains("CLIENT_ID") && strInputIndex.contains("CLIENT_ID")) ||
                                                            (strIndexSAP.contains("MANDT") && strInputIndex.contains("MANDT"))) {

                                                            String strIndexNoClient = strIndexSAP.replaceAll("(\\sCLIENT_ID$)|(\\sMANDT$)","");
                                                            strIndexNoClient = strIndexNoClient.trim();

                                                            if(strIndexNoClient.equalsIgnoreCase(strIndexSAP)) {
                                                                strIndexNoClient = strIndexSAP.replaceAll("(^CLIENT_ID\\s)|(^MANDT\\s)","");
                                                                strIndexNoClient = strIndexNoClient.trim();
                                                            }

                                                            String strInputIndexNoClient = strInputIndex.replaceAll("(\\sCLIENT_ID$)|(\\sMANDT$)","");
                                                            strInputIndexNoClient = strInputIndexNoClient.trim();

                                                            if(strInputIndexNoClient.equalsIgnoreCase(strInputIndex)) {
                                                                strInputIndexNoClient = strInputIndexNoClient.replaceAll("(^CLIENT_ID\\s)|(^MANDT\\s)","");
                                                                strInputIndexNoClient = strInputIndexNoClient.trim();
                                                            }

                                                            if(strIndexNoClient.equalsIgnoreCase(strInputIndexNoClient)) {
                                                                //matchFound = true;
                                                                matchFoundForOptimization = true;
                                                                System.out.println("---- MATCH FOUND FOR OPTIMIZATION ----");
                                                                break;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            // Reset strIndex
                                            strIndexSAP = sapData.get(2);
                                            strIndexSAP += " ";
                                            matchFound = false;
                                            matchFoundForOptimization = false;
                                        }
                                        else {
                                            strIndexSAP += sapData.get(2);
                                            strIndexSAP += " ";
                                        }
                                        ctr++;
                                    }
                                }
                                if(matchFound==false && matchFoundForOptimization==false && !strIndexSAP.isEmpty()) {
                                    // Trim index string
                                    strIndexSAP = strIndexSAP.trim();
                                    System.out.println(strTableName+" INDEX "+indexID+": "+strIndexSAP.trim());
                                    //If retrieved index is not empty
                                    if(!strIndexSAP.isEmpty()) {
                                        strInputIndex = strInputIndex.trim();
                                        //If retrieved index is same with the input index or it retrieved starts with input index
                                        if(strIndexSAP.equalsIgnoreCase(strInputIndex) || strIndexSAP.startsWith(strInputIndex)) {
                                            matchFound = true;
                                            System.out.println("---- MATCH FOUND ----");
                                        }
                                        else {
                                            // Check if condition is viable for optimization by removing CLIENT ID or MANDT

                                            if( (strIndexSAP.contains("CLIENT_ID") && strInputIndex.contains("CLIENT_ID")) ||
                                                (strIndexSAP.contains("MANDT") && strInputIndex.contains("MANDT"))) {

                                                String strIndexNoClient = strIndexSAP.replaceAll("(\\sCLIENT_ID$)|(\\sMANDT$)","");
                                                strIndexNoClient = strIndexNoClient.trim();

                                                if(strIndexNoClient.equalsIgnoreCase(strIndexSAP)) {
                                                    strIndexNoClient = strIndexSAP.replaceAll("(^CLIENT_ID\\s)|(^MANDT\\s)","");
                                                    strIndexNoClient = strIndexNoClient.trim();
                                                }

                                                String strInputIndexNoClient = strInputIndex.replaceAll("(\\sCLIENT_ID$)|(\\sMANDT$)","");
                                                strInputIndexNoClient = strInputIndexNoClient.trim();

                                                if(strInputIndexNoClient.equalsIgnoreCase(strInputIndex)) {
                                                    strInputIndexNoClient = strInputIndexNoClient.replaceAll("(^CLIENT_ID\\s)|(^MANDT\\s)","");
                                                    strInputIndexNoClient = strInputIndexNoClient.trim();
                                                }

                                                if(strIndexNoClient.equalsIgnoreCase(strInputIndexNoClient)) {
                                                    //matchFound = true;
                                                    matchFoundForOptimization = true;
                                                    System.out.println("---- MATCH FOUND FOR OPTIMIZATION ----");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            
                            if(matchFound) {
                                System.out.println(strTableName+" ALREADY INDEXED: "+strInputIndex);
                            }
                            else if(matchFoundForOptimization) {
                                System.out.println(strTableName+" FOR OPTIMIZATION: "+strInputIndex+" | JOB NAME: "+row.getCell(0).toString());
                            }
                            else {
                                List<String> indexDataRow = new ArrayList<>();
                                indexDataRow.add(strRepository);
                                indexDataRow.add(strTableName);
                                indexDataRow.add(strInputIndex);
                                indexData.add(indexDataRow);
                            }
                            System.out.println("------------ END ------------");
                        }
                    }
                }
            }
            
            // Sort List before deleting duplicates
            Collections.sort(indexData, comparator);
         
            // Delete duplicates using a Linked Hash Set to retain sort order
            Set<List<String>> hs = new LinkedHashSet<>();
            hs.addAll(indexData);
            indexData.clear();
            indexData.addAll(hs);
            
            // Output
            for (List<String> list : indexData) System.out.println(list);
            
            sapTableData.clear();
            //indexData.clear(); // Do not clear for generateIndexScript()
            sqlData.clear();
            System.gc();
        }catch(IOException | ClassNotFoundException | SQLException e) {
            e.printStackTrace();
        }
    }
    
    protected void generateIndexScript() {
        int indexCount = 0;
        
        for(List<String> indexData : indexData) {
            //System.out.println(indexData.get(0)+" "+indexData.get(1)+" "+indexData.get(2));
            
            String qry = "";
            
            if(indexData.get(0).equalsIgnoreCase("RM_TARGET")) {
                indexCount++;
                
                qry = "CREATE INDEX "+"BIDX"+indexCount;
                qry += " ON dbo."+indexData.get(1);
                qry += " ("+indexData.get(2).replaceAll(" ", ", ")+")";
                System.out.println(qry);
            }
        }
    }
}
