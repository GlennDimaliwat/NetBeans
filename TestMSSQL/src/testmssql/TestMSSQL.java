/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package testmssql;

import java.sql.Connection;
import java.sql.DatabaseMetaData;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

/**
 *
 * @author INDRA
 */
public class TestMSSQL {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws SQLException {
        Connection conn = null;
        String url = "jdbc:jtds:sqlserver://"; // Removed for GitHub
        String driver = "net.sourceforge.jtds.jdbc.Driver";
        String userName = ""; // Removed for GitHub
        String password = ""; // Removed for GitHub
        try {
            Class.forName(driver);
            conn = DriverManager.getConnection(url, userName, password);
            /*System.out.println("Connected to the database!!! Getting table list...");
            DatabaseMetaData dbm = conn.getMetaData();
            rs = dbm.getTables(null, null, "%", new String[] { "TABLE" });
            while (rs.next()) { System.out.println(rs.getString("TABLE_NAME")); }*/
            
            try (Statement stmt = conn.createStatement();
                ResultSet rs = stmt.executeQuery("SELECT mrID, mrTitle, mrSUBMITDATE, mrUpdateDate, SLA__bDue__bDate, Type__bof__bTicket, Category, Sub__bCategory, mrAssignees, mrStatus, Pending__bReason, Resolution__bState " +
                                                 "FROM dbo.MASTER1 WHERE mrID IN ('770','854','913','927','929','933','991') ORDER BY mrID ASC")) {
                
                while( rs.next() ) {
                    System.out.println(rs.getString(1) + " " +
                                        rs.getString(2) + " " +
                                        rs.getString(3) + " " +
                                        rs.getString(4) + " " +
                                        rs.getString(5) + " " +
                                        rs.getString(6) + " " +
                                        rs.getString(7) + " " +
                                        rs.getString(8) + " " +
                                        rs.getString(9) + " " +
                                        rs.getString(10) + " " +
                                        rs.getString(11) + " " +
                                        rs.getString(12) + " ");
                }
                // Required for MS SQL Server
                stmt.cancel();
                stmt.close();
                rs.close();
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            conn.close();
        }
    }
}
