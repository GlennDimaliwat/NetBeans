/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package sdsreport;

import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.Toolkit;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.SQLWarning;
import java.sql.Statement;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import javafx.embed.swing.JFXPanel;
import javax.swing.JOptionPane;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import org.apache.commons.lang3.StringEscapeUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author Glenn Dimaliwat
 */
public class MainWindow extends javax.swing.JFrame {
    
    private static final String hostname = ""; // Removed for GitHub
    private static final String port = ""; // Removed for GitHub       
    private static final String databaseName = "Footprints";
    private static final SimpleDateFormat sdfYear = new SimpleDateFormat("yyyy");
    private static final SimpleDateFormat sdfMonth = new SimpleDateFormat("MMMM");
    private static final SimpleDateFormat sdfMM = new SimpleDateFormat("MM");
    private static final SimpleDateFormat sdfMMddyyyy = new SimpleDateFormat("MM/dd/yyyy");
    private static final SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
    private static final SimpleDateFormat sdfTimestamp = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
    private static final SimpleDateFormat sdfDatestamp = new SimpleDateFormat("yyyy-MM-dd");
    private static final SimpleDateFormat sdfHHmmss = new SimpleDateFormat("HH:mm:ss");
    private static final SimpleDateFormat sdfHH = new SimpleDateFormat("HH");
    private static final SimpleDateFormat sdfmm = new SimpleDateFormat("mm");
    private static final SimpleDateFormat sdfss = new SimpleDateFormat("ss");
    private List<String> holidayList = new ArrayList<>();
    private List<List<String>> personList = new ArrayList<>();
    private List<List<String>> reportList = new ArrayList<>();
    private List<String> reportItem;
    private List<String> person;
    
    private List<List<String>> teamStatList = new ArrayList<>();
    private List<String> teamStatItem = new ArrayList<>();
    
    private String monthFilter;
    private String yearFilter;
    private String ticketTypeFilter = "All";
    private Date dateFilter = null;
    private String reportFilename = "SDS_Report_Service_Desk.xls"; //Default filename
    private String ticketCounterFilename = "TicketCounter.xls"; //Default filename
    private String teamFilter = "All";
    private String personFilter = "All";
    
    /**
     * Creates new form MainWindow
     */
    public MainWindow() {
        initComponents();
        initCustomComponents();
    }
    
    private void initCustomComponents() {
        JFXPanel fxPanel = new JFXPanel(); // Needed to initialize JavaFX Objects
        Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
        this.setLocation(dim.width/2-this.getSize().width/2, dim.height/2-this.getSize().height/2);
        jButton2.setEnabled(false);
        jButton3.setEnabled(false);
            
        jTextField2.setText(reportFilename);
        initHolidays();
        initPersons();
    }
    
    private void initHolidays() {
        holidayList.clear();
        try (BufferedReader br = new BufferedReader(new FileReader("Holidays.txt"))) {
            String line = br.readLine();
            while (line != null) {
                if(!line.equalsIgnoreCase("")) {
                    holidayList.add(line);
                    line = br.readLine();
                }
                else {
                    line = br.readLine();
                }
            }
        }
        catch(Exception e) {
            System.out.println(e);
        }
        /*for(int x=0;x<holidayList.size();x++) {
            System.out.println(holidayList.get(x));
        }*/
    }
    
    private void initPersons() {
        personList.clear();
        jComboBox4.removeAllItems();
        jComboBox4.addItem("All");
        jButton2.setEnabled(false);
        jButton3.setEnabled(false);
        
        String url = "jdbc:jtds:sqlserver://"+hostname+":"+port+"/"+databaseName;
        String driver = "net.sourceforge.jtds.jdbc.Driver";
        String user = ""; //Removed for GitHub
        String pass = ""; //Removed for GitHub
        
        String selectedWorkspace = "1";
        if(jComboBox1.getSelectedIndex()==0) {
            selectedWorkspace = "1";
        }
        else if(jComboBox1.getSelectedIndex()==1) {
            selectedWorkspace = "2";
        }
        else if(jComboBox1.getSelectedIndex()==2) {
            selectedWorkspace = "3";
        }
        
        String query;
        if(!teamFilter.equalsIgnoreCase("All")) {
            query = "select distinct a.user_id, a.real_name from users as a left join dbo.MASTER"+selectedWorkspace+"_ASSIGNMENT as b on a.user_id = b.Assignee where (user_type = '2' or user_type = '4') and real_name is not null and b.TeamAssignee = '"+teamFilter+"'";
        }
        else {
            query = "select user_id, real_name from users where (user_type = '2' or user_type = '4') and real_name is not null";
        }
        
        try {
            Class.forName(driver);
            Connection conn;
            conn = DriverManager.getConnection(url, user, pass);
            for( SQLWarning warn = conn.getWarnings(); warn != null; warn = warn.getNextWarning() ) {
                    System.out.println( "SQL Warning:" ) ;
                    System.out.println( "State  : " + warn.getSQLState()  ) ;
                    System.out.println( "Message: " + warn.getMessage()   ) ;
                    System.out.println( "Error  : " + warn.getErrorCode() ) ;
            }
            
            try (Statement stmt = conn.createStatement(); ResultSet rs = stmt.executeQuery( query )) {
                
                //int recordCount = 0;
                while( rs.next() ) {
                    person = new ArrayList<>();
                    person.add(rs.getString(1));
                    person.add(rs.getString(2));
                    personList.add(person);
                }
                
                // Required for MS SQL Server
                stmt.cancel();
                stmt.close();
                rs.close();
            }
            conn.close();
            
            for(List<String> personRow : personList) {
                jComboBox4.addItem(personRow.get(1));
            }
            /*for(int x=0;x<personList.size();x++) {
                //System.out.println(personList.get(x).get(0));
                //System.out.println(personList.get(x).get(1));
                jComboBox4.addItem(personList.get(x).get(1));
            }*/
        }
        //catch(IOException | ParseException | ClassNotFoundException | SQLException e) {
        catch(ClassNotFoundException | SQLException e) {
            JOptionPane.showMessageDialog(null, e, "Error", JOptionPane.ERROR_MESSAGE);
            //e.printStackTrace();
        }
        
    }
    
    private void initTicketCounter() {
        /* Start of Ticket Counter */
        teamStatList.clear();
        teamStatItem = new ArrayList<>();
        teamStatItem.add("Person");
        teamStatItem.add("On Time");
        teamStatItem.add("Early");
        teamStatItem.add("Total Number of Tickets (of that Person)");
        teamStatList.add(teamStatItem);
        /* End of Ticket Counter */
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        buttonGroup1 = new javax.swing.ButtonGroup();
        jButton1 = new javax.swing.JButton();
        jButton2 = new javax.swing.JButton();
        jComboBox1 = new javax.swing.JComboBox();
        jComboBox2 = new javax.swing.JComboBox();
        jTextField1 = new javax.swing.JTextField();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        jComboBox3 = new javax.swing.JComboBox();
        jLabel5 = new javax.swing.JLabel();
        jTextField2 = new javax.swing.JTextField();
        jRadioButton1 = new javax.swing.JRadioButton();
        jRadioButton2 = new javax.swing.JRadioButton();
        jLabel6 = new javax.swing.JLabel();
        jLabel7 = new javax.swing.JLabel();
        jComboBox4 = new javax.swing.JComboBox();
        jLabel8 = new javax.swing.JLabel();
        jComboBox5 = new javax.swing.JComboBox();
        jButton3 = new javax.swing.JButton();
        jMenuBar1 = new javax.swing.JMenuBar();
        jMenu1 = new javax.swing.JMenu();
        jMenuItem3 = new javax.swing.JMenuItem();
        jMenuItem4 = new javax.swing.JMenuItem();
        jMenuItem1 = new javax.swing.JMenuItem();
        jMenu2 = new javax.swing.JMenu();
        jMenuItem2 = new javax.swing.JMenuItem();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("SDS Report");
        setResizable(false);

        jButton1.setText("Extract Report");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });

        jButton2.setText("Open Report");
        jButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton2ActionPerformed(evt);
            }
        });

        jComboBox1.setModel(new javax.swing.DefaultComboBoxModel(new String[] { "Service Desk", "Problem Management", "Change and Release Management" }));
        jComboBox1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jComboBox1ActionPerformed(evt);
            }
        });

        jComboBox2.setModel(new javax.swing.DefaultComboBoxModel(new String[] { "All", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" }));
        jComboBox2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jComboBox2ActionPerformed(evt);
            }
        });

        jTextField1.setCursor(new java.awt.Cursor(java.awt.Cursor.TEXT_CURSOR));
        jTextField1.setName(""); // NOI18N

        jLabel1.setFont(new java.awt.Font("Tahoma", 1, 12)); // NOI18N
        jLabel1.setText("Workspace");

        jLabel2.setFont(new java.awt.Font("Tahoma", 1, 12)); // NOI18N
        jLabel2.setText("Month");

        jLabel4.setFont(new java.awt.Font("Tahoma", 1, 12)); // NOI18N
        jLabel4.setText("Year");

        jLabel3.setFont(new java.awt.Font("Tahoma", 1, 12)); // NOI18N
        jLabel3.setText("Type of Ticket");

        jComboBox3.setModel(new javax.swing.DefaultComboBoxModel(new String[] { "All", "Request", "Incident", "ITSM" }));
        jComboBox3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jComboBox3ActionPerformed(evt);
            }
        });

        jLabel5.setFont(new java.awt.Font("Tahoma", 1, 12)); // NOI18N
        jLabel5.setText("Filename");

        jTextField2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField2ActionPerformed(evt);
            }
        });
        jTextField2.addFocusListener(new java.awt.event.FocusAdapter() {
            public void focusLost(java.awt.event.FocusEvent evt) {
                jTextField2FocusLost(evt);
            }
        });

        buttonGroup1.add(jRadioButton1);
        jRadioButton1.setSelected(true);
        jRadioButton1.setText("SLA Report");
        jRadioButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jRadioButton1ActionPerformed(evt);
            }
        });

        buttonGroup1.add(jRadioButton2);
        jRadioButton2.setText("KB Monitoring Report");
        jRadioButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jRadioButton2ActionPerformed(evt);
            }
        });

        jLabel6.setFont(new java.awt.Font("Tahoma", 1, 12)); // NOI18N
        jLabel6.setText("Report Type");

        jLabel7.setFont(new java.awt.Font("Tahoma", 1, 12)); // NOI18N
        jLabel7.setText("Person");

        jComboBox4.setModel(new javax.swing.DefaultComboBoxModel(new String[] { "All" }));
        jComboBox4.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jComboBox4ActionPerformed(evt);
            }
        });

        jLabel8.setFont(new java.awt.Font("Tahoma", 1, 12)); // NOI18N
        jLabel8.setText("Team");

        jComboBox5.setModel(new javax.swing.DefaultComboBoxModel(new String[] { "All", "Asset Management", "Business Analyst", "Developer", "Field Support", "Network and Communications", "Operations Management", "SMO", "Service Desk", "System Administrator" }));
        jComboBox5.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jComboBox5ActionPerformed(evt);
            }
        });

        jButton3.setText("Open Ticket Counter");
        jButton3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton3ActionPerformed(evt);
            }
        });

        jMenu1.setMnemonic('F');
        jMenu1.setText("File");

        jMenuItem3.setText("Configure Holidays");
        jMenuItem3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem3ActionPerformed(evt);
            }
        });
        jMenu1.add(jMenuItem3);

        jMenuItem4.setText("Reload Holidays");
        jMenuItem4.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem4ActionPerformed(evt);
            }
        });
        jMenu1.add(jMenuItem4);

        jMenuItem1.setMnemonic('x');
        jMenuItem1.setText("Exit");
        jMenuItem1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem1ActionPerformed(evt);
            }
        });
        jMenu1.add(jMenuItem1);

        jMenuBar1.add(jMenu1);

        jMenu2.setMnemonic('H');
        jMenu2.setText("Help");

        jMenuItem2.setMnemonic('b');
        jMenuItem2.setText("About");
        jMenuItem2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem2ActionPerformed(evt);
            }
        });
        jMenu2.add(jMenuItem2);

        jMenuBar1.add(jMenu2);

        setJMenuBar(jMenuBar1);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(layout.createSequentialGroup()
                                        .addGap(10, 10, 10)
                                        .addComponent(jLabel6))
                                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                                        .addGap(19, 19, 19)
                                        .addComponent(jLabel1)))
                                .addComponent(jLabel3, javax.swing.GroupLayout.Alignment.TRAILING)
                                .addComponent(jLabel2, javax.swing.GroupLayout.Alignment.TRAILING)
                                .addComponent(jLabel7, javax.swing.GroupLayout.Alignment.TRAILING)
                                .addComponent(jLabel8, javax.swing.GroupLayout.Alignment.TRAILING))
                            .addComponent(jLabel5, javax.swing.GroupLayout.Alignment.TRAILING))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                            .addComponent(jTextField2, javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jComboBox4, javax.swing.GroupLayout.Alignment.LEADING, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addComponent(jComboBox5, javax.swing.GroupLayout.Alignment.LEADING, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(jComboBox2, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(jLabel4)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, 51, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(javax.swing.GroupLayout.Alignment.LEADING, layout.createSequentialGroup()
                                .addComponent(jRadioButton1)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                .addComponent(jRadioButton2)
                                .addGap(0, 0, Short.MAX_VALUE))
                            .addComponent(jComboBox1, javax.swing.GroupLayout.Alignment.LEADING, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addComponent(jComboBox3, javax.swing.GroupLayout.Alignment.LEADING, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(17, 17, 17)
                        .addComponent(jButton1)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jButton2, javax.swing.GroupLayout.PREFERRED_SIZE, 105, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jButton3)
                        .addGap(22, 22, 22)))
                .addContainerGap())
        );

        layout.linkSize(javax.swing.SwingConstants.HORIZONTAL, new java.awt.Component[] {jButton1, jButton2, jButton3});

        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jRadioButton1)
                    .addComponent(jRadioButton2)
                    .addComponent(jLabel6))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel1)
                    .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel2)
                    .addComponent(jComboBox2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel4)
                    .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jComboBox3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel3))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel8)
                    .addComponent(jComboBox5, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel7)
                    .addComponent(jComboBox4, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel5)
                    .addComponent(jTextField2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.CENTER)
                    .addComponent(jButton3)
                    .addComponent(jButton2)
                    .addComponent(jButton1))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void beginExtract() {
        if(validateFilters()) {
            if(!jTextField2.getText().isEmpty()) {
                // Fix filename extension
                fixFilenameExtension();
                // Validate Date
                if(!yearFilter.isEmpty()) {
                    try {
                        Calendar calendar = Calendar.getInstance();
                        if(monthFilter.equalsIgnoreCase("All")) {
                            calendar.setTime(sdfTimestamp.parse(yearFilter+"-01-01 11:59:59"));
                        }
                        else {
                            calendar.setTime(sdfTimestamp.parse(yearFilter+"-"+monthFilter+"-01 11:59:59"));
                        }
                        calendar.set(Calendar.DAY_OF_MONTH, calendar.getActualMaximum(Calendar.DAY_OF_MONTH));
                        calendar.set(Calendar.HOUR_OF_DAY, 23);
                        calendar.set(Calendar.MINUTE, 59);
                        calendar.set(Calendar.SECOND, 59);
                        calendar.set(Calendar.MILLISECOND, 999);
                        dateFilter = calendar.getTime();
                        //System.out.println(dateFilter);
                        //System.out.println((yearFilter+"-"+monthFilter+"-"+dayFilter+" 11:59:59"));
                        //dateFilter = sdfTimestamp.parse(sdfTimestamp.format(yearFilter+"-"+monthFilter+"-"+dayFilter+" 11:59:59"));
                    }
                    catch(ParseException pe) {
                        JOptionPane.showMessageDialog(null, pe, "Error", JOptionPane.WARNING_MESSAGE);
                    }
                    
                    if(jRadioButton1.isSelected()) {
                        extractReport();
                    }
                    else if(jRadioButton2.isSelected()) {
                        extractKBReport();
                    }
                }
                else {
                    if(jRadioButton1.isSelected()) {
                        extractReport();
                    }
                    else if(jRadioButton2.isSelected()) {
                        extractKBReport();
                    }
                }
            }
            else {
                JOptionPane.showMessageDialog(null, "Please provide a filename for the report", "Error", JOptionPane.WARNING_MESSAGE);
            }
        }
        else {
            JOptionPane.showMessageDialog(null, "Please check the year filter", "Error", JOptionPane.WARNING_MESSAGE);
        }
    }
    
    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        beginExtract();
        System.gc();
    }//GEN-LAST:event_jButton1ActionPerformed

    private void jMenuItem1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem1ActionPerformed
        System.exit(0);
    }//GEN-LAST:event_jMenuItem1ActionPerformed

    private void jMenuItem2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem2ActionPerformed
        JOptionPane.showMessageDialog(null, "SDS Report\nCopyright 2015\nIndra Philippines Inc. and Maynilad Water Services Inc.", "About SDS Report", JOptionPane.INFORMATION_MESSAGE);
    }//GEN-LAST:event_jMenuItem2ActionPerformed

    private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton2ActionPerformed
        try {
            Desktop dt = Desktop.getDesktop();
            reportFilename = jTextField2.getText();
            dt.open(new File(reportFilename));
            Process p = Runtime.getRuntime().exec("rundll32 url.dll,FileProtocolHandler " + reportFilename);
        }
        catch(Exception e) {
            JOptionPane.showMessageDialog(null, "Report file not found. Please run the report extraction.", "Warning", JOptionPane.WARNING_MESSAGE);
        }
    }//GEN-LAST:event_jButton2ActionPerformed

    private void jComboBox2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jComboBox2ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jComboBox2ActionPerformed

    private void uiChanged() {
        if(jRadioButton1.isSelected()) {
            if(jComboBox1.getSelectedIndex()==0) {
                reportFilename = "SDS_Report_Service_Desk.xls";
            }
            else if(jComboBox1.getSelectedIndex()==1) {
                reportFilename = "SDS_Report_Problem_Management.xls";
            }
            else if(jComboBox1.getSelectedIndex()==2) {
                reportFilename = "SDS_Report_Change_and_Release_Management.xls";
            }
            jComboBox3.setSelectedIndex(0); //All
            jComboBox3.setEnabled(true);
        }
        else if(jRadioButton2.isSelected()) {
            if(jComboBox1.getSelectedIndex()==0) {
                reportFilename = "SDS_KB_Monitoring_Service_Desk.xls";
            }
            else if(jComboBox1.getSelectedIndex()==1) {
                reportFilename = "SDS_KB_Monitoring_Problem_Management.xls";
            }
            else if(jComboBox1.getSelectedIndex()==2) {
                reportFilename = "SDS_KB_Monitoring_Change_and_Release_Management.xls";
            }
            jComboBox3.setSelectedIndex(2); //Incident
            jComboBox3.setEnabled(false);
        }
        jTextField2.setText(reportFilename);
        initPersons();
    }
    
    private void jComboBox1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jComboBox1ActionPerformed
        uiChanged();
    }//GEN-LAST:event_jComboBox1ActionPerformed

    private void jTextField2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField2ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField2ActionPerformed

    private void jTextField2FocusLost(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_jTextField2FocusLost
        // Fix filename extension
        if(!jTextField2.getText().isEmpty()) {
            fixFilenameExtension();
        }
    }//GEN-LAST:event_jTextField2FocusLost

    private void jRadioButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jRadioButton1ActionPerformed
        uiChanged();
    }//GEN-LAST:event_jRadioButton1ActionPerformed

    private void jRadioButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jRadioButton2ActionPerformed
        uiChanged();
    }//GEN-LAST:event_jRadioButton2ActionPerformed

    private void jMenuItem3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem3ActionPerformed
        try {
            Runtime.getRuntime().exec("notepad.exe Holidays.txt");
        }
        catch(FileNotFoundException fnf) {
            JOptionPane.showMessageDialog(null, fnf, "File not found", JOptionPane.ERROR_MESSAGE);
        }
        catch(IOException e) {
            JOptionPane.showMessageDialog(null, e, "Error", JOptionPane.ERROR_MESSAGE);
        }
    }//GEN-LAST:event_jMenuItem3ActionPerformed

    private void jMenuItem4ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem4ActionPerformed
        initHolidays();
        JOptionPane.showMessageDialog(null, "Holidays have been reloaded", "Holidays", JOptionPane.INFORMATION_MESSAGE);
    }//GEN-LAST:event_jMenuItem4ActionPerformed

    private void jComboBox5ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jComboBox5ActionPerformed
        if(jComboBox5.getSelectedItem()!=null) {
            teamFilter = String.valueOf(jComboBox5.getSelectedItem());
            teamFilter = Utilities.unfixString(teamFilter);
            initPersons();
        }
    }//GEN-LAST:event_jComboBox5ActionPerformed

    private void jComboBox3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jComboBox3ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jComboBox3ActionPerformed

    private void jComboBox4ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jComboBox4ActionPerformed
        if(jComboBox4.getSelectedItem()!=null) {
            personFilter = String.valueOf(jComboBox4.getSelectedItem());
            for(List<String> personRow : personList) {
                if(personFilter.equalsIgnoreCase(personRow.get(1))) {
                    personFilter = personRow.get(0);
                    break;
                }
            }
            /*for(int x=0;x<personList.size();x++) {
                if(personFilter.equalsIgnoreCase(personList.get(x).get(1))) {
                    personFilter = personList.get(x).get(0);
                    break;
                }
            }*/
        }
    }//GEN-LAST:event_jComboBox4ActionPerformed

    private void jButton3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton3ActionPerformed
        try {
            Desktop dt = Desktop.getDesktop();
            reportFilename = jTextField2.getText();
            dt.open(new File(reportFilename));
            Process p = Runtime.getRuntime().exec("rundll32 url.dll,FileProtocolHandler TicketCounter.xls");
        }
        catch(Exception e) {
            JOptionPane.showMessageDialog(null, "Ticket counter file not found. Please run the report extraction.", "Warning", JOptionPane.WARNING_MESSAGE);
        }
    }//GEN-LAST:event_jButton3ActionPerformed
    
    private void fixFilenameExtension() {
        if(!jTextField2.getText().endsWith(".xls")) {
            jTextField2.setText(jTextField2.getText()+".xls");
        }
    }
    
    private boolean validateFilters() {
        int year = 0;
        String monthString = String.valueOf(jComboBox2.getSelectedItem());
        String ticketTypeString = String.valueOf(jComboBox3.getSelectedItem());
        
        if(monthString.equalsIgnoreCase("January")) {
            monthString = "01";
        }
        else if(monthString.equalsIgnoreCase("February")) {
            monthString = "02";
        }
        else if(monthString.equalsIgnoreCase("March")) {
            monthString = "03";
        }
        else if(monthString.equalsIgnoreCase("April")) {
            monthString = "04";
        }
        else if(monthString.equalsIgnoreCase("May")) {
            monthString = "05";
        }
        else if(monthString.equalsIgnoreCase("June")) {
            monthString = "06";
        }
        else if(monthString.equalsIgnoreCase("July")) {
            monthString = "07";
        }
        else if(monthString.equalsIgnoreCase("August")) {
            monthString = "08";
        }
        else if(monthString.equalsIgnoreCase("September")) {
            monthString = "09";
        }
        else if(monthString.equalsIgnoreCase("October")) {
            monthString = "10";
        }
        else if(monthString.equalsIgnoreCase("November")) {
            monthString = "11";
        }
        else if(monthString.equalsIgnoreCase("December")) {
            monthString = "12";
        }
        
        if(!monthString.equalsIgnoreCase("All") || !jTextField1.getText().isEmpty()) {
            try { 
                year = Integer.parseInt(jTextField1.getText()); 
                if(year<1900 || year>9999) {
                    return false;
                }
            } catch(NumberFormatException e) { 
                return false; 
            }
        }
        
        /*if(ticketTypeString.equalsIgnoreCase("All")) {
            ticketTypeString = "";
        }*/
        
        monthFilter = monthString;
        yearFilter = String.valueOf(year);
        ticketTypeFilter = ticketTypeString;
        
        return true;
    }
    
    private void extractReport() {
        reportList.clear();
        String url = "jdbc:jtds:sqlserver://"+hostname+":"+port+"/"+databaseName;
        String driver = "net.sourceforge.jtds.jdbc.Driver";
        String user = ""; // Removed for GitHub
        String pass = ""; // Removed for GitHub
        
        String selectedWorkspace = "1";
        if(jComboBox1.getSelectedIndex()==0) {
            selectedWorkspace = "1";
        }
        else if(jComboBox1.getSelectedIndex()==1) {
            selectedWorkspace = "2";
        }
        else if(jComboBox1.getSelectedIndex()==2) {
            selectedWorkspace = "3";
        }
        
        String query = "SELECT a.mrID, a.mrTitle, a.mrStatus, a.mrUpdateDate, a.mrAssignees, a.mrSUBMITDATE, a.Type__bof__bTicket, a.SLA__bDue__bDate, a.Priority, b.VIP__bTAG, a.Impact, a.Resolution__bCode, a.Category, a.Sub__bCategory, a.Ticket__bSummary" +
                       " FROM MASTER"+selectedWorkspace+" a"+ // i.e. MASTER1, MASTER2, MASTER3
                       " INNER JOIN MASTER"+selectedWorkspace+"_ABDATA b ON a.mrID = b.mrID";
        query += " WHERE (a.mrStatus = 'Resolved' OR a.mrStatus = 'Closed')";
        
        if(!ticketTypeFilter.equalsIgnoreCase("All")) {
            if(ticketTypeFilter.equalsIgnoreCase("ITSM")) {
                query += " AND a.Sub__bCategory = 'ITSM__bDocumentation'";
            }
            else {
                query += " AND a.Type__bof__bTicket = '"+ticketTypeFilter+"'";
            }
        }
        
        if(!teamFilter.equalsIgnoreCase("All") && !personFilter.equalsIgnoreCase("All")) {
            query += " AND ((a.mrAssignees LIKE '"+teamFilter+"%'";
            query += " OR a.mrAssignees LIKE '%"+teamFilter+"'";
            query += " OR a.mrAssignees LIKE '%"+teamFilter+"%')";
            query += " AND (a.mrAssignees LIKE '"+personFilter+"%'";
            query += " OR a.mrAssignees LIKE '%"+personFilter+"'";
            query += " OR a.mrAssignees LIKE '%"+personFilter+"%'))";
        }
        else if(!teamFilter.equalsIgnoreCase("All")) {
            query += " AND (a.mrAssignees LIKE '"+teamFilter+"%'";
            query += " OR a.mrAssignees LIKE '%"+teamFilter+"'";
            query += " OR a.mrAssignees LIKE '%"+teamFilter+"%')";
        }
        else if(!personFilter.equalsIgnoreCase("All")) {
            query += " AND (a.mrAssignees LIKE '"+personFilter+"%'";
            query += " OR a.mrAssignees LIKE '%"+personFilter+"'";
            query += " OR a.mrAssignees LIKE '%"+personFilter+"%')";
        }
        
        query += " ORDER BY a.mrID ASC";
        try {
            Class.forName(driver);
            Connection conn;
            conn = DriverManager.getConnection(url, user, pass);
            for( SQLWarning warn = conn.getWarnings(); warn != null; warn = warn.getNextWarning() ) {
                    System.out.println( "SQL Warning:" ) ;
                    System.out.println( "State  : " + warn.getSQLState()  ) ;
                    System.out.println( "Message: " + warn.getMessage()   ) ;
                    System.out.println( "Error  : " + warn.getErrorCode() ) ;
            }

            reportItem = new ArrayList<>();
            reportItem.add("Ticket ID");
            reportItem.add("Title");
            reportItem.add("New Status");
            reportItem.add("Old Status");
            reportItem.add("Date Resolved");
            reportItem.add("Assignees");
            reportItem.add("Date Created");
            reportItem.add("Type of Ticket");
            reportItem.add("Date Due");
            reportItem.add("Severity");
            reportItem.add("Resolution Code");
            reportItem.add("Team Assigned");
            reportItem.add("Category");
            reportItem.add("Sub-Category");
            reportItem.add("Ticket Summary");
            reportItem.add("Age of Ticket");
            reportItem.add("RESOLVED on Time");
            reportItem.add("RESOLVED Early");
            reportList.add(reportItem);
            try (Statement stmt = conn.createStatement(); ResultSet rs = stmt.executeQuery( query )) {
                
                //int recordCount = 0;
                while( rs.next() ) {
                    String ticketID = rs.getString(1);
                    String title = StringEscapeUtils.unescapeHtml4(rs.getString(2));
                    String oldStatus = "";
                    String newStatus = rs.getString(3);
                    if(newStatus!=null) {
                        newStatus = newStatus.replaceAll("_PENDING_SOLUTION_", "Pending Solution");
                        newStatus = newStatus.replaceAll("_SOLVED_", "Solved");
                        newStatus = Utilities.fixString(newStatus);
                    }
                    String dateResolved = "";
                    /*if(rs.getString(4)!=null) {
                        dateResolved = sdf.format(sdf.parse(rs.getString(4).substring(0,10)));
                    }*/
                    String assignees = Utilities.fixString(rs.getString(5));
                    //assignees = assignees.split(" CC:")[0]; // Get only string before CC:
                    assignees = assignees.replaceAll("(\\bCC:)\\S+\\b","").trim();
                    assignees = Utilities.removeOrganization(assignees);                    
                    String dateCreated = "";
                    if(rs.getString(6)!=null) {
                        dateCreated = sdfTimestamp.format(sdfTimestamp.parse(rs.getString(6).substring(0,19)));
                    }
                    String typeOfTicket = rs.getString(7);
                    typeOfTicket = Utilities.fixString(typeOfTicket);
                    String dateDue = "";
                    if(rs.getString(8)!=null) {
                        dateDue = sdfTimestamp.format(sdfTimestamp.parse(rs.getString(8).substring(0,19)));
                    }
                    String priority =  Utilities.fixString(rs.getString(9));
                    String vip = Utilities.fixString(rs.getString(10));
                    String impact = Utilities.fixString(rs.getString(11));
                    String severity;   
                    String resolutionCode = Utilities.fixString(rs.getString(12));
                    String teamAssigned = Utilities.fixString(rs.getString(5));
                    String teams = "";
                    String category = Utilities.fixString(rs.getString(13));
                    String subCategory = Utilities.fixString(rs.getString(14));
                    String ticketSummary = Utilities.fixString(rs.getString(15));
                    String ageOfTicket = "";
                    String resolvedOnTime = "";
                    String resolvedEarly = "";
                    
                    /* Start of Team Assigned */
                    if(teamAssigned.contains("Asset Management")) {
                        teams += "Asset Management ";
                    }
                    if(teamAssigned.contains("Business Analyst")) {
                        teams += "Business Analyst ";
                    }
                    if(teamAssigned.contains("Developer")) {
                        teams += "Developer ";
                    }
                    if(teamAssigned.contains("Field Support")) {
                        teams += "Field Support ";
                    }
                    if(teamAssigned.contains("Network and Communications")) {
                        teams += "Network and Communications ";
                    }
                    if(teamAssigned.contains("Operations Management")) {
                        teams += "Operations Management ";
                    }
                    if(teamAssigned.contains("SMO")) {
                        teams += "SMO ";
                    }
                    if(teamAssigned.contains("Service Desk")) {
                        teams += "Service Desk ";
                    }
                    if(teamAssigned.contains("System Administrator")) {
                        teams += "System Administrator ";
                    }
                    if(!teams.isEmpty()) {
                        teamAssigned = teams.trim();
                    }
                    /* End of Team Assigned */
                    
                    /* Start of Severity */                 
                    if(vip.equalsIgnoreCase("TMT")) {
                        severity = "Critical";
                    }
                    else if(vip.equalsIgnoreCase("Yes")) {
                        severity = "Urgent";
                    }
                    else {
                        if(vip.equalsIgnoreCase("")) {
                            vip = "No";
                        }
                        
                        if(priority.equalsIgnoreCase("High") && vip.equalsIgnoreCase("No") && impact.equalsIgnoreCase("Enterprise-Wide")) {
                            severity = "Critical";
                        }
                        else if(priority.equalsIgnoreCase("High") && vip.equalsIgnoreCase("No") && impact.equalsIgnoreCase("Whole BA")) {
                            severity = "Urgent";
                        }
                        else if(priority.equalsIgnoreCase("High") && vip.equalsIgnoreCase("No") && impact.equalsIgnoreCase("Whole Office")) {
                            severity = "Important";
                        }
                        else if(priority.equalsIgnoreCase("High") && vip.equalsIgnoreCase("No") && impact.equalsIgnoreCase("Single User")) {
                            severity = "Important";
                        }
                        else if(priority.equalsIgnoreCase("Medium") && vip.equalsIgnoreCase("No") && impact.equalsIgnoreCase("Enterprise-Wide")) {
                            severity = "Urgent";
                        }
                        else if(priority.equalsIgnoreCase("Medium") && vip.equalsIgnoreCase("No") && impact.equalsIgnoreCase("Whole BA")) {
                            severity = "Important";
                        }
                        else if(priority.equalsIgnoreCase("Medium") && vip.equalsIgnoreCase("No") && impact.equalsIgnoreCase("Whole Office")) {
                            severity = "Minor";
                        }
                        else if(priority.equalsIgnoreCase("Medium") && vip.equalsIgnoreCase("No") && impact.equalsIgnoreCase("Single User")) {
                            severity = "Minor";
                        }
                        else if(priority.equalsIgnoreCase("Low") && vip.equalsIgnoreCase("No") && impact.equalsIgnoreCase("Enterprise-Wide")) {
                            severity = "Important";
                        }
                        else if(priority.equalsIgnoreCase("Low") && vip.equalsIgnoreCase("No") && impact.equalsIgnoreCase("Whole BA")) {
                            severity = "Minor";
                        }
                        else if(priority.equalsIgnoreCase("Low") && vip.equalsIgnoreCase("No") && impact.equalsIgnoreCase("Whole Office")) {
                            severity = "Minor";
                        }
                        else if(priority.equalsIgnoreCase("Low") && vip.equalsIgnoreCase("No") && impact.equalsIgnoreCase("Single User")) {
                            severity = "Minor";
                        }
                        else {
                            severity = "Important";
                        }
                    }
                    /* End of Severity */
                    
                    /* Start of Old Status */
                    String query2 = "select top 1 mrHISTORY from MASTER"+selectedWorkspace+"_HISTORY ";
                    query2 += " where mrID = '"+ticketID+"'";
                    query2 += " and not (mrHistory like '%Resolved%')";
                    query2 += " order by mrGENERATION desc";
                    
                    try (Statement stmt2 = conn.createStatement(); ResultSet rs2 = stmt2.executeQuery( query2 )) {
                        oldStatus = "";
                        while( rs2.next() ) {
                            String oldStatusString = Utilities.fixString(rs2.getString(1));
                            oldStatusString = oldStatusString.split("\\s+")[3];
                            if(oldStatusString != null) {
                                if(oldStatusString.equalsIgnoreCase("_PENDING_SOLUTION_")) {
                                    oldStatus = "Pending Solution";
                                }
                                else if(oldStatusString.equalsIgnoreCase("_SOLVED_")) {
                                    oldStatus = "Solved";
                                }
                                else if(oldStatusString.equalsIgnoreCase("Accepted")) {
                                    oldStatus = "Accepted";
                                }
                                else if(oldStatusString.equalsIgnoreCase("Assigned")) {
                                    oldStatus = "Assigned";
                                }
                                else if(oldStatusString.equalsIgnoreCase("Cancelled")) {
                                    oldStatus = "Cancelled";
                                }
                                else if(oldStatusString.equalsIgnoreCase("Closed")) {
                                    oldStatus = "Closed";
                                }
                                else if(oldStatusString.equalsIgnoreCase("Open")) {
                                    oldStatus = "Open";
                                }
                                else if(oldStatusString.equalsIgnoreCase("Pending")) {
                                    oldStatus = "Pending";
                                }
                                else if(oldStatusString.equalsIgnoreCase("Resolved")) {
                                    oldStatus = "Resolved";
                                }
                                else if(oldStatusString.equalsIgnoreCase("Escalated")) {
                                    oldStatus = "Escalated";
                                }
                                else if(oldStatusString.equalsIgnoreCase("Re-opened")) {
                                    oldStatus = "Re-opened";
                                }
                                else {
                                    oldStatus = "Unknown";
                                }
                            }
                        }
                        stmt2.cancel();
                        stmt2.close();
                        rs2.close();
                    }
                    catch(SQLException e) {
                        JOptionPane.showMessageDialog(null, e, "Error ", JOptionPane.ERROR_MESSAGE);
                    }
                    /* End of Old Status */
                    
                    /* Start of Date Resolved */
                    String query3 = "select mrHISTORY, mrGeneration from MASTER"+selectedWorkspace+"_HISTORY ";
                    query3 += " where mrID = '"+ticketID+"'";
                    query3 += " and mrHistory like '%Resolved%'";
                    query3 += " order by mrGENERATION desc";
                    //System.out.println("ticketID: "+ticketID);
                    int mrGeneration = 0;
                    try (Statement stmt3 = conn.createStatement(); ResultSet rs3 = stmt3.executeQuery( query3 )) {
                        while( rs3.next() ) {
                            String dateResolvedString = Utilities.fixString(rs3.getString(1));
                            if(mrGeneration == 0) {
                                mrGeneration = rs3.getInt(2);
                                dateResolved = dateResolvedString.split("\\s+")[0]+" "+dateResolvedString.split("\\s+")[1];
                            }
                            else {
                                if(mrGeneration-1 == rs3.getInt(2)) {
                                    mrGeneration = rs3.getInt(2);
                                    dateResolved = dateResolvedString.split("\\s+")[0]+" "+dateResolvedString.split("\\s+")[1];
                                }
                                else {
                                    break; //Date Resolved already found
                                }
                            }
                        }
                        stmt3.cancel();
                        stmt3.close();
                        rs3.close();
                    }
                    if(dateResolved.isEmpty()) {
                        /* Retry query as Closed since Ticket has no Resolved status */
                        query3 = "select mrHISTORY, mrGeneration from MASTER"+selectedWorkspace+"_HISTORY ";
                        query3 += " where mrID = '"+ticketID+"'";
                        query3 += " and mrHistory like '%Closed%'";
                        query3 += " order by mrGENERATION desc";
                        //System.out.println("ticketID: "+ticketID);
                        mrGeneration = 0;
                        try (Statement stmt3 = conn.createStatement(); ResultSet rs3 = stmt3.executeQuery( query3 )) {
                            while( rs3.next() ) {
                                String dateResolvedString = Utilities.fixString(rs3.getString(1));
                                if(mrGeneration == 0) {
                                    mrGeneration = rs3.getInt(2);
                                    dateResolved = dateResolvedString.split("\\s+")[0]+" "+dateResolvedString.split("\\s+")[1];
                                }
                                else {
                                    if(mrGeneration-1 == rs3.getInt(2)) {
                                        mrGeneration = rs3.getInt(2);
                                        dateResolved = dateResolvedString.split("\\s+")[0]+" "+dateResolvedString.split("\\s+")[1];
                                    }
                                    else {
                                        break; //Date Closed already found
                                    }
                                }
                            }
                            stmt3.cancel();
                            stmt3.close();
                            rs3.close();
                        }
                    }
                    //System.out.println("mrGeneration: "+mrGeneration);
                    //System.out.println("dateResolved: "+dateResolved);
                    /* End of Date Resolved */
                    /* Start of Age of Ticket */
                    String query4 = "select mrHISTORY from MASTER"+selectedWorkspace+"_HISTORY ";
                    query4 += " where mrID = '"+ticketID+"'";
                    query4 += " order by mrGENERATION asc";
                    
                    List<List<String>> historyList = new ArrayList<>();
                    List<String> historyItem;
                    historyList.clear();
                    Date d1;// = null;
                    Date d2 = null;
                    Date dResolved = sdfTimestamp.parse(dateResolved);
                    String d1_Status;// = "";
                    String d2_Status = "";
                    long totalSeconds = 0;
                    long totalMinutes = 0;
                    long totalHours = 0;
                    long totalDays = 0;
                    boolean pending = true;
                    boolean holiday = false;
                    boolean weekend = false;
                    long diffSeconds = 0;
                    long diffMinutes = 0;
                    long diffHours = 0;
                    long diffDays = 0;
                    
                    try (Statement stmt4 = conn.createStatement(); ResultSet rs4 = stmt4.executeQuery( query4 )) {
                        while( rs4.next() ) {
                            String history = Utilities.fixString(rs4.getString(1));
                            String timestamp = history.split("\\s+")[0]+" "+history.split("\\s+")[1]; 
                            String historyStatus = history.split("\\s+")[3];
                            
                            historyItem = new ArrayList<>();
                            historyItem.add(timestamp);
                            historyItem.add(historyStatus);
                            historyList.add(historyItem);
                        }
                        stmt4.cancel();
                        stmt4.close();
                        rs4.close();
                        
                        for(List<String> historyRow : historyList) {
                            //System.out.println(sdfTimestamp.parse(historyList.get(x).get(0)) + " vs "+dResolved);
                            // Exit loop if date of first resolution is found
                            if(sdfTimestamp.parse(historyRow.get(0)).after(dResolved)) {
                                break;
                            }
                            
                            d1 = d2;
                            d2 = sdfTimestamp.parse(historyRow.get(0));
                            d1_Status = d2_Status;
                            d2_Status = historyRow.get(1);
                            
                            /* Time Trim -- Start */
                            /* If time is less than 8am or greater than 5pm, cut to 8am or cut to 5pm */
                            int hour = Integer.valueOf(sdfHH.format(d2));
                            int minute = Integer.valueOf(sdfmm.format(d2));
                            int seconds = Integer.valueOf(sdfss.format(d2));
                            if(hour<8) {
                                d2 = sdfTimestamp.parse(sdfDatestamp.format(sdfDatestamp.parse(historyRow.get(0))) + " 08:00:00");
                            }
                            else if(hour==17) {
                                if(minute>0) {
                                    d2 = sdfTimestamp.parse(sdfDatestamp.format(sdfDatestamp.parse(historyRow.get(0))) + " 17:00:00");
                                }
                                else if(minute==0) {
                                    if(seconds>0) {
                                        d2 = sdfTimestamp.parse(sdfDatestamp.format(sdfDatestamp.parse(historyRow.get(0))) + " 17:00:00");
                                    }
                                }
                            }
                            else if(hour>17) {
                                d2 = sdfTimestamp.parse(sdfDatestamp.format(sdfDatestamp.parse(historyRow.get(0))) + " 17:00:00");
                            }
                            else if(hour==12 && minute>0)  {
                                d2 = sdfTimestamp.parse(sdfDatestamp.format(sdfDatestamp.parse(historyRow.get(0))) + " 12:00:00");
                            }
                            /* Time Trim -- End */
                            /*System.out.println(x+": "+sdfTimestamp.parse(historyList.get(x).get(0)));
                            System.out.println(x+": "+d2);
                            System.out.println("");*/
                            
                            if(d1!=null) {
                                if(!pending && !holiday && !weekend) { // Add last compute
                                    holiday = false;
                                    if(diffSeconds > 0) {
                                        totalSeconds += diffSeconds;
                                    }
                                    if(diffMinutes > 0) {
                                        totalMinutes += diffMinutes;
                                    }
                                    if(diffHours > 0) {
                                        totalHours += diffHours;
                                    }
                                    if(diffDays > 0) {
                                        totalDays += diffDays;
                                    }
                                    //System.out.println(totalDays + " days, " + totalHours + " hours, "+ totalMinutes + " minutes, "+totalSeconds + " seconds.");
                                }

                                long diff = d2.getTime() - d1.getTime();
                                diffSeconds = diff / 1000 % 60;
                                diffMinutes = diff / (60 * 1000) % 60;
                                diffHours = diff / (60 * 60 * 1000) % 24;
                                diffDays = diff / (24 * 60 * 60 * 1000);
                                
                                /*System.out.println(x+": "+d1+" vs "+d2);
                                System.out.println("Diff Seconds: "+diffSeconds);
                                System.out.println("Diff Minutes: "+diffMinutes);
                                System.out.println("Diff Hours: "+diffHours);
                                System.out.println("Diff Days: "+diffDays);
                                System.out.println("");*/
                                //if(historyList.get(x).get(1).equalsIgnoreCase("Pending")) {
                                if(d1_Status.equalsIgnoreCase("Pending")) {
                                    //System.out.println(x+": "+d1);
                                    pending = true;
                                }
                                else {
                                    pending = false;
                                }
                                
                                for(String holidayRow : holidayList) {
                                    if(sdfMMddyyyy.format(d1).equalsIgnoreCase(holidayRow)) {
                                        holiday = true;
                                        break;
                                    }
                                    else {
                                        holiday = false;
                                    }
                                }
                                /*for(int y=0;y<holidayList.size();y++) {
                                    if(sdfMMddyyyy.format(d1).equalsIgnoreCase(holidayList.get(y))) {
                                        holiday = true;
                                        break;
                                    }
                                    else {
                                        holiday = false;
                                    }
                                }*/
                                
                                Calendar cal = Calendar.getInstance();
                                cal.setTime(d1);
                                if (cal.get(Calendar.DAY_OF_WEEK) == Calendar.SUNDAY ||
                                    cal.get(Calendar.DAY_OF_WEEK) == Calendar.SATURDAY) {
                                    weekend = true;
                                }
                                else {
                                    weekend = false;
                                }
                            }
                        }
                        
                        if(!pending && !holiday) { // Add last compute
                            holiday = false;
                            if(diffSeconds > 0) {
                                totalSeconds += diffSeconds;
                            }
                            if(diffMinutes > 0) {
                                totalMinutes += diffMinutes;
                            }
                            if(diffHours > 0) {
                                totalHours += diffHours;
                            }
                            if(diffDays > 0) {
                                totalDays += diffDays;
                            }
                        }
                        while(totalSeconds >= 60) {
                            totalSeconds = totalSeconds - 60;
                            totalMinutes = totalMinutes + 1;
                        }
                        while(totalMinutes >= 60) {
                            totalMinutes = totalMinutes - 60;
                            totalHours = totalHours + 1;
                        }
                        while(totalHours >= 24) {
                            totalHours = totalHours - 24;
                            totalDays = totalDays + 1;
                        }
                        /* Subtract non-working time and lunch breaks */
                        long subtractHours = 0;
                        if(totalDays>2) {
                            subtractHours = ((totalDays-2) * 15) + 1;
                        }
                        else if(totalDays>1) {
                            subtractHours = 16; //Lunch break (1 hour) + Non-working hours (15 hours)
                        }
                        else if(totalDays==1 && totalHours>1) {
                            subtractHours = 16; //Lunch break (1 hour) + Non-working hours (15 hours)
                        }
                        else if(totalDays==1 && totalHours==0 && totalMinutes>1) {
                            totalDays = 0;
                            totalHours = 24; // Convert 1 day to 24 hours
                            subtractHours = 16; //Lunch break (1 hour) + Non-working hours (15 hours)
                        }
                        if(subtractHours > 0) {
                            if(subtractHours>=24) {
                                if(totalDays>0) {
                                    while(subtractHours>=24) {
                                        subtractHours = subtractHours - 24;
                                        totalDays = totalDays - 1;
                                    }
                                    totalDays = totalDays - 1;
                                }
                            }
                            if(totalHours >= subtractHours) {
                                totalHours = totalHours - subtractHours;
                            }
                            else {
                                totalDays = totalDays - 1;
                                totalHours = (totalHours + 24) - subtractHours;
                                if(totalHours > 24) {
                                    totalHours = totalHours - 24;
                                    totalDays = totalDays + 1;
                                }
                            }
                        }
                        /* End -- Subtract non-working time and lunch breaks */
                        
                        //System.out.println(totalDays + " days, " + totalHours + " hours, "+ totalMinutes + " minutes, "+totalSeconds + " seconds.");
                        ageOfTicket = totalDays + " days, " + totalHours + " hours, "+ totalMinutes + " minutes, "+totalSeconds + " seconds.";
                        //System.out.println(ageOfTicket);
                    }
                    catch(SQLException e) {
                        JOptionPane.showMessageDialog(null, e, "Error ", JOptionPane.ERROR_MESSAGE);
                    }
                    /* End of Age of Ticket */

                    /* Start of RESOLVED on Time / Early */
                    try {
                        Calendar calDateResolved = Calendar.getInstance();
                        Calendar calDateDue = Calendar.getInstance();
                        Calendar calDateResolved_DateOnly = Calendar.getInstance();
                        Calendar calDateDue_DateOnly = Calendar.getInstance();
                        
                        Calendar calDateResolved_Plus = Calendar.getInstance();
                        
                        calDateResolved.setTime(sdfTimestamp.parse(dateResolved));
                        calDateDue.setTime(sdfTimestamp.parse(dateDue));
                        
                        calDateResolved_DateOnly.setTime(sdfDatestamp.parse(dateResolved));
                        calDateDue_DateOnly.setTime(sdfDatestamp.parse(dateDue));
                        
                        /*System.out.println("calDateResolved: "+calDateResolved.getTime());
                        System.out.println("calDateDue: "+calDateDue.getTime());
                        System.out.println("calDateResolved_DateOnly: "+calDateResolved_DateOnly.getTime());
                        System.out.println("calDateDue_DateOnly: "+calDateDue_DateOnly.getTime());*/
                        
                        if(subCategory.equalsIgnoreCase("ITSM Documentation")) {
                            if(calDateDue_DateOnly.before(calDateResolved_DateOnly)) {
                                resolvedOnTime = "FAIL";
                                resolvedEarly = "NO";
                            }
                            else {
                                if(calDateDue.before(calDateResolved)) {
                                    resolvedOnTime = "PASS (Due Date breached)";
                                    resolvedEarly = "NO";
                                }
                                else {
                                    resolvedOnTime = "PASS";
                                    // 1 day earlier than due date
                                    calDateResolved_Plus.setTime(calDateResolved.getTime());
                                    calDateResolved_Plus.add(Calendar.DATE, 1);
                                    if(calDateDue.before(calDateResolved_Plus)) {
                                        resolvedEarly = "NO";
                                    }
                                    else {
                                        resolvedEarly = "YES";
                                    }
                                }
                            }
                        }
                        else if(typeOfTicket.equalsIgnoreCase("Incident")) {
                            if(severity.equalsIgnoreCase("Minor")) {
                                if(calDateDue_DateOnly.before(calDateResolved_DateOnly)) {
                                    resolvedOnTime = "FAIL";
                                    resolvedEarly = "NO";
                                }
                                else {
                                    if(calDateDue.before(calDateResolved)) {
                                        resolvedOnTime = "PASS (Due Date breached)";
                                        resolvedEarly = "NO";
                                    }
                                    else {
                                        resolvedOnTime = "PASS";
                                        // 6 hours earlier than due date
                                        calDateResolved_Plus.setTime(calDateResolved.getTime());
                                        calDateResolved_Plus.add(Calendar.HOUR_OF_DAY, 6);
                                        if(calDateDue.before(calDateResolved_Plus)) {
                                            resolvedEarly = "NO";
                                        }
                                        else {
                                            resolvedEarly = "YES";
                                        }
                                    }
                                }
                                
                            }
                            else if(severity.equalsIgnoreCase("Important")) {
                                if(calDateDue_DateOnly.before(calDateResolved_DateOnly)) {
                                    resolvedOnTime = "FAIL";
                                    resolvedEarly = "NO";
                                }
                                else {
                                    if(calDateDue.before(calDateResolved)) {
                                        resolvedOnTime = "PASS (Due Date breached)";
                                        resolvedEarly = "NO";
                                    }
                                    else {
                                        resolvedOnTime = "PASS";
                                        // 4 hours earlier than due date
                                        calDateResolved_Plus.setTime(calDateResolved.getTime());
                                        calDateResolved_Plus.add(Calendar.HOUR_OF_DAY, 4);
                                        if(calDateDue.before(calDateResolved_Plus)) {
                                            resolvedEarly = "NO";
                                        }
                                        else {
                                            resolvedEarly = "YES";
                                        }
                                    }
                                }
                            }
                            else if(severity.equalsIgnoreCase("Urgent")) {
                                if(calDateDue.before(calDateResolved)) {
                                    resolvedOnTime = "FAIL";
                                    resolvedEarly = "NO";
                                }
                                else {
                                    resolvedOnTime = "PASS";
                                    // 2 hours earlier than due date
                                    calDateResolved_Plus.setTime(calDateResolved.getTime());
                                    calDateResolved_Plus.add(Calendar.HOUR_OF_DAY, 2);
                                    if(calDateDue.before(calDateResolved_Plus)) {
                                        resolvedEarly = "NO";
                                    }
                                    else {
                                        resolvedEarly = "YES";
                                    }
                                }
                            }
                            else if(severity.equalsIgnoreCase("Critical")) {
                                if(calDateDue.before(calDateResolved)) {
                                    resolvedOnTime = "FAIL";
                                    resolvedEarly = "NO";
                                }
                                else {
                                    resolvedOnTime = "PASS";
                                    // 1 hour earlier than due date
                                    calDateResolved_Plus.setTime(calDateResolved.getTime());
                                    calDateResolved_Plus.add(Calendar.HOUR_OF_DAY, 1);
                                    if(calDateDue.before(calDateResolved_Plus)) {
                                        resolvedEarly = "NO";
                                    }
                                    else {
                                        resolvedEarly = "YES";
                                    }
                                }
                            }
                        }
                        else if(typeOfTicket.equalsIgnoreCase("Request")) {
                            if(calDateDue_DateOnly.before(calDateResolved_DateOnly)) {
                                resolvedOnTime = "FAIL";
                                resolvedEarly = "NO";
                            }
                            else {
                                if(calDateDue.before(calDateResolved)) {
                                    resolvedOnTime = "PASS (Due Date breached)";
                                    resolvedEarly = "NO";
                                }
                                else {
                                    resolvedOnTime = "PASS";
                                    // 2 days earlier than due date
                                    calDateResolved_Plus.setTime(calDateResolved.getTime());
                                    calDateResolved_Plus.add(Calendar.DATE, 2);
                                    
                                    if(calDateDue.before(calDateResolved_Plus)) {
                                        resolvedEarly = "NO";
                                    }
                                    else {
                                        resolvedEarly = "YES";
                                    }
                                }
                            }
                        }
                    }
                    catch(ParseException pe) {
                        //JOptionPane.showMessageDialog(null, pe, "Error ", JOptionPane.ERROR_MESSAGE);
                        resolvedOnTime = "PASS (No Due Date)";
                        resolvedEarly = "YES (No Due Date)";
                    }
                    /* Start of RESOLVED on Time / Early */
                    
                    
                    /*System.out.println("ticketTypeFilter: "+ticketTypeFilter);
                    System.out.println("yearFilter: "+yearFilter);*/
                    //if(ticketTypeFilter.equalsIgnoreCase("") || ticketTypeFilter.equalsIgnoreCase(typeOfTicket) || (ticketTypeFilter.equalsIgnoreCase("ITSM") && subCategory.equalsIgnoreCase("ITSM Documentation")) ) {
                    if((monthFilter.equalsIgnoreCase("All") && yearFilter.equalsIgnoreCase("0")) ||
                            (monthFilter.equalsIgnoreCase("All") && sdfYear.format(dateFilter).equals(sdfYear.format(sdfTimestamp.parse(dateResolved)))) ||
                            (!monthFilter.equalsIgnoreCase("All") && sdfMM.format(dateFilter).equals(sdfMM.format(sdfTimestamp.parse(dateResolved))) && sdfYear.format(dateFilter).equals(sdfYear.format(sdfTimestamp.parse(dateResolved)))) ) {
                            if(!personFilter.equalsIgnoreCase("All") && assignees.contains(personFilter)) {
                                reportItem = new ArrayList<>();
                                reportItem.add(ticketID);
                                reportItem.add(title);
                                reportItem.add(newStatus);
                                reportItem.add(oldStatus);
                                reportItem.add(dateResolved);
                                reportItem.add(assignees);
                                reportItem.add(dateCreated);
                                reportItem.add(typeOfTicket);
                                reportItem.add(dateDue);
                                reportItem.add(severity);
                                reportItem.add(resolutionCode);
                                reportItem.add(teamAssigned);
                                reportItem.add(category);
                                reportItem.add(subCategory);
                                reportItem.add(ticketSummary);
                                reportItem.add(ageOfTicket);
                                reportItem.add(resolvedOnTime);
                                reportItem.add(resolvedEarly);
                                reportList.add(reportItem);
                            }
                            else if(personFilter.equalsIgnoreCase("All")) {
                                reportItem = new ArrayList<>();
                                reportItem.add(ticketID);
                                reportItem.add(title);
                                reportItem.add(newStatus);
                                reportItem.add(oldStatus);
                                reportItem.add(dateResolved);
                                reportItem.add(assignees);
                                reportItem.add(dateCreated);
                                reportItem.add(typeOfTicket);
                                reportItem.add(dateDue);
                                reportItem.add(severity);
                                reportItem.add(resolutionCode);
                                reportItem.add(teamAssigned);
                                reportItem.add(category);
                                reportItem.add(subCategory);
                                reportItem.add(ticketSummary);
                                reportItem.add(ageOfTicket);
                                reportItem.add(resolvedOnTime);
                                reportItem.add(resolvedEarly);
                                reportList.add(reportItem);
                            }
                    }
                    //}
                    
                    /*System.out.println(ticketID + " " +
                                       title + " " +
                                       newStatus + " " +
                                       oldStatus + " " +
                                       dateModified + " " +
                                       assignees + " " +
                                       dateCreated + " " +
                                       typeOfTicket + " " +
                                       dateDue + " " +
                                       severity + " " +
                                       resolutionCode + " " +
                                       teamAssigned + " " +
                                       ageOfTicket);
                    recordCount++;*/
                }
                
                // Required for MS SQL Server
                stmt.cancel();
                stmt.close();
                rs.close();
            }
            conn.close();
            
            Workbook wb = new HSSFWorkbook();
            Sheet sheet = wb.createSheet("SLA Report");
            int rowCtr = 0;
            for(List<String> reportRow : reportList) {
            //for(int x=0;x<reportList.size();x++) {
                // Create a row and put some cells in it. Rows are 0 based.
                Row row = sheet.createRow((short)rowCtr++);
                // Create a cell and put a value in it.
                row.createCell(0).setCellValue(reportRow.get(0));
                row.createCell(1).setCellValue(reportRow.get(1));
                row.createCell(2).setCellValue(reportRow.get(2));
                row.createCell(3).setCellValue(reportRow.get(3));
                row.createCell(4).setCellValue(reportRow.get(4));
                row.createCell(5).setCellValue(reportRow.get(5));
                row.createCell(6).setCellValue(reportRow.get(6));
                row.createCell(7).setCellValue(reportRow.get(7));
                row.createCell(8).setCellValue(reportRow.get(8));
                row.createCell(9).setCellValue(reportRow.get(9));
                row.createCell(10).setCellValue(reportRow.get(10));
                row.createCell(11).setCellValue(reportRow.get(11));
                row.createCell(12).setCellValue(reportRow.get(12));
                row.createCell(13).setCellValue(reportRow.get(13));
                row.createCell(14).setCellValue(reportRow.get(14));
                row.createCell(15).setCellValue(reportRow.get(15));
                row.createCell(16).setCellValue(reportRow.get(16));
                row.createCell(17).setCellValue(reportRow.get(17));
            }
            // Write the output to a file
            reportFilename = jTextField2.getText();
            FileOutputStream fileOut = new FileOutputStream(reportFilename);
            wb.write(fileOut);
            wb.close();
            fileOut.close();
            
            // Write ticket counter
            String currentPerson;// = "";
            initTicketCounter();
            
            //if(!teamFilter.equalsIgnoreCase("All") || !personFilter.equalsIgnoreCase("All")) {

            for(List<String> personRow : personList) {
            //for(int y=0;y<personList.size();y++) {
                int totalOntime = 0;
                int totalEarly = 0;
                int totalTickets = 0;

                currentPerson = personRow.get(0);
                if(personFilter.equalsIgnoreCase("All")  || (!personFilter.equalsIgnoreCase("All") && (personFilter.equalsIgnoreCase(currentPerson)))) {
                    for(List<String> reportRow : reportList) {
                    //for(int x=0;x<reportList.size();x++) {
                        if(reportRow.get(5).indexOf(currentPerson)>-1) {
                            totalTickets++;
                            if(reportRow.get(16).startsWith("PASS")) {
                                totalOntime++;
                            }
                            if(reportRow.get(17).startsWith("YES")) {
                                totalEarly++;
                            }
                        }
                    }

                    teamStatItem = new ArrayList<>();
                    teamStatItem.add(personRow.get(1));
                    teamStatItem.add(String.valueOf(totalOntime));
                    teamStatItem.add(String.valueOf(totalEarly));
                    teamStatItem.add(String.valueOf(totalTickets));
                    teamStatList.add(teamStatItem);
                }
                //System.out.println(currentPerson+": "+totalOntime+" - "+totalEarly+" - "+totalTickets);
            //}
                Workbook wb2 = new HSSFWorkbook();
                Sheet sheet2 = wb2.createSheet("Ticket Counter");
                int rowNum=0;
                for(List<String> teamStatRow : teamStatList) {
                //for(int x=0;x<teamStatList.size();x++) {
                    
                    // Create a row and put some cells in it. Rows are 0 based.
                    Row row = sheet2.createRow((short)rowNum++);
                    // Create a cell and put a value in it.
                    row.createCell(0).setCellValue(teamStatRow.get(0));
                    row.createCell(1).setCellValue(teamStatRow.get(1));
                    row.createCell(2).setCellValue(teamStatRow.get(2));
                    row.createCell(3).setCellValue(teamStatRow.get(3));
                    //System.out.println(teamStatList.get(x).get(0)+": "+teamStatList.get(x).get(1)+" "+teamStatList.get(x).get(2)+" "+teamStatList.get(x).get(3));
                }
                // Write the output to a file
                FileOutputStream fileOut2 = new FileOutputStream(ticketCounterFilename);
                wb2.write(fileOut2);
                wb2.close();
                fileOut2.close();
            }
            JOptionPane.showMessageDialog(null, "Report extraction complete!", "Done", JOptionPane.INFORMATION_MESSAGE);
            jButton2.setEnabled(true);
            jButton3.setEnabled(true);
        }
        catch(IOException | ParseException | ClassNotFoundException | SQLException e) {
            JOptionPane.showMessageDialog(null, e, "Error", JOptionPane.ERROR_MESSAGE);
            //e.printStackTrace();
        }
    }

    private void extractKBReport() {
        reportList.clear();
        String url = "jdbc:jtds:sqlserver://"+hostname+":"+port+"/"+databaseName;
        String driver = "net.sourceforge.jtds.jdbc.Driver";
        String user = ""; // Removed for GitHub
        String pass = ""; // Removed for GitHub
        
        String selectedWorkspace = "1";
        if(jComboBox1.getSelectedIndex()==0) {
            selectedWorkspace = "1";
        }
        else if(jComboBox1.getSelectedIndex()==1) {
            selectedWorkspace = "2";
        }
        else if(jComboBox1.getSelectedIndex()==2) {
            selectedWorkspace = "3";
        }
        
        String query = "SELECT a.mrID, a.mrALLDESCRIPTIONS, a.mrAssignees";
            query += " FROM MASTER"+selectedWorkspace+" AS a";
            query += " WHERE (a.mrStatus = 'Resolved' OR a.mrStatus = 'Closed')";
            query += " AND a.Type__bof__bTicket = 'Incident'";
            
            if(!teamFilter.equalsIgnoreCase("All") && !personFilter.equalsIgnoreCase("All")) {
                query += " AND ((a.mrAssignees LIKE '"+teamFilter+"%'";
                query += " OR a.mrAssignees LIKE '%"+teamFilter+"'";
                query += " OR a.mrAssignees LIKE '%"+teamFilter+"%')";
                query += " AND (a.mrAssignees LIKE '"+personFilter+"%'";
                query += " OR a.mrAssignees LIKE '%"+personFilter+"'";
                query += " OR a.mrAssignees LIKE '%"+personFilter+"%'))";
            }
            else if(!teamFilter.equalsIgnoreCase("All")) {
                query += " AND (a.mrAssignees LIKE '"+teamFilter+"%'";
                query += " OR a.mrAssignees LIKE '%"+teamFilter+"'";
                query += " OR a.mrAssignees LIKE '%"+teamFilter+"%')";
            }
            else if(!personFilter.equalsIgnoreCase("All")) {
                query += " AND (a.mrAssignees LIKE '"+personFilter+"%'";
                query += " OR a.mrAssignees LIKE '%"+personFilter+"'";
                query += " OR a.mrAssignees LIKE '%"+personFilter+"%')";
            }
            
            query += " ORDER BY mrID ASC";
         
        try {
            Class.forName(driver);
            Connection conn;
            conn = DriverManager.getConnection(url, user, pass);
            for( SQLWarning warn = conn.getWarnings(); warn != null; warn = warn.getNextWarning() ) {
                    System.out.println( "SQL Warning:" ) ;
                    System.out.println( "State  : " + warn.getSQLState()  ) ;
                    System.out.println( "Message: " + warn.getMessage()   ) ;
                    System.out.println( "Error  : " + warn.getErrorCode() ) ;
            }

            reportItem = new ArrayList<>();
            reportItem.add("Ticket Number");
            reportItem.add("KB Solution Number");
            reportItem.add("Date Linked to Solution");
            reportItem.add("Date Resolved (Ticket)");
            reportItem.add("Team Responsible");
            reportItem.add("Age of KB");
            reportItem.add("Resolved On Time");
            reportItem.add("Resolved Early");
            reportItem.add("Assignees");
            reportList.add(reportItem);
            
            try (Statement stmt = conn.createStatement(); ResultSet rs = stmt.executeQuery( query )) {
                
                while( rs.next() ) {
                    String ticketNumber = rs.getString(1);
                    String kbSolutionNumber = "";
                    String dateLinkedToSolution = "";
                    String dateResolved = "";
                    String teamResponsible = "";
                    String ageOfKB = "";
                    String resolvedOnTime;// = "";
                    String resolvedEarly;// = "";
                    String description = rs.getString(2);
                    
                    String assignees = Utilities.fixString(rs.getString(3));
                    //assignees = assignees.split(" CC:")[0]; // Get only string before CC:
                    assignees = assignees.replaceAll("(\\bCC:)\\S+\\b","").trim();
                    assignees = Utilities.removeOrganization(assignees);                    
                    
                    boolean docsFound = false;
                    if(description.contains("http://itsi.mayniladwater.com.ph/appmgnt/Internal%20Documents/")) {
                        docsFound = true;
                    }
                    //System.out.println("Ticket Number: "+ticketNumber);
                    
                    /* Start of Date Resolved */
                    String query2 = "select mrHISTORY, mrGeneration from MASTER"+selectedWorkspace+"_HISTORY ";
                    query2 += " where mrID = '"+ticketNumber+"'";
                    query2 += " and mrHistory like '%Resolved%'";
                    query2 += " order by mrGENERATION desc";
                    //System.out.println("ticketID: "+ticketID);
                    int mrGeneration = 0;
                    try (Statement stmt3 = conn.createStatement(); ResultSet rs3 = stmt3.executeQuery( query2 )) {
                        while( rs3.next() ) {
                            String dateResolvedString = Utilities.fixString(rs3.getString(1));
                            if(mrGeneration == 0) {
                                mrGeneration = rs3.getInt(2);
                                dateResolved = dateResolvedString.split("\\s+")[0]+" "+dateResolvedString.split("\\s+")[1];
                            }
                            else {
                                if(mrGeneration-1 == rs3.getInt(2)) {
                                    mrGeneration = rs3.getInt(2);
                                    dateResolved = dateResolvedString.split("\\s+")[0]+" "+dateResolvedString.split("\\s+")[1];
                                }
                                else {
                                    break; //Date Resolved already found
                                }
                            }
                        }
                        stmt3.cancel();
                        stmt3.close();
                        rs3.close();
                    }
                    if(dateResolved.isEmpty()) {
                        /* Retry query as Closed since Ticket has no Resolved status */
                        query2 = "select mrHISTORY, mrGeneration from MASTER"+selectedWorkspace+"_HISTORY ";
                        query2 += " where mrID = '"+ticketNumber+"'";
                        query2 += " and mrHistory like '%Closed%'";
                        query2 += " order by mrGENERATION desc";
                        //System.out.println("ticketID: "+ticketID);
                        mrGeneration = 0;
                        try (Statement stmt2 = conn.createStatement(); ResultSet rs2 = stmt2.executeQuery( query2 )) {
                            while( rs2.next() ) {
                                String dateResolvedString = Utilities.fixString(rs2.getString(1));
                                if(mrGeneration == 0) {
                                    mrGeneration = rs2.getInt(2);
                                    dateResolved = dateResolvedString.split("\\s+")[0]+" "+dateResolvedString.split("\\s+")[1];
                                }
                                else {
                                    if(mrGeneration-1 == rs2.getInt(2)) {
                                        mrGeneration = rs2.getInt(2);
                                        dateResolved = dateResolvedString.split("\\s+")[0]+" "+dateResolvedString.split("\\s+")[1];
                                    }
                                    else {
                                        break; //Date Closed already found
                                    }
                                }
                            }
                            stmt2.cancel();
                            stmt2.close();
                            rs2.close();
                        }
                    }
                    /* End of Date Resolved */
                    
                    /* Start of Team Responsible */
                    String query3 = "SELECT mrAssignees";
                        query3 += " FROM MASTER"+selectedWorkspace; // i.e. MASTER1, MASTER2, MASTER3
                        query3 += " WHERE mrID = '"+ticketNumber+"'";
                    
                    //String typeOfTicket = "";
                    try (Statement stmt3 = conn.createStatement(); ResultSet rs3 = stmt3.executeQuery( query3 )) {
                        while( rs3.next() ) {
                            teamResponsible = Utilities.fixString(rs3.getString(1));//.split(" CC:")[0]);
                            //typeOfTicket = Utilities.fixString(rs3.getString(2));
                        }
                        stmt3.cancel();
                        stmt3.close();
                        rs3.close();
                    }
                    String teams = "";
                    if(teamResponsible.contains("Asset Management")) {
                        teams += "Asset Management ";
                    }
                    if(teamResponsible.contains("Business Analyst")) {
                        teams += "Business Analyst ";
                    }
                    if(teamResponsible.contains("Developer")) {
                        teams += "Developer ";
                    }
                    if(teamResponsible.contains("Field Support")) {
                        teams += "Field Support ";
                    }
                    if(teamResponsible.contains("Network and Communications")) {
                        teams += "Network and Communications ";
                    }
                    if(teamResponsible.contains("Operations Management")) {
                        teams += "Operations Management ";
                    }
                    if(teamResponsible.contains("SMO")) {
                        teams += "SMO ";
                    }
                    if(teamResponsible.contains("Service Desk")) {
                        teams += "Service Desk ";
                    }
                    if(teamResponsible.contains("System Administrator")) {
                        teams += "System Administrator ";
                    }
                    if(!teams.isEmpty()) {
                        teamResponsible = teams.trim();
                    }
                    /* End of Team Responsible, Type of Ticket  */
                    
                    /* Start of KB Solution Number, Date Linked to Solution */
                    boolean kbFound = false;
                    boolean rfcFound = false;
                    
                    /*String queryKB = "select mrHISTORY";
                           queryKB += " FROM MASTER"+selectedWorkspace+"_HISTORY"; // i.e. MASTER1_HISTORY, MASTER2_HISTORY, MASTER3_HISTORY
                           queryKB += " WHERE mrID = '"+ticketNumber+"'";
                           queryKB += " AND mrHistory like '%Copied to ticket%'";
                           
                    try (Statement stmtKB = conn.createStatement(); ResultSet rsKB = stmtKB.executeQuery( queryKB )) {
                        while( rsKB.next() ) {
                            kbSolutionNumber = rsKB.getString(1).split("\\s+")[7];
                            kbSolutionNumber = "SOLUTION "+kbSolutionNumber;
                            dateLinkedToSolution = rsKB.getString(1).split("\\s+")[0]+" "+rsKB.getString(1).split("\\s+")[1];
                            kbFoundIn1 = true;
                            
                            if(!dateResolved.isEmpty()) {
                                if((monthFilter.equalsIgnoreCase("All") && yearFilter.equalsIgnoreCase("0")) ||
                                        (monthFilter.equalsIgnoreCase("All") && sdfYear.format(dateFilter).equals(sdfYear.format(sdfTimestamp.parse(dateResolved)))) ||
                                        (!monthFilter.equalsIgnoreCase("All") && sdfMM.format(dateFilter).equals(sdfMM.format(sdfTimestamp.parse(dateResolved))) && sdfYear.format(dateFilter).equals(sdfYear.format(sdfTimestamp.parse(dateResolved)))) ) {
                                    reportItem = new ArrayList<>();
                                    reportItem.add(ticketNumber);
                                    reportItem.add(kbSolutionNumber);
                                    reportItem.add(dateLinkedToSolution);
                                    reportItem.add(dateResolved);
                                    reportItem.add(teamResponsible);
                                    reportItem.add(ageOfKB);
                                    reportList.add(reportItem);
                                }
                            }
                        }
                        stmtKB.cancel();
                        stmtKB.close();
                        rsKB.close();
                    }*/
                    
                    String queryKB = "select mrHISTORY";
                    queryKB += " FROM MASTER"+selectedWorkspace+"_HISTORY"; // i.e. MASTER1_HISTORY, MASTER2_HISTORY, MASTER3_HISTORY
                    queryKB += " WHERE mrID = '"+ticketNumber+"'";
                    queryKB += " AND mrHistory like '%Copied or Selected from add%'";

                    //System.out.println("queryKB: "+queryKB);
                    try (Statement stmtKB = conn.createStatement(); ResultSet rsKB = stmtKB.executeQuery( queryKB )) {
                        while( rsKB.next() ) {
                            String kbSolutionNumberTemp = rsKB.getString(1).split("\\s+")[9];
                            String kbSolutionString = rsKB.getString(1);
                            kbSolutionNumber = "";
                            
                            try {
                                int kbSolutionNumberInt = Integer.parseInt(kbSolutionNumberTemp);
                                String queryKBSolutionOnly = "select mrID";
                                queryKBSolutionOnly += " FROM MASTER"+selectedWorkspace; // i.e. MASTER1_HISTORY, MASTER2_HISTORY, MASTER3_HISTORY
                                queryKBSolutionOnly += " WHERE mrID = '"+kbSolutionNumberInt+"'";
                                queryKBSolutionOnly += " AND mrASSIGNEES = ''";
                                queryKBSolutionOnly += " AND mrURGENT = '0'";

                                try (Statement stmtKBSolutionOnly = conn.createStatement(); ResultSet rsKBSolutionOnly = stmtKBSolutionOnly.executeQuery( queryKBSolutionOnly )) {
                                    while( rsKBSolutionOnly.next() ) {
                                        kbSolutionNumber = rsKBSolutionOnly.getString(1);
                                    }
                                }
                            }
                            catch(Exception e) {
                                kbSolutionNumber = kbSolutionNumberTemp; // Restore original
                            }
                            
                            if(!kbSolutionNumber.isEmpty()) {
                                dateLinkedToSolution = kbSolutionString.split("\\s+")[0]+" "+kbSolutionString.split("\\s+")[1];
                                try {
                                    Integer.parseInt(kbSolutionNumber);
                                    //kbSolutionNumber = "SOLUTION "+kbSolutionNumber;
                                    kbFound = true;

                                    /* Start of Age of KB */
                                    if(!dateResolved.isEmpty() && !dateLinkedToSolution.isEmpty()) {
                                        Date d1 = sdfTimestamp.parse(dateResolved);
                                        Date d2 = sdfTimestamp.parse(dateLinkedToSolution);

                                        long diff = d2.getTime() - d1.getTime();
                                        long diffSeconds = diff / 1000 % 60;
                                        long diffMinutes = diff / (60 * 1000) % 60;
                                        long diffHours = diff / (60 * 60 * 1000) % 24;
                                        long diffDays = diff / (24 * 60 * 60 * 1000);

                                        if(diffSeconds<0) {
                                            diffSeconds = diffSeconds * -1;
                                        }
                                        if(diffMinutes<0) {
                                            diffMinutes = diffMinutes * -1;
                                        }
                                        if(diffHours<0) {
                                            diffHours = diffHours * -1;
                                        }
                                        if(diffDays<0) {
                                            diffDays = diffDays * -1;
                                        }

                                        ageOfKB = diffDays + " days, " + diffHours + " hours, "+ diffMinutes + " minutes, "+diffSeconds + " seconds.";
                                    }
                                    /* End of Age of KB */

                                    if(!dateResolved.isEmpty()) {
                                        if((monthFilter.equalsIgnoreCase("All") && yearFilter.equalsIgnoreCase("0")) ||
                                                (monthFilter.equalsIgnoreCase("All") && sdfYear.format(dateFilter).equals(sdfYear.format(sdfTimestamp.parse(dateResolved)))) ||
                                                (!monthFilter.equalsIgnoreCase("All") && sdfMM.format(dateFilter).equals(sdfMM.format(sdfTimestamp.parse(dateResolved))) && sdfYear.format(dateFilter).equals(sdfYear.format(sdfTimestamp.parse(dateResolved)))) ) {

                                            /* Check if Resolved On Time / Early */
                                            if(isResolvedOnTime(dateResolved, dateLinkedToSolution).equalsIgnoreCase("PASS")) {
                                                resolvedOnTime = "PASS";
                                                resolvedEarly = "NO";
                                            }
                                            else if(isResolvedOnTime(dateResolved, dateLinkedToSolution).equalsIgnoreCase("EARLY")) {
                                                resolvedOnTime = "PASS";
                                                resolvedEarly = "YES";
                                            }
                                            else if(isResolvedOnTime(dateResolved, dateLinkedToSolution).equalsIgnoreCase("FAIL")) {
                                                resolvedOnTime = "FAIL";
                                                resolvedEarly = "NO";
                                            }
                                            else {
                                                resolvedOnTime = "";
                                                resolvedEarly = "";
                                            }
                                            /* End -- Check if Resolved On Time / Early */

                                            reportItem = new ArrayList<>();
                                            reportItem.add(ticketNumber);
                                            reportItem.add(kbSolutionNumber);
                                            reportItem.add(dateLinkedToSolution);
                                            reportItem.add(dateResolved);
                                            reportItem.add(teamResponsible);
                                            reportItem.add(ageOfKB);
                                            reportItem.add(resolvedOnTime);
                                            reportItem.add(resolvedEarly);
                                            reportItem.add(assignees);
                                            reportList.add(reportItem);
                                        }
                                    }

                                }catch(NumberFormatException nfe) {
                                    /*if(kbSolutionNumber.contains("X1")) {
                                        kbSolutionNumber = "SOLUTION "+newKbSolutionNumber;
                                    }
                                    else if(kbSolutionNumber.contains("X2")) {
                                        kbSolutionNumber = "PROBLEM "+newKbSolutionNumber;
                                    }
                                    else if(kbSolutionNumber.contains("X3")) {
                                        kbSolutionNumber = "RFC "+newKbSolutionNumber;
                                    }
                                    else if(kbSolutionNumber.contains("X4")) {
                                        kbSolutionNumber = "SURVEY "+newKbSolutionNumber;
                                    }*/
                                    if(kbSolutionNumber.contains("X3")) {
                                        rfcFound = true;
                                        kbSolutionNumber = kbSolutionNumber.replaceAll("", " ");
                                        //System.out.println(kbSolutionNumber);
                                        String [] chars = kbSolutionNumber.split(" ");
                                        String newKbSolutionNumber = "";
                                        int charCounter = 0;
                                        for(String str:chars ) {
                                            //System.out.println(str);
                                            if(charCounter>1) {
                                                //System.out.println("Greater than 0");
                                                try {
                                                    Integer.parseInt(str);
                                                    newKbSolutionNumber = newKbSolutionNumber.concat(str);
                                                    //System.out.println(newKbSolutionNumber);
                                                }
                                                catch(NumberFormatException nfe2) {
                                                    break;
                                                }
                                            }
                                            charCounter++;
                                        }
                                        kbSolutionNumber = "RFC "+newKbSolutionNumber;
                                        /* Start of Age of RFC */
                                        if(!dateResolved.isEmpty() && !dateLinkedToSolution.isEmpty()) {
                                            Date d1 = sdfTimestamp.parse(dateResolved);
                                            Date d2 = sdfTimestamp.parse(dateLinkedToSolution);

                                            long diff = d2.getTime() - d1.getTime();
                                            long diffSeconds = diff / 1000 % 60;
                                            long diffMinutes = diff / (60 * 1000) % 60;
                                            long diffHours = diff / (60 * 60 * 1000) % 24;
                                            long diffDays = diff / (24 * 60 * 60 * 1000);

                                            if(diffSeconds<0) {
                                                diffSeconds = diffSeconds * -1;
                                            }
                                            if(diffMinutes<0) {
                                                diffMinutes = diffMinutes * -1;
                                            }
                                            if(diffHours<0) {
                                                diffHours = diffHours * -1;
                                            }
                                            if(diffDays<0) {
                                                diffDays = diffDays * -1;
                                            }

                                            ageOfKB = diffDays + " days, " + diffHours + " hours, "+ diffMinutes + " minutes, "+diffSeconds + " seconds.";
                                        }
                                        /* End of Age of RFC */

                                        if(!dateResolved.isEmpty()) {
                                            if((monthFilter.equalsIgnoreCase("All") && yearFilter.equalsIgnoreCase("0")) ||
                                                    (monthFilter.equalsIgnoreCase("All") && sdfYear.format(dateFilter).equals(sdfYear.format(sdfTimestamp.parse(dateResolved)))) ||
                                                    (!monthFilter.equalsIgnoreCase("All") && sdfMM.format(dateFilter).equals(sdfMM.format(sdfTimestamp.parse(dateResolved))) && sdfYear.format(dateFilter).equals(sdfYear.format(sdfTimestamp.parse(dateResolved)))) ) {

                                                /* Check if Resolved On Time / Early */
                                                if(isResolvedOnTime(dateResolved, dateLinkedToSolution).equalsIgnoreCase("PASS")) {
                                                    resolvedOnTime = "PASS";
                                                    resolvedEarly = "NO";
                                                }
                                                else if(isResolvedOnTime(dateResolved, dateLinkedToSolution).equalsIgnoreCase("EARLY")) {
                                                    resolvedOnTime = "PASS";
                                                    resolvedEarly = "YES";
                                                }
                                                else if(isResolvedOnTime(dateResolved, dateLinkedToSolution).equalsIgnoreCase("FAIL")) {
                                                    resolvedOnTime = "FAIL";
                                                    resolvedEarly = "NO";
                                                }
                                                else {
                                                    resolvedOnTime = "";
                                                    resolvedEarly = "";
                                                }
                                                /* End -- Check if Resolved On Time / Early */

                                                reportItem = new ArrayList<>();
                                                reportItem.add(ticketNumber);
                                                reportItem.add(kbSolutionNumber);
                                                reportItem.add(dateLinkedToSolution);
                                                reportItem.add(dateResolved);
                                                reportItem.add(teamResponsible);
                                                reportItem.add(ageOfKB);
                                                reportItem.add(resolvedOnTime);
                                                reportItem.add(resolvedEarly);
                                                reportItem.add(assignees);
                                                reportList.add(reportItem);
                                            }
                                        }
                                    }
                                    else { 
                                        kbSolutionNumber = "";
                                        dateLinkedToSolution = "";
                                    }
                                }
                            }
                        }
                        stmtKB.cancel();
                        stmtKB.close();
                        rsKB.close();
                        
                        if(!kbFound && !rfcFound && docsFound) { // If TA/TP documents found, still output as incident
                            if(!dateResolved.isEmpty()) {
                                if((monthFilter.equalsIgnoreCase("All") && yearFilter.equalsIgnoreCase("0")) ||
                                        (monthFilter.equalsIgnoreCase("All") && sdfYear.format(dateFilter).equals(sdfYear.format(sdfTimestamp.parse(dateResolved)))) ||
                                        (!monthFilter.equalsIgnoreCase("All") && sdfMM.format(dateFilter).equals(sdfMM.format(sdfTimestamp.parse(dateResolved))) && sdfYear.format(dateFilter).equals(sdfYear.format(sdfTimestamp.parse(dateResolved)))) ) {
                                    reportItem = new ArrayList<>();
                                    reportItem.add(ticketNumber);
                                    reportItem.add("TA/TP");
                                    reportItem.add("n/a");
                                    reportItem.add(dateResolved);
                                    reportItem.add(teamResponsible);
                                    reportItem.add("n/a");
                                    reportItem.add("PASS (No link date)");
                                    reportItem.add("NO");
                                    reportItem.add(assignees);
                                    reportList.add(reportItem);
                                }
                            }
                        }
                    }
                    /* End of KB Solution Number, Date Linked to Solution */
                    
                    
                    /*System.out.println("ticketNumber: "+ticketNumber);
                    System.out.println("kbSolutionNumber: "+kbSolutionNumber);
                    System.out.println("dateLinkedToSolution: "+dateLinkedToSolution);
                    System.out.println("dateResolved: "+dateResolved);
                    System.out.println("teamResponsible: "+teamResponsible);
                    System.out.println("ageOfKB: "+ageOfKB);*/
                    /*System.out.println("kbFound: "+kbFound);
                    System.out.println("docsFound: "+docsFound);
                    System.out.println("rfcFound: "+rfcFound);*/
                    if(!kbFound && !docsFound && !rfcFound) { // If no KB, no Docs, and no RFCs were found, still output as Incident
                        /* Start of Age of KB */
                        if(!dateResolved.isEmpty() && !dateLinkedToSolution.isEmpty()) {
                            Date d1 = sdfTimestamp.parse(dateResolved);
                            Date d2 = sdfTimestamp.parse(dateLinkedToSolution);

                            long diff = d2.getTime() - d1.getTime();
                            long diffSeconds = diff / 1000 % 60;
                            long diffMinutes = diff / (60 * 1000) % 60;
                            long diffHours = diff / (60 * 60 * 1000) % 24;
                            long diffDays = diff / (24 * 60 * 60 * 1000);

                            if(diffSeconds<0) {
                                diffSeconds = diffSeconds * -1;
                            }
                            if(diffMinutes<0) {
                                diffMinutes = diffMinutes * -1;
                            }
                            if(diffHours<0) {
                                diffHours = diffHours * -1;
                            }
                            if(diffDays<0) {
                                diffDays = diffDays * -1;
                            }

                            ageOfKB = diffDays + " days, " + diffHours + " hours, "+ diffMinutes + " minutes, "+diffSeconds + " seconds.";
                        }
                        /* End of Age of KB */
                        
                        if(!dateResolved.isEmpty()) {
                            if((monthFilter.equalsIgnoreCase("All") && yearFilter.equalsIgnoreCase("0")) ||
                                    (monthFilter.equalsIgnoreCase("All") && sdfYear.format(dateFilter).equals(sdfYear.format(sdfTimestamp.parse(dateResolved)))) ||
                                    (!monthFilter.equalsIgnoreCase("All") && sdfMM.format(dateFilter).equals(sdfMM.format(sdfTimestamp.parse(dateResolved))) && sdfYear.format(dateFilter).equals(sdfYear.format(sdfTimestamp.parse(dateResolved)))) ) {
                                
                                /* Check if Resolved On Time / Early */
                                if(isResolvedOnTime(dateResolved, dateLinkedToSolution).equalsIgnoreCase("PASS")) {
                                    resolvedOnTime = "PASS";
                                    resolvedEarly = "NO";
                                }
                                else if(isResolvedOnTime(dateResolved, dateLinkedToSolution).equalsIgnoreCase("EARLY")) {
                                    resolvedOnTime = "PASS";
                                    resolvedEarly = "YES";
                                }
                                else if(isResolvedOnTime(dateResolved, dateLinkedToSolution).equalsIgnoreCase("FAIL")) {
                                    resolvedOnTime = "FAIL";
                                    resolvedEarly = "NO";
                                }
                                else {
                                    resolvedOnTime = "";
                                    resolvedEarly = "";
                                }
                                /* End -- Check if Resolved On Time / Early */
                                
                                reportItem = new ArrayList<>();
                                reportItem.add(ticketNumber);
                                reportItem.add(kbSolutionNumber);
                                reportItem.add(dateLinkedToSolution);
                                reportItem.add(dateResolved);
                                reportItem.add(teamResponsible);
                                reportItem.add(ageOfKB);
                                reportItem.add(resolvedOnTime);
                                reportItem.add(resolvedEarly);
                                reportItem.add(assignees);
                                reportList.add(reportItem);
                            }
                        }
                    }
                }
                // Required for MS SQL Server
                stmt.cancel();
                stmt.close();
                rs.close();
            }
            catch(SQLException e) {
                JOptionPane.showMessageDialog(null, e, "Error ", JOptionPane.ERROR_MESSAGE);
            }
            conn.close();
            
            Workbook wb = new HSSFWorkbook();
            Sheet sheet = wb.createSheet("KB Monitoring Report");
            int rowCtr = 0;
            for(List<String> reportRow : reportList) {
            //for(int x=0;x<reportList.size();x++) {
                // Create a row and put some cells in it. Rows are 0 based.
                Row row = sheet.createRow((short)rowCtr++);
                // Create a cell and put a value in it.
                row.createCell(0).setCellValue(reportRow.get(0));
                row.createCell(1).setCellValue(reportRow.get(1));
                row.createCell(2).setCellValue(reportRow.get(2));
                row.createCell(3).setCellValue(reportRow.get(3));
                row.createCell(4).setCellValue(reportRow.get(4));
                row.createCell(5).setCellValue(reportRow.get(5));
                row.createCell(6).setCellValue(reportRow.get(6));
                row.createCell(7).setCellValue(reportRow.get(7));
                row.createCell(8).setCellValue(reportRow.get(8));
            }
            // Write the output to a file
            reportFilename = jTextField2.getText();
            FileOutputStream fileOut = new FileOutputStream(reportFilename);
            wb.write(fileOut);
            wb.close();
            fileOut.close();
            
            // Write ticket counter
            String currentPerson;// = "";
            initTicketCounter();
            
            //if(!teamFilter.equalsIgnoreCase("All") || !personFilter.equalsIgnoreCase("All")) {
              
            for(List<String> personRow : personList) {
            //for(int y=0;y<personList.size();y++) {
                int totalOntime = 0;
                int totalEarly = 0;
                int totalTickets = 0;

                currentPerson = personRow.get(0);
                if(personFilter.equalsIgnoreCase("All")  || (!personFilter.equalsIgnoreCase("All") && (personFilter.equalsIgnoreCase(currentPerson)))) {
                    for(List<String> reportRow : reportList) {
                    //for(int x=0;x<reportList.size();x++) {
                        if(reportRow.get(8).indexOf(currentPerson)>-1) {
                            totalTickets++;
                            if(reportRow.get(6).startsWith("PASS")) {
                                totalOntime++;
                            }
                            if(reportRow.get(7).startsWith("YES")) {
                                totalEarly++;
                            }
                        }
                    }

                    teamStatItem = new ArrayList<>();
                    teamStatItem.add(personRow.get(1));
                    teamStatItem.add(String.valueOf(totalOntime));
                    teamStatItem.add(String.valueOf(totalEarly));
                    teamStatItem.add(String.valueOf(totalTickets));
                    teamStatList.add(teamStatItem);
                }
                //System.out.println(currentPerson+": "+totalOntime+" - "+totalEarly+" - "+totalTickets);
            //}
                Workbook wb2 = new HSSFWorkbook();
                Sheet sheet2 = wb2.createSheet("Ticket Counter");
                int rowNum=0;
                for(List<String> teamStatRow : teamStatList) {
                //for(int x=0;x<teamStatList.size();x++) {
                    // Create a row and put some cells in it. Rows are 0 based.
                    Row row = sheet2.createRow((short)rowNum++);
                    // Create a cell and put a value in it.
                    row.createCell(0).setCellValue(teamStatRow.get(0));
                    row.createCell(1).setCellValue(teamStatRow.get(1));
                    row.createCell(2).setCellValue(teamStatRow.get(2));
                    row.createCell(3).setCellValue(teamStatRow.get(3));
                    //System.out.println(teamStatList.get(x).get(0)+": "+teamStatList.get(x).get(1)+" "+teamStatList.get(x).get(2)+" "+teamStatList.get(x).get(3));
                }
                // Write the output to a file
                FileOutputStream fileOut2 = new FileOutputStream(ticketCounterFilename);
                wb2.write(fileOut2);
                wb2.close();
                fileOut2.close();
            }
            
            JOptionPane.showMessageDialog(null, "Report extraction complete!", "Done", JOptionPane.INFORMATION_MESSAGE);
            jButton2.setEnabled(true);
            jButton3.setEnabled(true);
        }
        catch(IOException | ParseException | ClassNotFoundException | SQLException e) {
            JOptionPane.showMessageDialog(null, e, "Error", JOptionPane.ERROR_MESSAGE);
            //e.printStackTrace();
        }
    }
    
    private String isResolvedOnTime(String dateResolved, String dateLinked) {
        /*
         * Returns: PASS  = If date linked is on or before resolved date
         *          EARLY = If date linked is 2 hours early than resolved date
         *          FAIL  = If date linked is after resolved date
         *          FAIL  = If date linked is empty
         */
         
        if(dateLinked.isEmpty()) {
            return "FAIL";
        }
        /* Start of RESOLVED on Time / Early */
        try {
            Calendar calDateResolved = Calendar.getInstance();
            Calendar calDateLinked = Calendar.getInstance();
            Calendar calDateResolved_DateOnly = Calendar.getInstance();
            Calendar calDateLinked_DateOnly = Calendar.getInstance();

            Calendar calDateLinked_Plus;// = Calendar.getInstance();


            calDateResolved.setTime(sdfTimestamp.parse(dateResolved));
            calDateLinked.setTime(sdfTimestamp.parse(dateLinked));

            calDateResolved_DateOnly.setTime(sdfDatestamp.parse(dateResolved));
            calDateLinked_DateOnly.setTime(sdfDatestamp.parse(dateLinked));

            if(calDateResolved_DateOnly.before(calDateLinked_DateOnly)) {
                return "FAIL";
            }
            else {
                // 2 hours earlier than resolution date
                calDateLinked_Plus = calDateLinked;
                calDateLinked_Plus.add(Calendar.HOUR_OF_DAY, 2);
                if(calDateResolved.before(calDateLinked_Plus)) {
                    return "PASS";
                }
                else {
                    return "EARLY";
                }
            }
        }
        catch(ParseException pe) {
            return "PASS";
        }
    }
    
    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(MainWindow.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(MainWindow.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(MainWindow.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(MainWindow.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Set the look and feel */
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | UnsupportedLookAndFeelException e) {
        }
        
        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new MainWindow().setVisible(true);
            }
        });
    }
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.ButtonGroup buttonGroup1;
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton3;
    private javax.swing.JComboBox jComboBox1;
    private javax.swing.JComboBox jComboBox2;
    private javax.swing.JComboBox jComboBox3;
    private javax.swing.JComboBox jComboBox4;
    private javax.swing.JComboBox jComboBox5;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JLabel jLabel8;
    private javax.swing.JMenu jMenu1;
    private javax.swing.JMenu jMenu2;
    private javax.swing.JMenuBar jMenuBar1;
    private javax.swing.JMenuItem jMenuItem1;
    private javax.swing.JMenuItem jMenuItem2;
    private javax.swing.JMenuItem jMenuItem3;
    private javax.swing.JMenuItem jMenuItem4;
    private javax.swing.JRadioButton jRadioButton1;
    private javax.swing.JRadioButton jRadioButton2;
    private javax.swing.JTextField jTextField1;
    private javax.swing.JTextField jTextField2;
    // End of variables declaration//GEN-END:variables
}
