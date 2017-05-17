package servicecloud;

import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.HeadlessException;
import java.awt.Image;
import java.awt.Point;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.io.Writer;
import java.net.URI;
import java.net.URISyntaxException;
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
import java.util.Collections;
import java.util.Comparator;
import java.util.HashSet;
import java.util.List;
import java.util.Properties;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.embed.swing.JFXPanel;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPopupMenu;
import javax.swing.JTable;
import javax.swing.ListSelectionModel;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.event.TableModelEvent;
import javax.swing.event.TableModelListener;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableColumn;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringEscapeUtils;
import org.apache.commons.lang3.StringUtils;

/**
 *
 * @author Glenn Dimaliwat
 */
public class Main extends javax.swing.JFrame {

    private static final String appVersion = "3.1";
    private static final String binary = "\\\\mayniladfs\\IT APPS MGT\\AS File Server\\ServiceCloud\\ServiceCloud.rar";
    private static final String source = "\\\\mayniladfs\\IT APPS MGT\\AS File Server\\ServiceCloud\\Updates";
    private static final String destination = System.getProperty("user.dir");
    private static final String latestVersionFile = source+"\\version.txt";
    private static final String localVersionFile = destination+"\\version.txt";
    //private boolean updateServerResolved = false;
    //private boolean updateSuccessful = false;
    
    private static DefaultTableModel tableModel;
    private DefaultTableModel model;
    private static final String hostname = "172.18.2.83";
    private static final String port = "1433";        
    private static final String databaseName = "Footprints";
    private String selectedTicketNumber = "";
    private String selectedTicketNumberModifiedDate = "";
    private String selectedTicketNumberResolvedDate = "";
    private String selectedSolutionNumber = "";
    private String searchKey = "";
    private String lastSearchKey = "";
    private boolean repeatSearch = false;
    private String sortOrder = "0 ASC";
    private String sourceValue = "";
    private int sourceSelection = 0;
    private boolean sortDisabled = false;
    
    //Initialize 2D Array Lists
    private List<List<String>> sqlData = new ArrayList<>();
    private List<List<String>> data = new ArrayList<>();
    
    //Initialize other Windows
    Comments frameComments = new Comments();
    CommentsHTML frameCommentsHTML = new CommentsHTML();
    
    private static final SimpleDateFormat sdfYear = new SimpleDateFormat("yyyy");
    private static final SimpleDateFormat sdfMonth = new SimpleDateFormat("MMMM");
    private static final SimpleDateFormat sdfMM = new SimpleDateFormat("MM");
    private static final SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
    //private static final SimpleDateFormat sdfTimestamp = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
    private static final SimpleDateFormat sdfWords = new SimpleDateFormat("MMMMM dd, yyyy");
    
    private List<List<String>> ticketsList = new ArrayList<>();
    private List<List<String>> progressList = new ArrayList<>();
    //private List<List<String>> commentsList = new ArrayList<>();
    private String conditionMaster1 = "";
    private String conditionMaster2 = "";
    private String conditionMaster3 = "";
    private String conditionSearch = "";
    private boolean errorAlreadyEncountered = false;
    private int ticketCount = 0;
    private int incidentCount = 0;
    
    final int columnTicketSummary = 0;
    final int columnProgress = 1;
    final int columnCreated = 2;
    final int columnDueDate = 3;
    final int columnCategory = 4;
    final int columnRequestor = 5;
    final int columnTicket = 6;
    final int columnDescription = 7;
    final int columnStatus = 8;
    final int columnRemarks = 9;
    final int columnClassification = 10;
    final int columnAssignees = 11;
    final int columnModified = 12;
    final int columnResolved = 13;
    final int columnKB = 14;
    final int columnDateKB = 15;
    final int columnTCode = 16;
    
    /**
     * Creates new form Main
     */
    public Main() {
        initComponents();
        initCustomComponents();
        initTicketsFile();
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jScrollPane1 = new javax.swing.JScrollPane();
        jTable1 = new javax.swing.JTable();
        jPanel1 = new javax.swing.JPanel();
        jButton1 = new javax.swing.JButton();
        jToggleButton1 = new javax.swing.JToggleButton();
        jButton3 = new javax.swing.JButton();
        jButton6 = new javax.swing.JButton();
        jButton5 = new javax.swing.JButton();
        jButton7 = new javax.swing.JButton();
        jComboBox2 = new javax.swing.JComboBox();
        jLabel1 = new javax.swing.JLabel();
        jTextField1 = new javax.swing.JTextField();
        jButton4 = new javax.swing.JButton();
        jButton2 = new javax.swing.JButton();
        jPanel2 = new javax.swing.JPanel();
        jStatusBar = new javax.swing.JLabel();
        jMenuBar1 = new javax.swing.JMenuBar();
        jMenu2 = new javax.swing.JMenu();
        jMenuItem8 = new javax.swing.JMenuItem();
        jMenuItem9 = new javax.swing.JMenuItem();
        jMenuItem6 = new javax.swing.JMenuItem();
        jMenuItem7 = new javax.swing.JMenuItem();
        jMenuItem1 = new javax.swing.JMenuItem();
        jMenu1 = new javax.swing.JMenu();
        jMenuItem4 = new javax.swing.JMenuItem();
        jMenuItem5 = new javax.swing.JMenuItem();
        jMenu4 = new javax.swing.JMenu();
        jMenuItem12 = new javax.swing.JMenuItem();
        jMenuItem11 = new javax.swing.JMenuItem();
        jMenu3 = new javax.swing.JMenu();
        jMenuItem3 = new javax.swing.JMenuItem();
        jMenuItem10 = new javax.swing.JMenuItem();
        jMenuItem2 = new javax.swing.JMenuItem();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Service Cloud");
        setIconImage(getIconImage());

        jScrollPane1.setBackground(new java.awt.Color(255, 255, 255));

        jTable1.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null}
            },
            new String [] {
                "Title 1", "Title 2", "Title 3", "Title 4"
            }
        ));
        jTable1.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jTable1MouseClicked(evt);
            }
            public void mousePressed(java.awt.event.MouseEvent evt) {
                jTable1MousePressed(evt);
            }
        });
        jTable1.addKeyListener(new java.awt.event.KeyAdapter() {
            public void keyReleased(java.awt.event.KeyEvent evt) {
                jTable1KeyReleased(evt);
            }
        });
        jScrollPane1.setViewportView(jTable1);

        jButton1.setLabel("Refresh");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });

        jToggleButton1.setText("Cell mode");
        jToggleButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jToggleButton1ActionPerformed(evt);
            }
        });

        jButton3.setText("Go to Ticket");
        jButton3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton3ActionPerformed(evt);
            }
        });

        jButton6.setText("Go to Solution");
        jButton6.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton6ActionPerformed(evt);
            }
        });

        jButton5.setText("Go to Ticket Docs");
        jButton5.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton5ActionPerformed(evt);
            }
        });

        jButton7.setText("Go to KB Docs");
        jButton7.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton7ActionPerformed(evt);
            }
        });

        jComboBox2.setModel(new javax.swing.DefaultComboBoxModel(new String[] { "Tickets File", "Username or Group" }));
        jComboBox2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jComboBox2ActionPerformed(evt);
            }
        });

        jLabel1.setFont(new java.awt.Font("Tahoma", 1, 12)); // NOI18N
        jLabel1.setText("Data:");

        jButton4.setText("Save");
        jButton4.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton4ActionPerformed(evt);
            }
        });

        jButton2.setText("Edit File");
        jButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton2ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addComponent(jButton1)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jToggleButton1, javax.swing.GroupLayout.PREFERRED_SIZE, 79, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jButton3)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jButton6)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jButton5)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jButton7)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 111, Short.MAX_VALUE)
                .addComponent(jLabel1)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jComboBox2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, 146, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jButton4)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jButton2))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                .addComponent(jButton2)
                .addComponent(jButton4)
                .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addComponent(jComboBox2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addComponent(jLabel1))
            .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                .addComponent(jButton1)
                .addComponent(jToggleButton1)
                .addComponent(jButton3)
                .addComponent(jButton6)
                .addComponent(jButton5)
                .addComponent(jButton7))
        );

        jPanel2.setBorder(javax.swing.BorderFactory.createEmptyBorder(1, 1, 1, 1));

        jStatusBar.setFont(new java.awt.Font("Tahoma", 1, 11)); // NOI18N
        jStatusBar.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jStatusBar.setText("TADASHI");

        javax.swing.GroupLayout jPanel2Layout = new javax.swing.GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jStatusBar, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, 1171, Short.MAX_VALUE)
        );
        jPanel2Layout.setVerticalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jStatusBar)
        );

        jMenu2.setMnemonic('F');
        jMenu2.setText("File");

        jMenuItem8.setMnemonic('A');
        jMenuItem8.setText("Add Ticket");
        jMenuItem8.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem8ActionPerformed(evt);
            }
        });
        jMenu2.add(jMenuItem8);

        jMenuItem9.setMnemonic('R');
        jMenuItem9.setText("Remove Ticket");
        jMenuItem9.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem9ActionPerformed(evt);
            }
        });
        jMenu2.add(jMenuItem9);

        jMenuItem6.setAccelerator(javax.swing.KeyStroke.getKeyStroke(java.awt.event.KeyEvent.VK_F5, 0));
        jMenuItem6.setMnemonic('f');
        jMenuItem6.setLabel("Refresh");
        jMenuItem6.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem6ActionPerformed(evt);
            }
        });
        jMenu2.add(jMenuItem6);

        jMenuItem7.setAccelerator(javax.swing.KeyStroke.getKeyStroke(java.awt.event.KeyEvent.VK_Y, java.awt.event.InputEvent.CTRL_MASK));
        jMenuItem7.setMnemonic('C');
        jMenuItem7.setText("Cell mode");
        jMenuItem7.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem7ActionPerformed(evt);
            }
        });
        jMenu2.add(jMenuItem7);

        jMenuItem1.setMnemonic('x');
        jMenuItem1.setText("Exit");
        jMenuItem1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem1ActionPerformed(evt);
            }
        });
        jMenu2.add(jMenuItem1);

        jMenuBar1.add(jMenu2);

        jMenu1.setMnemonic('S');
        jMenu1.setText("Search");

        jMenuItem4.setAccelerator(javax.swing.KeyStroke.getKeyStroke(java.awt.event.KeyEvent.VK_F, java.awt.event.InputEvent.CTRL_MASK));
        jMenuItem4.setMnemonic('F');
        jMenuItem4.setText("Find");
        jMenuItem4.setToolTipText("");
        jMenuItem4.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem4ActionPerformed(evt);
            }
        });
        jMenu1.add(jMenuItem4);

        jMenuItem5.setAccelerator(javax.swing.KeyStroke.getKeyStroke(java.awt.event.KeyEvent.VK_F3, 0));
        jMenuItem5.setText("Find Next");
        jMenuItem5.setToolTipText("N");
        jMenuItem5.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem5ActionPerformed(evt);
            }
        });
        jMenu1.add(jMenuItem5);

        jMenuBar1.add(jMenu1);

        jMenu4.setMnemonic('T');
        jMenu4.setText("Tools");

        jMenuItem12.setMnemonic('G');
        jMenuItem12.setText("Go to a Ticket Page");
        jMenuItem12.setToolTipText("Specify a ticket number and go to its page");
        jMenuItem12.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem12ActionPerformed(evt);
            }
        });
        jMenu4.add(jMenuItem12);

        jMenuItem11.setMnemonic('C');
        jMenuItem11.setText("Clean Progress File");
        jMenuItem11.setToolTipText("Removes deleted tickets from the Progress File");
        jMenuItem11.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem11ActionPerformed(evt);
            }
        });
        jMenu4.add(jMenuItem11);

        jMenuBar1.add(jMenu4);

        jMenu3.setMnemonic('H');
        jMenu3.setText("Help");
        jMenu3.setToolTipText("");

        jMenuItem3.setMnemonic('U');
        jMenuItem3.setText("How to use this App");
        jMenuItem3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem3ActionPerformed(evt);
            }
        });
        jMenu3.add(jMenuItem3);

        jMenuItem10.setText("Check for Updates");
        jMenuItem10.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem10ActionPerformed(evt);
            }
        });
        jMenu3.add(jMenuItem10);

        jMenuItem2.setMnemonic('b');
        jMenuItem2.setText("About");
        jMenuItem2.setToolTipText("");
        jMenuItem2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem2ActionPerformed(evt);
            }
        });
        jMenu3.add(jMenuItem2);

        jMenuBar1.add(jMenu3);

        setJMenuBar(jMenuBar1);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jScrollPane1)
                    .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 493, Short.MAX_VALUE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jPanel2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void initCustomComponents() {
        JFXPanel fxPanel = new JFXPanel(); // Needed to initialize JavaFX Objects
        
        // Status Bar
        jStatusBar.setText("Hello "+System.getProperty("user.name")+"!"+" Today is "+sdfWords.format(Calendar.getInstance().getTime()));
        
        // Version Info
        saveVersionInfo();
        updateApplication(true);
        
        // Table Model
        tableModel = new DefaultTableModel(new Object[] { "TICKET SUMMARY", "PROGRESS", "CREATED", "DUE DATE", "CATEGORY", "REQUESTOR", "TICKET", "DESCRIPTION", "STATUS", "REMARKS", "CLASSIFICATION", "ASSIGNEES", "MODIFIED", "RESOLVED", "KB", "DATE LINKED TO KB", "T-CODE" }, 0);
        jTable1.setModel(tableModel);  
        jTable1.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        jTable1.getTableHeader().setReorderingAllowed(false);
        jTable1.getTableHeader().addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
              
              if(sortDisabled) {
                  int confirmation = JOptionPane.showConfirmDialog (null, "You've recently updated your progress. Please press OK to save your progress.\nYou will not be allowed to sort before save your progress or until you press Refresh.", "Save Progress?", JOptionPane.OK_CANCEL_OPTION);
        
                  if(confirmation == JOptionPane.OK_OPTION) {
                      sortDisabled = false;
                      refresh();
                  }
                  else {
                      sortDisabled = true;
                  }
              }
              else {
                jTable1.getColumnModel().getColumn(columnTicketSummary).setHeaderValue("TICKET SUMMARY");
                jTable1.getColumnModel().getColumn(columnProgress).setHeaderValue("PROGRESS");
                jTable1.getColumnModel().getColumn(columnCreated).setHeaderValue("CREATED");
                jTable1.getColumnModel().getColumn(columnDueDate).setHeaderValue("DUE DATE");
                jTable1.getColumnModel().getColumn(columnCategory).setHeaderValue("CATEGORY");
                jTable1.getColumnModel().getColumn(columnRequestor).setHeaderValue("REQUESTOR");
                jTable1.getColumnModel().getColumn(columnTicket).setHeaderValue("TICKET");
                jTable1.getColumnModel().getColumn(columnDescription).setHeaderValue("DESCRIPTION");         
                jTable1.getColumnModel().getColumn(columnStatus).setHeaderValue("STATUS");
                jTable1.getColumnModel().getColumn(columnRemarks).setHeaderValue("REMARKS");
                jTable1.getColumnModel().getColumn(columnClassification).setHeaderValue("CLASSIFICATION");    
                jTable1.getColumnModel().getColumn(columnAssignees).setHeaderValue("ASSIGNEES");
                jTable1.getColumnModel().getColumn(columnModified).setHeaderValue("MODIFIED");
                jTable1.getColumnModel().getColumn(columnResolved).setHeaderValue("RESOLVED");
                jTable1.getColumnModel().getColumn(columnKB).setHeaderValue("KB");
                jTable1.getColumnModel().getColumn(columnDateKB).setHeaderValue("DATE LINKED TO KB");
                jTable1.getColumnModel().getColumn(columnTCode).setHeaderValue("T-CODE");

                int col = jTable1.columnAtPoint(new Point(e.getX(), e.getY()));
                if(jTable1.getRowCount()>0) {
                  switch(col) {
                      case columnTicketSummary:
                          if(sortOrder.equalsIgnoreCase(col+" ASC")) {
                              Collections.sort(data, new ArrayList2DComparator(col,"DESC"));
                              sortOrder = col+" DESC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▼ TICKET SUMMARY");
                          }
                          else {
                              Collections.sort(data, new ArrayList2DComparator(col,"ASC"));
                              sortOrder = col+" ASC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▲ TICKET SUMMARY");
                          }
                          break;
                      case columnProgress:
                          if(sortOrder.equalsIgnoreCase(col+" ASC")) {
                              Collections.sort(data, new ArrayList2DComparator(col,"DESC"));
                              sortOrder = col+" DESC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▼ PROGRESS");
                          }
                          else {
                              Collections.sort(data, new ArrayList2DComparator(col,"ASC"));
                              sortOrder = col+" ASC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▲ PROGRESS");
                          }
                          break;
                      case columnCreated:
                          if(sortOrder.equalsIgnoreCase(col+" ASC")) {
                              Collections.sort(data, new ArrayList2DComparator(col,"DESC"));
                              sortOrder = col+" DESC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▼ CREATED");
                          }
                          else {
                              Collections.sort(data, new ArrayList2DComparator(col,"ASC"));
                              sortOrder = col+" ASC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▲ CREATED");
                          }
                          break;
                      case columnDueDate: 
                          if(sortOrder.equalsIgnoreCase(col+" ASC")) {
                              Collections.sort(data, new ArrayList2DComparator(col,"DESC"));
                              sortOrder = col+" DESC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▼ DUE DATE");
                          }
                          else {
                              Collections.sort(data, new ArrayList2DComparator(col,"ASC"));
                              sortOrder = col+" ASC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▲ DUE DATE");
                          }
                          break;
                      case columnCategory: 
                          if(sortOrder.equalsIgnoreCase("4 ASC")) {
                              Collections.sort(data, new ArrayList2DComparator(col,"DESC"));
                              sortOrder = col+" DESC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▼ CATEGORY");
                          }
                          else {
                              Collections.sort(data, new ArrayList2DComparator(col,"ASC"));
                              sortOrder = col+" ASC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▲ CATEGORY");
                          }
                          break;
                      case columnRequestor: 
                          if(sortOrder.equalsIgnoreCase("5 ASC")) {
                              Collections.sort(data, new ArrayList2DComparator(col,"DESC"));
                              sortOrder = col+" DESC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▼ REQUESTOR");
                          }
                          else {
                              Collections.sort(data, new ArrayList2DComparator(col,"ASC"));
                              sortOrder = col+" ASC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▲ REQUESTOR");
                          }
                          break;
                      case columnTicket: 
                          if(sortOrder.equalsIgnoreCase("6 ASC")) {
                              Collections.sort(data, new ArrayList2DComparator(col,"DESC"));
                              sortOrder = col+" DESC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▼ TICKET");
                          }
                          else {
                              Collections.sort(data, new ArrayList2DComparator(col,"ASC"));
                              sortOrder = col+" ASC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▲ TICKET");
                          }
                          break;
                      case columnDescription: 
                          if(sortOrder.equalsIgnoreCase("7 ASC")) {
                              Collections.sort(data, new ArrayList2DComparator(col,"DESC"));
                              sortOrder = col+" DESC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▼ DESCRIPTION");
                          }
                          else {
                              Collections.sort(data, new ArrayList2DComparator(col,"ASC"));
                              sortOrder = col+" ASC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▲ DESCRIPTION");
                          }
                          break;
                      case columnStatus: 
                          if(sortOrder.equalsIgnoreCase("8 ASC")) {
                              Collections.sort(data, new ArrayList2DComparator(col,"DESC"));
                              sortOrder = col+" DESC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▼ STATUS");
                          }
                          else {
                              Collections.sort(data, new ArrayList2DComparator(col,"ASC"));
                              sortOrder = col+" ASC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▲ STATUS");
                          }
                          break;
                      case columnRemarks: 
                          if(sortOrder.equalsIgnoreCase("9 ASC")) {
                              Collections.sort(data, new ArrayList2DComparator(col,"DESC"));
                              sortOrder = col+" DESC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▼ REMARKS");
                          }
                          else {
                              Collections.sort(data, new ArrayList2DComparator(col,"ASC"));
                              sortOrder = col+" ASC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▲ REMARKS");
                          }
                          break;
                      case columnClassification:
                          if(sortOrder.equalsIgnoreCase("10 ASC")) {
                              Collections.sort(data, new ArrayList2DComparator(col,"DESC"));
                              sortOrder = col+" DESC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▼ CLASSIFICATION");
                          }
                          else {
                              Collections.sort(data, new ArrayList2DComparator(col,"ASC"));
                              sortOrder = col+" ASC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▲ CLASSIFICATION");
                          }
                          break;
                      case columnAssignees:
                          if(sortOrder.equalsIgnoreCase("11 ASC")) {
                              Collections.sort(data, new ArrayList2DComparator(col,"DESC"));
                              sortOrder = col+" DESC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▼ ASSIGNEES");
                          }
                          else {
                              Collections.sort(data, new ArrayList2DComparator(col,"ASC"));
                              sortOrder = col+" ASC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▲ ASSIGNEES");
                          }
                          break;
                      case columnModified:
                          if(sortOrder.equalsIgnoreCase("12 ASC")) {
                              Collections.sort(data, new ArrayList2DComparator(col,"DESC"));
                              sortOrder = col+" DESC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▼ MODIFIED");
                          }
                          else {
                              Collections.sort(data, new ArrayList2DComparator(col,"ASC"));
                              sortOrder = col+" ASC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▲ MODIFIED");
                          }
                          break;
                      case columnResolved:
                          if(sortOrder.equalsIgnoreCase("13 ASC")) {
                              Collections.sort(data, new ArrayList2DComparator(col,"DESC"));
                              sortOrder = col+" DESC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▼ RESOLVED");
                          }
                          else {
                              Collections.sort(data, new ArrayList2DComparator(col,"ASC"));
                              sortOrder = col+" ASC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▲ RESOLVED");
                          }
                          break;
                      case columnKB:
                          if(sortOrder.equalsIgnoreCase("14 ASC")) {
                              Collections.sort(data, new ArrayList2DComparator(col,"DESC"));
                              sortOrder = col+" DESC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▼ KB");
                          }
                          else {
                              Collections.sort(data, new ArrayList2DComparator(col,"ASC"));
                              sortOrder = col+" ASC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▲ KB");
                          }
                          break;
                      case columnDateKB:
                          if(sortOrder.equalsIgnoreCase("15 ASC")) {
                              Collections.sort(data, new ArrayList2DComparator(col,"DESC"));
                              sortOrder = col+" DESC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▼ DATE LINKED TO KB");
                          }
                          else {
                              Collections.sort(data, new ArrayList2DComparator(col,"ASC"));
                              sortOrder = col+" ASC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▲ DATE LINKED TO KB");
                          }
                          break;
                      case columnTCode:
                          if(sortOrder.equalsIgnoreCase("16 ASC")) {
                              Collections.sort(data, new ArrayList2DComparator(col,"DESC"));
                              sortOrder = col+" DESC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▼ T-CODE");
                          }
                          else {
                              Collections.sort(data, new ArrayList2DComparator(col,"ASC"));
                              sortOrder = col+" ASC";
                              jTable1.getColumnModel().getColumn(col).setHeaderValue("▲ T-CODE");
                          }
                          break;
                  }
                  jTable1.getTableHeader().resizeAndRepaint();
                  loadTableContents();
                }
              }
            };
        });
        
        jTable1.getModel().addTableModelListener(new TableModelListener() {
            @Override
            public void tableChanged(TableModelEvent e) {
                //System.out.println("Row: "+e.getFirstRow()+" Column: "+e.getColumn());
                if(e.getColumn()==columnProgress) {
                    //System.out.println(jTable1.getModel().getValueAt(e.getFirstRow(), columnProgress));
                    //System.out.println(jTable1.getModel().getValueAt(e.getFirstRow(), columnTicket));
                    String ticketNumber = jTable1.getModel().getValueAt(e.getFirstRow(), columnTicket).toString();
                    String progressData = jTable1.getModel().getValueAt(e.getFirstRow(), columnProgress).toString();
                    
                    if(!progressData.isEmpty()) {
                        progressData = progressData.trim();
                    }
                    if(!progressData.isEmpty()) {
                        File progressFile = new File("Progress.txt");
                        File tempFile = new File("Progress.tmp");
                        try {
                            BufferedReader readFileExists = new BufferedReader(new FileReader(progressFile));
                            readFileExists.close();
                        }
                        catch(IOException ioe1) {
                            try {
                                BufferedWriter writeFileExists = new BufferedWriter(new FileWriter(progressFile));
                                writeFileExists.close();
                            }
                            catch(IOException ioe2) {}
                        }
                                
                        try {
                            try (BufferedReader reader = new BufferedReader(new FileReader(progressFile)); BufferedWriter writer = new BufferedWriter(new FileWriter(tempFile))) {
                                
                                String currentLine;
                                while((currentLine = reader.readLine()) != null) {
                                    String trimmedLine = currentLine.trim();
                                    if(trimmedLine.indexOf(ticketNumber)>=0) {
                                        continue;
                                    }
                                    else {
                                        writer.write(currentLine + System.getProperty("line.separator"));
                                    }
                                }
                                
                                reader.close();
                                writer.close();
                            }
                            FileUtils.deleteQuietly(progressFile);
                            FileUtils.moveFile(tempFile, progressFile);
                            
                            Writer output;
                            output = new BufferedWriter(new FileWriter(progressFile, true));
                            output.append(ticketNumber+" "+progressData+System.getProperty("line.separator"));
                            output.close();
                        }catch(IOException ioe3) {}
                        
                        // Disable table sorting temporarily
                        sortDisabled = true;
                    }
                }
            }
        });
        
        // set preferred column widths
        TableColumn a;
        
        final int widthTicketSummary = 110;
        final int widthProgress = 70;
        final int widthCreated = 70;
        final int widthDueDate = 70;
        final int widthCategory = 120;
        final int widthRequestor = 90;
        final int widthTicket = 80;
        final int widthDescription = 200;        
        final int widthStatus = 90;
        final int widthRemarks = 125;
        final int widthClassification = 110;
        final int widthAssignees = 100;
        final int widthModified = 70;
        final int widthResolved = 70;
        final int widthKB = 90;
        final int widthDateKB = 70;
        final int widthTCode = 100;
        
        a = jTable1.getColumnModel().getColumn(columnTicketSummary);
        a.setPreferredWidth(widthTicketSummary);
        a.setCellRenderer(new CustomTableCellRenderer());
        a = jTable1.getColumnModel().getColumn(columnProgress);
        a.setPreferredWidth(widthProgress);
        a.setCellRenderer(new CustomTableCellRenderer());
        a = jTable1.getColumnModel().getColumn(columnCreated);
        a.setPreferredWidth(widthCreated);
        a.setCellRenderer(new CustomTableCellRenderer());
        a = jTable1.getColumnModel().getColumn(columnDueDate);
        a.setPreferredWidth(widthDueDate);
        a.setCellRenderer(new CustomTableCellRenderer());
        a = jTable1.getColumnModel().getColumn(columnCategory);
        a.setPreferredWidth(widthCategory);
        a.setCellRenderer(new CustomTableCellRenderer());
        a = jTable1.getColumnModel().getColumn(columnRequestor);
        a.setPreferredWidth(widthRequestor);
        a.setCellRenderer(new CustomTableCellRenderer());
        a = jTable1.getColumnModel().getColumn(columnTicket);
        a.setPreferredWidth(widthTicket);
        a.setCellRenderer(new CustomTableCellRenderer());
        a = jTable1.getColumnModel().getColumn(columnDescription);
        a.setPreferredWidth(widthDescription);
        a.setCellRenderer(new CustomTableCellRenderer());
        a = jTable1.getColumnModel().getColumn(columnStatus);
        a.setPreferredWidth(widthStatus);
        a.setCellRenderer(new CustomTableCellRenderer());
        a = jTable1.getColumnModel().getColumn(columnRemarks);
        a.setPreferredWidth(widthRemarks);
        a.setCellRenderer(new CustomTableCellRenderer());
        a = jTable1.getColumnModel().getColumn(columnAssignees);
        a.setPreferredWidth(widthAssignees);
        a.setCellRenderer(new CustomTableCellRenderer());
        a = jTable1.getColumnModel().getColumn(columnClassification);
        a.setPreferredWidth(widthClassification);
        a.setCellRenderer(new CustomTableCellRenderer());      
        a = jTable1.getColumnModel().getColumn(columnModified);
        a.setPreferredWidth(widthModified);
        a.setCellRenderer(new CustomTableCellRenderer());
        a = jTable1.getColumnModel().getColumn(columnResolved);
        a.setPreferredWidth(widthResolved);
        a.setCellRenderer(new CustomTableCellRenderer());
        a = jTable1.getColumnModel().getColumn(columnKB);
        a.setPreferredWidth(widthKB);
        a.setCellRenderer(new CustomTableCellRenderer());
        a = jTable1.getColumnModel().getColumn(columnDateKB);
        a.setPreferredWidth(widthDateKB);
        a.setCellRenderer(new CustomTableCellRenderer());
        a = jTable1.getColumnModel().getColumn(columnTCode);
        a.setPreferredWidth(widthTCode);
        a.setCellRenderer(new CustomTableCellRenderer());
        
        Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
        this.setLocation(dim.width/2-this.getSize().width/2, dim.height/2-this.getSize().height/2);
        
        java.net.URL url = ClassLoader.getSystemResource("servicecloud/images/icon.png");
        Toolkit kit = Toolkit.getDefaultToolkit();
        Image img = kit.createImage(url);
        this.setIconImage(img);
    }
    
    private void saveVersionInfo() {
        try (PrintWriter versionWriter = new PrintWriter("version.txt", "UTF-8")) {
            versionWriter.println(appVersion);
            versionWriter.close();
        } catch (FileNotFoundException | UnsupportedEncodingException ex) {
            Logger.getLogger(Main.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    private void initTicketsFile() {
        try {
            Properties prop = new Properties();
            try (InputStream in = new FileInputStream("config.properties")) {
                prop.load(in);
            }
            sourceValue = prop.getProperty("sourceValue");
            jTextField1.setText(sourceValue);
            sourceSelection = Integer.valueOf(prop.getProperty("sourceSelection"));
            jComboBox2.setSelectedIndex(sourceSelection);
        }
        catch(IOException | NumberFormatException e) {
            System.out.println(e);
        }
    }
    
    private void addRows(String ticketSummary, String progress, String created, String dueDate, String category, String requestor, String ticketID, String description,String status, String remarks,  String classification, String assignees, String modified, String resolved, String kb, String dateKB, String tCode) {
        model = (DefaultTableModel) jTable1.getModel();
        if(!status.equalsIgnoreCase("Pending") && !status.equalsIgnoreCase("Resolved")) {
            remarks = "";
        }
        // Add the rows of the table
        model.addRow(new Object[]{ticketSummary, progress, created, dueDate, category, requestor, ticketID, description,status, remarks,  classification, assignees, modified, resolved, kb, dateKB, tCode});  
    }
    
    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
       refresh();
    }//GEN-LAST:event_jButton1ActionPerformed

    private void cleanProgressFile() {
        int confirmation = JOptionPane.showConfirmDialog (null, "Are you sure you want to continue?\nThis will delete all progress which are not on your current tickets list.","Warning", JOptionPane.YES_NO_OPTION);
        
        if(confirmation == JOptionPane.YES_OPTION) {
        
            try {
                List<String> progessFileList = new ArrayList<>();
                File progressFile = new File("Progress.txt");
                File tempFile = new File("Progress.tmp");

                try (BufferedReader reader = new BufferedReader(new FileReader(progressFile))) {
                    String currentLine;
                    while((currentLine = reader.readLine()) != null) {
                        if(!currentLine.isEmpty()) {
                            progessFileList.add(currentLine.toUpperCase().trim());
                        }
                    }          
                    reader.close();

                    // Clear Duplicates
                    Set<String> hs = new HashSet<>();
                    hs.addAll(progessFileList);
                    progessFileList.clear();
                    progessFileList.addAll(hs);

                    // Sort progress file
                    Collections.sort(progessFileList);
                }
                try (BufferedWriter writer = new BufferedWriter(new FileWriter(tempFile))) {
                    for(String p : progessFileList) {
                        for(List<String> s : ticketsList) {
                            if(p.indexOf(s.get(0)+" "+s.get(1))<0) {
                                continue;
                            }
                            else {
                                writer.write(p + System.getProperty("line.separator"));
                                break;
                            }
                        }
                    }
                    writer.close();
                }
                FileUtils.deleteQuietly(progressFile);
                FileUtils.moveFile(tempFile, progressFile);

                jStatusBar.setText("Progress File Cleaned!");
                JOptionPane.showMessageDialog(null, "Progress File Cleaned!", "Progress File", JOptionPane.INFORMATION_MESSAGE);
            }
            catch(IOException ioe) {}
        }
    }
    
    private void loadTicketConditions() {
        try {
            if(!jTextField1.getText().isEmpty()) {
                if(jComboBox2.getSelectedIndex()==0) {
                    try (BufferedReader br = new BufferedReader(new FileReader(jTextField1.getText()))) {
                        StringBuilder sb1 = new StringBuilder();
                        StringBuilder sb2 = new StringBuilder();
                        StringBuilder sb3 = new StringBuilder();
                        
                        String line = br.readLine();
                        
                        while (line != null) {
                            try {
                                List<String> ticketsRow = new ArrayList<>();
                                ticketsRow.add(line.split(" ")[0]);
                                ticketsRow.add(line.split(" ")[1]);
                                
                                ticketsList.add(ticketsRow);
                                line = br.readLine();
                            }
                            catch(ArrayIndexOutOfBoundsException | NullPointerException e) {}
                        }
                        
                        /*while (line != null) {
                            if(sb.length()!=0) {
                                sb.append(",");
                            }
                            sb.append("'").append(line).append("'");
                            line = br.readLine();
                            
                            
                        }*/
                        //condition = sb.toString();
                        
                        for(List<String> s : ticketsList) {
                            if(s.get(0).equalsIgnoreCase("SD")) {
                                if(sb1.length()!=0) {
                                    sb1.append(",");
                                }
                                sb1.append("'").append(s.get(1)).append("'");
                            }
                            else if(s.get(0).equalsIgnoreCase("PM")) {
                                if(sb2.length()!=0) {
                                    sb2.append(",");
                                }
                                sb2.append("'").append(s.get(1)).append("'");
                            }
                            else if(s.get(0).equalsIgnoreCase("RFC")) {
                                if(sb3.length()!=0) {
                                    sb3.append(",");
                                }
                                sb3.append("'").append(s.get(1)).append("'");
                            }
                        }
                        conditionMaster1 = sb1.toString();
                        conditionMaster2 = sb2.toString();
                        conditionMaster3 = sb3.toString();
                        
                        /*System.out.println("SD: "+ conditionMaster1);
                        System.out.println("PM: "+ conditionMaster2);
                        System.out.println("RFC: "+ conditionMaster3);*/
                    }
                    catch(FileNotFoundException fnf) {
                        conditionMaster1 = "";
                        conditionMaster2 = "";
                        conditionMaster3 = "";
                        conditionSearch = "";
                        if(!errorAlreadyEncountered) {
                            JOptionPane.showMessageDialog(null, "The file you specified could not be located", "File not found", JOptionPane.ERROR_MESSAGE);
                            errorAlreadyEncountered = true;
                        }
                    }
                }
                else if(jComboBox2.getSelectedIndex()==1) {
                    conditionSearch = jTextField1.getText();
                }
                
                conditionMaster1 = conditionMaster1.trim();
                conditionMaster2 = conditionMaster2.trim();
                conditionMaster3 = conditionMaster3.trim();
                conditionSearch = conditionSearch.trim();
            }
            else {
                conditionMaster1 = "";
                conditionMaster2 = "";
                conditionMaster3 = "";
                conditionSearch = "";
                if(!errorAlreadyEncountered) {
                    JOptionPane.showMessageDialog(null, "Please provide a valid data source", "Warning", JOptionPane.WARNING_MESSAGE);
                    errorAlreadyEncountered = true;
                }
            }
        }
        catch(IOException | HeadlessException e) {
            if(!errorAlreadyEncountered) {
                JOptionPane.showMessageDialog(null, e, "Error", JOptionPane.ERROR_MESSAGE);
                errorAlreadyEncountered = true;
            }
        }
    }
    
    private void loadProgress() {
        File progressFile = new File("Progress.txt");
        
        try (BufferedReader br = new BufferedReader(new FileReader(progressFile))) {
            String line = br.readLine();
            while (line != null) {
                try {
                    List<String> progressRow = new ArrayList<>();
                    progressRow.add(line.split(" ")[0]+" "+line.split(" ")[1]);
                    progressRow.add(line.split(" ")[2]);

                    progressList.add(progressRow);
                    line = br.readLine();
                }
                catch(ArrayIndexOutOfBoundsException | NullPointerException e) {}
            }
        }
        catch(Exception e) { }
    }
    
    private void refresh() {
        long tStart = System.currentTimeMillis();
        String url = "jdbc:jtds:sqlserver://"+hostname+":"+port+"/"+databaseName;
        String driver = "net.sourceforge.jtds.jdbc.Driver";
        String user = ""; // Removed for GitHub
        String pass = ""; // Removed for GitHub
        
        selectedTicketNumber = "";
        sortDisabled = false;
        int recordCount = 0;
        //int commentCount = 0;
        //commentsList.clear();
        
        try {
            ticketsList.clear();
            loadTicketConditions();
            loadProgress();
            
            if( (!conditionMaster1.equalsIgnoreCase("") && !conditionMaster1.isEmpty()) ||
                (!conditionMaster2.equalsIgnoreCase("") && !conditionMaster2.isEmpty()) ||
                (!conditionMaster3.equalsIgnoreCase("") && !conditionMaster3.isEmpty()) ||
                (!conditionSearch.equalsIgnoreCase("")  && !conditionSearch.isEmpty())) {
                
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
                
                // Note: MASTER1 and MASTER2 due date field is a.SLA__bDue__bDate
                // Note: MASTER3 due date field is a.Due__bDate
                if(jComboBox2.getSelectedIndex()==0) {
                    if(!conditionMaster1.isEmpty()) {
                        query += " SELECT a.Ticket__bSummary, a.mrSUBMITDATE, a.SLA__bDue__bDate, a.Category, a.Sub__bCategory, b.Display__bName, 'SD '+CAST(a.mrID AS VARCHAR), a.mrTitle, a.mrStatus, a.Pending__bReason, a.Resolution__bState, a.Type__bof__bTicket, a.mrAssignees, a.mrUpdateDate, a.mrALLDESCRIPTIONS, a.T__ucode"+
                                 " FROM MASTER1 a"+ // i.e. MASTER1, MASTER2, MASTER3
                                 " INNER JOIN MASTER1_ABDATA b ON a.mrID = b.mrID";
                        query += " WHERE a.mrID IN ("+conditionMaster1+")";
                    }
                    if(!conditionMaster2.isEmpty()) {
                        if(!conditionMaster1.isEmpty()) {
                            query += " UNION ALL";
                        }
                        query += " SELECT a.Ticket__bSummary, a.mrSUBMITDATE, a.SLA__bDue__bDate, a.Category, a.Sub__bCategory, b.Display__bName, 'PM '+CAST(a.mrID AS VARCHAR), a.mrTitle, a.mrStatus, a.Pending__bReason, a.Resolution__bState, a.Type__bof__bTicket, a.mrAssignees, a.mrUpdateDate, a.mrALLDESCRIPTIONS, a.T__ucode"+
                                 " FROM MASTER2 a"+ // i.e. MASTER1, MASTER2, MASTER3
                                 " INNER JOIN MASTER2_ABDATA b ON a.mrID = b.mrID";
                        query += " WHERE a.mrID IN ("+conditionMaster2+")";
                    }
                    if(!conditionMaster3.isEmpty()) {
                        if(!conditionMaster1.isEmpty() || !conditionMaster2.isEmpty()) {
                            query += " UNION ALL";
                        }
                        query += " SELECT a.Ticket__bSummary, a.mrSUBMITDATE, a.Due__bDate, a.Category, a.Sub__bCategory, b.Display__bName, 'RFC '+CAST(a.mrID AS VARCHAR), a.mrTitle, a.mrStatus, a.Pending__bReason, a.Resolution__bState, a.Type__bof__bTicket, a.mrAssignees, a.mrUpdateDate, a.mrALLDESCRIPTIONS, a.T__ucode"+
                                 " FROM MASTER3 a"+ // i.e. MASTER1, MASTER2, MASTER3
                                 " INNER JOIN MASTER3_ABDATA b ON a.mrID = b.mrID";
                        query += " WHERE a.mrID IN ("+conditionMaster3+")";
                    }
                }
                else if(jComboBox2.getSelectedIndex()==1) {
                    //conditionSearch = Utilities.unfixString(conditionSearch);
                    query += " SELECT a.Ticket__bSummary, a.mrSUBMITDATE, a.SLA__bDue__bDate, a.Category, a.Sub__bCategory, b.Display__bName, 'SD '+CAST(a.mrID AS VARCHAR), a.mrTitle, a.mrStatus, a.Pending__bReason, a.Resolution__bState, a.Type__bof__bTicket, a.mrAssignees, a.mrUpdateDate, a.mrALLDESCRIPTIONS, a.T__ucode"+
                             " FROM MASTER1 a"+ // i.e. MASTER1, MASTER2, MASTER3
                             " INNER JOIN MASTER1_ABDATA b ON a.mrID = b.mrID";
                    query += " WHERE a.mrAssignees LIKE '%"+conditionSearch+"%' OR a.mrAssignees LIKE '"+conditionSearch+"%' OR a.mrAssignees LIKE '%"+conditionSearch+"'";
                    query += " UNION ALL";
                    query += " SELECT a.Ticket__bSummary, a.mrSUBMITDATE, a.SLA__bDue__bDate, a.Category, a.Sub__bCategory, b.Display__bName, 'PM '+CAST(a.mrID AS VARCHAR), a.mrTitle, a.mrStatus, a.Pending__bReason, a.Resolution__bState, a.Type__bof__bTicket, a.mrAssignees, a.mrUpdateDate, a.mrALLDESCRIPTIONS, a.T__ucode"+
                             " FROM MASTER2 a"+ // i.e. MASTER1, MASTER2, MASTER3
                             " INNER JOIN MASTER2_ABDATA b ON a.mrID = b.mrID";
                    query += " WHERE a.mrAssignees LIKE '%"+conditionSearch+"%' OR a.mrAssignees LIKE '"+conditionSearch+"%' OR a.mrAssignees LIKE '%"+conditionSearch+"'";
                    query += " UNION ALL";
                    query += " SELECT a.Ticket__bSummary, a.mrSUBMITDATE, a.Due__bDate, a.Category, a.Sub__bCategory, b.Display__bName, 'RFC '+CAST(a.mrID AS VARCHAR), a.mrTitle, a.mrStatus, a.Pending__bReason, a.Resolution__bState, a.Type__bof__bTicket, a.mrAssignees, a.mrUpdateDate, a.mrALLDESCRIPTIONS, a.T__ucode"+
                             " FROM MASTER3 a"+ // i.e. MASTER1, MASTER2, MASTER3
                             " INNER JOIN MASTER3_ABDATA b ON a.mrID = b.mrID";
                    query += " WHERE a.mrAssignees LIKE '%"+conditionSearch+"%' OR a.mrAssignees LIKE '"+conditionSearch+"%' OR a.mrAssignees LIKE '%"+conditionSearch+"'";
                }
                query += " ORDER BY a.Ticket__bSummary ASC, a.mrSUBMITDATE ASC";
                
                sqlData.clear();
                try (Statement stmt = conn.createStatement(); ResultSet rs = stmt.executeQuery( query )) {
                    // Loop through the result set
                    while( rs.next() ) {
                        List<String> sqlDataRow = new ArrayList<>();
                        sqlDataRow.add(rs.getString(1));
                        sqlDataRow.add(rs.getString(2));
                        sqlDataRow.add(rs.getString(3));
                        sqlDataRow.add(rs.getString(4));
                        sqlDataRow.add(rs.getString(5));
                        sqlDataRow.add(rs.getString(6));
                        sqlDataRow.add(rs.getString(7));
                        sqlDataRow.add(rs.getString(8));
                        sqlDataRow.add(rs.getString(9));
                        sqlDataRow.add(rs.getString(10));
                        sqlDataRow.add(rs.getString(11));
                        sqlDataRow.add(rs.getString(12));
                        sqlDataRow.add(rs.getString(13));
                        sqlDataRow.add(rs.getString(14));
                        sqlDataRow.add(rs.getString(15));
                        sqlDataRow.add(rs.getString(16));
                        sqlData.add(sqlDataRow);
                    }
                    
                    // Required for MS SQL Server
                    stmt.cancel();
                    stmt.close();
                    rs.close();
                }
                
                try {
                    data.clear();
                    ticketCount = 0;
                    incidentCount = 0;
                    for(List<String> ticket : sqlData) {
                        ticketCount++;
                        String ticketSummary = ticket.get(0);
                        String dateCreated = "";
                        String dateDue = "";
                        String category = "";
                        if(ticket.get(4)!=null) {
                            category = ticket.get(4);
                        }
                        else if(ticket.get(3)!=null) {
                            category = ticket.get(4);
                        }
                        String requestor = ticket.get(5);
                        String ticketNumber = ticket.get(6);
                        String description = StringEscapeUtils.unescapeHtml4(ticket.get(7));

                        String progress = "";
                        for(List<String> progressStr : progressList) {
                            if(ticketNumber.equalsIgnoreCase(progressStr.get(0))) {
                                progress = progressStr.get(1);
                            }
                        }
                        
                        String status = ticket.get(8);
                        String remarks = "";
                        String classification = ticket.get(11);
                        if(classification==null) {
                            classification = "";
                        }
                        if(classification.equalsIgnoreCase("Incident")) {
                            incidentCount++;
                        }
                        String assignees = ticket.get(12);
                        String dateModified = "";
                        String dateResolved = "";
                        String kbSolutionNumber = "";
                        String dateLinkedToSolution = "";
                        String comments = ticket.get(14);
                        String tcode = ticket.get(15);
                        
                        //commentsList.add(new ArrayList<String>());
                        //commentsList.get(commentCount).add(ticketNumber);
                        //commentsList.get(commentCount).add(comments);
                        //commentCount++;

                        if(status!=null) {
                            status = status.replaceAll("_PENDING_SOLUTION_", "Pending Solution");
                            status = status.replaceAll("_SOLVED_", "Solved");
                            status = Utilities.fixString(status);

                            if(status.equalsIgnoreCase("Pending") || status.equalsIgnoreCase("Pending Solution")) {
                                remarks = ticket.get(9);
                            }
                            else {
                                remarks = ticket.get(10);
                            }
                        }

                        if(ticket.get(1)!=null) {
                            dateCreated = sdf.format(sdf.parse(ticket.get(1).substring(0,10)));
                        }
                        if(ticket.get(2)!=null) {
                            dateDue = sdf.format(sdf.parse(ticket.get(2).substring(0,10)));
                        }
                        if(ticket.get(13)!=null) {
                            dateModified = sdf.format(sdf.parse(ticket.get(13).substring(0,10)));
                        }

                        assignees = Utilities.fixString(assignees);
                        //assignees = assignees.split(" CC:")[0]; // Get only string before CC:
                        assignees = assignees.replaceAll("(\\bCC:)\\S+\\b","").trim();
                        assignees = Utilities.removeOrganization(assignees);
                        ticketSummary = Utilities.fixString(ticketSummary);
                        category = Utilities.fixString(category);
                        remarks = Utilities.fixString(remarks);

                        String selectedWorkspace = getSelectedWorkspace(ticketNumber);

                        /* Start of Date Resolved */
                        String query2 = "select mrHISTORY, mrGeneration from MASTER"+selectedWorkspace+"_HISTORY ";
                        query2 += " where mrID = '"+ticketNumber.split(" ")[1]+"'";
                        query2 += " and mrHistory like '%Resolved%'";
                        query2 += " order by mrGENERATION desc";
                        //System.out.println("ticketID: "+ticketID);
                        int mrGeneration = 0;
                        try (Statement stmt3 = conn.createStatement(); ResultSet rs3 = stmt3.executeQuery( query2 )) {
                            while( rs3.next() ) {
                                String dateResolvedString = Utilities.fixString(rs3.getString(1));
                                if(mrGeneration == 0) {
                                    mrGeneration = rs3.getInt(2);
                                    //dateResolved = dateResolvedString.split("\\s+")[0]+" "+dateResolvedString.split("\\s+")[1];
                                    dateResolved = dateResolvedString.split("\\s+")[0];
                                }
                                else {
                                    if(mrGeneration-1 == rs3.getInt(2)) {
                                        mrGeneration = rs3.getInt(2);
                                        //dateResolved = dateResolvedString.split("\\s+")[0]+" "+dateResolvedString.split("\\s+")[1];
                                        dateResolved = dateResolvedString.split("\\s+")[0];
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
                            query2 += " where mrID = '"+ticketNumber.split(" ")[1]+"'";
                            query2 += " and mrHistory like '%Closed%'";
                            query2 += " order by mrGENERATION desc";
                            //System.out.println("ticketID: "+ticketID);
                            mrGeneration = 0;
                            try (Statement stmt2 = conn.createStatement(); ResultSet rs2 = stmt2.executeQuery( query2 )) {
                                while( rs2.next() ) {
                                    String dateResolvedString = Utilities.fixString(rs2.getString(1));
                                    if(mrGeneration == 0) {
                                        mrGeneration = rs2.getInt(2);
                                        //dateResolved = dateResolvedString.split("\\s+")[0]+" "+dateResolvedString.split("\\s+")[1];
                                        dateResolved = dateResolvedString.split("\\s+")[0];
                                    }
                                    else {
                                        if(mrGeneration-1 == rs2.getInt(2)) {
                                            mrGeneration = rs2.getInt(2);
                                            //dateResolved = dateResolvedString.split("\\s+")[0]+" "+dateResolvedString.split("\\s+")[1];
                                            dateResolved = dateResolvedString.split("\\s+")[0];
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
                        /* Start of KB Solution Number, Date Linked to Solution */
                        boolean kbFound = false;
                        boolean rfcFound = false;

                        String queryKB = "select mrHISTORY";
                        queryKB += " FROM MASTER"+selectedWorkspace+"_HISTORY"; // i.e. MASTER1_HISTORY, MASTER2_HISTORY, MASTER3_HISTORY
                        queryKB += " WHERE mrID = '"+ticketNumber.split(" ")[1]+"'";
                        queryKB += " AND mrHistory like '%Copied or Selected from add%'";

                        //System.out.println("queryKB: "+queryKB);
                        try (Statement stmtKB = conn.createStatement(); ResultSet rsKB = stmtKB.executeQuery( queryKB )) {
                            while( rsKB.next() ) {
                                kbSolutionNumber = rsKB.getString(1).split("\\s+")[9];
                                dateLinkedToSolution = rsKB.getString(1).split("\\s+")[0];
                                try {
                                    Integer.parseInt(kbSolutionNumber);
                                    kbSolutionNumber = "SOLUTION "+kbSolutionNumber;
                                    kbFound = true;

                                }catch(NumberFormatException nfe) {
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
                                    }
                                    else { 
                                        kbSolutionNumber = "";
                                        dateLinkedToSolution = "";
                                    }
                                }
                            }
                            stmtKB.cancel();
                            stmtKB.close();
                            rsKB.close();

                            boolean docsFound = false;
                            if(comments.contains("http://itsi.mayniladwater.com.ph/appmgnt/Internal%20Documents/")) {
                                docsFound = true;
                            }
                            if(!kbFound && !rfcFound && docsFound) { // If TA/TP documents found, still output as incident
                                kbSolutionNumber = "TA/TP";
                            }
                        }
                        /* End of KB Solution Number, Date Linked to Solution */

                        data.add(new ArrayList<String>());
                        data.get(recordCount).add(ticketSummary);
                        data.get(recordCount).add(progress);
                        data.get(recordCount).add(dateCreated);
                        data.get(recordCount).add(dateDue);
                        data.get(recordCount).add(category);
                        data.get(recordCount).add(requestor);
                        data.get(recordCount).add(ticketNumber);
                        data.get(recordCount).add(description);
                        data.get(recordCount).add(status);
                        data.get(recordCount).add(remarks);
                        data.get(recordCount).add(classification);
                        data.get(recordCount).add(assignees);
                        data.get(recordCount).add(dateModified);
                        data.get(recordCount).add(dateResolved);
                        data.get(recordCount).add(kbSolutionNumber);
                        data.get(recordCount).add(dateLinkedToSolution);
                        data.get(recordCount).add(tcode);
                        recordCount++;
                    }
                    sqlData.clear();
                }
                catch(ParseException pe) {
                    System.out.println(pe);
                }
                
                /***** BEGIN OUTPUT TO GUI TABLE *****/
                // Reset the sort order
                sortOrder = "0 ASC";
                jTable1.getColumnModel().getColumn(columnTicketSummary).setHeaderValue("TICKET SUMMARY");
                jTable1.getColumnModel().getColumn(columnProgress).setHeaderValue("PROGRESS");
                jTable1.getColumnModel().getColumn(columnCreated).setHeaderValue("CREATED");
                jTable1.getColumnModel().getColumn(columnDueDate).setHeaderValue("DUE DATE");
                jTable1.getColumnModel().getColumn(columnCategory).setHeaderValue("CATEGORY");
                jTable1.getColumnModel().getColumn(columnRequestor).setHeaderValue("REQUESTOR");
                jTable1.getColumnModel().getColumn(columnTicket).setHeaderValue("TICKET");
                jTable1.getColumnModel().getColumn(columnDescription).setHeaderValue("DESCRIPTION");         
                jTable1.getColumnModel().getColumn(columnStatus).setHeaderValue("STATUS");
                jTable1.getColumnModel().getColumn(columnRemarks).setHeaderValue("REMARKS");
                jTable1.getColumnModel().getColumn(columnClassification).setHeaderValue("CLASSIFICATION");    
                jTable1.getColumnModel().getColumn(columnAssignees).setHeaderValue("ASSIGNEES");
                jTable1.getColumnModel().getColumn(columnModified).setHeaderValue("MODIFIED");
                jTable1.getColumnModel().getColumn(columnResolved).setHeaderValue("RESOLVED");
                jTable1.getColumnModel().getColumn(columnKB).setHeaderValue("KB");
                jTable1.getColumnModel().getColumn(columnDateKB).setHeaderValue("DATE LINKED TO KB");
                jTable1.getColumnModel().getColumn(columnTCode).setHeaderValue("T-CODE");
                jTable1.getTableHeader().resizeAndRepaint();
                // Set table contents
                loadTableContents();
                if(recordCount==0) {
                    clearTableContents();
                    addRows("","","","","","","","","","","","","","","","","");
                    if(!errorAlreadyEncountered) {
                        JOptionPane.showMessageDialog(null, "Your data source does not contain any valid tickets or is empty", "Warning", JOptionPane.WARNING_MESSAGE);
                        errorAlreadyEncountered = true;
                    }
                }
                /***** END OUTPUT TO GUI TABLE *****/
                
                conn.close();
            }
            else {
                clearTableContents();
                addRows("","","","","","","","","","","","","","","","","");
                if(!errorAlreadyEncountered) {
                    JOptionPane.showMessageDialog(null, "Your data source does not contain any valid tickets or is empty", "Warning", JOptionPane.WARNING_MESSAGE);
                    errorAlreadyEncountered = true;
                }
            }
        }
        catch(ClassNotFoundException | SQLException e) {
            if(!errorAlreadyEncountered) {
                JOptionPane.showMessageDialog(null, e, "Error", JOptionPane.ERROR_MESSAGE);
                errorAlreadyEncountered = true;
            }
        }
        System.gc();
        long tEnd = System.currentTimeMillis();        
        long tDelta = tEnd - tStart;
        double elapsedSeconds = tDelta / 1000.0;
        if(incidentCount == 1) {
            jStatusBar.setText("Extracted "+ticketCount+" tickets in "+elapsedSeconds+" seconds. You have 1 incident ticket!");
        }
        else if(incidentCount > 1) {
            jStatusBar.setText("Extracted "+ticketCount+" tickets in "+elapsedSeconds+" seconds. You have "+incidentCount+" incident tickets!");
        }
        else {
            if(elapsedSeconds < 4) {
                jStatusBar.setText("Extracted "+ticketCount+" tickets in "+elapsedSeconds+" seconds.");
            }
            else {
                jStatusBar.setText("Extracted "+ticketCount+" tickets in "+elapsedSeconds+" seconds.");
            }
        }
    }
    
    private void clearTableContents() {
        jTable1.clearSelection();
        model = (DefaultTableModel) jTable1.getModel();
        // Remove all contents of the table
        model.getDataVector().removeAllElements();
    }
    
    private void loadTableContents() {
        clearTableContents();
        
        // Add contents to the table
        for(List<String> s : data) {
            addRows(s.get(columnTicketSummary),s.get(columnProgress),s.get(columnCreated),s.get(columnDueDate),s.get(columnCategory),s.get(columnRequestor),s.get(columnTicket),s.get(columnDescription),s.get(columnStatus),s.get(columnRemarks),s.get(columnClassification),s.get(columnAssignees),s.get(columnModified),s.get(columnResolved),s.get(columnKB),s.get(columnDateKB),s.get(columnTCode));
        }
    }
    
    private void jMenuItem1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem1ActionPerformed
        System.exit(0);
    }//GEN-LAST:event_jMenuItem1ActionPerformed

    private void jMenuItem2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem2ActionPerformed
        JOptionPane.showMessageDialog(null, "Service Cloud v"+appVersion+"\n\nInstall Directory: "+destination+"\nUpdate Server: "+source+"\n\nService Cloud is a free non-open source ticket management application.\nCopyright 2016 © Glenn Dimaliwat", "About", JOptionPane.INFORMATION_MESSAGE);
    }//GEN-LAST:event_jMenuItem2ActionPerformed

    private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton2ActionPerformed
        try {
            if(jComboBox2.getSelectedIndex()==0) {
                Runtime.getRuntime().exec("notepad.exe " + jTextField1.getText());
            }
            else {
                JOptionPane.showMessageDialog(null, "This feature is only available if your data source is a Tickets File.", "Information", JOptionPane.INFORMATION_MESSAGE);
            }
        }
        catch(FileNotFoundException fnf) {
            JOptionPane.showMessageDialog(null, fnf, "File not found", JOptionPane.ERROR_MESSAGE);
        }
        catch(IOException e) {
            JOptionPane.showMessageDialog(null, e, "Error", JOptionPane.ERROR_MESSAGE);
        }
    }//GEN-LAST:event_jButton2ActionPerformed

    private void jMenuItem3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem3ActionPerformed
        JOptionPane.showMessageDialog(null, "1. Provide a data source either by specifying a File Path or your BMC Footprints Username.\n2. If you specified a file, clicking Edit File will open Notepad containing the file you indicated. Type your ticket numbers separated by a new line (enter). Save the file and close the Notepad.\n3. If you provided a username, the program will search all possible combinations of that username from the Assignees section.\n4. Select a workspace and click refresh.\n5. To search for a specific keyword in your tickets, press Ctrl+F\n6. You may also add new tickets to your file by going to the File menu > Add/Remove Ticket.", "How to use this App", JOptionPane.INFORMATION_MESSAGE);
    }//GEN-LAST:event_jMenuItem3ActionPerformed

    private void jToggleButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jToggleButton1ActionPerformed
        toggleCellMode();
    }//GEN-LAST:event_jToggleButton1ActionPerformed

    private void toggleCellMode() {
        if(jToggleButton1.isSelected()) {
            jTable1.setCellSelectionEnabled(true);
        }
        else {
            jTable1.setCellSelectionEnabled(false);
            jTable1.setRowSelectionAllowed(true);
        }
    }
    
    private void jTable1MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jTable1MouseClicked
        
    }//GEN-LAST:event_jTable1MouseClicked

    private void jButton3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton3ActionPerformed
        goToTicket();
    }//GEN-LAST:event_jButton3ActionPerformed

    private void goToTicket() {
        if(!selectedTicketNumber.equalsIgnoreCase("")) {
            try {
                String selectedWorkspace = getSelectedWorkspace(selectedTicketNumber);
                URI uri = new URI("http://servicedesk.mayniladwater.com.ph/MRcgi/MRlogin.pl?DL="+selectedTicketNumber.split(" ")[1]+"DA"+selectedWorkspace);
                Desktop dt = Desktop.getDesktop();
                dt.browse(uri);
                
            } catch ( URISyntaxException | IOException e ) {
                JOptionPane.showMessageDialog(null, "Unable to connect to URL.", "Error", JOptionPane.ERROR_MESSAGE);
            }
        }
        else {
            JOptionPane.showMessageDialog(null, "Please select a ticket number from the list.", "Information", JOptionPane.INFORMATION_MESSAGE);
        }
    }
    
    private void goToSolution() {
        if(!selectedTicketNumber.equalsIgnoreCase("")) {
            try {
                String selectedWorkspace = getSelectedWorkspace(selectedTicketNumber);
                
                if(selectedSolutionNumber.split(" ")[0].equalsIgnoreCase("SOLUTION")) {
                    URI uri = new URI("http://servicedesk.mayniladwater.com.ph/MRcgi/MRlogin.pl?DL="+selectedSolutionNumber.split(" ")[1]+"DA"+selectedWorkspace);
                    Desktop dt = Desktop.getDesktop();
                    dt.browse(uri);
                }
                else if(selectedSolutionNumber.split(" ")[0].equalsIgnoreCase("RFC")) {
                    URI uri = new URI("http://servicedesk.mayniladwater.com.ph/MRcgi/MRlogin.pl?DL="+selectedSolutionNumber.split(" ")[1]+"DA3");
                    Desktop dt = Desktop.getDesktop();
                    dt.browse(uri);
                }
                else {
                    JOptionPane.showMessageDialog(null, "Unable to connect to URL.", "Error", JOptionPane.ERROR_MESSAGE);
                }                
            } catch ( ArrayIndexOutOfBoundsException | URISyntaxException | IOException e ) {
                JOptionPane.showMessageDialog(null, "Unable to connect to URL.", "Error", JOptionPane.ERROR_MESSAGE);
            }
        }
        else {
            JOptionPane.showMessageDialog(null, "Please select a ticket number from the list.", "Information", JOptionPane.INFORMATION_MESSAGE);
        }
    }
    
    private void jMenuItem4ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem4ActionPerformed
        lastSearchKey = searchKey;
        try {
            searchKey = "";
            searchKey = JOptionPane.showInputDialog("Find What", lastSearchKey);
            searchKey = searchKey.toLowerCase();
            System.out.println(searchKey);
            
            if(!searchKey.equalsIgnoreCase("") && model!=null) {
                for(int x=0;x<=model.getRowCount()-1;x++) {                        
                    if(model.getValueAt(x, 0)!=null && model.getValueAt(x, 0).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(0, 0);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 0);
                        break;
                    }
                    else if(model.getValueAt(x, 1)!=null && model.getValueAt(x, 1).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(1, 1);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 1);
                        break;
                    }
                    else if(model.getValueAt(x, 2)!=null && model.getValueAt(x, 2).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(2, 2);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 2);
                        break;
                    }
                    else if(model.getValueAt(x, 3)!=null && model.getValueAt(x, 3).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(3, 3);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 3);
                        break;
                    }
                    else if(model.getValueAt(x, 4)!=null && model.getValueAt(x, 4).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(4, 4);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 4);
                        break;
                    }
                    else if(model.getValueAt(x, 5)!=null && model.getValueAt(x, 5).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(5, 5);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 5);
                        break;
                    }
                    else if(model.getValueAt(x, 6)!=null && model.getValueAt(x, 6).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(6, 6);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 6);
                        break;
                    }
                    else if(model.getValueAt(x, 7)!=null && model.getValueAt(x, 7).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(7, 7);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 7);
                        break;
                    }
                    else if(model.getValueAt(x, 8)!=null && model.getValueAt(x, 8).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(8, 8);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 8);
                        break;
                    }
                    else if(model.getValueAt(x, 9)!=null && model.getValueAt(x, 9).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(9, 9);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 9);
                        break;
                    }
                    else if(model.getValueAt(x, 10)!=null && model.getValueAt(x, 10).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(10, 10);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 10);
                        break;
                    }
                    else if(model.getValueAt(x, 11)!=null && model.getValueAt(x, 11).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(11, 11);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 11);
                        break;
                    }
                    else if(model.getValueAt(x, 12)!=null && model.getValueAt(x, 12).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(12, 12);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 12);
                        break;
                    }
                    else if(model.getValueAt(x, 13)!=null && model.getValueAt(x, 13).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(13, 13);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 13);
                        break;
                    }
                    else if(model.getValueAt(x, 14)!=null && model.getValueAt(x, 14).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(14, 14);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 14);
                        break;
                    }
                    else if(model.getValueAt(x, 15)!=null && model.getValueAt(x, 15).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(15, 15);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 15);
                        break;
                    }
                    else if(model.getValueAt(x, 16)!=null && model.getValueAt(x, 16).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(16, 16);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 16);
                        break;
                    }
                    
                    if(x==model.getRowCount()-1) {
                        JOptionPane.showMessageDialog(null, "Search string not found.", "Information", JOptionPane.INFORMATION_MESSAGE);
                    }
                }
                jTable1.requestFocus();
            }
        }
        catch(NullPointerException e) {
            searchKey = lastSearchKey;
        }
        catch(Exception e) {
            System.out.println(e);
        }
    }//GEN-LAST:event_jMenuItem4ActionPerformed
    
    private void jMenuItem5ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem5ActionPerformed
        int searchRow;
        int lastRowSearched = 0;
        try {            
            if(repeatSearch) {
                searchRow = 0;
            }
            else {
                searchRow = jTable1.getSelectedRow();
                if(searchRow==model.getRowCount()-1) {
                    searchRow = 0;
                }
                else {
                    searchRow = searchRow + 1;
                }
            }
            
            if(!searchKey.equalsIgnoreCase("") && model!=null) {
                for(int x=searchRow;x<=model.getRowCount()-1;x++) {
                    if(model.getValueAt(x, 0)!=null && model.getValueAt(x, 0).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(0, 0);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 0);
                        lastRowSearched = x;
                        break;
                    }
                    else if(model.getValueAt(x, 1)!=null && model.getValueAt(x, 1).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(1, 1);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 1);
                        lastRowSearched = x;
                        break;
                    }
                    else if(model.getValueAt(x, 2)!=null && model.getValueAt(x, 2).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(2, 2);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 2);
                        lastRowSearched = x;
                        break;
                    }
                    else if(model.getValueAt(x, 3)!=null && model.getValueAt(x, 3).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(3, 3);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 3);
                        lastRowSearched = x;
                        break;
                    }
                    else if(model.getValueAt(x, 4)!=null && model.getValueAt(x, 4).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(4, 4);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 4);
                        lastRowSearched = x;
                        break;
                    }
                    else if(model.getValueAt(x, 5)!=null && model.getValueAt(x, 5).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(5, 5);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 5);
                        lastRowSearched = x;
                        break;
                    }
                    else if(model.getValueAt(x, 6)!=null && model.getValueAt(x, 6).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(6, 6);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 6);
                        lastRowSearched = x;
                        break;
                    }
                    else if(model.getValueAt(x, 7)!=null && model.getValueAt(x, 7).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(7, 7);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 7);
                        lastRowSearched = x;
                        break;
                    }
                    else if(model.getValueAt(x, 8)!=null && model.getValueAt(x, 8).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(8, 8);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 8);
                        lastRowSearched = x;
                        break;
                    }
                    else if(model.getValueAt(x, 9)!=null && model.getValueAt(x, 9).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(9, 9);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 9);
                        lastRowSearched = x;
                        break;
                    }
                    else if(model.getValueAt(x, 10)!=null && model.getValueAt(x, 10).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(10, 10);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 10);
                        lastRowSearched = x;
                        break;
                    }
                    else if(model.getValueAt(x, 11)!=null && model.getValueAt(x, 11).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(11, 11);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 11);
                        lastRowSearched = x;
                        break;
                    }
                    else if(model.getValueAt(x, 12)!=null && model.getValueAt(x, 12).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(12, 12);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 12);
                        lastRowSearched = x;
                        break;
                    }
                    else if(model.getValueAt(x, 13)!=null && model.getValueAt(x, 13).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(13, 13);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 13);
                        lastRowSearched = x;
                        break;
                    }
                    else if(model.getValueAt(x, 14)!=null && model.getValueAt(x, 14).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(14, 14);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 14);
                        lastRowSearched = x;
                        break;
                    }
                    else if(model.getValueAt(x, 15)!=null && model.getValueAt(x, 15).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(15, 15);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 15);
                        lastRowSearched = x;
                        break;
                    }
                    else if(model.getValueAt(x, 16)!=null && model.getValueAt(x, 16).toString().toLowerCase().contains(searchKey)) {
                        jTable1.setRowSelectionInterval(x, x);
                        jTable1.setColumnSelectionInterval(16, 16);
                        selectedTicketNumber = (String) jTable1.getModel().getValueAt(x, 16);
                        lastRowSearched = x;
                        break;
                    }
                }
                
                for(int x=lastRowSearched;x<=model.getRowCount()-1;x++) {
                    for(int y=0;y<=model.getColumnCount()-1;y++) {
                        if(model.getValueAt(x, y)!=null && model.getValueAt(x, y).toString().toLowerCase().contains(searchKey)) {
                            if(lastRowSearched==x) {
                                repeatSearch = true;
                            }
                            else {
                                repeatSearch = false;
                            }
                            break;
                        }
                    }
                }
                
                jTable1.requestFocus();
            }
        }
        catch(Exception e) {
            System.out.println(e);
        }
    }//GEN-LAST:event_jMenuItem5ActionPerformed

    private void jMenuItem6ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem6ActionPerformed
        refresh();
    }//GEN-LAST:event_jMenuItem6ActionPerformed

    private void jMenuItem7ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem7ActionPerformed
        jToggleButton1.doClick();
        toggleCellMode();
    }//GEN-LAST:event_jMenuItem7ActionPerformed

    private void jButton4ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton4ActionPerformed
        Properties prop = new Properties();
        try {
            if(!jTextField1.getText().equalsIgnoreCase("")) {
                prop.setProperty("sourceValue",jTextField1.getText());
                prop.store(new FileOutputStream("config.properties"),null);
                sourceValue = jTextField1.getText();
                prop.setProperty("sourceSelection",String.valueOf(jComboBox2.getSelectedIndex()));
                prop.store(new FileOutputStream("config.properties"),null);
                sourceSelection = jComboBox2.getSelectedIndex();
                JOptionPane.showMessageDialog(null, "The data source has been updated", "Information", JOptionPane.INFORMATION_MESSAGE);
            }
            else {
                JOptionPane.showMessageDialog(null, "Please provide a data source", "Warning", JOptionPane.WARNING_MESSAGE);
                jTextField1.setText(sourceValue);
                jComboBox2.setSelectedIndex(sourceSelection);
            }
        }
        catch (IOException | HeadlessException e) {
            System.out.println(e);
        }
    }//GEN-LAST:event_jButton4ActionPerformed

    private void jButton5ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton5ActionPerformed
        goToTicketDocs();
    }//GEN-LAST:event_jButton5ActionPerformed

    private void goToTicketDocs() {
        if(!selectedTicketNumber.equalsIgnoreCase("")) {
            try {
                String strMonth;
                String strYear;
                String strMM;
                if(!selectedTicketNumberResolvedDate.isEmpty()) { // If ticket is not yet resolved, use modified date
                    strMonth = sdfMonth.format(sdf.parse(selectedTicketNumberResolvedDate.substring(0, 10)));
                    strYear = sdfYear.format(sdf.parse(selectedTicketNumberResolvedDate.substring(0, 10)));
                    strMM   = sdfMM.format(sdf.parse(selectedTicketNumberResolvedDate.substring(0, 10)));
                }
                else {
                    strMonth = sdfMonth.format(sdf.parse(selectedTicketNumberModifiedDate.substring(0, 10)));
                    strYear = sdfYear.format(sdf.parse(selectedTicketNumberModifiedDate.substring(0, 10)));
                    strMM   = sdfMM.format(sdf.parse(selectedTicketNumberModifiedDate.substring(0, 10)));
                }

                if(Integer.parseInt(strYear)>=2015) {
                    if(selectedTicketNumber.split(" ")[0].equalsIgnoreCase("SD")) {
                        URI uri = new URI("http://itsi.mayniladwater.com.ph/appmgnt/Internal%20Documents/Forms/AllItems.aspx?RootFolder=%2Fappmgnt%2FInternal%20Documents%2F"+strYear+"%2F"+strMM+"%20"+strMonth+"%2FSD%20"+selectedTicketNumber.split(" ")[1]);
                        Desktop dt = Desktop.getDesktop();
                        dt.browse(uri);
                    }
                    else if(selectedTicketNumber.split(" ")[0].equalsIgnoreCase("PM")) {
                        URI uri = new URI("http://itsi.mayniladwater.com.ph/appmgnt/Internal%20Documents/Forms/AllItems.aspx?RootFolder=%2Fappmgnt%2FInternal%20Documents%2F"+strYear+"%2F"+strMM+"%20"+strMonth+"%2FPM%20"+selectedTicketNumber.split(" ")[1]);
                        Desktop dt = Desktop.getDesktop();
                        dt.browse(uri);
                    }
                    else if(selectedTicketNumber.split(" ")[0].equalsIgnoreCase("RFC")) {
                        URI uri = new URI("http://itsi.mayniladwater.com.ph/appmgnt/Internal%20Documents/Forms/AllItems.aspx?RootFolder=%2Fappmgnt%2FInternal%20Documents%2F"+strYear+"%2F"+strMM+"%20"+strMonth+"%2FRFC%20"+selectedTicketNumber.split(" ")[1]);
                        Desktop dt = Desktop.getDesktop();
                        dt.browse(uri);
                    }
                }
                else {
                    URI uri = new URI("http://itsi.mayniladwater.com.ph/appmgnt/Internal%20Documents/Forms/AllItems.aspx?RootFolder=%2Fappmgnt%2FInternal%20Documents%2F"+strYear+"%2F"+strMM+"%20"+strMonth+"%2FTN"+selectedTicketNumber.split(" ")[1]);
                    Desktop dt = Desktop.getDesktop();
                    dt.browse(uri);
                }
                
            } catch ( ParseException | URISyntaxException | IOException e ) {
                JOptionPane.showMessageDialog(null, "Unable to connect to URL.", "Error", JOptionPane.ERROR_MESSAGE);
            }
        }
        else {
            JOptionPane.showMessageDialog(null, "Please select a ticket number from the list.", "Information", JOptionPane.INFORMATION_MESSAGE);
        }
    }
    
    private void goToKBDocs() {
        if(!selectedTicketNumber.equalsIgnoreCase("")) {
            try {
                String strMonth;
                String strYear;
                String strMM;
                if(!selectedTicketNumberResolvedDate.isEmpty()) { // If ticket is not yet resolved, use modified date
                    strMonth = sdfMonth.format(sdf.parse(selectedTicketNumberResolvedDate.substring(0, 10)));
                    strYear = sdfYear.format(sdf.parse(selectedTicketNumberResolvedDate.substring(0, 10)));
                    strMM   = sdfMM.format(sdf.parse(selectedTicketNumberResolvedDate.substring(0, 10)));
                }
                else {
                    strMonth = sdfMonth.format(sdf.parse(selectedTicketNumberModifiedDate.substring(0, 10)));
                    strYear = sdfYear.format(sdf.parse(selectedTicketNumberModifiedDate.substring(0, 10)));
                    strMM   = sdfMM.format(sdf.parse(selectedTicketNumberModifiedDate.substring(0, 10)));
                }
                
                if(selectedSolutionNumber.split(" ")[0].equalsIgnoreCase("RFC")) {
                    URI uri = new URI("http://itsi.mayniladwater.com.ph/appmgnt/Internal%20Documents/Forms/AllItems.aspx?RootFolder=%2Fappmgnt%2FInternal%20Documents%2F"+strYear+"%2F"+strMM+"%20"+strMonth+"%2FRFC%20"+selectedSolutionNumber.split(" ")[1]);
                    Desktop dt = Desktop.getDesktop();
                    dt.browse(uri);
                }
                else if(selectedSolutionNumber.split(" ")[0].equalsIgnoreCase("TA/TP")) {
                    goToTicketDocs();
                }
                else {
                    JOptionPane.showMessageDialog(null, "KB Docs are only available for RFC or SD TA/TP.", "Error", JOptionPane.ERROR_MESSAGE);
                }
                
            } catch ( ArrayIndexOutOfBoundsException | ParseException | URISyntaxException | IOException e ) {
                JOptionPane.showMessageDialog(null, "Unable to connect to URL.", "Error", JOptionPane.ERROR_MESSAGE);
            }
        }
        else {
            JOptionPane.showMessageDialog(null, "Please select a ticket number from the list.", "Information", JOptionPane.INFORMATION_MESSAGE);
        }
    }
    
    private void jTable1KeyReleased(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_jTable1KeyReleased
        int keyCode = evt.getKeyCode();
        if(keyCode == KeyEvent.VK_UP||keyCode == KeyEvent.VK_DOWN||keyCode == KeyEvent.VK_LEFT||keyCode == KeyEvent.VK_RIGHT) {
            //selectedTicketNumber = (String) jTable1.getModel().getValueAt(jTable1.getSelectedRow(), 5);
            //selectedTicketNumberResolvedDate = (String) jTable1.getModel().getValueAt(jTable1.getSelectedRow(), 13);
            selectRowItem();
        }
    }//GEN-LAST:event_jTable1KeyReleased

    private void jMenuItem8ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem8ActionPerformed
        addTicketNumber();
    }//GEN-LAST:event_jMenuItem8ActionPerformed

    private void addTicketNumber() {
        if(jComboBox2.getSelectedIndex()==0) {
            if(!jTextField1.getText().isEmpty()) {
                String ticketNumber = JOptionPane.showInputDialog(null, "Please input the ticket number", "Add Ticket", JOptionPane.QUESTION_MESSAGE);
                
                if(ticketNumber!=null) {
                    if(!ticketNumber.isEmpty()) {
                        if(ticketNumber.split(" ").length>1 && ticketNumber.split(" ").length<3) {
                            addTicket(ticketNumber);
                            jStatusBar.setText("You just added "+ticketNumber);
                        }
                        else {
                            JOptionPane.showMessageDialog(null, "Invalid Ticket Number.", "Error", JOptionPane.ERROR_MESSAGE);
                        }
                    }
                    else {
                        JOptionPane.showMessageDialog(null, "Invalid Ticket Number.", "Error", JOptionPane.ERROR_MESSAGE);
                    }
                }
            }
            else {
                JOptionPane.showMessageDialog(null, "Please provide a ticket file path", "Warning", JOptionPane.WARNING_MESSAGE);
            }
        }
        else {
            JOptionPane.showMessageDialog(null, "This feature is only available if your data source is a Tickets File.", "Information", JOptionPane.INFORMATION_MESSAGE);
        }
    }
    
    private void jMenuItem9ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem9ActionPerformed
        removeTicketNumber();
    }//GEN-LAST:event_jMenuItem9ActionPerformed

    private void removeTicketNumber() {
        if(jComboBox2.getSelectedIndex()==0) {
            if(!selectedTicketNumber.equalsIgnoreCase("")) {
                int answer = JOptionPane.showConfirmDialog(null, "Remove ticket "+selectedTicketNumber+"?", "Remove Ticket", JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE);
                if(answer == 0) {
                    String removedTicket = selectedTicketNumber;
                    removeTicket(selectedTicketNumber);
                    jStatusBar.setText("You just deleted "+removedTicket);
                }
            }
            else {
                JOptionPane.showMessageDialog(null, "Please select a ticket number from the list.", "Information", JOptionPane.INFORMATION_MESSAGE);
            }
        }
        else {
            JOptionPane.showMessageDialog(null, "This feature is only available if your data source is a Tickets File.", "Information", JOptionPane.INFORMATION_MESSAGE);
        }
    }
    
    private void jComboBox2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jComboBox2ActionPerformed
        initSourceValue();
    }//GEN-LAST:event_jComboBox2ActionPerformed

    private void jMenuItem10ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem10ActionPerformed
        updateApplication(false);
    }//GEN-LAST:event_jMenuItem10ActionPerformed

    private void jTable1MousePressed(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jTable1MousePressed
        if(SwingUtilities.isRightMouseButton(evt)){
            Point p = evt.getPoint();
            int rowNumber = jTable1.rowAtPoint( p );
            ListSelectionModel lsModel;
            lsModel = jTable1.getSelectionModel();
            lsModel.setSelectionInterval( rowNumber, rowNumber );
            
            selectRowItem();
            
            PopupMenu menu = new PopupMenu();
            menu.show(evt.getComponent(), evt.getX(), evt.getY());
        }
        else {
            selectRowItem();
        }
    }//GEN-LAST:event_jTable1MousePressed

    private void selectRowItem() {
        selectedTicketNumber = (String) jTable1.getModel().getValueAt(jTable1.getSelectedRow(), columnTicket);
        selectedTicketNumberModifiedDate = (String) jTable1.getModel().getValueAt(jTable1.getSelectedRow(), columnModified);
        selectedTicketNumberResolvedDate = (String) jTable1.getModel().getValueAt(jTable1.getSelectedRow(), columnResolved);
        selectedSolutionNumber = (String) jTable1.getModel().getValueAt(jTable1.getSelectedRow(), columnKB);
        
        if(selectedSolutionNumber.isEmpty()) {
            jStatusBar.setText("You just selected "+selectedTicketNumber+".");
        }
        else {
            jStatusBar.setText("You just selected "+selectedTicketNumber+". This is linked to "+selectedSolutionNumber+".");
        }
    }
    
    private void jButton6ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton6ActionPerformed
        goToSolution();
    }//GEN-LAST:event_jButton6ActionPerformed

    private void jButton7ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton7ActionPerformed
        goToKBDocs();
    }//GEN-LAST:event_jButton7ActionPerformed

    private void jMenuItem11ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem11ActionPerformed
        cleanProgressFile();
    }//GEN-LAST:event_jMenuItem11ActionPerformed

    private void jMenuItem12ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem12ActionPerformed
        try {
            String ticketNumber = JOptionPane.showInputDialog(null, "Please input the ticket number", "Add Ticket", JOptionPane.QUESTION_MESSAGE);

            if(ticketNumber!=null) {
                if(!ticketNumber.isEmpty()) {
                    if(ticketNumber.split(" ").length>1 && ticketNumber.split(" ").length<3) {
                        String selectedWorkspace = getSelectedWorkspace(ticketNumber);
                        URI uri = new URI("http://servicedesk.mayniladwater.com.ph/MRcgi/MRlogin.pl?DL="+ticketNumber.split(" ")[1]+"DA"+selectedWorkspace);
                        Desktop dt = Desktop.getDesktop();
                        dt.browse(uri);
                    }
                    else {
                        JOptionPane.showMessageDialog(null, "Invalid Ticket Number.", "Error", JOptionPane.ERROR_MESSAGE);
                    }
                }
                else {
                    JOptionPane.showMessageDialog(null, "Invalid Ticket Number.", "Error", JOptionPane.ERROR_MESSAGE);
                }
            }

        } catch ( URISyntaxException | IOException e) {
            JOptionPane.showMessageDialog(null, "Unable to connect to URL.", "Error", JOptionPane.ERROR_MESSAGE);
        }
    }//GEN-LAST:event_jMenuItem12ActionPerformed
    
    /*private void updateApplicationThread(final boolean auto) {
        Thread thread = new Thread(new Runnable() {
            @Override
            public void run() {
                updateServerResolved = false;
                updateApplication(auto);
                if(updateSuccessful) {
                    System.exit(0);
                }
            }
        });
        thread.start();
        long endTimeMillis = System.currentTimeMillis() + 5000;
        while (thread.isAlive()) {
            if (System.currentTimeMillis() >= endTimeMillis) { // Thread ends if process reached 5 secs
                break;
            }
            else if(updateServerResolved) { // Thread ends if process finished before 5 secs
                break;
            }
            
            try {
                Thread.sleep(500);
            }
            catch (InterruptedException t) {}
        }
    }*/
    
    private void updateApplication(boolean auto) {
        try {
            File f = new File (latestVersionFile);
            BufferedReader br = new BufferedReader(new FileReader(f));
            StringBuilder sb = new StringBuilder();
            String line = br.readLine();
            while (line != null) {
                sb.append(line);
                line = br.readLine();
            }
            String appVersionFromServer = sb.toString();
            br.close();
            //updateServerResolved = true;

            if(Double.valueOf(appVersionFromServer)>Double.valueOf(appVersion)) {
                int startUpdate = JOptionPane.showConfirmDialog(null, "A new version of Service Cloud is available. Would you like to update Service Cloud to v"+appVersionFromServer+"? (ADMINISTRATOR ONLY)\n\nYou may also copy the latest version from the following file:\n"+binary, "Update Available", JOptionPane.YES_NO_OPTION);
                if(startUpdate==0) {
                    Runtime.getRuntime().exec("cmd /c robocopy \""+source+"\" \""+destination+"\" /e /copyall");
                    
                    // Recheck the version to check if Update was successful
                    f = new File (localVersionFile);
                    br = new BufferedReader(new FileReader(f));
                    sb = new StringBuilder();
                    line = br.readLine();

                    while (line != null) {
                        sb.append(line);
                        line = br.readLine();
                    }
                    appVersionFromServer = sb.toString();
                    br.close();
                    
                    if(Double.valueOf(appVersionFromServer)>Double.valueOf(appVersion)) {
                        JOptionPane.showMessageDialog(null, "The update was not successful.\nPlease check if you have write access to your Service Cloud install directory.\nYou may re-try the update by using Help > Check for Updates.", "Update failed", JOptionPane.INFORMATION_MESSAGE);
                    }
                    else {
                        JOptionPane.showMessageDialog(null, "The update has successfully completed!\nService Cloud will now close. Please re-run the application.", "Update complete", JOptionPane.INFORMATION_MESSAGE);
                        System.exit(0);
                    }
                }
            }
            else {
                if(!auto) {
                    JOptionPane.showMessageDialog(null, "No new updates available at this time", "Check for Updates", JOptionPane.INFORMATION_MESSAGE);
                }
            }
        }
        catch(IOException | NumberFormatException | HeadlessException e) {
            if(!auto) {
                JOptionPane.showMessageDialog(null, "The update server could not be resolved at this time.\nPlease try again later.", "Check for Updates", JOptionPane.WARNING_MESSAGE);
            }
        }
    }
    
    private void initSourceValue() {
        try {
            Properties prop = new Properties();
            try (InputStream in = new FileInputStream("config.properties")) {
                prop.load(in);
            }
            sourceSelection = Integer.valueOf(prop.getProperty("sourceSelection"));
            if(jComboBox2.getSelectedIndex()==sourceSelection) {
                sourceValue = prop.getProperty("sourceValue");
                jTextField1.setText(sourceValue);
            }
            else {
                jTextField1.setText("");
            }
            
        }
        catch(IOException | NumberFormatException e) {
            System.out.println(e);
        }
    }
    private void addTicket(String ticketNumber) {
        try {
            try {
                
                String ticketWorkspace = ticketNumber.split(" ")[0];
                if(ticketWorkspace.equalsIgnoreCase("SD") || ticketWorkspace.equalsIgnoreCase("PM") || ticketWorkspace.equalsIgnoreCase("RFC")) {
                    int ticketNum = Integer.parseInt(ticketNumber.split(" ")[1]);
                    if(ticketNum >= 0) { // MIN ticket number in Table
                        Writer output;
                        output = new BufferedWriter(new FileWriter(jTextField1.getText(), true));
                        output.append("\r\n"+ticketNumber.toUpperCase());
                        output.close();

                        try {
                            sortTicketFile();
                            refresh();
                        } catch (Exception ex) {
                            System.err.println(ex);
                        }
                    }
                    else {
                        JOptionPane.showMessageDialog(null, "Please provide a valid ticket number (e.g. SD 1000)", "Error", JOptionPane.ERROR_MESSAGE);
                    }
                }
                else {
                    JOptionPane.showMessageDialog(null, "Please provide a valid ticket workspace (e.g. SD 1000)", "Error", JOptionPane.ERROR_MESSAGE);
                }
            }
            catch(NumberFormatException nfe) {
                JOptionPane.showMessageDialog(null, "Please provide a valid ticket number (e.g. SD 1000)", "Error", JOptionPane.ERROR_MESSAGE);
            }
        }
        catch(NullPointerException | HeadlessException | IOException e) {
            JOptionPane.showMessageDialog(null, "Error adding ticket number", "Error", JOptionPane.ERROR_MESSAGE);
        }
    }
    
    private void sortTicketFile() throws Exception {
        List<String> ticketList = new ArrayList<>();
        String line;
        
        if(!jTextField1.getText().isEmpty()) {
            try (FileReader fileReader = new FileReader(jTextField1.getText())) {
                BufferedReader bufferedReader = new BufferedReader(fileReader);
                while ((line = bufferedReader.readLine()) != null) {
                    if(!line.equalsIgnoreCase("")) {
                        ticketList.add(line.toUpperCase());
                    }
                }
            }
            
            // Clear Duplicates
            Set<String> hs = new HashSet<>();
            hs.addAll(ticketList);
            ticketList.clear();
            ticketList.addAll(hs);

            Collections.sort(ticketList);
            
            try (FileWriter fileWriter = new FileWriter(jTextField1.getText()); PrintWriter out = new PrintWriter(fileWriter)) {
                for (String outputLine : ticketList) {
                    out.println(outputLine);
                }
                out.flush();
            }
        }
    }
    
    private void removeTicket(String ticketNumber) {
        List<String> ticketList = new ArrayList<>();
        String line;
        
        try {
            if(!jTextField1.getText().isEmpty()) {
                try (FileReader fileReader = new FileReader(jTextField1.getText())) {
                    BufferedReader bufferedReader = new BufferedReader(fileReader);
                    while ((line = bufferedReader.readLine()) != null) {
                        if(!line.equalsIgnoreCase(ticketNumber) && !line.equalsIgnoreCase("")) {
                            ticketList.add(line.toUpperCase());
                        }
                    }
                }
                
                // Clear Duplicates
                Set<String> hs = new HashSet<>();
                hs.addAll(ticketList);
                ticketList.clear();
                ticketList.addAll(hs);

                // Sort list
                Collections.sort(ticketList);
                
                try (FileWriter fileWriter = new FileWriter(jTextField1.getText()); PrintWriter out = new PrintWriter(fileWriter)) {
                    for (String outputLine : ticketList) {
                        out.println(outputLine);
                    }
                    out.flush();
                }
                
                refresh();
            }
        }
        catch(Exception e) {
            JOptionPane.showMessageDialog(null, "Error removing ticket number", "Error", JOptionPane.ERROR_MESSAGE);
        }
    }
    
    private String getSelectedWorkspace(String ticketNumber) {
        if(ticketNumber.split(" ")[0].equalsIgnoreCase("SD")) {
            return "1";
        }
        else if(ticketNumber.split(" ")[0].equalsIgnoreCase("PM")) {
            return "2";
        }
        else if(ticketNumber.split(" ")[0].equalsIgnoreCase("RFC")) {
            return "3";
        }
        else {
            return "1";
        }
    }
    
    private void quickView() {
        /*for(List<String> c : commentsList) {
            if(selectedTicketNumber.equalsIgnoreCase(c.get(0))) {
                frameComments.loadWindow(c.get(0), c.get(1));
                break;
            }
        }*/
        
        String url = "jdbc:jtds:sqlserver://"+hostname+":"+port+"/"+databaseName;
        String driver = "net.sourceforge.jtds.jdbc.Driver";
        String user = ""; //Removed for GitHub
        String pass = ""; //Removed for GitHub
        
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

            String query = "";
            
            String selectedWorkspace = getSelectedWorkspace(selectedTicketNumber);
            String selectedTicketNo = selectedTicketNumber.split(" ")[1];
            
            if(selectedWorkspace.equalsIgnoreCase("1")) {
                query += " SELECT mrALLDESCRIPTIONS FROM MASTER1";
                query += " WHERE mrID IN ("+selectedTicketNo+")";
            }
            else if(selectedWorkspace.equalsIgnoreCase("2")) {
                query += " SELECT mrALLDESCRIPTIONS FROM MASTER2";
                query += " WHERE mrID IN ("+selectedTicketNo+")";
            }
            else if(selectedWorkspace.equalsIgnoreCase("3")) {
                query += " SELECT mrALLDESCRIPTIONS FROM MASTER3";
                query += " WHERE mrID IN ("+selectedTicketNo+")";
            }

            sqlData.clear();
            try (Statement stmt = conn.createStatement(); ResultSet rs = stmt.executeQuery( query )) {
                // Loop through the result set
                while( rs.next() ) {
                    List<String> sqlDataRow = new ArrayList<>();
                    sqlDataRow.add(rs.getString(1));
                    sqlData.add(sqlDataRow);
                }

                // Required for MS SQL Server
                stmt.cancel();
                stmt.close();
                rs.close();
            }
            
            String commentString = "";
            for(List<String> ticket : sqlData) {
                if(ticket.get(0)!=null) {
                    commentString = commentString + ticket.get(0)+System.getProperty("line.separator");
                }
            }
            frameComments.loadWindow(selectedTicketNumber, commentString);
            sqlData.clear();
        }
        catch(ClassNotFoundException | SQLException e) {
        }
    }
    
    private void quickViewHTML() {
        
        String url = "jdbc:jtds:sqlserver://"+hostname+":"+port+"/"+databaseName;
        String driver = "net.sourceforge.jtds.jdbc.Driver";
        String user = ""; // Removed for GitHub
        String pass = ""; // Removed for GitHub
        
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

            String query = "";
            
            String selectedWorkspace = getSelectedWorkspace(selectedTicketNumber);
            String selectedTicketNo = selectedTicketNumber.split(" ")[1];
            
            if(selectedWorkspace.equalsIgnoreCase("1")) {
                query += " SELECT mrDESCRIPTION FROM MASTER1_DESCRIPTIONS";
                query += " WHERE mrID IN ("+selectedTicketNo+")";
            }
            else if(selectedWorkspace.equalsIgnoreCase("2")) {
                query += " SELECT mrDESCRIPTION FROM MASTER2_DESCRIPTIONS";
                query += " WHERE mrID IN ("+selectedTicketNo+")";
            }
            else if(selectedWorkspace.equalsIgnoreCase("3")) {
                query += " SELECT mrDESCRIPTION FROM MASTER3_DESCRIPTIONS";
                query += " WHERE mrID IN ("+selectedTicketNo+")";
            }

            sqlData.clear();
            try (Statement stmt = conn.createStatement(); ResultSet rs = stmt.executeQuery( query )) {
                // Loop through the result set
                while( rs.next() ) {
                    List<String> sqlDataRow = new ArrayList<>();
                    sqlDataRow.add(rs.getString(1));
                    sqlData.add(sqlDataRow);
                }

                // Required for MS SQL Server
                stmt.cancel();
                stmt.close();
                rs.close();
            }
            
            String commentString = "";
            for(List<String> ticket : sqlData) {
                if(ticket.get(0)!=null) {
                    commentString = commentString + ticket.get(0)+"<br/>";
                }
            }
            frameCommentsHTML.loadWindow(selectedTicketNumber, commentString);
            sqlData.clear();
        }
        catch(ClassNotFoundException | SQLException e) {
        }
    }
    
    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the look and feel */
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | UnsupportedLookAndFeelException e) {
        }

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            @Override
            public void run() {
                new Main().setVisible(true);
            }
        });
    }
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton3;
    private javax.swing.JButton jButton4;
    private javax.swing.JButton jButton5;
    private javax.swing.JButton jButton6;
    private javax.swing.JButton jButton7;
    private javax.swing.JComboBox jComboBox2;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JMenu jMenu1;
    private javax.swing.JMenu jMenu2;
    private javax.swing.JMenu jMenu3;
    private javax.swing.JMenu jMenu4;
    private javax.swing.JMenuBar jMenuBar1;
    private javax.swing.JMenuItem jMenuItem1;
    private javax.swing.JMenuItem jMenuItem10;
    private javax.swing.JMenuItem jMenuItem11;
    private javax.swing.JMenuItem jMenuItem12;
    private javax.swing.JMenuItem jMenuItem2;
    private javax.swing.JMenuItem jMenuItem3;
    private javax.swing.JMenuItem jMenuItem4;
    private javax.swing.JMenuItem jMenuItem5;
    private javax.swing.JMenuItem jMenuItem6;
    private javax.swing.JMenuItem jMenuItem7;
    private javax.swing.JMenuItem jMenuItem8;
    private javax.swing.JMenuItem jMenuItem9;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JLabel jStatusBar;
    private javax.swing.JTable jTable1;
    private javax.swing.JTextField jTextField1;
    private javax.swing.JToggleButton jToggleButton1;
    // End of variables declaration//GEN-END:variables

    private static class ArrayList2DComparator implements Comparator {
        private int column = 0;
        private String sortOrder = "ASC";
        
        ArrayList2DComparator(int column, String sortOrder) {
            this.column = column;
            this.sortOrder = sortOrder;
        }
        @Override
        public int compare(Object obj1, Object obj2) {  
            if (! (obj1 instanceof ArrayList) || ! (obj2 instanceof ArrayList)) {  
                throw new ClassCastException(  
                            "compared objects must be instances of ArrayList");  
            }
            String str1 = (String) ((ArrayList) obj1).get(column);
            String str2 = (String) ((ArrayList) obj2).get(column);
           
            if(isInteger(str1) && isInteger(str2)) { // Pad 7 zeroes if String is all integer
                str1 = StringUtils.leftPad(str1, 7, "0");
                str2 = StringUtils.leftPad(str2, 7, "0");
            }
            
            if(sortOrder.equalsIgnoreCase("ASC")) {
                if(str1==null) {
                    return -1;
                }
                else if(str2==null) {
                    return 1;
                }
                return str1.compareTo(str2);
            }
            else {
                if(str2==null) {
                    return -1;
                }
                else if(str1==null) {
                    return 1;
                }
                return str2.compareTo(str1);
            }
        }
        
        public static boolean isInteger(String s) {
            try { 
                Integer.parseInt(s); 
            } catch(NumberFormatException e) { 
                return false; 
            }
            return true;
        }
    }
    
    private class PopupMenu extends JPopupMenu {
        private JMenuItem menuItem1;
        private JMenuItem menuItem2;
        private JMenuItem menuItem3;
        private JMenuItem menuItem4;
        private JMenuItem menuItem5;
        private JMenuItem menuItem6;
        private JMenuItem menuItem7;

        public PopupMenu() {
            menuItem1 = new JMenuItem("Refresh List");
            menuItem2 = new JMenuItem("Quick View in Plain Text");
            menuItem3 = new JMenuItem("Quick View in HTML");
            menuItem4 = new JMenuItem("Go to Ticket");
            menuItem5 = new JMenuItem("Add Ticket");
            menuItem6 = new JMenuItem("Remove Ticket");
            menuItem7 = new JMenuItem("Clean Progress File");
            
            menuItem1.setMnemonic('e');
            menuItem2.setMnemonic('P');
            menuItem3.setMnemonic('H');
            menuItem4.setMnemonic('T');
            menuItem5.setMnemonic('A');
            menuItem6.setMnemonic('R');
            menuItem7.setMnemonic('C');
            
            menuItem1.addActionListener(new MenuActionListener());
            menuItem2.addActionListener(new MenuActionListener());
            menuItem3.addActionListener(new MenuActionListener());
            menuItem4.addActionListener(new MenuActionListener());
            menuItem5.addActionListener(new MenuActionListener());
            menuItem6.addActionListener(new MenuActionListener());
            menuItem7.addActionListener(new MenuActionListener());
            menuItem7.setToolTipText("Removes deleted tickets from the Progress File");

            add(menuItem1);
            add(menuItem2);
            add(menuItem3);
            add(menuItem4);
            add(menuItem5);
            add(menuItem6);
            add(menuItem7);
        }
    }

    private class MenuActionListener implements ActionListener {
      @Override
      public void actionPerformed(ActionEvent e) {
        switch(e.getActionCommand()) {
            case "Refresh List":
                refresh();
                break;
            case "Quick View in Plain Text":
                quickView();
                break;
            case "Quick View in HTML":
                quickViewHTML();
                break;
            case "Go to Ticket":
                goToTicket();
                break;
            case "Add Ticket":
                addTicketNumber();
                break;
            case "Remove Ticket":
                removeTicketNumber();
                break;
            case "Clean Progress File":
                cleanProgressFile();
                break;
        }
      }
    }
}
