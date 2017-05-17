/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package servicecloud;

import java.awt.Color;
import java.awt.Component;
import java.util.Date;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import javax.swing.JLabel;
import javax.swing.JTable;
import javax.swing.table.DefaultTableCellRenderer;

/**
 *
 * @author Glenn Dimaliwat
 */
public class CustomTableCellRenderer extends DefaultTableCellRenderer {
    @Override
    public Component getTableCellRendererComponent(JTable table,
                                                 Object value,
                                                 boolean isSelected,
                                                 boolean hasFocus,
                                                 int row,
                                                 int column) {
    Component c = super.getTableCellRendererComponent(table, value,
                                          isSelected, hasFocus,
                                          row, column);
    
    //final int columnTicketSummary = 0;
    final int columnProgress = 1;
    //final int columnCreated = 2;
    final int columnDueDate = 3;
    //final int columnCategory = 4;
    //final int columnRequestor = 5;
    //final int columnTicket = 6;
    //final int columnDescription = 7;
    final int columnStatus = 8;
    final int columnRemarks = 9;
    final int columnClassification = 10;
    //final int columnAssignees = 11;
    //final int columnModified = 12;
    final int columnResolved = 13;
    final int columnKB = 14;
    final int columnDateKB = 15;
    
    try {
        //System.out.println("row: "+row+" col: "+column+" value: "+value);
        
        /* Center Progress Column */
        if(column==columnProgress) {
            ((JLabel)c).setHorizontalAlignment( JLabel.CENTER );
        }
        
        /* CHECK IF ROW IS SELECTED */
        if(isSelected) {
            c.setBackground(Color.BLUE);
            c.setForeground(Color.WHITE);
        }
        else if(!isSelected) {
            c.setBackground(Color.WHITE);
            c.setForeground(Color.BLACK);
        }
        
        /* CHECK STATUS */
        if(table.getValueAt(row, columnStatus).toString().equalsIgnoreCase("Pending") || table.getValueAt(row, columnStatus).toString().equalsIgnoreCase("Accepted") || table.getValueAt(row, columnStatus).toString().equalsIgnoreCase("Assigned") || table.getValueAt(row, columnStatus).toString().equalsIgnoreCase("Open") || table.getValueAt(row, columnStatus).toString().equalsIgnoreCase("Pending Solution") || table.getValueAt(row, columnStatus).toString().equalsIgnoreCase("Solved") || table.getValueAt(row, columnStatus).toString().equalsIgnoreCase("Re-opened")) {
            if(column==columnStatus || column==columnRemarks) {
                if(table.getCellSelectionEnabled()==false) {
                    c.setBackground(Color.BLUE.darker()); 
                    c.setForeground(Color.WHITE);
                }
            }
        }
        else if(table.getValueAt(row, columnStatus).toString().contains("Resolved") || table.getValueAt(row, columnStatus).toString().toString().contains("Closed") || table.getValueAt(row, columnStatus).toString().toString().contains("Cancelled") || table.getValueAt(row, columnStatus).toString().contains("Solved")) {
            if(column==columnStatus || column==columnRemarks) {
                if(table.getCellSelectionEnabled()==false) {
                    c.setBackground(Color.GREEN.darker().darker().darker());
                    c.setForeground(Color.WHITE);
                }
            }
            
            if(table.getValueAt(row, columnStatus).toString().contains("Resolved") || table.getValueAt(row, columnStatus).toString().contains("Closed")) {
                if(column==columnKB || column==columnDateKB) {
                    if(table.getCellSelectionEnabled()==false) {
                        if(table.getValueAt(row, columnKB).toString().equalsIgnoreCase("")) {
                            if(table.getValueAt(row, columnClassification).toString().contains("Incident")) {
                                c.setBackground(Color.RED);
                                c.setForeground(Color.WHITE);
                            }
                        }
                        else {
                            c.setBackground(Color.GREEN.darker().darker().darker());
                            c.setForeground(Color.WHITE);
                        }
                    }
                }
                if(column==columnResolved) {
                    if(table.getCellSelectionEnabled()==false) {
                        c.setBackground(Color.GREEN.darker().darker().darker());
                        c.setForeground(Color.WHITE);
                    }
                }
                
            }
        }
        else if(isSelected){// && column!=columnStatus && column!=columnRemarks) {
            c.setBackground(Color.BLUE);
            c.setForeground(Color.WHITE);
        }
        else {
            c.setBackground(Color.WHITE);
            c.setForeground(Color.BLACK);
        }
        
        /* CHECK IF TICKET IS ALREADY DUE */
        if(!table.getValueAt(row, columnStatus).toString().contains("Resolved") && !table.getValueAt(row, columnStatus).toString().toString().contains("Closed") && !table.getValueAt(row, columnStatus).toString().toString().contains("Cancelled") && !table.getValueAt(row, columnStatus).toString().contains("Solved")) {
            if(table.getCellSelectionEnabled()==false) {
                if(column==columnStatus || column==columnRemarks) {
                    if(!table.getValueAt(row, columnDueDate).toString().equalsIgnoreCase("")) {
                        Date systemDate = Calendar.getInstance().getTime();
                        String systemDateString = new SimpleDateFormat("yyyy-MM-dd").format(systemDate);
                        Date systemDateNoTime = new SimpleDateFormat("yyyy-MM-dd").parse(systemDateString.substring(0, 10));
                        Date dueDate = new SimpleDateFormat("yyyy-MM-dd").parse(table.getValueAt(row, columnDueDate).toString());//.substring(0, 10));
                        //System.out.println("row: "+row+" col: "+column+"--"+systemDate +" vs "+ dueDate);
                        if(systemDateNoTime.after(dueDate)) {
                            c.setBackground(Color.RED.darker());
                            //c.setForeground(Color.RED);
                        }
                        else if(systemDateNoTime.equals(dueDate)) {
                            //System.out.println(table.getValueAt(row, columnStatus).toString());
                            c.setBackground(Color.YELLOW.darker().darker());
                        }
                    }
                }
            }   
        }
        
        /* CHECK IF TICKET IS AN INCIDENT */
        if(table.getValueAt(row, columnClassification)!=null) {
            if(table.getValueAt(row, columnClassification).toString().contains("Incident")) {
                if(column==columnClassification) {
                    if(table.getCellSelectionEnabled()==false) {
                        c.setBackground(Color.RED);
                        c.setForeground(Color.WHITE);
                    }
                }
            }
        }
    }catch(Exception e) {
        System.out.println("CustomTableCellRenderer: "+e);
    }
   
    return c;
  }
}
