/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package kboxviewer2;

import java.awt.Color;
import java.awt.Component;
import java.util.Date;
import java.text.SimpleDateFormat;
import java.util.Calendar;
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

    final int dueDateField = 4;
    final int classificationField = 6;
    final int statusField = 8;
    final int reasonForPendingField = 9;
    
    try {
        //System.out.println("row: "+row+" col: "+column+" value: "+value);
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
        if(table.getValueAt(row, statusField).toString().equalsIgnoreCase("Pending") || table.getValueAt(row, statusField).toString().equalsIgnoreCase("Accepted") || table.getValueAt(row, statusField).toString().equalsIgnoreCase("Assigned")) {
            if(column==statusField || column==reasonForPendingField) {
                if(table.getCellSelectionEnabled()==false) {
                    c.setBackground(Color.BLUE.darker()); 
                    c.setForeground(Color.WHITE);
                }
            }
        }
        else if(table.getValueAt(row, statusField).toString().contains("Resolved") || table.getValueAt(row, statusField).toString().toString().contains("Closed") || table.getValueAt(row, statusField).toString().toString().contains("Cancelled")) {
            if(column==statusField || column==reasonForPendingField) {
                if(table.getCellSelectionEnabled()==false) {
                    c.setBackground(Color.GREEN.darker().darker().darker());
                    c.setForeground(Color.WHITE);
                }
            }
        }
        else if(isSelected){// && column!=statusField && column!=reasonForPendingField) {
            c.setBackground(Color.BLUE);
            c.setForeground(Color.WHITE);
        }
        else {
            c.setBackground(Color.WHITE);
            c.setForeground(Color.BLACK);
        }
        
        /* CHECK IF TICKET IS ALREADY DUE */
        if(!table.getValueAt(row, statusField).toString().contains("Resolved") && !table.getValueAt(row, statusField).toString().toString().contains("Closed") && !table.getValueAt(row, statusField).toString().toString().contains("Cancelled")) {
            if(table.getCellSelectionEnabled()==false) {
                if(column==statusField || column==reasonForPendingField) {
                    Date systemDate = Calendar.getInstance().getTime();
                    String systemDateString = new SimpleDateFormat("yyyy-MM-dd").format(systemDate);
                    Date systemDateNoTime = new SimpleDateFormat("yyyy-MM-dd").parse(systemDateString.substring(0, 10));
                    Date dueDate = new SimpleDateFormat("yyyy-MM-dd").parse(table.getValueAt(row, dueDateField).toString().substring(0, 10));
                    //System.out.println("row: "+row+" col: "+column+"--"+systemDate +" vs "+ dueDate);
                    if(systemDateNoTime.after(dueDate)) {
                        c.setBackground(Color.RED.darker());
                        //c.setForeground(Color.RED);
                    }
                    else if(systemDateNoTime.equals(dueDate)) {
                        //System.out.println(table.getValueAt(row, statusField).toString());
                        c.setBackground(Color.YELLOW.darker().darker());
                    }
                }
            }   
        }
        
        /* CHECK IF TICKET IS AN INCIDENT */
        if(table.getValueAt(row, classificationField).toString().contains("Incident")) {
            if(column==classificationField) {
                if(table.getCellSelectionEnabled()==false) {
                    c.setBackground(Color.RED);
                    c.setForeground(Color.WHITE);
                }
            }
        }
    }catch(Exception e) {
        System.out.println(e);
    }
   
    return c;
  }
}
