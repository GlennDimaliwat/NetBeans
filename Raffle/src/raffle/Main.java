/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package raffle;

import java.awt.Dimension;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedReader;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.List;
import javafx.embed.swing.JFXPanel;
import javax.swing.JOptionPane;
import javax.swing.Timer;
import java.util.Random;
import javax.swing.DefaultListCellRenderer;
import javax.swing.DefaultListModel;
import javax.swing.JLabel;
/**
 *
 * @author Glenn Dimaliwat
 */
public class Main extends javax.swing.JFrame {

    private int randomNumber;
    private int rotation = 0;
    private int delay = 10;
    private int maxRotation = 5;
    /**
     * Creates new form Main
     */
    public Main() {
        initComponents();
        initCustomComponents();
    }

    private void initCustomComponents() {
        JFXPanel fxPanel = new JFXPanel(); // Needed to initialize JavaFX Objects
        
        // Center Names on the List
        DefaultListCellRenderer renderer =  (DefaultListCellRenderer)raffleList.getCellRenderer();  
        renderer.setHorizontalAlignment(JLabel.CENTER);  

        // Initialize list
        loadValuesFromFile();
        
        // Center window
        Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
        this.setLocation(dim.width/2-this.getSize().width/2, dim.height/2-this.getSize().height/2);
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
        raffleList = new javax.swing.JList();
        jButton1 = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Raffle");

        raffleList.setModel(new javax.swing.AbstractListModel() {
            String[] strings = { "Item 1", "Item 2", "Item 3", "Item 4", "Item 5" };
            public int getSize() { return strings.length; }
            public Object getElementAt(int i) { return strings[i]; }
        });
        raffleList.setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));
        raffleList.setLayoutOrientation(javax.swing.JList.VERTICAL_WRAP);
        raffleList.setVisibleRowCount(17);
        jScrollPane1.setViewportView(raffleList);
        raffleList.getAccessibleContext().setAccessibleName("");

        jButton1.setFont(new java.awt.Font("Comic Sans MS", 1, 18)); // NOI18N
        jButton1.setText("DRAW RAFFLE");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jScrollPane1, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, 624, Short.MAX_VALUE)
            .addGroup(layout.createSequentialGroup()
                .addGap(227, 227, 227)
                .addComponent(jButton1)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 363, Short.MAX_VALUE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(jButton1)
                .addGap(6, 6, 6))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    
    private void loadValuesFromFile() {
            // Initialize list
            DefaultListModel raffleModel = new DefaultListModel();
            
            try (BufferedReader br = new BufferedReader(new FileReader("raffle.txt"))) {
                String line = br.readLine();
                while (line != null) {
                    try {
                        if(!line.equalsIgnoreCase("")) {
                            raffleModel.addElement(line);
                        }
                        line = br.readLine();
                    }
                    catch(ArrayIndexOutOfBoundsException | NullPointerException e) {}
                }
            }
            catch(IOException ioe) {
                ioe.printStackTrace();
            }
            
            // Set list model
            raffleList.setModel(raffleModel);
    }
    
    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        try {
            // Generate Random Number
            Random rand = new Random();
            int min = 0;
            int max = raffleList.getLastVisibleIndex();
            
            if(max>=0) {
                randomNumber = rand.nextInt( ( max - min ) + 1 ) + min;
                //System.out.println("RNG: "+randomNumber);

                // Begin the animation
                Timer timer = new Timer(delay, new ActionListener() {
                    @Override
                    public void actionPerformed(ActionEvent evt) {
                        //System.out.println("raffleList.getSelectedIndex() :"+raffleList.getSelectedIndex());
                        //System.out.println("raffleList.getLastVisibleIndex() :"+raffleList.getLastVisibleIndex());
                        int selectedIndex = raffleList.getSelectedIndex();
                        
                        //Disable Button
                        jButton1.setEnabled(false);
                        
                        // If selected row is last row
                        if(selectedIndex==raffleList.getLastVisibleIndex()) {
                            raffleList.setSelectedIndex(0);
                            rotation++;
                            delay+=10;
                            ((Timer)evt.getSource()).setDelay(delay);
                        }
                        else {
                            raffleList.setSelectedIndex(selectedIndex+1);
                        }

                        // Refresg GUI
                        raffleList.repaint();
                        raffleList.ensureIndexIsVisible(selectedIndex);

                        // If selected row is equal to RNG
                        if(selectedIndex==randomNumber && rotation == maxRotation) {
                            // Announce winner
                            JOptionPane.showMessageDialog(null, "Congratulations "+raffleList.getSelectedValue(), "We have a winner!", JOptionPane.PLAIN_MESSAGE);
                            String winnerName = raffleList.getSelectedValue().toString();
                            
                            // Remove winner from list
                            DefaultListModel model = (DefaultListModel) raffleList.getModel();
                            int winner = raffleList.getSelectedIndex();
                            if (winner != -1) {
                                model.remove(winner);
                            }
                            
                            // Remove winner from file
                            removeWinner(winnerName);
                            
                            // Reset rotation and delay
                            rotation = 0;
                            delay = 10;
                            
                            // Enable Button
                            jButton1.setEnabled(true);
                            
                            // Stop the timer
                            ((Timer)evt.getSource()).stop();
                        }
                    }
                });
                timer.start();
            }
            else {
                JOptionPane.showMessageDialog(null, "No more available winners. Please restart the app.", "Raffle", JOptionPane.ERROR_MESSAGE);
            }
        }catch(Exception e) {
            e.printStackTrace();
        }
    }//GEN-LAST:event_jButton1ActionPerformed

    private void removeWinner(String value) {
        List<String> valueList = new ArrayList<>();
        String line;
        
        try (FileReader fileReader = new FileReader("raffle.txt")) {
            BufferedReader bufferedReader = new BufferedReader(fileReader);
            while ((line = bufferedReader.readLine()) != null) {
                if(!line.equalsIgnoreCase(value) && !line.equalsIgnoreCase("")) {
                    valueList.add(line);
                }
            }
        }
        catch(IOException ioe) {
            ioe.printStackTrace();
        }
        
        try (FileWriter fileWriter = new FileWriter("raffle.txt"); PrintWriter out = new PrintWriter(fileWriter)) {
            for (String outputLine : valueList) {
                out.println(outputLine);
            }
            out.flush();
        }
        catch(IOException ioe) {
            ioe.printStackTrace();
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
            java.util.logging.Logger.getLogger(Main.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(Main.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(Main.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Main.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new Main().setVisible(true);
            }
        });
    }
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JList raffleList;
    // End of variables declaration//GEN-END:variables
}
