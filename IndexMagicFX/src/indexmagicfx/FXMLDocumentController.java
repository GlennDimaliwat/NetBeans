/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package indexmagicfx;

import java.net.URL;
import java.util.ResourceBundle;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.TextField;

/**
 *
 * @author Glenn Dimaliwat
 */
public class FXMLDocumentController implements Initializable {
    
    IndexMagicMethods imm = new IndexMagicMethods();
    
    @FXML
    private TextField filePath = new TextField();
    
    @FXML
    private void checkIndexesAction(ActionEvent event) {
        System.out.println("checkIndexesAction");
        imm.checkIndex(filePath.getText());
    }
    
    @FXML
    private void checkTableFieldsRapidMartAction(ActionEvent event) {
        System.out.println("checkTableFieldsRapidMartAction");
        imm.checkFields("RM_TARGET",filePath.getText());
    }
    
    @FXML
    private void checkTableFieldsDEFSAction(ActionEvent event) {
        System.out.println("checkTableFieldsDEFSAction");
        imm.checkFields("DWH_DEF_TARGET_PHASE_2",filePath.getText());
    }
    
    @FXML
    private void generateIndexScriptAction(ActionEvent event) {
        System.out.println("generateIndexScriptAction");
        System.out.println(filePath.getText());
        
        imm.generateIndexScript();
    }
    
    @Override
    public void initialize(URL url, ResourceBundle rb) {
        // TODO
    }    
}
