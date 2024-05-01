/*
 * Created by JFormDesigner on Fri Apr 26 17:20:51 MSK 2024
 */

package Lab2;

import javax.swing.*;
import javax.swing.GroupLayout;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.File;
import java.util.List;

/**
 * @author SystemX
 */
public class GUI extends JFrame {
    private String filePath;
    private int sheetNumber;

    public GUI() {
        initComponents();
        chooseFileButton.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            int returnValue = fileChooser.showOpenDialog(GUI.this);
            if (returnValue == JFileChooser.APPROVE_OPTION) {
                File selectedFile = fileChooser.getSelectedFile();
                if (selectedFile.getName().endsWith(".xlsx")) {
                    filePath = selectedFile.getAbsolutePath();
                    fileErrorLabel.setText(selectedFile.getName());
                    JOptionPane.showMessageDialog(GUI.this, "Excel file uploaded successfully!", "Success", JOptionPane.INFORMATION_MESSAGE);
                }
            } else
                JOptionPane.showMessageDialog(GUI.this, "Error with uploading file", "Error", JOptionPane.ERROR_MESSAGE);
        });
        chooseSheetNumberButton.addActionListener(e -> {
            try {
                sheetNumber = Integer.parseInt(sheetField.getText());
                ExcelParser.parse(filePath, sheetNumber);
                sheetErrorLabel.setText("Sheet number " + (sheetNumber));
            } catch (Exception ex) {
                sheetErrorLabel.setText("Error");
                JOptionPane.showMessageDialog(GUI.this, "Error with sheet number: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                sheetNumber = 0;
            }
        });
        generateExcelButton.addActionListener(e -> {
            if (sheetNumber != 0) try {
                JFileChooser saveFileChooser = new JFileChooser();
                saveFileChooser.setDialogTitle("Save Excel File");
                saveFileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
                saveFileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files (.xlsx)", "xlsx"));

                int saveReturnValue = saveFileChooser.showSaveDialog(GUI.this);
                if (saveReturnValue == JFileChooser.APPROVE_OPTION) {
                    File saveFile = saveFileChooser.getSelectedFile();
                    String saveFilePath = saveFile.getAbsolutePath().concat(".xlsx");

                    ExcelCreator.writeSamplesToExcel(ExcelParser.parse(filePath, sheetNumber), saveFilePath);

                    JOptionPane.showMessageDialog(GUI.this, "Excel file generated successfully!", "Success", JOptionPane.INFORMATION_MESSAGE);
                }
            } catch (Exception ex) {
                ex.printStackTrace();
                JOptionPane.showMessageDialog(GUI.this, "Error generating Excel file: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            }
            else
                JOptionPane.showMessageDialog(GUI.this, "Choose file and sheet number first", "Error", JOptionPane.ERROR_MESSAGE);
        });

        exitButton.addActionListener(e -> System.exit(0));
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        this.setVisible(true);

    }

    private void initComponents() {
        // JFormDesigner - Component initialization - DO NOT MODIFY  //GEN-BEGIN:initComponents  @formatter:off
        // Generated using JFormDesigner Evaluation license - Leonid
        chooseFileButton = new JButton();
        chooseSheetNumberButton = new JButton();
        generateExcelButton = new JButton();
        sheetField = new JTextField();
        fileErrorLabel = new JLabel();
        sheetErrorLabel = new JLabel();
        exitButton = new JButton();

        //======== this ========
        var contentPane = getContentPane();

        //---- chooseFileButton ----
        chooseFileButton.setText("Choose Excel file");

        //---- chooseSheetNumberButton ----
        chooseSheetNumberButton.setText("Choose sheet number");

        //---- generateExcelButton ----
        generateExcelButton.setText("Generate Excel file");

        //---- fileErrorLabel ----
        fileErrorLabel.setText("File isn't uploaded");

        //---- sheetErrorLabel ----
        sheetErrorLabel.setText("Sheet isn't choosen");

        //---- exitButton ----
        exitButton.setText("EXIT");

        GroupLayout contentPaneLayout = new GroupLayout(contentPane);
        contentPane.setLayout(contentPaneLayout);
        contentPaneLayout.setHorizontalGroup(
            contentPaneLayout.createParallelGroup()
                .addGroup(GroupLayout.Alignment.TRAILING, contentPaneLayout.createSequentialGroup()
                    .addGap(25, 25, 25)
                    .addGroup(contentPaneLayout.createParallelGroup()
                        .addComponent(chooseFileButton, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(chooseSheetNumberButton, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(generateExcelButton, GroupLayout.Alignment.TRAILING, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(sheetField, GroupLayout.Alignment.TRAILING))
                    .addGap(18, 18, 18)
                    .addGroup(contentPaneLayout.createParallelGroup()
                        .addComponent(fileErrorLabel)
                        .addComponent(sheetErrorLabel)
                        .addComponent(exitButton, GroupLayout.PREFERRED_SIZE, 123, GroupLayout.PREFERRED_SIZE))
                    .addGap(10, 10, 10))
        );
        contentPaneLayout.setVerticalGroup(
            contentPaneLayout.createParallelGroup()
                .addGroup(contentPaneLayout.createSequentialGroup()
                    .addGroup(contentPaneLayout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                        .addComponent(chooseFileButton, GroupLayout.PREFERRED_SIZE, 48, GroupLayout.PREFERRED_SIZE)
                        .addComponent(fileErrorLabel))
                    .addGap(18, 18, 18)
                    .addGroup(contentPaneLayout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                        .addComponent(chooseSheetNumberButton, GroupLayout.DEFAULT_SIZE, 54, Short.MAX_VALUE)
                        .addComponent(sheetErrorLabel))
                    .addGap(18, 18, 18)
                    .addComponent(sheetField, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
                    .addGap(30, 30, 30)
                    .addGroup(contentPaneLayout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                        .addComponent(generateExcelButton, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE)
                        .addComponent(exitButton, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE))
                    .addGap(22, 22, 22))
        );
        pack();
        setLocationRelativeTo(getOwner());
        // JFormDesigner - End of component initialization  //GEN-END:initComponents  @formatter:on
    }

    // JFormDesigner - Variables declaration - DO NOT MODIFY  //GEN-BEGIN:variables  @formatter:off
    // Generated using JFormDesigner Evaluation license - Leonid
    private JButton chooseFileButton;
    private JButton chooseSheetNumberButton;
    private JButton generateExcelButton;
    private JTextField sheetField;
    private JLabel fileErrorLabel;
    private JLabel sheetErrorLabel;
    private JButton exitButton;
    // JFormDesigner - End of variables declaration  //GEN-END:variables  @formatter:on
}
