/*
 * Created by JFormDesigner on Fri Apr 26 17:20:51 MSK 2024
 */

package Lab2;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.GroupLayout;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Objects;

/**
 * @author SystemX
 */
public class GUI extends JFrame {
    private String filePath;
    private String sheetName;


    public GUI() {
        initComponents();
        chooseFileButton.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setCurrentDirectory(new File("."));
            int returnValue = fileChooser.showOpenDialog(GUI.this);
            if (returnValue == JFileChooser.APPROVE_OPTION) {
                File selectedFile = fileChooser.getSelectedFile();
                if (selectedFile.getName().endsWith(".xlsx")) {
                    filePath = selectedFile.getAbsolutePath();
                    fileErrorLabel.setText(selectedFile.getName());
                    JOptionPane.showMessageDialog(GUI.this, "Excel file uploaded successfully!", "Success", JOptionPane.INFORMATION_MESSAGE);
                    try (FileInputStream fis = new FileInputStream(filePath); Workbook workbook = new XSSFWorkbook(fis)) {
                        for (Sheet sheet : workbook)
                            sheetBox.addItem(sheet.getSheetName());
                    } catch (IOException ex) {
                        ex.printStackTrace();
                        JOptionPane.showMessageDialog(GUI.this, "Error with uploading file", "Error", JOptionPane.ERROR_MESSAGE);
                    }
                }
            } else {
                sheetErrorLabel.setText(sheetName);
                JOptionPane.showMessageDialog(GUI.this, "Error with uploading file", "Error", JOptionPane.ERROR_MESSAGE);
            }
            });
        chooseSheetNumberButton.addActionListener(e -> {
            try {
                sheetName = (String) sheetBox.getSelectedItem();
                sheetErrorLabel.setText(sheetName);
            } catch (Exception ex) {
                JOptionPane.showMessageDialog(GUI.this, "Error with sheet number: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                sheetName = " ";
            }

        });
        generateExcelButton.addActionListener(e -> {
            if (!Objects.equals(sheetName, " ")) try {
                JFileChooser saveFileChooser = new JFileChooser();
                saveFileChooser.setCurrentDirectory(new File("."));
                saveFileChooser.setDialogTitle("Save Excel File");
                saveFileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
                saveFileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files (.xlsx)", "xlsx"));

                int saveReturnValue = saveFileChooser.showSaveDialog(GUI.this);
                if (saveReturnValue == JFileChooser.APPROVE_OPTION) {
                    File saveFile = saveFileChooser.getSelectedFile();
                    String saveFilePath = saveFile.getAbsolutePath().concat(".xlsx");

                    ExcelCreator.writeSamplesToExcel(ExcelParser.parse(filePath, sheetName), saveFilePath);

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
        fileErrorLabel = new JLabel();
        sheetErrorLabel = new JLabel();
        exitButton = new JButton();
        sheetBox = new JComboBox<>();

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
                        .addComponent(generateExcelButton, GroupLayout.Alignment.TRAILING, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(chooseSheetNumberButton, GroupLayout.Alignment.TRAILING, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(chooseFileButton, GroupLayout.Alignment.TRAILING, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                    .addGap(18, 18, 18)
                    .addGroup(contentPaneLayout.createParallelGroup()
                        .addComponent(exitButton, GroupLayout.PREFERRED_SIZE, 123, GroupLayout.PREFERRED_SIZE)
                        .addComponent(sheetErrorLabel)
                        .addComponent(fileErrorLabel))
                    .addGap(10, 10, 10))
                .addGroup(contentPaneLayout.createSequentialGroup()
                    .addGap(57, 57, 57)
                    .addComponent(sheetBox, GroupLayout.PREFERRED_SIZE, 109, GroupLayout.PREFERRED_SIZE)
                    .addContainerGap(GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        contentPaneLayout.setVerticalGroup(
            contentPaneLayout.createParallelGroup()
                .addGroup(contentPaneLayout.createSequentialGroup()
                    .addContainerGap()
                    .addGroup(contentPaneLayout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                        .addComponent(chooseFileButton, GroupLayout.PREFERRED_SIZE, 48, GroupLayout.PREFERRED_SIZE)
                        .addComponent(fileErrorLabel))
                    .addGap(18, 18, 18)
                    .addComponent(sheetBox, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
                    .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                    .addGroup(contentPaneLayout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                        .addComponent(sheetErrorLabel)
                        .addComponent(chooseSheetNumberButton, GroupLayout.DEFAULT_SIZE, 54, Short.MAX_VALUE))
                    .addGap(22, 22, 22)
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
    private JLabel fileErrorLabel;
    private JLabel sheetErrorLabel;
    private JButton exitButton;
    private JComboBox<String> sheetBox;
    // JFormDesigner - End of variables declaration  //GEN-END:variables  @formatter:on


}
