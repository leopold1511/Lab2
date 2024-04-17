package Lab2;


import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;


import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class GUI extends JFrame {

    private JPanel rootPane;
    private JButton chooseFileButton;
    private JLabel fileErrorLabel;
    private JButton chooseSheetNumberButton;
    private JTextField sheetField;
    private JLabel sheetErrorLabel;
    private JButton generateExcelButton;
    private int sheetNumber;
    private String filePath;

    public GUI() {
        super("Statistics");
        setContentPane(rootPane);
        chooseFileButton.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            int returnValue = fileChooser.showOpenDialog(GUI.this);
            if (returnValue == JFileChooser.APPROVE_OPTION) {
                File selectedFile = fileChooser.getSelectedFile();
                if (selectedFile.getName().endsWith(".xlsx")) {
                    filePath = selectedFile.getAbsolutePath();
                    fileErrorLabel.setText("Success");
                }
                fileErrorLabel.setText("Wrong format");
            } else fileErrorLabel.setText("Error");
        });
        chooseSheetNumberButton.addActionListener(e -> {
            try {
                ExcelParser.parse(filePath, Integer.parseInt(sheetField.getText()));
                sheetNumber = Integer.parseInt(sheetField.getText());
                sheetErrorLabel.setText("Success");
            } catch (Exception ex) {
                sheetErrorLabel.setText("Error");
            }
        });
        generateExcelButton.addActionListener(e -> {
            try {
                JFileChooser saveFileChooser = new JFileChooser();
                saveFileChooser.setDialogTitle("Save Excel File");
                saveFileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
                saveFileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files (xlsx)", "xlsx"));

                int saveReturnValue = saveFileChooser.showSaveDialog(GUI.this);
                if (saveReturnValue == JFileChooser.APPROVE_OPTION) {
                    File saveFile = saveFileChooser.getSelectedFile();
                    String saveFilePath = saveFile.getAbsolutePath().concat(".xlsx");

                    Map<Integer, List<Double>> samples = ExcelParser.parse(filePath, sheetNumber);
                    List<Sample> samples1 = new ArrayList<>();
                    for (int i = 0; i < samples.size(); i++) {
                        samples1.add(new Sample(samples.get(i), i));
                    }
                    ExcelCreator writer = new ExcelCreator();
                    writer.writeSamplesToExcel(samples1, saveFilePath);

                    JOptionPane.showMessageDialog(GUI.this, "Excel file generated successfully!", "Success", JOptionPane.INFORMATION_MESSAGE);
                }
            } catch (Exception ex) {
                ex.printStackTrace();
                JOptionPane.showMessageDialog(GUI.this, "Error generating Excel file: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            }
        });


        pack();

        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        //this.setSize(200, 200);
        this.setVisible(true);
    }

}



