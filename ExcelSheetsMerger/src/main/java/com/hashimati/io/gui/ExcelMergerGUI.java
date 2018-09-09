package com.hashimati.io.gui;

import com.hashimati.io.merger.Merger;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.util.ArrayList;

public class ExcelMergerGUI extends JFrame
{

    JFrame me = this;
    JFileChooser inputFilesDialog = new JFileChooser("");
    JTextField source = new JTextField("Please, choose files "), destination = new JTextField("Please, click save " +
            "button");

    JTextArea logs = new JTextArea("");
    JScrollPane scrollPane = new JScrollPane(logs);

    JProgressBar progress = new JProgressBar(JProgressBar.HORIZONTAL, 0, 100);
    JFileChooser outFiles = new JFileChooser("");
    JButton selectButton = new JButton("Choose Files");
    JButton saveButton = new JButton( "save");
    JButton startProcessing = new JButton("Start");
    private ArrayList<File> inputFiles = new ArrayList<File>();
    private String outputPath;
    private String outFile;
    JPanel panel = new JPanel();
    public ExcelMergerGUI() throws ClassNotFoundException, UnsupportedLookAndFeelException, InstantiationException, IllegalAccessException {

        init();
        this.setSize(600, 600);
        this.setTitle("Excel Merger");
        this.setVisible(true);
        this.setResizable(false);

        inputFilesDialog.removeChoosableFileFilter(inputFilesDialog.getAcceptAllFileFilter());
        inputFilesDialog.addChoosableFileFilter(new FileNameExtensionFilter("Spreadsheets", "xls","xlsx"));
        outFiles.removeChoosableFileFilter(outFiles.getAcceptAllFileFilter()); 
        outFiles.addChoosableFileFilter(new FileNameExtensionFilter("xls", "xls"));
        outFiles.addChoosableFileFilter(new FileNameExtensionFilter("xlsx", "xlsx"));

        UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        SwingUtilities.updateComponentTreeUI(me);

        me.setIconImage(new BufferedImage(3,3,3));
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    }
    public void init(){

        inputFilesDialog.setMultiSelectionEnabled(true);
        JPanel buttonPanel = new JPanel();
        buttonPanel.add(selectButton);
        buttonPanel.add(saveButton);
        buttonPanel.add(startProcessing);

        JPanel panelLabels = new JPanel();

        panelLabels.setLayout(new BoxLayout(panelLabels, BoxLayout.Y_AXIS));
        JPanel sourcePanel = new JPanel();
        sourcePanel.setLayout(new FlowLayout(FlowLayout.LEADING));
        sourcePanel.add(new JLabel("Source"))
                ;
        source.setEditable(false);
        sourcePanel.add(source);

        JPanel desinationPanel = new JPanel();

        desinationPanel.setLayout(new FlowLayout(FlowLayout.LEADING));

        desinationPanel.add(new JLabel("Destination"));
        destination.setEditable(false);
        desinationPanel.add(destination);




        panelLabels.add(sourcePanel);
        panelLabels.add(desinationPanel);


        saveButton.addActionListener(a->{
try {
    outFiles.showSaveDialog(me);

    SwingUtilities.updateComponentTreeUI(outFiles);


    outputPath = outFiles.getSelectedFile().getParent();
    outFile = outFiles.getFileFilter().getDescription().equalsIgnoreCase("xls") ? (

            outFiles.getSelectedFile().getName().endsWith(".xls") ? outFiles.getSelectedFile().getName() :
                    outFiles.getSelectedFile().getName() + ".xls") : (
            outFiles.getSelectedFile().getName().endsWith(".xlsx") ?
                    outFiles.getSelectedFile().getName() : outFiles.getSelectedFile().getName() + ".xlsx");
    ;
    System.out.println(outputPath + "\\" + outFile);

    destination.setVisible(true);
    destination.setSize(400, source.getPreferredSize().height);
    destination.setText(outFiles.getSelectedFile().toString());

}
catch(Exception ex)
{

    JOptionPane.showMessageDialog(me, "Error");
}
        });
        selectButton.addActionListener(x->{

            inputFilesDialog.showOpenDialog(me);
            SwingUtilities.updateComponentTreeUI(inputFilesDialog);
            inputFiles = new ArrayList<File>();
            String inputString = "[";
            for(File f : inputFilesDialog.getSelectedFiles())
            {
                inputFiles.add(f);
                inputString +=f.toString() + ",";
                System.out.println(f);
            }
            inputString= inputString.substring(0, inputString.length()-1)+ "]";
            System.out.println(inputString);
            source.setVisible(true);
            source.setSize(400, source.getPreferredSize().height);
            source.setText(inputString);
            System.out.println(source.getText());
        });


        startProcessing.addActionListener(a->{
            try {
                if(!ableToStart())
                {
                    JOptionPane.showMessageDialog(me, "Specify the " +
                            "excel files and destination files","Error", JOptionPane.ERROR_MESSAGE);
                    return;
                }

                Merger merger = new Merger(outputPath, outFile);

                System.out.println(inputFiles);

                new Thread(new Runnable() {
                    @Override
                    public void run() {
                         merger.merger(inputFiles, 0);
                    }
                }).start();
                new Thread(new Runnable() {
                    @Override
                    public void run() {
                        selectButton.setEnabled(false);
                        saveButton.setEnabled(false);
                        startProcessing.setEnabled(false);
                        while(!merger.isDone()){

                            logs.setText(merger.getLogs());
                            scrollPane.getVerticalScrollBar().setValue(scrollPane.getVerticalScrollBar().getMaximum());
                            progress.setValue(merger.getProgress());
                        }
                        selectButton.setEnabled(true);
                        saveButton.setEnabled(true);
                        startProcessing.setEnabled(true);
                    }
                }).start();
            }
            catch (Exception ex)
            {
                ex.printStackTrace();
            }
        });
        panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));
        panel.add(buttonPanel);
        panel.add(panelLabels);
        this.setLayout(new BorderLayout());
        this.add(panel, BorderLayout.NORTH);


        logs.setEditable(false);
        scrollPane.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS);

        this.add(scrollPane, BorderLayout.CENTER);

        this.add(progress, BorderLayout.SOUTH);
    }

    private boolean ableToStart() {
        return (outputPath!= null) && !outputPath.trim().isEmpty()
                && outFile!=null && !outFile.trim().isEmpty()
                && inputFiles != null && !inputFiles.isEmpty()
                ;
    }
}
