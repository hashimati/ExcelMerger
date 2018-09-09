package com.hashimati.io.merger;

import com.hashimati.io.util.Util;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;

public class Merger
{

    public boolean isOXML(String file){
        return file.toLowerCase().endsWith("xlsx");
    }
    private String outputFolder, outputFile, absolutePath;

    //this will be used to write to the rows.
    private AtomicInteger outputRowIterator = new AtomicInteger(1);
    private FileOutputStream fos ;
    private ArrayList<String> failedToMergeFiles = new ArrayList<>();

    private int totalFiles=0, completeedFiles = 0;
    private int progress = 0;
    private boolean done = false;
    private String logs = "";

    private ArrayList<String> metadata;
    public Merger(String outputFolder, String outputFile){
         setOutputFile(outputFile);
         setOutputFolder(outputFolder);
         absolutePath =  Util.isOsWin()?outputFolder+"\\" + outputFile: outputFolder +"/" + outputFile;
        try {
            System.out.println("This Abslute Path" + absolutePath);
            fos = new FileOutputStream(absolutePath);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }

    public Merger(String outputPath){
        absolutePath =  outputPath;
    try {
            fos = new FileOutputStream(absolutePath);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }

    public Workbook createNewWorkBook(){



        Workbook wb = isOXML(absolutePath)?new XSSFWorkbook(): new HSSFWorkbook();
        Sheet sheet = wb.createSheet("Output");

        return wb;

    }


    public boolean compareHeaders(ArrayList<File> filesPaths, int sheetNo)
    {

        totalFiles = filesPaths.size();
        AtomicBoolean endResult = new AtomicBoolean(true);

        logs +="Start Comparing Files\n";
        try {
            FileInputStream fi0= new FileInputStream(filesPaths.get(0));
            Workbook workbook = isOXML(filesPaths.get(0).getAbsolutePath())?
                    new XSSFWorkbook(fi0):new HSSFWorkbook(fi0);
            Sheet sheet = workbook.getSheetAt(sheetNo);
            metadata = new ArrayList<String>(){
                {
                    sheet.getRow(0).forEach(cell -> add(cell.getStringCellValue()));
                }

            };
            fi0.close();

            AtomicBoolean endJob = new AtomicBoolean(false);
            filesPaths.stream().forEach(
                    file-> {
                        try{

                            FileInputStream fi =null;
                            if(!endJob.get())



                        fi = new FileInputStream(file);
                            Workbook workbook1 = isOXML(file.getAbsolutePath())?
                                    new XSSFWorkbook(fi):new HSSFWorkbook(fi);
                            Sheet sheet0 = workbook1.getSheetAt(sheetNo);
                            AtomicInteger index = new AtomicInteger(0);

                            sheet0.getRow(0).forEach(cell->{

                                if(!cell.getStringCellValue().equalsIgnoreCase(metadata.get(index.getAndIncrement()))){
                                    endJob.set(true);
                                    logs +="The format of File " + file.toString() +" matches with files\n";
                                }
                                if(endJob.get())
                                {
                                    endResult.set(false);
                                    logs +="The format of File " + file.toString() +" doesn't match with files\n";
                                }

                            }
                            );
                            fi.close();
                        } catch (FileNotFoundException e) {
                            e.printStackTrace();
                        } catch (IOException e) {
                            e.printStackTrace();
                        }


                    }
            );
        } catch (IOException e) {
            e.printStackTrace();
        }

        return endResult.get();
    }

    public boolean merger(ArrayList<File> filesPaths,int sheetNo)
    {
        System.out.println("The Proccess will start on these files:\n*********************");
        logs +="Starting the process of files\n";

        filesPaths.forEach(System.out::println);
        System.out.println(absolutePath);
        System.out.println("**************************");
        AtomicBoolean result = new AtomicBoolean(true);
        if(compareHeaders(filesPaths,sheetNo))
        {


            Workbook output = createNewWorkBook();
            if( output != null){
                Row header = output.getSheet("Output").createRow(0);
                logs +="Creating the consolidated sheet with name \"Output\".\n";

                AtomicInteger headerIndex= new AtomicInteger();
                System.out.println(metadata);
                for(String md:metadata) {
                    logs +="Createing the headers of the cosolidated sheet.\n";

                    Cell cell = header.createCell(headerIndex.getAndIncrement());

                    cell.setCellValue(md);
                    logs +="Creating the headers is completed.\n";


                }
                filesPaths.stream().forEach(file-> {
                    logs +="Starting merging files\n";

                    try {
                        logs +="Staring to merge " + file.toString() +"\n";

                        result.set(
                                        result.get() && singleFileMerger(file, sheetNo, output)
                                        );
                        logs +="Complete merging " + file.toString() +".\n";

                    } catch (IOException e) {
                                e.printStackTrace();
                        logs +="Failed to merge file " + file.toString() +" \n";

                    }
                        }
                );
                try {
                    logs +="Start writing the output to the destinatmion file " + absolutePath+ "\n";

                    output.write(fos);
                    fos.close();
                    logs +="The process is completed";

                } catch (IOException e) {
                    e.printStackTrace();
                }
            }else
            {

                result.set(false);
                done = true;
                logs +="Failed \n";

            }

        }
        try {
            fos.close();
            done = true;
        } catch (IOException e) {
            e.printStackTrace();
        }
        return result.get();
    }
    public boolean singleFileMerger(File filePath, int sheetNo, Workbook output) throws IOException {
        try {
            AtomicInteger innerInc = new AtomicInteger(0);
            Workbook inputWorkBook = isOXML(filePath.getAbsolutePath())?
                    new XSSFWorkbook(new FileInputStream(filePath)):new HSSFWorkbook(new FileInputStream(filePath));

            Sheet inputSheet = inputWorkBook.getSheetAt(sheetNo);

            inputSheet.forEach(r-> {
                if(innerInc.get()> 0)
                {
                    Row outputRow = output.getSheet("Output").createRow(outputRowIterator.getAndIncrement());
                    AtomicInteger cIndex = new AtomicInteger();
                    r.forEach(c->{
                        Cell outputCell = outputRow.createCell(c.getColumnIndex(), c.getCellType());
//                        outputCell.setCellStyle(c.getCellStyle());
                      //  outputCell.setCellFormula(c.getCellFormula());
                        switch (outputCell.getCellType())
                        {
                            case Cell.CELL_TYPE_BOOLEAN:
                                ;
                                outputCell.setCellValue(c.getBooleanCellValue());

                                break;
                             case Cell.CELL_TYPE_NUMERIC:
                                outputCell.setCellValue(c.getNumericCellValue());
                                break;
                            case Cell.CELL_TYPE_FORMULA:
                                outputCell.setCellValue(c.getCellFormula());
                                break;
                            case Cell.CELL_TYPE_STRING:
                                outputCell.setCellValue(c.getStringCellValue());
                                break;
                            case Cell.CELL_TYPE_ERROR:
                                outputCell.setCellValue(c.getErrorCellValue());
                                break;
                            case Cell.CELL_TYPE_BLANK:


                        }
                    });
                }
                innerInc.incrementAndGet();
            });

            //fos.close();
            completeedFiles++;
            progress= totalFiles == completeedFiles?100:(int)(((double)completeedFiles/(double) totalFiles)*100);

            return true;
        } catch (IOException e) {
            failedToMergeFiles.add(filePath.toString());
            e.printStackTrace();
        }
       // fos.close();
        return false;
    }
    public String getOutputFolder() {
        return outputFolder;
    }
    public void setOutputFolder(String outputFolder) {
        this.outputFolder = outputFolder;
    }

    public String getOutputFile() {
        return outputFile;
    }

    public void setOutputFile(String outputFile) {
        this. outputFile = outputFile;
    }

    public String getLogs() {
        return logs;
    }

    public void setLogs(String logs) {
        this.logs = logs;
    }

    public boolean isDone() {
        return done;
    }

    public void setDone(boolean done) {
        this.done = done;
    }

    public int getProgress() {
        return progress;
    }
}
