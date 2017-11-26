package com.insightr.filesplitter;

import com.monitorjbl.xlsx.StreamingReader;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public class XlsxFileSplitter {
    private static final Logger logger = Logger.getLogger(XlsxFileSplitter.class);

    public static void main(String[] args) {
        String log4jConfPath = "log4j.properties";
        PropertyConfigurator.configure(log4jConfPath);

        String sourceFile = "D:/dev/projects/FileSplitter/src/main/resources/large-excel-file.xlsx";
        String targetDirectory = "D:/TEMP";
        try {
            XlsxFileSplitter xlsxFileSplitter = new XlsxFileSplitter();
            xlsxFileSplitter.split(sourceFile, targetDirectory, 1500);
        } catch (FileNotFoundException fnfe) {
            logger.error("File " + sourceFile + " not found !", fnfe);
        } catch (IOException ioe) {
            logger.error("Error while readinf the file !", ioe);
        }
    }

    private void split(String sourceFile, String targetDirectory, int rowsByFile) throws IOException {
        InputStream is = new FileInputStream(new File(sourceFile));
        Workbook workbook = StreamingReader.builder()
                .rowCacheSize(100)    // number of rows to keep in memory (defaults to 10)
                .bufferSize(4096)     // buffer size to use when reading InputStream to file (defaults to 1024)
                .open(is);            // InputStream or File for XLSX file (required)

        Sheet sheet = workbook.getSheet("Feuille numero 1");
        logger.info(sheet.getSheetName());

        if (targetDirectory.endsWith("/")) {
            targetDirectory += "/";
        }
        Path path = Paths.get(targetDirectory + "file0.csv");
        BufferedWriter writer = Files.newBufferedWriter(path);
        int counter = 1;
        int fileCounter = 1;
        for (Row r : sheet) {
            if (counter % rowsByFile == 0) {
                writer.close();
                writer = Files.newBufferedWriter(Paths.get("D:/TEMP/file0" + fileCounter + ".csv"));
                logger.info("génération du fichier " + "D:/TEMP/file0" + fileCounter + ".csv");
                fileCounter++;
            }
            StringBuilder string = new StringBuilder();
            for (Cell c : r) {
                string.append(c.getStringCellValue());
                string.append(";");
            }
            writer.write(string.toString());
            writer.newLine();
            counter++;
        }
    }

}
