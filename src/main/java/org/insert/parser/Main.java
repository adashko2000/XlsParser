package org.insert.parser;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.insert.parser.sheets.TemplateSheet;
import org.insert.parser.xlsParser.XlsParser;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.util.Arrays;
import java.util.List;

public class Main {
    public static void main(String[] args) {
        final long startTime = System.currentTimeMillis();
        final ObjectMapper om = new ObjectMapper();
        try {
            File folder = new File("./conf");
            if (!folder.exists()) {
                folder = new File("./src/main/resources");
            }
            final List<File> fileList = Arrays.asList(folder.listFiles());

            for (final File file : fileList) {
                final TemplateSheet sheet = om.readValue(file, TemplateSheet.class);
                XlsParser reader = new XlsParser();

                final String insertList = reader.convertToInsert(sheet);
                final BufferedWriter writer = new BufferedWriter(new FileWriter(sheet.outputFile));

                System.out.println("Table inserts \n" + insertList);

                writer.write(insertList);
                writer.close();
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        System.out.println("Run time" + (System.currentTimeMillis() - startTime));
    }
}