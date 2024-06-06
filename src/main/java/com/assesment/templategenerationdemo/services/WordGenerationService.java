package com.assesment.templategenerationdemo.services;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.List;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import jakarta.xml.bind.JAXBException;

/*
 * This service helps create a word from a template. 
 */
@Service
public class WordGenerationService {

    @Value("${template.templateUrlString}")
    private String templateUrlString;
    @Value("${template.pdfPathUrl}")
    private String pdfPathUrl;

    public void executeOperation() {

        WordprocessingMLPackage wordMLPackage = readTemplate(templateUrlString);
        

        if (wordMLPackage == null) {
            System.out.println("Template not found");
            return;
        }
        insertTable(wordMLPackage, 5, 6);
        SaveToDisk(wordMLPackage, pdfPathUrl);

    }

    public WordprocessingMLPackage readTemplate(String template) {

        try {
            File file = new ClassPathResource(template).getFile();
            System.out.println("Reading template: " + file.getAbsolutePath());

            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(file);

            // wordMLPackage.getDocPropsCorePart().getContents().setTitle("My title");
            wordMLPackage.setTitle("My title");

            return wordMLPackage;

        } catch (Docx4JException e) {
            System.out.println("Error reading template: " + e.getMessage());
        } catch (Exception e) {
            System.out.println("File not found: " + template);
        }
        return null;

    }

    public static void insertTable(WordprocessingMLPackage wordMLPackage, int rows, int cols) {
        MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();

        Tbl table = createTable(rows, cols);
        mainDocumentPart.addObject(table);
    }

    public static void modifyTable(MainDocumentPart mainDocumentPart, int tableIndex, int rows, int cols) {

        try {
            // Find all tables in the document
            List<Object> tables = mainDocumentPart.getJAXBNodesViaXPath("//w:tbl", false);

            if (tableIndex < tables.size()) {
                Tbl table = (Tbl) tables.get(tableIndex);
                System.out.println("Modifying table " + table.getContent());

                // Clear existing table content (optional, depending on whether you want to
                // completely replace the content)
                table.getContent().clear();

                // Add new rows and columns
                ObjectFactory factory = new ObjectFactory();
                for (int i = 0; i < rows; i++) {
                    Tr tableRow = factory.createTr();
                    for (int j = 0; j < cols; j++) {
                        Tc tableCell = factory.createTc();
                        P paragraph = factory.createP();
                        R run = factory.createR();
                        Text text = factory.createText();

                        text.setValue("Cell " + (i + 1) + "," + (j + 1));
                        run.getContent().add(text);
                        paragraph.getContent().add(run);
                        tableCell.getContent().add(paragraph);
                        tableRow.getContent().add(tableCell);
                    }
                    table.getContent().add(tableRow);
                }
            } else {
                System.out.println("Table index out of bounds.");
            }
        } catch (JAXBException e) {
            System.out.println("Error modifying table JaxB: "+e.getMessage());
        }catch (Exception e) {
            System.out.println("Error modifying table: " + e.getMessage());
        }
    }

    public static Tbl createTable(int rows, int cols) {

        ObjectFactory factory = new ObjectFactory();
        Tbl table = factory.createTbl();

        for (int i = 0; i < rows; i++) {
            Tr tableRow = factory.createTr();

            for (int j = 0; j < cols; j++) {

                Tc tableCell = factory.createTc();

                // Create a paragraph for the cell
                P paragraph = factory.createP();

                // Create a run for the text
                R run = factory.createR();

                // Create the text element and set the content
                Text text = factory.createText();
                text.setValue("Cell " + (i + 1) + "," + (j + 1));

                // Add the text to the run
                run.getContent().add(text);

                // Add the run to the paragraph
                paragraph.getContent().add(run);

                // Add the paragraph to the table cell
                tableCell.getContent().add(paragraph);

                // Add the cell to the row
            }
            table.getContent().add(tableRow);
        }

        return table;
    }

    public void SaveToDisk(WordprocessingMLPackage wordMLPackage, String destinationUrl) {

        try {

            File file = new File(pdfPathUrl);
            FileOutputStream fos = new FileOutputStream(file);

            System.out.println("Saving to PDF: " + file.getAbsolutePath());

            wordMLPackage.save(fos);

            System.out.println("File saved to: " + destinationUrl);
        } catch (FileNotFoundException e) {
            System.out.println("File not found: " + destinationUrl);
        } catch (Docx4JException e) {
            System.out.println("Error saving file: " + e.getMessage());
        }

    }

}
