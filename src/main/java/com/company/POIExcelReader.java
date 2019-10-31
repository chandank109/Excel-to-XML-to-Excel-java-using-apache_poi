package com.company;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class POIExcelReader {

    private Element rootElement;
    private Document doc;

    public POIExcelReader() {

    }

    /**
     * Read Excel File
     *
     * @param xlsPath
     */
    public void displayFromExcel(String xlsPath) {
        //obtaining input bytes from a file
        try {
            File myFile = new File(xlsPath);
            FileInputStream fis = new FileInputStream(myFile);
            // Finds the workbook instance for XLSX file
            XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);
            // Return first sheet from the XLSX workbook
            XSSFSheet mySheet = myWorkBook.getSheetAt(0);
            // Get iterator to all the rows in current sheet
            Iterator<Row> rowIterator = mySheet.iterator();

            //start of xml
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder;

            dBuilder = dbFactory.newDocumentBuilder();
            doc = dBuilder.newDocument();
            rootElement =
                    doc.createElementNS("", "resources");
            //append root element to document
            doc.appendChild(rootElement);


            // Traversing over each row of XLSX file
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                // For each row, iterate through each columns
                Iterator<Cell> cellIterator = row.cellIterator();
                int index = 0;
                String key = "";
                String value = "";
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    if (index == 0) {
                        index = 1;
                        key = cell.getStringCellValue();
                        System.out.print(cell.getStringCellValue() + "-->");
                    } else {
                        index = 0;
                        System.out.print(cell.getStringCellValue() + "\n");
                        value = cell.getStringCellValue();
                    }
                }

                //Appending key value for xml with node
                rootElement.appendChild(getEmployeeElements(doc, key, value));
            }


            //For Writing in xml file
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer transformer = transformerFactory.newTransformer();
            //for pretty print
            transformer.setOutputProperty(OutputKeys.INDENT, "yes");
            transformer.setOutputProperty(OutputKeys.METHOD, "xml");
            DOMSource source = new DOMSource(doc);

            //write to console or file
            StreamResult console = new StreamResult(System.out);
            StreamResult file = new StreamResult(new File("D:/Chandan/RE Documents/strings.xml"));

            //write data
            transformer.transform(source, console);
            transformer.transform(source, file);
            System.out.println("DONE");


        } catch (IOException e) {
            e.printStackTrace();
        } catch (TransformerConfigurationException e) {
            e.printStackTrace();
        } catch (TransformerException e) {
            e.printStackTrace();
        } catch (ParserConfigurationException e) {
            e.printStackTrace();
        }
    }

    //utility method to create text node
    private static Node getEmployeeElements(Document doc, String key, String value) {
        Element node = doc.createElement("string");
        node.setAttribute("name", key);
        node.appendChild(doc.createTextNode(value));
        return node;
    }
}
