package com.company;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.w3c.dom.Document;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.io.FileOutputStream;

public class Main {

    public static void main(String[] args) {

        POIExcelReader poiExcelReader=new POIExcelReader();
        poiExcelReader.displayFromExcel("D:/Chandan/RE Documents/LanguageXML.xls");

        //        xmlToExcel();

    }

    private static void xmlToExcel() {

        String FILE_NAME="LanguageXML";
        String STRING_FILE_LOCATION="D:/IntellijProjects/Documents/strings.xml";

        try {

            XSSFWorkbook workbook = new XSSFWorkbook();
            //Create a blank sheet
            XSSFSheet spreadsheet = workbook.createSheet(FILE_NAME);
            //Create row object
            XSSFRow row;

            File fXmlFile = new File(STRING_FILE_LOCATION);
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            Document doc = dBuilder.parse(fXmlFile);

            //optional, but recommended
            //read this - http://stackoverflow.com/questions/13786607/normalization-in-dom-parsing-with-java-how-does-it-work
            doc.getDocumentElement().normalize();

            System.out.println("Root element :" + doc.getDocumentElement().getNodeName());

            NodeList nList = doc.getElementsByTagName("string");
            System.out.println("List Size:::: " + nList.getLength());

            row = spreadsheet.createRow(0);

            XSSFCell cell = row.createCell(0);
            cell.setCellValue("KEY");
            XSSFCell cell1 = row.createCell(1);
            cell1.setCellValue("VALUE");

            makeRowBold(workbook,row);

            for (int temp = 0; temp < nList.getLength(); temp++) {

                Node nNode = nList.item(temp);
                System.out.println("\nCurrent Element :" + nNode.getNodeName());
                String node_key = "";
                String node_value = "";
                if (nNode.getNodeType() == Node.ELEMENT_NODE) {
                    if (nNode.hasAttributes()) {
                        // get attributes names and values
                        NamedNodeMap nodeMap = nNode.getAttributes();
                        for (int i = 0; i < nodeMap.getLength(); i++) {
                            Node node = nodeMap.item(i);
                            node_key = node.getNodeValue();
                            System.out.println("Node key : " + node.getNodeValue());
                        }
                        System.out.println("Node Value : " + nNode.getTextContent());
                        node_value = nNode.getTextContent();
                    }
                }
                row = spreadsheet.createRow(temp + 1);
                cell = row.createCell(0);
                cell.setCellValue(node_key);
                cell1 = row.createCell(1);
                cell1.setCellValue(node_value);
            }


            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(
                    new File("D:/Chandan/RE Documents/LanguageXML.xls"));

            workbook.write(out);
            out.close();

            System.out.println("LanguageXML.xls written successfully");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void makeRowBold(Workbook wb, Row row){
        CellStyle style = wb.createCellStyle();//Create style
        Font font = wb.createFont();//Create font
        font.setBold(true);//Make font bold
        style.setFont(font);//set it to bold

        for(int i = 0; i < row.getLastCellNum(); i++){//For each cell in the row
            row.getCell(i).setCellStyle(style);//Set the style
        }
    }

}
