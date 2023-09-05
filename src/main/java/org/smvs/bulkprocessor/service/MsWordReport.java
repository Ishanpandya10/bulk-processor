package org.smvs.bulkprocessor.service;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblLayoutType;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

public class MsWordReport {

    static List<String> columnNames = Arrays.asList("Member Id", "FID", "SG", "SC", "AG", "AP", "Full Name", "Mobile", "Whatsapp", "Gender", "House No");
    static List<String> columnKeys = Arrays.asList("ID", "FAMILY_ID", "SATSANG_GRADE", "SATSANG_CATEGORY", "AGE_GRADE", "AGE", "FULL_NAME", "MOBILE", "MOBILE", "GENDER", "FLAT_HOUSE_NO");
    //static List<String> rowHeaders = Arrays.asList("MemberID","FID","SG","SC","AG","FULL NAME", "FULL ADDRESS", "MOBILE", "CENTER");

    /*
    {"reportName":"General Member Report","user":"lkbhudiya","headerKey":"ID,FAMILY_ID,SATSANG_GRADE,SATSANG_CATEGORY,AGE_GRADE,AP,FULL_NAME,MOBILE,WHATSAPP,
    GENDER,FLAT_HOUSE_NO","headerName":"Member Id,FID,SG,SC,AG,AP,Full Name,Mobile,Whatsapp,Gender,House No",
    "columnWidth":"125,99,49,45,48,220,424,248,132,88,123","newPage":false,
    "landscape":false,"topHeader":"CENTER","subHeaderKey":"SOCIETY_NAME","methodName":"getTwoLevelPDFWithSingleHeader"}
    */
    private static final int numCols = columnNames.size();

    public static void main(String[] args) throws IOException {
        String jsonString = String.join("", Files.readAllLines(Paths.get("E:\\Projects\\Smvs-projects\\bulk-processor\\src\\main\\java\\org\\smvs\\bulkprocessor\\service\\test.json")));
        JSONObject jsonObject = new JSONObject(jsonString);
        JSONArray data = jsonObject.getJSONArray("DATA");
        //exportToWordReport(data);
        //List<Object> objects = data.toList();

        ObjectMapper objectMapper = new ObjectMapper();
        List<HashMap<String, String>> allData = objectMapper.readValue(data.toString(), new TypeReference<>() {
        });

        Map<String, List<HashMap<String, String>>> oneLevelGrouping = allData.stream()
                .collect(Collectors.groupingBy(hm -> hm.get("CENTER")));

        System.out.println("oneLevelGrouping: " + oneLevelGrouping);

        exportOneLevelGroupingToWordReport(oneLevelGrouping);

        Map<String, Map<String, List<HashMap<String, String>>>> twoLevelGrouping = allData.stream()
                .collect(Collectors.groupingBy(hm -> hm.get("CENTER"), Collectors.groupingBy(hm -> hm.get("SOCIETY_NAME"))));

        System.out.println("twoLevelGrouping: " + twoLevelGrouping);

        exportTwoLevelGroupingToWordReport(twoLevelGrouping);


    }


    private static void exportOneLevelGroupingToWordReport(Map<String, List<HashMap<String, String>>> data) throws IOException {
        XWPFDocument document = new XWPFDocument();
        //document.enforceReadonlyProtection();
        setDocumentMargins(document);
        createTablesWithOneLevelGrouping(document, data);

        FileOutputStream out = new FileOutputStream("report.docx");
        document.write(out);
        out.close();

    }

    private static void exportTwoLevelGroupingToWordReport(Map<String, Map<String, List<HashMap<String, String>>>> twoLevelGrouping) throws IOException {
        XWPFDocument document = new XWPFDocument();
        //document.enforceReadonlyProtection();
        setDocumentMargins(document);
        createTablesWithTwoLevelGrouping(document, twoLevelGrouping);

        FileOutputStream out = new FileOutputStream("report2.docx");
        document.write(out);
        out.close();
    }

    private static void setDocumentMargins(XWPFDocument document) {
        CTSectPr sectPr = document.getDocument().getBody().getSectPr();
        if (sectPr == null) sectPr = document.getDocument().getBody().addNewSectPr();
        CTPageMar pageMar = sectPr.getPgMar();
        if (pageMar == null) pageMar = sectPr.addNewPgMar();
        pageMar.setLeft(BigInteger.valueOf(300)); //720 TWentieths of an Inch Point (Twips) = 720/20 = 36 pt = 36/72 = 0.5"
        pageMar.setRight(BigInteger.valueOf(300));
        pageMar.setTop(BigInteger.valueOf(300));
        pageMar.setBottom(BigInteger.valueOf(0));
        pageMar.setFooter(BigInteger.valueOf(0));
        pageMar.setHeader(BigInteger.valueOf(0));
        pageMar.setGutter(BigInteger.valueOf(0));

        if (!sectPr.isSetPgSz()) {
            sectPr.addNewPgSz();
        }
        CTPageSz pageSize = sectPr.getPgSz();
        pageSize.setW(BigInteger.valueOf(595 * 20));
        pageSize.setH(BigInteger.valueOf(842 * 20));
    }

    private static void createTablesWithOneLevelGrouping(XWPFDocument document, Map<String, List<HashMap<String, String>>> data) throws JSONException {

        for (Map.Entry<String, List<HashMap<String, String>>> record : data.entrySet()) {
            List<HashMap<String, String>> firstGroup = record.getValue();
            int numRows = firstGroup.size() + 2; //For Table header


            XWPFTable table = createTable(document, numRows);
            AtomicInteger rowNum = new AtomicInteger(0);


            XWPFTableRow groupRow = table.getRow(rowNum.getAndIncrement());
            XWPFTableCell groupCell = groupRow.getCell(0);
            groupCell.setText(record.getKey());
            //spanCellsAcrossRow(table, rowNum.get(), 0, columnKeys.size());

            addTableHeaders(table, rowNum);

            for (HashMap<String, String> stringStringHashMap : firstGroup) {
                XWPFTableRow tableRow = table.getRow(rowNum.getAndIncrement());
                for (int col = 0; col < columnKeys.size(); col++) {
                    XWPFTableCell cell = tableRow.getCell(col);
                    cell.removeParagraph(0);

                    String val = stringStringHashMap.get(columnKeys.get(col));
                    cell.setText(val);

                }
            }
            spanCellsAcrossRow(table, 0);

            addSpace(document);
        }


        // Example: Set table width
        //CTTblWidth tableWidth = table.getCTTbl().addNewTblPr().addNewTblW();
        //tableWidth.setType(STTblWidth.DXA);
        //tableWidth.setW(BigInteger.valueOf(11400)); // Set table width to 5000 twips (1 inch = 1440 twips)
    }

    private static void createTablesWithTwoLevelGrouping(XWPFDocument document, Map<String, Map<String, List<HashMap<String, String>>>> data) throws JSONException {

        for (Map.Entry<String, Map<String, List<HashMap<String, String>>>> groupRecord : data.entrySet()) {
            boolean flag = true;
            for (Map.Entry<String, List<HashMap<String, String>>> record : groupRecord.getValue().entrySet()) {
                List<HashMap<String, String>> groupedData = record.getValue();
                int numRows = groupedData.size() + 3; //For Table header
                XWPFTable table = createTable(document, numRows);
                AtomicInteger rowNum = new AtomicInteger(0);

                if (flag) {
                    XWPFTableRow firstGroupRow = table.getRow(rowNum.getAndIncrement());
                    XWPFTableCell firstGroupCell = firstGroupRow.getCell(0);
                    firstGroupCell.setText(groupRecord.getKey());

                }

                XWPFTableRow secondGroupRow = table.getRow(rowNum.getAndIncrement());
                XWPFTableCell secondGroupCell = secondGroupRow.getCell(0);
                secondGroupCell.setText(record.getKey());
                //spanCellsAcrossRow(table, rowNum.get(), 0, columnKeys.size());

                addTableHeaders(table, rowNum);

                for (HashMap<String, String> stringStringHashMap : groupedData) {
                    XWPFTableRow tableRow = table.getRow(rowNum.getAndIncrement());
                    for (int col = 0; col < columnKeys.size(); col++) {
                        XWPFTableCell cell = tableRow.getCell(col);
                        cell.removeParagraph(0);

                        String val = stringStringHashMap.get(columnKeys.get(col));
                        cell.setText(val);

                    }
                }
                if (flag) {
                    spanCellsAcrossRow(table, 0);
                    spanCellsAcrossRow(table, 1);
                } else {
                    spanCellsAcrossRow(table, 0);
                }

                flag = false;

            }
            addSpace(document);
        }

    }

    private static void spanCellsAcrossRow(XWPFTable table, int rowNum) {

        /*IntStream.range(1, 3)
                .forEach(value -> {
                    XWPFTableRow row = table.getRow(rowNum);
                    XWPFTableCell removed = row.getCell(value);
                    removed.getCTTc().newCursor().removeXml();
                    row.removeCell(value);
                });*/

        CTHMerge hMerge = CTHMerge.Factory.newInstance();
        hMerge.setVal(STMerge.RESTART);
        for (int i = 0; i < columnKeys.size(); i++) {
            XWPFTableCell cell = table.getRow(rowNum).getCell(i);
            addTcPr(cell);
            cell.getCTTc().getTcPr().setHMerge(hMerge);
            hMerge.setVal(STMerge.CONTINUE);
        }


        // First Row
       /* CTHMerge hMerge = CTHMerge.Factory.newInstance();
        hMerge.setVal(STMerge.RESTART);
        XWPFTableCell cell = table.getRow(0).getCell(0);
        addTcPr(cell);
        cell.getCTTc().getTcPr().setHMerge(hMerge);*/

       /* XWPFTableCell cell1 = table.getRow(1).getCell(0);
        addTcPr(cell1);
        cell1.getCTTc().getTcPr().setHMerge(hMerge);*/

// Secound Row cell will be merged/"deleted"
        //CTHMerge hMerge1 = CTHMerge.Factory.newInstance();
      /*  hMerge.setVal(STMerge.CONTINUE);
        XWPFTableCell cell2 = table.getRow(0).getCell(1);
        addTcPr(cell2);
        cell2.getCTTc().getTcPr().setHMerge(hMerge);*/

        /*XWPFTableCell cell3 = table.getRow(1).getCell(1);
        addTcPr(cell3);
        cell3.getCTTc().getTcPr().setHMerge(hMerge1);*/

       /* XWPFTableCell cell = table.getRow(rowNum).getCell(colNum);
        if (cell.getCTTc().getTcPr() == null) cell.getCTTc().addNewTcPr();
        if (cell.getCTTc().getTcPr().getGridSpan() == null) cell.getCTTc().getTcPr().addNewGridSpan();
        cell.getCTTc().getTcPr().getGridSpan().setVal(BigInteger.valueOf(span));*/

        /*IntStream.range(1, columnKeys.size())
                .forEach(value -> {
                    XWPFTableRow row = table.getRow(rowNum);
                    //row.removeCell(value);
                    XWPFTableCell removed = row.getCell(colNum);
                    removed.getCTTc().newCursor().removeXml();
                    row.removeCell(value);
                });*/




    }

    private static void addTcPr(XWPFTableCell cell1) {
        if (cell1.getCTTc().getTcPr() == null) cell1.getCTTc().addNewTcPr();
    }

    private static void addTableHeaders(XWPFTable table, AtomicInteger rowNum) {
        XWPFTableRow tableRow = table.getRow(rowNum.getAndIncrement());
        IntStream.range(0, columnNames.size())
                .forEach(col -> {
                    XWPFTableCell cell = tableRow.getCell(col);
                    //cell.getParagraphs().get(0).getRuns().get(0).setBold(true);
                    cell.removeParagraph(0);
                    XWPFParagraph paragraph = cell.addParagraph();
                    //removeParagraphSpacing(paragraph);

                    // Set the alignment of the paragraph
                    //paragraph.setAlignment(ParagraphAlignment.LEFT);

                    // Create a new run within the paragraph
                    XWPFRun run = paragraph.createRun();
                    run.setBold(true);
                    run.setText(columnNames.get(col));
                    /*cell.getParagraphs().get(0).getRuns().get(0).setBold(true);
                    cell.setText(rowHeaders.get(col));*/
                });
    }

    private static void addSpace(XWPFDocument document) {
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun xwpfRun = paragraph.createRun();
        xwpfRun.addBreak();

        /*XWPFTable table123 = document.getTableArray(0);
        org.apache.xmlbeans.XmlCursor cursor = table123.getCTTbl().newCursor();
        cursor.toEndToken(); //now we are at end of the CTTbl
        //there always must be a next start token. Either a p or at least sectPr.
        while (cursor.toNextToken() != org.apache.xmlbeans.XmlCursor.TokenType.START) ;
        XWPFParagraph newParagraph = document.insertNewParagraph(cursor);
        newParagraph.createRun();*/
    }

    private static void removeParagraphSpacing(XWPFParagraph paragraph) {
        paragraph.setSpacingBefore(0);
        paragraph.setSpacingBeforeLines(0);
        paragraph.setSpacingAfter(0);
        paragraph.setSpacingAfterLines(0);
    }

    private static XWPFTable createTable(XWPFDocument document, int numRows) {
        XWPFTable table = document.createTable(numRows, numCols);
        table.setCellMargins(0, 0, 0, 0);
        CTTblLayoutType type = table.getCTTbl().getTblPr().addNewTblLayout();
        //type.setType(STTblLayoutType.FIXED);
        type.setType(STTblLayoutType.AUTOFIT);

        //table.removeBorders();
        return table;
    }

    private static String getFullName(JSONObject jsonObject) throws JSONException {
        return String.join(" ", jsonObject.getString("FIRST_NAME_GUJ"),
                jsonObject.getString("MIDDLE_NAME_GUJ"),
                jsonObject.getString("LAST_NAME_GUJ"));
    }
}
