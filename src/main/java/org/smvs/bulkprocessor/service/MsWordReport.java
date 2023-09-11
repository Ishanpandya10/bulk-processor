package org.smvs.bulkprocessor.service;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
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
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

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

    private List<String> headerKey;
    private List<String> headerName;
    private List<String> columnWidth;
    private boolean landscape;
    private String topHeader;
    private String subHeaderKey;
    private String methodName;
    private boolean isNewPage;

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

    public MsWordReport(JSONObject inputJsonData) {
        extractInputJson(inputJsonData);
    }

    public static void main(String[] args) throws IOException {
        MsWordReport msWordReport = new MsWordReport(getInputJSON());
        String jsonString = String.join("", Files.readAllLines(Paths.get("E:\\Projects\\Smvs-projects\\bulk-processor\\src\\main\\java\\org\\smvs\\bulkprocessor\\service\\test.json")));
        JSONObject jsonObject = new JSONObject(jsonString);
        JSONArray data = jsonObject.getJSONArray("DATA");
        msWordReport.exportReportToWord(data);
    }

    private void extractInputJson(JSONObject jsonObject) {
        String headerKey = jsonObject.getString("headerKey");
        this.headerKey = getListFromCommaSeparatedString(headerKey);

        String headerName = jsonObject.getString("headerName");
        this.headerName = getListFromCommaSeparatedString(headerName);

        String width = jsonObject.getString("columnWidth");
        this.columnWidth = getListFromCommaSeparatedString(width);

        this.landscape = jsonObject.getBoolean("landscape");

        this.topHeader = jsonObject.getString("topHeader");

        this.subHeaderKey = jsonObject.optString("subHeaderKey", "");
        this.methodName = jsonObject.getString("methodName");

        this.isNewPage = jsonObject.getBoolean("newPage");

    }

    private static List<String> getListFromCommaSeparatedString(String width) {
        return Arrays.asList(width.split(","));
    }

    private static JSONObject getInputJSON() {
        try {
            String inputJson = String.join(" ", Files.readAllLines(Paths.get("src/main/java/org/smvs/bulkprocessor/service/input_config.json")));
            return new JSONObject(inputJson);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private void exportReportToWord(JSONArray data) throws IOException {

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


    private void exportOneLevelGroupingToWordReport(Map<String, List<HashMap<String, String>>> data) throws IOException {
        XWPFDocument document = new XWPFDocument();
        //document.enforceReadonlyProtection();
        setDocumentMargins(document);
        setDocumentFooter(document, "lalitB");
        createTablesWithOneLevelGrouping(document, data);

        FileOutputStream out = new FileOutputStream("report_one_level.docx");
        document.write(out);
        out.close();

    }

    private void exportTwoLevelGroupingToWordReport(Map<String, Map<String, List<HashMap<String, String>>>> twoLevelGrouping) throws IOException {
        XWPFDocument document = new XWPFDocument();
        //document.enforceReadonlyProtection();
        setDocumentMargins(document);
        setDocumentFooter(document, "lalitB");
        createTablesWithTwoLevelGrouping(document, twoLevelGrouping);

        FileOutputStream out = new FileOutputStream("report_two_level.docx");
        document.write(out);
        out.close();
    }

    private void setDocumentMargins(XWPFDocument document) {
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

    private void setDocumentFooter(XWPFDocument document, String currentUser) {
        XWPFFooter footer = document.createFooter(HeaderFooterType.DEFAULT);

        XWPFParagraph paragraph = footer.getParagraphArray(0);
        if (paragraph == null) paragraph = footer.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.LEFT);
        //addTabStop(paragraph, "CENTER", 3.25);

        XWPFRun run = paragraph.createRun();
        run.setText("Time: ");
        paragraph.getCTP().addNewFldSimple().setInstr("TIME \\@ \"dd.MM.yyyy HH:mm:ss\" \\* MERGEFORMAT");

        run = paragraph.createRun();
        run.addTab();

        run = paragraph.createRun();
        run.setText("User: " + currentUser);

        run = paragraph.createRun();
        run.addTab();

        run = paragraph.createRun();
        run.setText("Page ");
        paragraph.getCTP().addNewFldSimple().setInstr("PAGE \\* MERGEFORMAT");
        run = paragraph.createRun();
        run.setText(" of ");
        paragraph.getCTP().addNewFldSimple().setInstr("NUMPAGES \\* MERGEFORMAT");

    }

    private void createTablesWithOneLevelGrouping(XWPFDocument document, Map<String, List<HashMap<String, String>>> data) throws JSONException {

        for (Map.Entry<String, List<HashMap<String, String>>> record : data.entrySet()) {
            List<HashMap<String, String>> firstGroup = record.getValue();
            int numRows = firstGroup.size() + 2; //For Table header


            XWPFTable table = createTable(document, numRows);
            AtomicInteger rowNum = new AtomicInteger(0);


            XWPFTableRow groupRow = table.getRow(rowNum.getAndIncrement());
            XWPFTableCell groupCell = groupRow.getCell(0);
            groupCell.setText(record.getKey());
            setTopHeaderTextFormat(groupCell);
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

    private void createTablesWithTwoLevelGrouping(XWPFDocument document, Map<String, Map<String, List<HashMap<String, String>>>> data) throws JSONException {

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
                    setTopHeaderTextFormat(firstGroupCell);
                }

                XWPFTableRow secondGroupRow = table.getRow(rowNum.getAndIncrement());
                XWPFTableCell secondGroupCell = secondGroupRow.getCell(0);
                secondGroupCell.setText(record.getKey());
                secondGroupCell.getParagraphs().get(0).getRuns().get(0).setBold(true);
                secondGroupCell.getParagraphs().get(0).getRuns().get(0).setFontSize(13);
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

    private void setTopHeaderTextFormat(XWPFTableCell tableCell) {
        tableCell.getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
        tableCell.getParagraphs().get(0).getRuns().get(0).setBold(true);
        tableCell.getParagraphs().get(0).getRuns().get(0).setFontSize(15);
    }

    private void spanCellsAcrossRow(XWPFTable table, int rowNum) {

        CTHMerge hMerge = CTHMerge.Factory.newInstance();
        hMerge.setVal(STMerge.RESTART);
        for (int i = 0; i < columnKeys.size(); i++) {
            XWPFTableCell cell = table.getRow(rowNum).getCell(i);
            addTcPr(cell);
            cell.getCTTc().getTcPr().setHMerge(hMerge);
            hMerge.setVal(STMerge.CONTINUE);
        }

    }

    private static void addTcPr(XWPFTableCell cell1) {
        if (cell1.getCTTc().getTcPr() == null) cell1.getCTTc().addNewTcPr();
    }

    private void addTableHeaders(XWPFTable table, AtomicInteger rowNum) {
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

    private void addSpace(XWPFDocument document) {
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun xwpfRun = paragraph.createRun();
        xwpfRun.addBreak();
    }

    private XWPFTable createTable(XWPFDocument document, int numRows) {
        XWPFTable table = document.createTable(numRows, numCols);
        table.setCellMargins(0, 0, 0, 0);
        table.setWidth("100%");
        //CTTblLayoutType type = table.getCTTbl().getTblPr().addNewTblLayout();
        //type.setType(STTblLayoutType.FIXED);
        //type.setType(STTblLayoutType.AUTOFIT);
        return table;
    }

}
