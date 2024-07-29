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
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblLayoutType;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.IntStream;

import static java.util.stream.Collectors.groupingBy;
import static java.util.stream.Collectors.toList;

public class KutumbMahitiReport {

    public static final String DATA_JSON_FILE_PATH = "E:\\Projects\\Smvs_Global_Docs\\Kutumb_Darshan.json";
    private final List<String> headerKey = List.of("ID", "FULL_NAME_GUJ", "RELATION", "DOB", "SATSANG_GRADE", "AP", "GENDER", "MOBILE", "BLOOD_GROUP", "EMAIL");
    private List<String> headerName;
    private boolean landscape;
    private boolean isNewPage;
    private String user;

    private final int numCols = headerKey.size();

    int[] colsWidth = {
            500,
            4032,
            1728,
            1998,
            1040,
            1040,
            1040,
            2528,
            1728,
            3432
    };

    public static void main(String[] args) throws IOException {
        KutumbMahitiReport msWordReport = new KutumbMahitiReport();

        String jsonString = String.join("", Files.readAllLines(Paths.get(DATA_JSON_FILE_PATH)));
        JSONArray jsonArray = new JSONArray(jsonString);
        msWordReport.exportReportToWord(jsonArray);
    }

    private void exportReportToWord(JSONArray data) throws IOException {

        ObjectMapper objectMapper = new ObjectMapper();
        List<LinkedHashMap<String, String>> allData = objectMapper.readValue(data.toString(), new TypeReference<>() {
        });

        generateThreeLevelReport(allData);
    }


    private void generateThreeLevelReport(List<LinkedHashMap<String, String>> allData) throws IOException {
        LinkedHashMap<String, LinkedHashMap<String, LinkedHashMap<String, List<LinkedHashMap<String, String>>>>> threeLevelGroupedData = allData.stream()
                .collect(groupingBy(hm -> hm.get("ZONE_NAME"), LinkedHashMap::new,
                        groupingBy(hm -> hm.get("CENTER"), LinkedHashMap::new,
                                groupingBy(hm -> hm.get("FAMILY_ID"), LinkedHashMap::new, toList()
                                )
                        )
                ));
        System.out.println(threeLevelGroupedData);

        exportThreeLevelGroupingToWordReport(threeLevelGroupedData);
    }

    private void exportThreeLevelGroupingToWordReport(LinkedHashMap<String, LinkedHashMap<String, LinkedHashMap<String, List<LinkedHashMap<String, String>>>>> twoLevelGrouping) throws IOException {
        XWPFDocument document = new XWPFDocument();
        //document.enforceReadonlyProtection();
        setDocumentMargins(document);
        setDocumentFooter(document);
        createTablesWithThreeLevelGrouping(document, twoLevelGrouping);

        FileOutputStream out = new FileOutputStream("KutumbMahitiReport.docx");
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

        if (landscape) {
            pageSize.setOrient(STPageOrientation.LANDSCAPE);
            pageSize.setW(BigInteger.valueOf(842 * 20));
            pageSize.setH(BigInteger.valueOf(595 * 20));
        } else {
            pageSize.setOrient(STPageOrientation.PORTRAIT);
            pageSize.setH(BigInteger.valueOf(842 * 20));
            pageSize.setW(BigInteger.valueOf(595 * 20));
        }
    }


    private void setDocumentFooter(XWPFDocument document) {
        XWPFFooter footer = document.createFooter(HeaderFooterType.DEFAULT);

        XWPFParagraph paragraph = footer.getParagraphArray(0);
        if (paragraph == null) paragraph = footer.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.LEFT);
        //addTabStop(paragraph, "CENTER", 3.25);

        XWPFRun run = paragraph.createRun();
        run.setText("Swaminarayan Mandir Vasna Sanstha");

        run = paragraph.createRun();
        run.addTab();
        run.addTab();

        run = paragraph.createRun();
        run.setText("Page ");
        paragraph.getCTP().addNewFldSimple().setInstr("PAGE \\* MERGEFORMAT");
        run = paragraph.createRun();
        run.setText(" of ");
        paragraph.getCTP().addNewFldSimple().setInstr("NUMPAGES \\* MERGEFORMAT");

        run = paragraph.createRun();
        run.addTab();
        run.addTab();

        run = paragraph.createRun();
        run.setText("By: " + user);

        run = paragraph.createRun();
        run.addTab();

        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("EEE, d MMM yyyy HH:mm:ss");
        LocalDateTime currentTime = LocalDateTime.now();

        run.setText("At: " + currentTime.format(formatter));
        //paragraph.getCTP().addNewFldSimple().setInstr("TIME \\@ \"dd.MMM.yyyy HH:mm:ss\" \\* MERGEFORMAT");


    }

    private void createTablesWithThreeLevelGrouping(XWPFDocument document, LinkedHashMap<String, LinkedHashMap<String, LinkedHashMap<String, List<LinkedHashMap<String, String>>>>> data) throws JSONException {
        for (Map.Entry<String, LinkedHashMap<String, LinkedHashMap<String, List<LinkedHashMap<String, String>>>>> zoneLevel : data.entrySet()) { //Zone
            boolean zoneFlag = true;
            for (Map.Entry<String, LinkedHashMap<String, List<LinkedHashMap<String, String>>>> centerLevel : zoneLevel.getValue().entrySet()) { //Center
                boolean centerFlag = true;
                for (Map.Entry<String, List<LinkedHashMap<String, String>>> familyLevel : centerLevel.getValue().entrySet()) {

                    List<LinkedHashMap<String, String>> familyLevelGroupedData = familyLevel.getValue();
                    int additionalRows = zoneFlag ? 4 : centerFlag ? 3 : 2; //Adding rows for group headers
                    int numRows = familyLevelGroupedData.size() + additionalRows;
                    XWPFTable table = createTable(document, numRows);
                    AtomicInteger rowNum = new AtomicInteger(0);

                    if (zoneFlag) {
                        XWPFTableRow zoneGroupHeaderRow = table.getRow(rowNum.getAndIncrement());
                        XWPFTableCell firstGroupCell = zoneGroupHeaderRow.getCell(0);
                        firstGroupCell.setText("Zone: " + zoneLevel.getKey());
                        setTopHeaderTextFormat(firstGroupCell);
                    }
                    if (centerFlag) {
                        XWPFTableRow centerGroupHeaderRow = table.getRow(rowNum.getAndIncrement());
                        XWPFTableCell centerGroupCell = centerGroupHeaderRow.getCell(0);
                        centerGroupCell.setText("Center: " + centerLevel.getKey());
                        centerGroupCell.getParagraphs().get(0).getRuns().get(0).setBold(true);
                        centerGroupCell.getParagraphs().get(0).getRuns().get(0).setFontSize(16);
                        //spanCellsAcrossRow(table, rowNum.get(), 0, columnKeys.size());

                    }

                    XWPFTableRow familyGroupRow = table.getRow(rowNum.getAndIncrement());
                    XWPFTableCell familyGroupCell = familyGroupRow.getCell(0);
                    familyGroupCell.setText("FID: " + familyLevel.getKey() + ",     Address: " + familyLevelGroupedData.get(0).get("FULL_ADDRESS_GUJ"));
                    familyGroupCell.getParagraphs().get(0).getRuns().get(0).setBold(true);
                    familyGroupCell.getParagraphs().get(0).getRuns().get(0).setFontSize(11);

                    addTableHeaders(table, rowNum);

                    int row = 1; //Serial Number
                    for (HashMap<String, String> familyLevelData : familyLevelGroupedData) {
                        XWPFTableRow tableRow = table.getRow(rowNum.getAndIncrement());
                        for (int col = 0; col < headerKey.size(); col++) {
                            XWPFTableCell cell = tableRow.getCell(col);
                            cell.removeParagraph(0);

                            if (col == 0) {
                                cell.setText(String.valueOf(row));
                            } else {
                                String val = familyLevelData.get(headerKey.get(col));
                                cell.setText(val);
                            }

                            CTTcPr tcPr = getCtTcPr(cell);
                            CTTcBorders tblBorders = getCtTcBorders(tcPr);

                            setBorderForTable(col, tblBorders, headerKey.size());

                            CTTblWidth cellWidth = tcPr.addNewTcW();
                            CTTcPr pr = cell.getCTTc().addNewTcPr();
                            pr.addNewNoWrap();
                            cellWidth.setW(BigInteger.valueOf(colsWidth[col]));

                        }
                        row++;
                    }
                    if (zoneFlag) {
                        spanCellsAcrossRow(table, 0);
                        spanCellsAcrossRow(table, 1);
                        spanCellsAcrossRow(table, 2);
                    } else if (centerFlag) {
                        spanCellsAcrossRow(table, 0);
                        spanCellsAcrossRow(table, 1);
                        //spanCellsAcrossRow(table, 2);
                    } else {
                        spanCellsAcrossRow(table, 0);
                        //spanCellsAcrossRow(table, 1);
                    }
                    centerFlag = false;
                    zoneFlag = false;
                }
            }
            addSpace(document);
        }
    }

    private void setBorderForTable(int col, CTTcBorders tblBorders, int maxSize) {
        if (col == 0) {
            tblBorders.addNewRight().setVal(STBorder.NIL);
            tblBorders.addNewBottom().setVal(STBorder.NIL);
            tblBorders.addNewTop().setVal(STBorder.NIL);
        } else if (col == maxSize - 1) {
            tblBorders.addNewLeft().setVal(STBorder.NIL);
            tblBorders.addNewBottom().setVal(STBorder.NIL);
            tblBorders.addNewTop().setVal(STBorder.NIL);
        } else {
            tblBorders.addNewTop().setVal(STBorder.NIL);
            tblBorders.addNewRight().setVal(STBorder.NIL);
            tblBorders.addNewLeft().setVal(STBorder.NIL);
            tblBorders.addNewBottom().setVal(STBorder.NIL);
        }
    }

    private static CTTcBorders getCtTcBorders(CTTcPr tcPr) {
        CTHMerge hMerge = tcPr.addNewHMerge();
        hMerge.setVal(STMerge.RESTART);
        CTTcBorders tblBorders = tcPr.addNewTcBorders();
        return tblBorders;
    }

    private static CTTcPr getCtTcPr(XWPFTableCell cell) {
        CTTc ctTc = cell.getCTTc();
        CTTcPr tcPr = ctTc.addNewTcPr();
        return tcPr;
    }

    private void setTopHeaderTextFormat(XWPFTableCell tableCell) {
        tableCell.getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
        tableCell.getParagraphs().get(0).getRuns().get(0).setBold(true);
        tableCell.getParagraphs().get(0).getRuns().get(0).setFontSize(20);
    }

    private void spanCellsAcrossRow(XWPFTable table, int rowNum) {

        CTHMerge hMerge = CTHMerge.Factory.newInstance();
        hMerge.setVal(STMerge.RESTART);
        for (int i = 0; i < headerKey.size(); i++) {
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
        this.headerName = List.of("#", "Name", "Relation", "DOB", "S", "AP", "G", "Mobile", "BID G", "Email");
        XWPFTableRow tableRow = table.getRow(rowNum.getAndIncrement());
        IntStream.range(0, headerName.size())
                .forEach(col -> {
                    XWPFTableCell cell = tableRow.getCell(col);
                    //cell.getParagraphs().get(0).getRuns().get(0).setBold(true);
                    cell.removeParagraph(0);
                    XWPFParagraph paragraph = cell.addParagraph();

                    CTTcPr tcPr = getCtTcPr(cell);
                    CTTcBorders tblBorders = getCtTcBorders(tcPr);
                    setBorderForTable(col, tblBorders, headerName.size());
                    //removeParagraphSpacing(paragraph);

                    // Set the alignment of the paragraph
                    //paragraph.setAlignment(ParagraphAlignment.LEFT);

                    // Create a new run within the paragraph
                    XWPFRun run = paragraph.createRun();
                    run.setBold(true);
                    run.setText(headerName.get(col));
                    /*cell.getParagraphs().get(0).getRuns().get(0).setBold(true);
                    cell.setText(rowHeaders.get(col));*/
                });
    }

    private void addSpace(XWPFDocument document) {
        XWPFParagraph paragraph = document.createParagraph();
        if (isNewPage) {
            paragraph.setPageBreak(true);
        } else {
            XWPFRun xwpfRun = paragraph.createRun();
            xwpfRun.addBreak();
        }

    }

    private XWPFTable createTable(XWPFDocument document, int numRows) {
        XWPFTable table = document.createTable(numRows, numCols);
        table.setCellMargins(0, 0, 0, 0);
        table.setWidth("100%");
        CTTblLayoutType type = table.getCTTbl().getTblPr().addNewTblLayout();
        type.setType(STTblLayoutType.FIXED);
        //type.setType(STTblLayoutType.AUTOFIT);
        return table;
    }
}
