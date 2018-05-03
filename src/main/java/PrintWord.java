import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;

class PrintWord {

    static void createWord(String absPath, Workbook book) {
        try (FileOutputStream out = new FileOutputStream(new File(absPath))) {
            XWPFDocument document = addPangramList(book);
            document.write(out);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static XWPFDocument addPangramList(Workbook book) {
        Sheet sheet = book.getSheetAt(0);
        XWPFDocument document = new XWPFDocument();
        int wordId = 1;
        int rowNum = 0;
        int day = 1;
        insertDayTitle(document, day);

        while (true) {
            XWPFRun run = document.createParagraph().createRun();
            Row row = sheet.getRow(rowNum);
            if (row == null) return document;

            String kana = row.getCell(1).getStringCellValue();
            String kanji = row.getCell(2).getStringCellValue();

            String format = wordId < 10 ? "%s.    %s (%s) \n" : "%s.  %s (%s) \n";
            String text = String.format(format, wordId, kana, kanji);
            run.setText(text);

            if (wordId == 18) {
                document.createParagraph().createRun().addBreak();
                insertTable(document);
                document.createParagraph().setPageBreak(true);
                insertDayTitle(document, ++day);
                wordId = 1;
            } else {
                wordId++;
            }
            rowNum++;
        }
    }

    private static void insertDayTitle(XWPFDocument document, int day) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun runTitle = paragraph.createRun();
        runTitle.setFontSize(24);
        runTitle.setText("計測" + day + "日目");
    }

    private static void insertTable(XWPFDocument document){
        XWPFTable table = document.createTable(2, 1);
        XWPFTableCell timeCell = table.getRow(0).getCell(0);
        setCellWidth(timeCell);
        XWPFParagraph timeParagraph = timeCell.getParagraphs().get(0);
        XWPFRun run = timeParagraph.createRun();
        run.setText("開始時刻:");

        XWPFTableRow row = table.getRow(1);
        row.setHeight(2000);
        XWPFTableCell commentCell = row.getCell(0);
        setCellWidth(commentCell);
    }

    private static void setCellWidth(XWPFTableCell cell){
        cell.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(10500L));
    }
}
