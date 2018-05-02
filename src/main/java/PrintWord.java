import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

class PrintWord {

    static void createWord(String absPath, Workbook book){
        try (FileOutputStream out = new FileOutputStream(new File(absPath))) {
            XWPFDocument document = addPangramList(book);
            document.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static XWPFDocument addPangramList(Workbook book){
        XWPFDocument document = new XWPFDocument();
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();

        run.setText(book.getSheetAt(0).getRow(0).getCell(1).getStringCellValue());
        return document;
    }
}
