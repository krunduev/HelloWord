
import java.io.FileInputStream;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;

/**
 * Created by krunduev on 10/7/16.
 *
 * How to look inside doc file in java using Apache POI
 *
 */
public class Main {

    public static void main(String[] args)throws Exception
    {
        HWPFDocument doc = new HWPFDocument(
                new FileInputStream("src/main/resources/32-bekhukotay-1.doc"));


        Range range = doc.getRange();
        for (int i = 0; i < range.numParagraphs(); i++) {
            Paragraph par = range.getParagraph(i);

            if (!par.isInTable()) {
                System.out.println("Paragraph:" + par.text());
                int numberOfRuns = par.numCharacterRuns();
                for (int runIndex = 0; runIndex < numberOfRuns; runIndex++)
                {
                    CharacterRun run = par.getCharacterRun(runIndex);
                    System.out.println("Text: " + run.text());
                    System.out.println("Color: " + run.getColor());
                    System.out.println("Font: " + run.getFontName());
                    System.out.println("Font Size: " + run.getFontSize());
                    System.out.println("Is Bold: " + run.isBold());
                    System.out.println("Is Italic: " + run.isItalic());
                }
            } else {
                Table table = range.getTable(par);
                for (int rowIdx = 0; rowIdx < table.numRows(); rowIdx++) {
                    TableRow row = table.getRow(rowIdx);
                    for (int colIdx = 0; colIdx < row.numCells(); colIdx++) {
                        TableCell cell = row.getCell(colIdx);
                        System.out.print(" column=" + cell.getParagraph(0).text());
                        i++;
                    }
                    System.out.println();
                    i++;
                }
            }
        }

    }


}
