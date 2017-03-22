package test;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.math.BigInteger;

/**
 * Class for creating of MS Word documents
 *
 * @author Konstantin Valerievich Dichenko
 * @version 1.0
 */
public class DocConstructor
{

    public XWPFDocument getDoc()
    {
        XWPFDocument docx = new XWPFDocument();
        addCaption(docx, "Caption 1");
        XWPFParagraph paragraph = docx.createParagraph();
        addAbsatz(paragraph, loremIpsum);
        addAbsatz(paragraph, loremIpsum);

        addTable(docx);

        return docx;
    }

    private void sizeA4(XWPFDocument document)
    {
        CTBody body = document.getDocument().getBody();
        if (!body.isSetSectPr())
            body.addNewSectPr();
        CTSectPr section = body.getSectPr();
        if (!section.isSetPgSz())
            section.addNewPgSz();
        CTPageSz pageSize = section.getPgSz();
        pageSize.setW(BigInteger.valueOf(16840));
        pageSize.setH(BigInteger.valueOf(11900));
    }

    private void addCaption(XWPFDocument document, String text)
    {
        String captionStyleName = "CaptionStyle";
        final XWPFParagraph paragraph = document.createParagraph();
        addCustomHeadingStyle(document, captionStyleName, 1);
        final XWPFRun run = paragraph.createRun();
        run.setText(text);
        addFormats(run);
    }

    private void addFormats(XWPFRun run){
        run.setBold(true);
        run.setFontSize(18);
        run.setFontFamily("sans serif");
    }

    private static void addCustomHeadingStyle(XWPFDocument docxDocument, String styleId, int headingLevel)
    {

        CTStyle ctStyle = CTStyle.Factory.newInstance();
        ctStyle.setStyleId(styleId);

        CTString styleName = CTString.Factory.newInstance();
        styleName.setVal(styleId);
        ctStyle.setName(styleName);

        CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
        indentNumber.setVal(BigInteger.valueOf(headingLevel));

        // lower number > style is more prominent in the formats bar
        ctStyle.setUiPriority(indentNumber);

        CTOnOff onoffnull = CTOnOff.Factory.newInstance();
        ctStyle.setUnhideWhenUsed(onoffnull);

        // style shows up in the formats bar
        ctStyle.setQFormat(onoffnull);

        // style defines a heading of the given level
        CTPPr ppr = CTPPr.Factory.newInstance();
        ppr.setOutlineLvl(indentNumber);
        ctStyle.setPPr(ppr);

        XWPFStyle style = new XWPFStyle(ctStyle);

        // is a null op if already defined
        XWPFStyles styles = docxDocument.createStyles();

        style.setType(STStyleType.PARAGRAPH);
        styles.addStyle(style);
    }

    private void addAbsatz(XWPFParagraph paragraph, String text)
    {
        XWPFRun run = paragraph.createRun();
        run.setText(text);
        run.addBreak();
        run.setFontSize(14);
    }

    private void addTable(XWPFDocument document)
    {
        XWPFTable table = document.createTable();
        CTTblWidth width = table.getCTTbl().addNewTblPr().addNewTblW();
        width.setType(STTblWidth.DXA);
        width.setW(BigInteger.valueOf(8000));

        getCell(table, 0, 0).setText("0,0");
        getCell(table, 0, 1).setText("0,1");
        getCell(table, 0, 2).setText("0,2");
        getCell(table, 1, 0).setText("1,0");
        getCell(table, 1, 1).setText("1,1");
        getCell(table, 1, 2).setText("1,2");
        getCell(table, 1, 3).setText("1,3");
        getCell(table, 2, 0).setText("2,0");
        getCell(table, 2, 1).setText("2,1");
        getCell(table, 6, 5).setText("6,5");
    }

    private XWPFTableCell getCell(XWPFTable table, int rowIndex, int columnIndex)
    {
        XWPFTableRow row = getRowWithIndex(table, rowIndex);
        return getCellWithIndex(row, columnIndex);
    }

    private XWPFTableRow getRowWithIndex(XWPFTable table, int rowIndex)
    {
        XWPFTableRow row = table.getRow(rowIndex);
        while (row == null)
        {
            table.createRow();
            row = table.getRow(rowIndex);
        }
        return row;
    }

    private XWPFTableCell getCellWithIndex(XWPFTableRow row, int cellIndex)
    {
        XWPFTableCell cell = row.getCell(cellIndex);
        while (cell == null)
        {
            row.addNewTableCell();
            cell = row.getCell(cellIndex);
        }
        return cell;
    }

    private String loremIpsum = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nulla condimentum id nibh et ullamcorper." +
            " Curabitur tempus quam id ex molestie, sit amet tincidunt massa porta. Integer facilisis arcu sed nisl sollicitudin scelerisque." +
            " Aenean dictum nibh velit. Vivamus aliquam eget justo at fringilla. Suspendisse velit leo, varius vel maximus vel, feugiat ac nunc." +
            " Nam volutpat volutpat auctor. Donec dignissim eros ut malesuada hendrerit." +
            " Nulla sodales dui vitae pretium porta. Pellentesque fringilla dolor id justo eleifend vulputate." +
            " Duis ex nibh, mollis quis facilisis a, porta vel libero.";
}
