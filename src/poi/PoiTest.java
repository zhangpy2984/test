package poi;

import com.carrotsearch.sizeof.RamUsageEstimator;
import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Pattern;

/**
 * @author zhangpengyue
 * @date 2018/12/16
 */
public class PoiTest {
    public static void main(String[] args) {

        String filename = "D:\\幸福人寿201810.xlsx";
        OPCPackage pkg = null;
        InputStream sheetInputStream = null;
        try {
            pkg = OPCPackage.open(filename, PackageAccess.READ);
            XSSFReader xssfReader = new XSSFReader(pkg);
            StylesTable styles = xssfReader.getStylesTable();
            ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(pkg);
            Iterator<InputStream> inputStreamIterator = xssfReader.getSheetsData();
            while (inputStreamIterator.hasNext()) {
                sheetInputStream = inputStreamIterator.next();
                List<List<String>> sheetContent = processSheet(styles, strings, sheetInputStream);
                System.out.println(RamUsageEstimator.humanSizeOf(sheetContent));
            }
        } catch (Exception e) {
            throw new RuntimeException(e.getMessage(), e);
        } finally {
            if (sheetInputStream != null) {
                try {
                    sheetInputStream.close();
                } catch (IOException e) {
                    throw new RuntimeException(e.getMessage(), e);
                }
            }
            if (pkg != null) {
                try {
                    pkg.close();
                } catch (IOException e) {
                    throw new RuntimeException(e.getMessage(), e);
                }
            }
        }
    }

    private static List<List<String>> processSheet(StylesTable styles, ReadOnlySharedStringsTable strings, InputStream sheetInputStream)
            throws SAXException, ParserConfigurationException, IOException {
        XMLReader sheetParser = SAXHelper.newXMLReader();
        List<List<String>> sheetContents = new ArrayList<>();
        Pattern patternNumber = Pattern.compile("[0-9]*");
        XSSFSheetXMLHandler.SheetContentsHandler handler = new XSSFSheetXMLHandler.SheetContentsHandler() {
            List<String> rowContents = null;
            List<String> titleRow = new ArrayList<>();
            int currentRow = 0;
            @Override
            public void startRow(int rowNum) {
                rowContents = new ArrayList<>();
                currentRow = rowNum;
            }

            @Override
            public void endRow(int rowNum) {
                sheetContents.add(rowContents);
            }

            @Override
            public void cell(String cellReference, String formattedValue, XSSFComment comment) {
                if (0 == currentRow) {
                    titleRow.add(patternNumber.matcher(cellReference).replaceAll(""));
                    rowContents.add(formattedValue);
                }
                //判断空单元格，补齐为空字符串
                else {
                    String col = patternNumber.matcher(cellReference).replaceAll("");
                    int index = titleRow.indexOf(col);
                    if (index == rowContents.size()) {
                        rowContents.add(formattedValue);
                    } else {
                        int times = index - rowContents.size();
                        for (int i = 0; i < times; i++) {
                            rowContents.add("空");
                        }
                        rowContents.add(formattedValue);
                    }
                }
            }
        };

        sheetParser.setContentHandler(new XSSFSheetXMLHandler(styles, strings, handler, false));
        sheetParser.parse(new InputSource(sheetInputStream));
        return sheetContents;
    }
}
