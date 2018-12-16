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
import org.xml.sax.XMLReader;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

/**
 * 自定义ExcelReader，按照事件驱动模式来读取Excel
 *
 * @author zhangpengyue
 * @date 2018/12/16
 */
public class ExcelXlsxReader {

    private String file;

    OPCPackage pkg = null;
    InputStream sheetInputStream = null;
    StylesTable styles = null;
    ReadOnlySharedStringsTable stringsTable = null;
    XSSFReader xssfReader = null;

    private static final Pattern patternNumber = Pattern.compile("[0-9]*");
    private static final Pattern patternNotNumber = Pattern.compile("[^0-9]");

    private ExcelXlsxReader() {
    }

    public static void main(String[] args) {
        String filename = "D:\\幸福人寿201810.xlsx";
        try {
            ExcelXlsxReader reader = new ExcelXlsxReader(filename);
            Map<String, XSheet> sheetMap = reader.processExcel();
            System.out.println(RamUsageEstimator.humanSizeOf(sheetMap));
            Map<String, List<List<String>>> excelValue = reader.processExcelValue();
            System.out.println(RamUsageEstimator.humanSizeOf(excelValue));
            System.out.println(excelValue);
        } catch (ExcelReaderException e) {
            e.printStackTrace();
        }
    }

    /**
     * 构造方法
     *
     * @param file
     * @throws ExcelReaderException
     */
    public ExcelXlsxReader(String file) throws ExcelReaderException {
        this.file = file;
        try {
            pkg = OPCPackage.open(file, PackageAccess.READ);
            xssfReader = new XSSFReader(pkg);
            styles = xssfReader.getStylesTable();
            stringsTable = new ReadOnlySharedStringsTable(pkg);
        } catch (Exception e) {
            throw new ExcelReaderException("打开Excel异常", e);
        }
    }

    /**
     * 读取Excel，按照Excel的格式来获得输出，这个方法由于有Excel格式在内，因此会占用较大的内存空间
     *
     * @return
     * @throws ExcelReaderException
     */
    public Map<String, XSheet> processExcel() throws ExcelReaderException {
        XSSFReader.SheetIterator inputStreamIterator = null;
        Map<String, XSheet> sheetMap = new HashMap<>();
        try {
            inputStreamIterator = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        } catch (Exception e) {
            throw new ExcelReaderException("处理Excel异常", e);
        }
        while (inputStreamIterator.hasNext()) {
            sheetInputStream = inputStreamIterator.next();
            String sheetName = inputStreamIterator.getSheetName();
            XSheet sheet = processSheet(styles, stringsTable, sheetInputStream);
            sheet.setSheetName(sheetName);
            sheetMap.put(sheetName, sheet);
        }
        return sheetMap;
    }

    /**
     * 读取Excel，只获得Excel的值,占用空间较小
     *
     * @return
     * @throws ExcelReaderException
     */
    public Map<String, List<List<String>>> processExcelValue() throws ExcelReaderException {
        XSSFReader.SheetIterator inputStreamIterator = null;
        Map<String, List<List<String>>> sheetMap = new HashMap<>();
        try {
            inputStreamIterator = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        } catch (Exception e) {
            throw new ExcelReaderException("处理Excel异常", e);
        }
        while (inputStreamIterator.hasNext()) {
            sheetInputStream = inputStreamIterator.next();
            String sheetName = inputStreamIterator.getSheetName();
            List<List<String>> sheet = processSheetValue(styles, stringsTable, sheetInputStream);
            sheetMap.put(sheetName, sheet);
        }
        return sheetMap;
    }

    private static XSheet processSheet(StylesTable styles, ReadOnlySharedStringsTable strings, InputStream sheetInputStream)
            throws ExcelReaderException {
        try {
            XMLReader sheetParser = SAXHelper.newXMLReader();
            XSheet sheet = new XSheet();
            XSSFSheetXMLHandler.SheetContentsHandler handler = new XSSFSheetXMLHandler.SheetContentsHandler() {
                XRow rowContents = null;
                List<String> titleRow = new ArrayList<>();
                int currentRow = 0;

                @Override
                public void startRow(int rowNum) {
                    rowContents = new XRow(rowNum);
                    currentRow = rowNum;
                }

                @Override
                public void endRow(int rowNum) {
                    sheet.add(rowContents);
                }

                @Override
                public void cell(String cellReference, String formattedValue, XSSFComment comment) {
                    int colNo = Integer.parseInt(patternNotNumber.matcher(cellReference).replaceAll(""));
                    if (0 == currentRow) {
                        titleRow.add(patternNumber.matcher(cellReference).replaceAll(""));
                        rowContents.addCell(new XCell(currentRow, colNo, cellReference, formattedValue));
                    }
                    //判断空单元格，补齐为空字符串
                    else {
                        String col = patternNumber.matcher(cellReference).replaceAll("");
                        int index = titleRow.indexOf(col);
                        if (index == rowContents.getCellList().size()) {
                            rowContents.addCell(new XCell(currentRow, colNo, cellReference, formattedValue));
                        } else {
                            int cellSize = rowContents.getCellList().size();
                            int times = index - cellSize;
                            for (int i = 0; i < times; i++) {
                                rowContents.addCell(new XCell(currentRow, colNo, titleRow.get(cellSize + i) + colNo, ""));
                            }
                            rowContents.addCell(new XCell(currentRow, colNo, cellReference, formattedValue));
                        }
                    }
                }
            };

            sheetParser.setContentHandler(new XSSFSheetXMLHandler(styles, strings, handler, false));
            sheetParser.parse(new InputSource(sheetInputStream));
            return sheet;
        } catch (Exception e) {
            throw new ExcelReaderException("处理Excel异常", e);
        }

    }

    private static List<List<String>> processSheetValue(StylesTable styles, ReadOnlySharedStringsTable strings, InputStream sheetInputStream)
            throws ExcelReaderException {
        try {
            XMLReader sheetParser = SAXHelper.newXMLReader();
            List<List<String>> sheetContents = new ArrayList<>();
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
                                rowContents.add("");
                            }
                            rowContents.add(formattedValue);
                        }
                    }
                }
            };

            sheetParser.setContentHandler(new XSSFSheetXMLHandler(styles, strings, handler, false));
            sheetParser.parse(new InputSource(sheetInputStream));
            return sheetContents;
        } catch (Exception e) {
            throw new ExcelReaderException("处理Excel异常", e);
        }

    }


}
