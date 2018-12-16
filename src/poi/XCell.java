package poi;

import java.io.Serializable;

/**
 * 自定义的cell
 *
 * @author zhangpengyue
 * @date 2018/12/16
 */
public class XCell implements Serializable {

    private static final long serialVersionUID = 4031516554949038717L;
    /**
     * 行号
     */
    private int rowNum;
    /**
     * 列号
     */
    private int colNum;

    /**
     * 单元格别称
     */
    private String cellReference;

    private String cellValue;

    public XCell(int rowNum, int colNum, String cellReference, String cellValue) {
        this.rowNum = rowNum;
        this.colNum = colNum;
        this.cellReference = cellReference;
        this.cellValue = cellValue;
    }

    public int getRowNum() {
        return rowNum;
    }

    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    public int getColNum() {
        return colNum;
    }

    public void setColNum(int colNum) {
        this.colNum = colNum;
    }

    public String getCellValue() {
        return cellValue;
    }

    public void setCellValue(String cellValue) {
        this.cellValue = cellValue;
    }

    public String getCellReference() {
        return cellReference;
    }

    public void setCellReference(String cellReference) {
        this.cellReference = cellReference;
    }

    @Override
    public String toString() {
        return "XCell{" +
                "rowNum=" + rowNum +
                ", colNum=" + colNum +
                ", cellReference=" + cellReference +
                ", cellValue='" + cellValue + '\'' +
                '}';
    }
}
