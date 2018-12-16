package poi;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

/**
 * 自定义sheet
 *
 * @author zhangpengyue
 * @date 2018/12/16
 */
public class XSheet implements Serializable {
    private static final long serialVersionUID = -5809925631136243201L;

    private String sheetName;
    private List<XRow> rows;

    public XSheet() {
        rows = new ArrayList<>();
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public List<XRow> getRows() {
        return rows;
    }

    public void setRows(List<XRow> rows) {
        this.rows = rows;
    }

    public synchronized void add(XRow row) {
        if (null == rows) {
            rows = new ArrayList<>();
        }
        rows.add(row);
    }

    @Override
    public String toString() {
        return "XSheet{" +
                "sheetName='" + sheetName + '\'' +
                ", rows=" + rows +
                '}';
    }
}
