package itstack.demo.upload.common;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


public class PoiUtils {
    private static final Logger LOGGER = LoggerFactory
            .getLogger(PoiUtils.class);
    /* LONG */
    protected static final String LONG = "java.lang.Long";
    /* SHORT */
    protected static final String SHORT = "java.lang.Short";
    /* INT */
    protected static final String INT = "java.lang.Integer";
    /* STRING */
    protected static final String STRING = "java.lang.String";
    /* DATE */
    protected static final String DATE = "java.sql.Timestamp";
    /* BIG */
    protected static final String BIG = "java.math.BigDecimal";
    /* CLOB */
    protected static final String CLOB = "oracle.sql.CLOB";

    public static void main(String[] args) throws FileNotFoundException {

        String path = "C:\\Users\\fuzhengwei\\Desktop\\test.xls";
        File file = new File(path);
        InputStream inputStream = new FileInputStream(file);
        int count = getRecordsCountReadStream(inputStream, 0, false, 0);

        List<String[]> list = readRecordsInputPath(path, true, 1);
        System.out.println(GsonUtils.toJson(list));
    }


    /**
     * 通过文件路径获取Excel读取行数
     *
     * @param path        文件路径，只接受xls或xlsx结尾
     * @param isHeader    是否表头
     * @param headerCount 表头行数
     * @return count 如果文件路径为空，返回0；
     */
    public static int getRecordsCountReadPath(String path, boolean isHeader, int headerCount) {

        int count = 0;

        if (path == null) {
            return count;
        } else if (!path.endsWith("xls") && !path.endsWith("xlsx")
                && !path.endsWith("XLS") && !path.endsWith("XLSX")) {
            return count;
        }

        try {
            File file = new File(path);
            InputStream inputStream = new FileInputStream(file);
            Workbook hwb = null;
            if (path.endsWith("xls") || path.endsWith("XLS")) {
                hwb = new HSSFWorkbook(inputStream);
            } else if (path.endsWith("xlsx") || path.endsWith("XLSX")) {
                hwb = new XSSFWorkbook(inputStream);
            }

            if (null == hwb) {
                return count;
            }

            Sheet sheet = hwb.getSheetAt(0);//暂定只取首页签
            int begin = sheet.getFirstRowNum();
            if (isHeader) {
                begin += headerCount;
            }
            int end = sheet.getLastRowNum();
            for (int i = begin; i <= end; i++) {
                if (null == sheet.getRow(i)) {
                    continue;
                }
                count++;
            }

        } catch (FileNotFoundException e) {
            LOGGER.error("excel解析:", e);
            return 0;
        } catch (IOException e) {
            LOGGER.error("excel解析:", e);
            return 0;
        }
        return count;
    }

    /**
     * 通过文件流获取Excel读取行数
     *
     * @param inputStream 文件流
     * @param type        类型，0为xls，1为xlsx；
     * @param isHeader    是否表头
     * @param headerCount 表头行数
     * @return count 如果文件路径为空，返回0；
     */
    public static int getRecordsCountReadStream(InputStream inputStream, int type, boolean isHeader, int headerCount) {

        int count = 0;
        if (type != 0 && type != 1) {
            return count;
        }

        try {
            Workbook hwb = null;
            if (type == 0) {
                hwb = new HSSFWorkbook(inputStream);
            } else if (type == 1) {
                hwb = new XSSFWorkbook(inputStream);
            }

            if (null == hwb) {
                return count;
            }

            Sheet sheet = hwb.getSheetAt(0);
            int begin = sheet.getFirstRowNum();
            if (isHeader) {
                begin += headerCount;
            }
            int end = sheet.getLastRowNum();
            for (int i = begin; i <= end; i++) {
                if (null == sheet.getRow(i)) {
                    continue;
                }
                count++;
            }
        } catch (FileNotFoundException e) {
            LOGGER.error("excel解析:", e);
            return 0;
        } catch (IOException e) {
            LOGGER.error("excel解析:", e);
            return 0;
        }
        return count;
    }

    /**
     * 通过文件流获取Excel读取
     *
     * @param inputStream 文件流
     * @param type        类型，0为xls，1为xlsx；
     * @param isHeader    是否表头
     * @param headerCount 表头行数
     * @return poiList 如果文件路径为空，返回0；
     */
    public static List<String[]> readRecordsInputStream(InputStream inputStream, int type, boolean isHeader, int headerCount) {
        List<String[]> poiList = new ArrayList<String[]>();
        if (type != 0 && type != 1) {
            return null;
        }
        if (type == 0) {
            poiList = readXLSRecords(inputStream, isHeader, headerCount);
        } else if (type == 1) {
            poiList = readXLSXRecords(inputStream, isHeader, headerCount);
        }
        return poiList;
    }

    /**
     * 通过文件路径获取Excel读取
     *
     * @param path        文件路径，只接受xls或xlsx结尾
     * @param isHeader    是否表头
     * @param headerCount 表头行数
     * @return count 如果文件路径为空，返回0；
     */
    public static List<String[]> readRecordsInputPath(String path, boolean isHeader, int headerCount) {
        List<String[]> poiList = new ArrayList<String[]>();
        if (path == null) {
            return null;
        } else if (!path.endsWith("xls") && !path.endsWith("xlsx")
                && !path.endsWith("XLS") && !path.endsWith("XLSX")) {
            return null;
        }
        File file = new File(path);
        try {
            InputStream inputStream = new FileInputStream(file);

            if (path.endsWith("xls") || path.endsWith("XLS")) {
                poiList = readXLSRecords(inputStream, isHeader, headerCount);
            } else if (path.endsWith("xlsx") || path.endsWith("XLSX")) {
                poiList = readXLSXRecords(inputStream, isHeader, headerCount);
            }
        } catch (Exception e) {
            LOGGER.error("excel解析:", e);
            return null;
        }
        return poiList;
    }

    /**
     * 解析EXCEL2003文件流
     * 如果一行记录的行中或行尾出现空格，POI工具类可能会跳过空格不做处理，所以默认第一行是表头，所有待解析的记录都以表头为准
     *
     * @param inputStream 输入流
     * @param isHeader    是否要跳过表头
     * @param headerCount 表头占用行数
     * @return 返回一个字符串数组List
     */
    public static List<String[]> readXLSRecords(InputStream inputStream, boolean isHeader, int headerCount) {
        List<String[]> poiList = new ArrayList<String[]>();
        try {
            HSSFWorkbook wbs = new HSSFWorkbook(inputStream);
            HSSFSheet childSheet = wbs.getSheetAt(0);
            //获取表头
            int begin = childSheet.getFirstRowNum();
            HSSFRow firstRow = childSheet.getRow(begin);
            int cellTotal = firstRow.getPhysicalNumberOfCells();
            //是否跳过表头解析数据
            if (isHeader) {
                begin += headerCount;
            }
            //逐行获取单元格数据
            for (int i = begin; i <= childSheet.getLastRowNum(); i++) {
                HSSFRow row = childSheet.getRow(i); //一行的所有单元格格式都是常规的情况下，返回的row为null
                if (null != row) {
                    String[] cells = new String[cellTotal];
                    for (int k = 0; k < cellTotal; k++) {
                        HSSFCell cell = row.getCell(k);
                        cells[k] = getStringXLSCellValue(cell);
                    }
                    poiList.add(cells);
                }
            }
        } catch (Exception e) {
            LOGGER.error("excel解析:", e);
            return null;
        }
        return poiList;
    }

    /**
     * 解析EXCEL2003文件流
     * 如果一行记录的行中或行尾出现空格，POI工具类可能会跳过空格不做处理，所以默认第一行是表头，所有待解析的记录都以表头为准
     * 该解析方法只适用于表头占用一行的情况
     *
     * @param inputStream 输入流
     * @param isHeader    是否要跳过表头
     * @param headerCount 表头占用行数
     * @param maxColNum   最大列数，适用于多表头
     * @return 返回一个字符串数组List
     */
    public static List<String[]> readXLSRecords(InputStream inputStream, boolean isHeader, int headerCount, int maxColNum) {
        List<String[]> poiList = new ArrayList<String[]>();
        try {
            HSSFWorkbook wbs = new HSSFWorkbook(inputStream);
            HSSFSheet childSheet = wbs.getSheetAt(0);
            //获取表头
            int begin = childSheet.getFirstRowNum();
            //HSSFRow firstRow = childSheet.getRow(begin);
            //int cellTotal = firstRow.getPhysicalNumberOfCells();
            //是否跳过表头解析数据
            if (isHeader) {
                begin += headerCount;
            }
            //逐行获取单元格数据
            for (int i = begin; i <= childSheet.getLastRowNum(); i++) {
                HSSFRow row = childSheet.getRow(i); //一行的所有单元格格式都是常规的情况下，返回的row为null
                String[] cells = new String[maxColNum]; //空行对应空串数组
                for (int k = 0; k < maxColNum; k++) {
                    HSSFCell cell = row == null ? null : row.getCell(k);
                    cells[k] = getStringXLSCellValue(cell);
                }
                poiList.add(cells);
            }
        } catch (Exception e) {
            LOGGER.error("excel解析:", e);
            return null;
        }
        return poiList;
    }

    /**
     * 解析EXCEL2007文件流
     * 如果一行记录的行中或行尾出现空格，POI工具类可能会跳过空格不做处理，所以默认第一行是表头，所有待解析的记录都以表头为准
     * 该处理方法中，表头对应都占用一行
     *
     * @param inputStream 输入流
     * @param isHeader    是否要跳过表头
     * @param headerCount 表头占用行数
     * @return 返回一个字符串数组List
     */
    public static List<String[]> readXLSXRecords(InputStream inputStream, boolean isHeader, int headerCount) {
        List<String[]> poiList = new ArrayList<String[]>();
        try {
            XSSFWorkbook wbs = new XSSFWorkbook(inputStream);
            XSSFSheet childSheet = wbs.getSheetAt(0);
            //获取表头
            int begin = childSheet.getFirstRowNum();
            XSSFRow firstRow = childSheet.getRow(begin);
            int cellTotal = firstRow.getPhysicalNumberOfCells();
            //是否跳过表头解析数据
            if (isHeader) {
                begin += headerCount;
            }
            for (int i = begin; i <= childSheet.getLastRowNum(); i++) {
                XSSFRow row = childSheet.getRow(i);  //一行的所有单元格格式都是常规的情况下，返回的row为null
                if (null != row) {
                    String[] cells = new String[cellTotal];
                    for (int k = 0; k < cellTotal; k++) {
                        XSSFCell cell = row.getCell(k);
                        cells[k] = getStringXLSXCellValue(cell);
                    }
                    poiList.add(cells);
                }
            }
        } catch (Exception e) {
            LOGGER.error("excel解析:", e);
            return null;
        }
        return poiList;
    }

    /**
     * 解析EXCEL2007文件流
     * 如果一行记录的行中或行尾出现空格，POI工具类可能会跳过空格不做处理，所以默认第一行是表头，所有待解析的记录都以表头为准
     * 该处理方法中，表头对应都占用一行
     *
     * @param inputStream 输入流
     * @param isHeader    是否要跳过表头
     * @param headerCount 表头占用行数
     * @param maxColNum   最大列数，适用于多表头的情况
     * @return 返回一个字符串数组List
     */
    public static List<String[]> readXLSXRecords(InputStream inputStream, boolean isHeader, int headerCount, int maxColNum) {
        List<String[]> poiList = new ArrayList<String[]>();
        try {
            XSSFWorkbook wbs = new XSSFWorkbook(inputStream);
            XSSFSheet childSheet = wbs.getSheetAt(0);
            //获取表头
            int begin = childSheet.getFirstRowNum();
            //XSSFRow firstRow = childSheet.getRow(begin);
            //int cellTotal = firstRow.getPhysicalNumberOfCells();
            //是否跳过表头解析数据
            if (isHeader) {
                begin += headerCount;
            }
            for (int i = begin; i <= childSheet.getLastRowNum(); i++) {
                XSSFRow row = childSheet.getRow(i);  //一行的所有单元格格式都是常规的情况下，返回的row为null
                String[] cells = new String[maxColNum];  //空行对应空串数组
                for (int k = 0; k < maxColNum; k++) {
                    XSSFCell cell = row == null ? null : row.getCell(k);
                    cells[k] = getStringXLSXCellValue(cell);
                }
                poiList.add(cells);
            }
        } catch (Exception e) {
            LOGGER.error("excel解析:", e);
            return null;
        }
        return poiList;
    }

    /**
     * 获取单元格数据内容为字符串类型的数据
     *
     * @param cell Excel单元格
     * @return String 单元格数据内容
     */
    private static String getStringXLSCellValue(HSSFCell cell) {
        String strCell = "";
        if (cell == null) {
            return "";
        }

        //将数值型参数转成文本格式，该算法不能保证1.00这种类型数值的精确度
        DecimalFormat df = (DecimalFormat) NumberFormat.getPercentInstance();
        StringBuffer sb = new StringBuffer();
        sb.append("0");
        df.applyPattern(sb.toString());

        switch (cell.getCellType()) {
            case HSSFCell.CELL_TYPE_STRING:
                strCell = cell.getStringCellValue();
                break;
            case HSSFCell.CELL_TYPE_NUMERIC:
                double value = cell.getNumericCellValue();
                while (Double.parseDouble(df.format(value)) != value) {
                    if ("0".equals(sb.toString())) {
                        sb.append(".0");
                    } else {
                        sb.append("0");
                    }
                    df.applyPattern(sb.toString());
                }
                strCell = df.format(value);
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN:
                strCell = String.valueOf(cell.getBooleanCellValue());
                break;
            case HSSFCell.CELL_TYPE_BLANK:
                strCell = "";
                break;
            default:
                strCell = "";
                break;
        }
        if (strCell == null || "".equals(strCell)) {
            return "";
        }
        return strCell;
    }

    /**
     * 获取单元格数据内容为字符串类型的数据
     *
     * @param cell Excel单元格
     * @return String 单元格数据内容
     */
    private static String getStringXLSXCellValue(XSSFCell cell) {
        String strCell = "";
        if (cell == null) {
            return "";
        }
        //将数值型参数转成文本格式，该算法不能保证1.00这种类型数值的精确度
        DecimalFormat df = (DecimalFormat) NumberFormat.getPercentInstance();
        StringBuffer sb = new StringBuffer();
        sb.append("0");
        df.applyPattern(sb.toString());

        switch (cell.getCellType()) {
            case XSSFCell.CELL_TYPE_STRING:
                strCell = cell.getStringCellValue();
                break;
            case XSSFCell.CELL_TYPE_NUMERIC:
                double value = cell.getNumericCellValue();
                while (Double.parseDouble(df.format(value)) != value) {
                    if ("0".equals(sb.toString())) {
                        sb.append(".0");
                    } else {
                        sb.append("0");
                    }
                    df.applyPattern(sb.toString());
                }
                strCell = df.format(value);
                break;
            case XSSFCell.CELL_TYPE_BOOLEAN:
                strCell = String.valueOf(cell.getBooleanCellValue());
                break;
            case XSSFCell.CELL_TYPE_BLANK:
                strCell = "";
                break;
            default:
                strCell = "";
                break;
        }
        if (strCell == null || "".equals(strCell)) {
            return "";
        }
        return strCell;
    }

}