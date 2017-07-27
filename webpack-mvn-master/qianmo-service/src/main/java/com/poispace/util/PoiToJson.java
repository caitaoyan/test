package com.poispace.util;

import com.poispace.pojo.CellModel;
import net.sf.json.JSONObject;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class PoiToJson {

    private Sheet sheet; // 表格类实例
    LinkedList[] result; // 保存每个单元格的数据 ，使用的是一种链表数组的结构
    Workbook workBook;
    //String contents = "";
    int firstEmptyRow;
    int secondEmptyRow;
    List<CellModel> CellList = null;

    /**
     * 读取excel文件，创建表格实例
     *
     * @param filePath
     */
    private void loadExcel(String filePath) {
        CellList = new ArrayList();
        FileInputStream inStream = null;
        try {
            inStream = new FileInputStream(new File(filePath));
            workBook = create(inStream);
            sheet = workBook.getSheetAt(0);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (inStream != null) {
                    inStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 获取每个单元格的值
     *
     * @param cell
     * @param poiType
     * @return
     */
    private CellModel getCellValue(Cell cell, String poiType) {
        CellModel model = new CellModel();
        String cellValue = "";
        String cellType = "";
        CellAddress address = null;
        DataFormatter formatter = new DataFormatter();
        StringBuilder address_str;
        int firstColumn;
        int lastColumn;
        int firstRow;
        int lastRow;
        int firstColumn1 = 0;
        //int lastColumn1 = 0;
        int firstRow1 = 0;
        //int lastRow1 = 0;
        String area = "";
        //System.out.println("address:"+cell.getAddress());
        if (cell != null) {

            address = cell.getAddress();
            int row = cell.getRowIndex();
            int column = cell.getColumnIndex();
            int sheetMergeCount = sheet.getNumMergedRegions();

            //System.out.println("cell.getAddress():" + cell.getAddress());

            for (int i = 0; i < sheetMergeCount; i++) {
                CellRangeAddress ca = sheet.getMergedRegion(i);

                firstColumn = ca.getFirstColumn();
                lastColumn = ca.getLastColumn();
                firstRow = ca.getFirstRow();
                lastRow = ca.getLastRow();

                if (row >= firstRow && row <= lastRow && column >= firstColumn && column <= lastColumn) {
                    area = "" + excelColIndexToStr(firstColumn) + "_" + (firstRow + 1) + "_" + excelColIndexToStr(lastColumn)
                            + "_" + (lastRow + 1);
                    firstRow1 = firstRow;
                    firstColumn1 = firstColumn;
                    //lastRow1 = lastRow;
                    //lastColumn1 = lastColumn;
                }
            }

            // 判断单元格数据的类型，不同类型调用不同的方法
            //System.out.println("判断单元格数据的类型，不同类型调用不同的方法");
            //System.out.println("Cell.CELL_TYPE_NUMERIC:"+cell.getCellType());
            switch (cell.getCellType()) {
                // 数值类型
                case Cell.CELL_TYPE_NUMERIC:
                    // 进一步判断 ，单元格格式是日期格式
                    if (DateUtil.isCellDateFormatted(cell)) {
                        cellValue = formatter.formatCellValue(cell);
                        cellType = "date";
                    } else {
                        // 数值
                        String dataFormatString = cell.getCellStyle().getDataFormatString();
                        //System.out.println("dataFormatString:" + dataFormatString);
                        if (dataFormatString.indexOf("%") != -1) {
                            cellType = "percent";
                            String format = "0";
                            int j = 0;
                            for (int i = 0; i < dataFormatString.length(); i++) {
                                if ("0".equals(String.valueOf(dataFormatString.charAt(i)))) {
                                    if (j > 2) {
                                        format += "0";
                                    } else if (j == 2) {
                                        format += ".0";
                                    }
                                }
                                j++;
                            }
                            //System.out.println("format:" + format);
                            DecimalFormat df = new DecimalFormat(format);
                            cellValue = df.format(cell.getNumericCellValue() * 100) + "%";
                        } else {
                            cellType = "numerical";
                            double value = cell.getNumericCellValue();
                            int intValue = (int) value;
                            cellValue = value - intValue == 0 ? String.valueOf(intValue) : String.valueOf(value);
                            //cellValue = String.valueOf(value);
                            //System.out.println("cellValue:" + cellValue);
                        }
                    }
                    break;
                case Cell.CELL_TYPE_STRING:
                    cellType = "text";
                    cellValue = cell.getStringCellValue();
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    cellType = "boolean";
                    cellValue = String.valueOf(cell.getBooleanCellValue());
                    break;
                // 判断单元格是公式格式，需要做一种特殊处理来得到相应的值
                case Cell.CELL_TYPE_FORMULA: {
                    cellType = "formula";
                    try {
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    } catch (IllegalStateException e) {
                        // String.valueOf(cell.getRichStringCellValue());
                    }
                }
                break;
                case Cell.CELL_TYPE_BLANK:
                    cellType = "blank";
                    cellValue = "null";
                    break;
                case Cell.CELL_TYPE_ERROR:
                    cellType = "error";
                    cellValue = "null";
                    break;
                default:
                    cellValue = cell.toString().trim();
                    break;
            }

        }
        if (address != null) {
            HSSFCellStyle cellStyleHSSF = null;
            XSSFCellStyle cellStyleXSSF = null;
            if ("HSSF".equals(poiType)) {
                cellStyleHSSF = (HSSFCellStyle) cell.getCellStyle();
            } else if ("XSSF".equals(poiType)) {
                cellStyleXSSF = (XSSFCellStyle) cell.getCellStyle();
            }

            try {
                model.setFormula("=" + cell.getCellFormula());
            } catch (Exception e) {
                model.setFormula("");
            }

            address_str = new StringBuilder(address.toString());
            for (int k = 0; k < address_str.length(); k++) {
                int address_int = (int) address_str.charAt(k);
                if (address_int >= 48 && address_int <= 57) {
                    address_str.insert(k, "_");
                    break;
                }
            }

            if (address.getRow() == firstRow1 && address.getColumn() == firstColumn1) {
                model.setArea(area);
            } else {
                model.setArea(address_str.toString() + "_" + address_str.toString());
            }

            model.setCellName(address_str.toString());
            model.setContent(cellValue.replace("\n", "").replace("\r", "").replace("\t", ""));
            model.setFormat(cellType);

            //System.out.println("设置字体及背景色");
            //设置字体及背景色
            if ("HSSF".equals(poiType)) {
                HSSFCellStyle cellStyle = (HSSFCellStyle) cell.getCellStyle();
                model.setFont(cellStyle.getFont(workBook).getFontName());
                model.setFontSize(cellStyle.getFont(workBook).getFontHeightInPoints() + "pt");
                HSSFColor color = cellStyle.getFillForegroundColorColor();
                HSSFColor color2 = cellStyle.getFillBackgroundColorColor();
                model.setBackgroundColor(convertToStardColor(color));
                model.setForegroundColor(convertToStardColor(color2));
            } else if ("XSSF".equals(poiType)) {
                XSSFCellStyle cellStyle = (XSSFCellStyle) cell.getCellStyle();
                model.setFont(cellStyle.getFont().getFontName());
                model.setFontSize(cellStyle.getFont().getFontHeightInPoints() + "pt");
                if (cellStyle.getFillBackgroundXSSFColor() != null && cellStyle.getFillBackgroundXSSFColor().getARGBHex() != null) {
                    model.setForegroundColor(cellStyle.getFillBackgroundXSSFColor().getARGBHex().replaceFirst("FF", "#"));
                }
                if (cellStyle.getFillForegroundXSSFColor() != null && cellStyle.getFillForegroundXSSFColor().getARGBHex() != null) {
                    model.setBackgroundColor(cellStyle.getFillForegroundXSSFColor().getARGBHex().replaceFirst("FF", "#"));
                }
            }

            //设置缩进
            //System.out.println("设置缩进");
            model.setIndentation(cell.getCellStyle().getIndention());

            short bold = 0;
            short alignment = 0;
            short vertical = 0;
            boolean italic = false;
            CellStyle cellStyle2 = null;
            if ("HSSF".equals(poiType)) {
                bold = cellStyleHSSF.getFont(workBook).getBoldweight();
                alignment = cellStyleHSSF.getAlignment();
                vertical = cellStyleHSSF.getVerticalAlignment();
                italic = cellStyleHSSF.getFont(workBook).getItalic();
                cellStyle2 = cellStyleHSSF;
            } else if ("XSSF".equals(poiType)) {
                bold = cellStyleXSSF.getFont().getBoldweight();
                alignment = cellStyleXSSF.getAlignment();
                vertical = cellStyleXSSF.getVerticalAlignment();
                italic = cellStyleXSSF.getFont().getItalic();
                cellStyle2 = cellStyleXSSF;
            }

            //水平对齐方式
            //System.out.println("水平对齐方式");
            if (alignment == HSSFCellStyle.ALIGN_CENTER) {
                model.setAlignment("center");//水平居中对齐
            } else if (alignment == HSSFCellStyle.ALIGN_LEFT) {
                model.setAlignment("left");//居左对齐
            } else if (alignment == HSSFCellStyle.ALIGN_RIGHT) {
                model.setAlignment("right");//居右对齐
            }

            //垂直对齐方式
            if (vertical == HSSFCellStyle.VERTICAL_CENTER) {
                model.setVertical("center");//垂直居中
            } else if (vertical == HSSFCellStyle.VERTICAL_BOTTOM) {
                model.setVertical("bottom");//底端对齐
            } else if (vertical == HSSFCellStyle.VERTICAL_TOP) {
                model.setVertical("top");//顶端对齐
            }

            //加粗
            if (bold == HSSFFont.BOLDWEIGHT_BOLD) {
                model.setBold("true");
            } else {
                model.setBold("false");
            }

            //斜体
            if (italic) {
                model.setItalic("true");
            } else {
                model.setItalic("false");
            }

            //单元格格式
            //System.out.println("单元格格式");
            int dataFormat = cellStyle2.getDataFormat();
            model.setDataFormat(dataFormat);
            if (cellType.equals("formula") || cellType.equals("blank")) {//公式格式 默认格式
                if (dataFormat < 5) {
                    model.setFormat("numerical");//数值格式
                } else if (dataFormat < 9 && dataFormat > 4) {
                    model.setFormat("money");//货币格式
                } else if (dataFormat < 11 && dataFormat > 8) {
                    model.setFormat("percent");//百分比格式
                } else if (dataFormat == 11) {
                    model.setFormat("notation");//科学记数格式
                } else if (dataFormat < 14 && dataFormat > 11) {
                    model.setFormat("fraction");//分数格式
                } else if (dataFormat < 23 && dataFormat > 13) {
                    model.setFormat("date");//日期格式
                } else if (null != model.getContent()) {
                    Pattern pattern = Pattern.compile("([-\\+]?[1-9]([0-9]*)(\\.[0-9]+)?)|(^0$)");
                    Matcher isNum = pattern.matcher(model.getContent());
                    if (!isNum.matches()) {
                        model.setFormat("number");//数字格式
                        // System.err.println("数值格式");
                    }
                } else {
                    System.err.println(model.getContent());
                }
            }

            //单元格自动换行
            if (cell.getCellStyle().getWrapText()) {
                model.setWrapText("true");
            } else {
                model.setWrapText("false");
            }

            //设置行高列宽
            //int cellWidth = cell.getSheet().getColumnWidth(address.getColumn());
            //model.setWidth(String.valueOf(cellWidth));
            model.setWidth(String.valueOf(cell.getSheet().getColumnWidthInPixels(address.getColumn())));
            float cellHeight = cell.getRow().getHeightInPoints();
            model.setHeight(String.valueOf(cellHeight));
            //System.out.println("cellWidth:" + cellWidth);
            //System.out.println("width:" + cell.getSheet().getColumnWidthInPixels(address.getColumn()));
            //double width = SheetUtil.getCellWidth(cell,8,formatter,false);
            //System.out.println("width2:"+width);

            //设置单元格边框
            if (cell.getCellStyle().getBorderBottom() == 0) {
                model.setBottomFrame("false");
            } else {
                model.setBottomFrame("true");
            }
            if (cell.getCellStyle().getBorderTop() == 0) {
                model.setTopFrame("false");
            } else {
                model.setTopFrame("true");
            }
            if (cell.getCellStyle().getBorderLeft() == 0) {
                model.setLeftFrame("false");
            } else {
                model.setLeftFrame("true");
            }
            if (cell.getCellStyle().getBorderRight() == 0) {
                model.setRightFrame("false");
            } else {
                model.setRightFrame("true");
            }
        }
        return model;
    }

    public static boolean isNumber(String str) {
        for (int i = 0; i < str.length(); i++) {
            if (!Character.isDigit(str.charAt(i))) {
                return false;
            }
        }
        return true;
    }

    /**
     * HSSFCOLOR颜色转为16进制
     *
     * @param hc
     * @return
     */
    private String convertToStardColor(HSSFColor hc) {
        StringBuffer sb = new StringBuffer("");
        if (hc != null) {
            int a = HSSFColor.AUTOMATIC.index;
            int b = hc.getIndex();
            if (a == b) {
                return "";
            }
            sb.append("#");
            for (int i = 0; i < hc.getTriplet().length; i++) {
                String str;
                String str_tmp = Integer.toHexString(hc.getTriplet()[i]);
                if (str_tmp != null && str_tmp.length() < 2) {
                    str = "0" + str_tmp;
                } else {
                    str = str_tmp;
                }
                sb.append(str);
            }
        }
        return sb.toString();
    }

    /**
     * Excel列数字转为字母 Excel column index begin 1
     *
     * @param columnIndex
     * @return
     */
    public static String excelColIndexToStr(int columnIndex) {
        columnIndex = columnIndex + 1;
        if (columnIndex <= 0) {
            return null;
        }
        String columnStr = "";
        columnIndex--;
        do {
            if (columnStr.length() > 0) {
                columnIndex--;
            }
            columnStr = ((char) (columnIndex % 26 + (int) 'A')) + columnStr;
            columnIndex = (int) ((columnIndex - columnIndex % 26) / 26);
        } while (columnIndex > 0);
        return columnStr;
    }

    /**
     * 初始化表格中的每一行，并得到每一个单元格的值
     *
     * @param poiType
     * @return
     */
    public List<CellModel> init(String poiType) {
        int rowNum = sheet.getLastRowNum() + 1;
        result = new LinkedList[rowNum];
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < rowNum; i++) {
            int emptyrow = 0;
            int emptycolumn = 0;
            Row row = sheet.getRow(i);
            // 每有新的一行，创建一个新的LinkedList对象
            result[i] = new LinkedList();
            for (int j = 0; j < row.getLastCellNum(); j++) {
                Cell cell = row.getCell(j);
                // 获取单元格的值
                CellModel everyem = null;
                if ("HSSF".equals(poiType)) {
                    everyem = getCellValue(cell, "HSSF");
                } else if ("XSSF".equals(poiType)) {
                    everyem = getCellValue(cell, "XSSF");
                }
                // if(null==everyem.getFormat()||everyem.getFormat().equals("默认格式")&&everyem.getContent().trim().isEmpty()){
                if (null == everyem.getContent() || everyem.getContent().trim().isEmpty()) {
                    emptycolumn = emptycolumn + 1;
                }
                if (everyem.allEmpty()) {

                } else {
                    CellList.add(everyem);
                }
                if (emptycolumn == row.getLastCellNum()) {
                    emptyrow++;
                }
            }
            if (emptyrow == 1) {
                if (secondEmptyRow == 0) {
                    firstEmptyRow = 0;
                    secondEmptyRow = row.getRowNum();
                } else {
                    firstEmptyRow = secondEmptyRow;
                    secondEmptyRow = row.getRowNum();

                }
                if ((secondEmptyRow - firstEmptyRow) == 1) {

                    break;
                }

            }
        }
        return CellList;
    }

    /**
     * 读取excel并转为json
     *
     * @param readpath
     * @param writepath
     * @param name
     * @return
     */
    public String readExcel(String readpath, String writepath, String name) {
        // PoiToJson poi = new PoiToJson();
        //System.out.println(readpath);
        loadExcel(readpath);
        FileOutputStream fop = null;
        File file;

        Map<String, Object> map = new HashMap<String, Object>();
        if (name.indexOf(".xlsx") != -1) {
            map.put("CellList", init("XSSF"));
        } else {
            map.put("CellList", init("HSSF"));
        }

        try {
            file = new File(writepath + File.separator + name.substring(0, name.lastIndexOf(".")) + ".txt");
            //file = new File(writepath + "/" + name.substring(0, name.lastIndexOf(".")) + ".txt");
            if (!file.getParentFile().exists()) {
                // 如果目标文件所在的目录不存在，则创建父目录
                System.out.println("目标文件所在目录不存在，准备创建它！");
                if (!file.getParentFile().mkdirs()) {
                    System.out.println("创建目标文件所在目录失败！");
                }
            }
            fop = new FileOutputStream(file);

            if (!file.exists()) {
                file.createNewFile();
            }
            JSONObject jsonObjectFromMap = JSONObject.fromObject(map);
            String excel_json = jsonObjectFromMap.getString("CellList");
            //System.out.println("excel_json:" + excel_json);
            byte[] contentInBytes = excel_json.getBytes();

            fop.write(contentInBytes);
            fop.flush();
            fop.close();

            //System.out.println("Done");
            return excel_json;
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (fop != null) {
                    fop.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return null;
    }

    /**
     * 判断解析类型
     *
     * @param in
     * @return
     * @throws Exception
     */
    public static Workbook create(InputStream in) throws Exception {
        if (!in.markSupported()) {
            in = new PushbackInputStream(in, 8);
        }
        if (POIFSFileSystem.hasPOIFSHeader(in)) {
            return new HSSFWorkbook(in);
        }
        if (POIXMLDocument.hasOOXMLHeader(in)) {
            return new XSSFWorkbook(OPCPackage.open(in));
        }
        throw new IllegalArgumentException("你的excel版本目前poi解析不了");

    }

    /**
     * 解析excel生成json
     *
     * @param file
     * @param request
     * @return
     * @throws IOException
     */
    public static String excelTojson(MultipartFile file, HttpServletRequest request) {
        //String path = "/tmp/excel";
        //String jsonDestPath = "/tmp/txt";
        String path = request.getSession().getServletContext().getRealPath("excel/xls");
        String jsonDestPath = request.getSession().getServletContext().getRealPath("excel/txt");
        // 1. 获取文件完整名称
        String fileStr = file.getOriginalFilename();
        // 2. 使用原文件名+随机生成的字符串+源文件扩展名组成新的文件名称,防止文件重名
        Date now = new Date();
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMddHH24mmss");// 可以方便地修改日期格式
        String nowTime = dateFormat.format(now);
        String commonname = fileStr.substring(0, fileStr.lastIndexOf(".")) + "-" + nowTime;
        String newfileName = commonname + fileStr.substring(fileStr.lastIndexOf("."));
        String newjsonfileName = commonname + ".txt";
        // 3. 将文件保存到服务器
        String destPath = path + File.separator + newfileName;
        File destFile = new File(destPath);
        if (!destFile.getParentFile().exists()) {
            System.out.println("目标文件所在目录不存在，准备创建");
            if (!destFile.getParentFile().mkdirs()) {
                System.out.println("目录创建失败");
            }
        }
        try {
            file.transferTo(destFile);
        } catch (IOException e) {
            e.printStackTrace();
        }
        // 4. 解析文件
        PoiToJson poiToJson = new PoiToJson();
        poiToJson.readExcel(destPath, jsonDestPath, newfileName);
        // 5. 解析完删除文件
        destFile.delete();
        return newjsonfileName;
    }

    /**
     * 读取文件返回json数据
     *
     * @param request
     * @param response
     * @param url
     */
    public static void getContent(HttpServletRequest request, HttpServletResponse response, String url) {
        //String jsonDestPath = request.getSession().getServletContext().getRealPath("excel/txt");
        String jsonDestPath = "/tmp/txt";
        String path = jsonDestPath + File.separator + url;
        File file = new File(path);
        InputStreamReader read;
        BufferedReader br = null;
        PrintWriter out = null;
        StringBuilder result = new StringBuilder();
        try {
            read = new InputStreamReader(new FileInputStream(file));// 考虑到编码格式
            br = new BufferedReader(read);// 构造一个BufferedReader类来读取文件
            String s;
            while ((s = br.readLine()) != null) {// 使用readLine方法，一次读一行
                result.append(System.lineSeparator() + s);
            }

            request.setCharacterEncoding("utf-8");
            response.setContentType("text/html;charset=utf-8");
            response.setHeader("Cache-Control", "no-cache");
            response.setHeader("Access-Control-Allow-Origin", "*");

            out = response.getWriter();
            out.print(result.toString());

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (br != null) {
                    br.close();
                }
                if (out != null) {
                    out.flush();
                    out.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

}