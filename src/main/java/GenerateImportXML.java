import java.io.*;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class GenerateImportXML{
    static String keyword;
    static String title;
    static int step;
    static String detail;
    static StringBuffer action = new StringBuffer();
    static StringBuffer expect = new StringBuffer();
    static XSSFWorkbook targetWorkbook;
    static XSSFSheet targetSheet;
    static String[] headers = new String[]{"Case 标题", "状态", "创建者", "关键词", "脚本状态", "脚本地址", "步骤", "指派给", "优先级", "可用版本", "类型"};
    static String status = "Active";
    static String creator = "王婵";
    static String scriptstatus = "已完成";
    static String scriptaddress;
    static String assign = "王婵";
    static String priority = "高";
    static String version = "3.18";
    static String type = "功能";
    static InputStream inputStream;
    static OutputStream outputStream;
    static int rowCount = 1;
    public static void parseExcel(File file){

        try {
            inputStream = new FileInputStream(file);
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            for(int i = 1; i < workbook.getNumberOfSheets(); i++){
                boolean blankRow = false;
                keyword = workbook.getSheetName(i);
                if (keyword.equals("注册")){
                    scriptaddress = "Register.java";
                }else if (keyword.equals("登录")){
                    scriptaddress = "Login.java";
                }else if (keyword.equals("发帖分享")){
                    scriptaddress = "ShareAfterPost.java";
                }else if (keyword.equals("位置信息")){
                    scriptaddress = "GeoInPost.java";
                }else if (keyword.equals("发帖引导")){
                    scriptaddress = "TipForPost.java";
                }else if (keyword.equals("发帖")){
                    scriptaddress = "Post.java";
                }else if (keyword.equals("发帖插件")){
                    scriptaddress = "PostPlugin.java";
                }else if (keyword.equals("爱情公社")){
                    scriptaddress = "PostForLove.java";
                }else if (keyword.equals("点赞")){
                    scriptaddress = "LikePost.java";
                }else if (keyword.equals("发评论")){
                    scriptaddress = "Comment.java";
                }else if (keyword.equals("消息里的评论")){
                    scriptaddress = "CommentInMessage.java";
                }else if (keyword.equals("评论列表")){
                    scriptaddress = "CommentList.java";
                }else if (keyword.equals("评论操作")){
                    scriptaddress = "CommentAction.java";
                }else if (keyword.equals("提到我的")){
                    scriptaddress = "CommentMention.java";
                }else if (keyword.equals("送礼物")){
                    scriptaddress = "Gift.java";
                }else if (keyword.equals("关闭评论")){
                    scriptaddress = "CommentClose.java";
                }
                XSSFSheet sheet = workbook.getSheetAt(i);
                for (Row row : sheet){
                    if (row.getRowNum() != 0){
                        if (!isMergedRegion(sheet, row.getRowNum(), -1) && row.getFirstCellNum() == -1 || (row.getFirstCellNum() == 0 && !isMergedRegion(sheet, row.getRowNum(), 0) && (row.getCell(0).getStringCellValue() == null || row.getCell(0).getStringCellValue().equals(""))) || (row.getFirstCellNum() == 1 && (row.getCell(1).getStringCellValue() == null || row.getCell(1).getStringCellValue().equals("")))){
                            if (blankRow == true){
                                break;
                            }
                            detail = "【步骤】" + "\r\n" + action + "【期望】" + "\r\n" + expect;
                            fillTarget();
                            blankRow = true;
                        }else if (row.getFirstCellNum() != -1 && isMergedRegion(sheet, row.getRowNum(), 0) && row.getCell(0).getStringCellValue() != null && !row.getCell(0).getStringCellValue().equals("")){
                            title = row.getCell(row.getFirstCellNum()).getStringCellValue();
                            step = 1;
                            action = new StringBuffer();
                            expect = new StringBuffer();
                            blankRow = false;
                        }else if (!isMergedRegion(sheet, row.getRowNum(), 0)){
                            if (row.getCell(0) != null){
                                action.append(String.valueOf(step) + ". " + row.getCell(0).getStringCellValue() + "\r\n");
                            }else {
                                action.append(String.valueOf(step) + ". " + "\r\n");
                            }
                            if (row.getCell(1) != null){
                                expect.append(String.valueOf(step) + ". " + row.getCell(1).getStringCellValue() + "\r\n");
                            }else {
                                expect.append(String.valueOf(step) + ". " + "\r\n");
                            }
                            step++;
                            blankRow = false;
                        }
                    }
                }
            }
        }catch (IOException e){
            e.printStackTrace();
        }finally {
            if (inputStream != null){
                try {
                    inputStream.close();
                }catch (IOException e){
                    e.printStackTrace();
                }
            }
        }
    }

    public static void initTarget(){
        targetWorkbook = new XSSFWorkbook();
        targetSheet = targetWorkbook.createSheet();
        XSSFRow firstRow = targetSheet.createRow(0);
        for (int i = 0; i < headers.length; i++){
            Cell cell = firstRow.createCell(i);
            cell.setCellValue(headers[i]);
        }
    }

    public static void fillTarget(){
        XSSFRow row = targetSheet.createRow(rowCount++);
        Cell cell0 = row.createCell(0);
        cell0.setCellValue(title);
        Cell cell1 = row.createCell(1);
        cell1.setCellValue(status);
        Cell cell2 = row.createCell(2);
        cell2.setCellValue(creator);
        Cell cell3 = row.createCell(3);
        cell3.setCellValue(keyword);
        Cell cell4 = row.createCell(4);
        cell4.setCellValue(scriptstatus);
        Cell cell5 = row.createCell(5);
        cell5.setCellValue(scriptaddress);
        Cell cell6 = row.createCell(6);
        cell6.setCellValue(detail);
        Cell cell7 = row.createCell(7);
        cell7.setCellValue(assign);
        Cell cell8 = row.createCell(8);
        cell8.setCellValue(priority);
        Cell cell9 = row.createCell(9);
        cell9.setCellValue(version);
        Cell cell10 = row.createCell(10);
        cell10.setCellValue(type);
    }

    public static void writeExcel(File file){
        try {
            outputStream = new FileOutputStream(file);
            targetWorkbook.write(outputStream);
        }catch (IOException e){
            e.printStackTrace();
        }finally {
            if (outputStream != null){
                try{
                    outputStream.close();
                }catch (IOException e){
                    e.printStackTrace();
                }
            }
        }
    }

    private static boolean isMergedRegion(XSSFSheet sheet,int row ,int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if(row >= firstRow && row <= lastRow && column >= firstColumn && column <= lastColumn){
                return true;
            }
        }
        return false;
    }

    public static void main(String[] args){
        initTarget();
        File sourceFile = new File("src\\main\\resources\\自动化测试用例.xlsx");
        File targetFile = new File("src\\main\\resources\\result.xlsx");
        if (targetFile.exists()){
            targetFile.delete();
        }
        parseExcel(sourceFile);
        writeExcel(targetFile);
    }
}
