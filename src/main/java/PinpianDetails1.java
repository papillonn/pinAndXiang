import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

//速度很慢，不过能用

public class PinpianDetails1 {

    String filePath = "E:/20210528_1_details/1.xlsx";
    int kaiShiHangShu = 123;
    static int ILeft = 1, IInterval = 13;
    static int QLeft = 105, QInterval = 13;
    public static void main(String[] args) throws Exception {
        PinpianDetails1 pinpian1 = new PinpianDetails1();
        List<String> listI = new ArrayList<>();
        List<String> listQ = new ArrayList<>();
        pinpian1.readFileToList(listI, ILeft, IInterval);
        pinpian1.readFileToList(listQ, QLeft, QInterval);
        pinpian1.ListToTxt(listI, "XI");
        pinpian1.ListToTxt(listQ, "XQ");
        pinpian1.ListToMergeTxt(listI, "XI");
        pinpian1.ListToMergeTxt(listQ, "XQ");
    }

    public void readFileToList(List<String> list, int left, int Interval) throws Exception {
        File file = new File(filePath);
        InputStream is = new FileInputStream(file.getAbsoluteFile());
        Workbook wk = StreamingReader.builder()
                .rowCacheSize(100)  //缓存到内存中的行数，默认是10
                .bufferSize(8192)  //读取资源时，缓存到内存的字节大小，默认是1024
                .open(is);  //打开资源，必须，可以是InputStream或者是File，注意：只能打开XLSX格式的文件
        Sheet sheet = wk.getSheetAt(0);
//        XSSFWorkbook sheets = new XSSFWorkbook(is);
//        XSSFSheet sheet = sheets.getSheetAt(0);
        //获取最后一行的num，即总行数。此处从0开始计数
        int maxRow = sheet.getLastRowNum();
        System.out.println("总行数为：" + maxRow);
        for (Row row : sheet) {
            if (row.getRowNum() >= kaiShiHangShu){
                int index = 0;
                for (Cell cell : row) {
                    if (cell.getColumnIndex() % Interval == left){
                        list.add(cell.getStringCellValue());
                        index ++;
                        if(index == 8){
                            break;
                        }
                    }
                }
            }
        }
    }

    public void ListToTxt(List<String> list, String filename) throws IOException {
        File file = new File(filePath);
        for (int i = 0;i < 8; i ++){
            int index = i;
            if(!filename.equals("fre_out")){
                index = i + 1;
            }
            File outputFile = new File(file.getParent() + "/" + filename + index  + ".txt");
            BufferedWriter writer = null;
            writer = new BufferedWriter(new FileWriter(outputFile));
            for (int j = i ;j < list.size() ; j = j + 8){
                String s = list.get(j);
                writer.write(s);
                writer.newLine();

            }
            writer.close();
        }

        System.out.println("输出完成！");
    }


    public void ListToMergeTxt(List<String> list, String filename) throws IOException {
        //如果放到matlab中，输入数据可能要删除几个值
        File file = new File(filePath);
        File outputFile = new File(file.getParent() + "/" + filename + ".txt");
        BufferedWriter writer = null;
        writer = new BufferedWriter(new FileWriter(outputFile));
        for (String s : list) {
            writer.write(s);
            writer.newLine();
        }
        writer.close();

        System.out.println("输出完成！");
    }

}
