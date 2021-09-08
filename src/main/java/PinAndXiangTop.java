import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;


public class PinAndXiangTop {

    static String filePath = "F:\\830\\3.xlsx";
//    static String filePath = "C:\\Users\\10469\\Desktop\\628/2.xlsx";
    //第一个数据是38
    int kaiShiHangShu = 54;
    static int ILeft = 1, IInterval = 13;
    static int QLeft = 105, QInterval = 13;
    //    下面是输出文件
    static int IOutLeft = 209, IOutInterval = 21;
    static int QOutLeft = 377, QOutInterval = 21;

    //频偏输出文件
    static int IOutFreLeft = 545, IOutFreInterval = 20;
    static int QOutFreLeft = 705, QOutFreInterval = 20;
    public static void main(String[] args) throws Exception {
        PinAndXiangTop pinpianTop = new PinAndXiangTop();
        List<String> listI = new ArrayList<>();
        List<String> listQ = new ArrayList<>();
        pinpianTop.readFileToList(listI, ILeft, IInterval);
        pinpianTop.readFileToList(listQ, QLeft, QInterval);
        System.out.println("已输入到list");
        pinpianTop.ListToTxt(listI,"XI");
        pinpianTop.ListToTxt(listQ,"XQ");
        System.out.println("已输出到txt文件中");
        pinpianTop.ListToMergeTxt(listI,"XI");
        pinpianTop.ListToMergeTxt(listQ,"XQ");
        System.out.println("已输出到合并的txt文件中");
        //输出文件
        List<String> listIOutPhase = new ArrayList<>();
        List<String> listQOutPhase = new ArrayList<>();
        pinpianTop.readFileToList(listIOutPhase, IOutLeft, IOutInterval);
        pinpianTop.readFileToList(listQOutPhase, QOutLeft, QOutInterval);
        pinpianTop.ListToMergeTxt(listIOutPhase,"XI_out_phase");
        pinpianTop.ListToMergeTxt(listQOutPhase,"XQ_out_phase");

        List<String> listIOutFre = new ArrayList<>();
        List<String> listQOutFre = new ArrayList<>();
        pinpianTop.readFileToList(listIOutFre, IOutFreLeft, IOutFreInterval);
        pinpianTop.readFileToList(listQOutFre, QOutFreLeft, QOutFreInterval);
        pinpianTop.ListToMergeTxt(listIOutFre,"XI_out_fre");
        pinpianTop.ListToMergeTxt(listQOutFre,"XQ_out_fre");
//        pinpianTop.ListToTxt(listIOutFre,"XI_out_fre");
//        pinpianTop.ListToTxt(listQOutFre,"XQ_out_fre");


        System.out.println("已输出输出文件 : " + filePath);

    }
    public void readFileToList(List<String> list, int left, int Interval) throws Exception {
        File file = new File(filePath);
        InputStream is = new FileInputStream(file.getAbsoluteFile());
        XSSFWorkbook sheets = new XSSFWorkbook(is);
        XSSFSheet sheet = sheets.getSheetAt(0);
        //获取最后一行的num，即总行数。此处从0开始计数
        int maxRow = sheet.getLastRowNum();
        System.out.println("总行数为：" + maxRow);
        for (int row = kaiShiHangShu; row <= maxRow; row++) {
            //获取最后单元格num，即总单元格数 ***注意：此处从1开始计数***
            int maxRol = sheet.getRow(row).getLastCellNum();
            int index = 0;
            for (int rol = left; rol < maxRol; rol = rol + Interval){
                XSSFCell cell = sheet.getRow(row).getCell(rol);
                list.add(cell.getRawValue());
                index ++;
                if(index == 8){
                    break;
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
