package com.example.demo.controller;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.*;
import jxl.write.biff.RowsExceededException;
import org.springframework.stereotype.Controller;

import java.io.File;
import java.io.IOException;

/**
 * @Author wmy
 * @Date 2020/5/18 17:21
 * @Version 1.0
 */
@Controller
public class ExcelController {

    public static void main(String[] args) throws IOException, WriteException {
        makeExcelFile();
    }
    /**注意：实际项目要写到service业务实现层，测试写道test里单元测试进行测试*/


    /**创建excel文件（设置复杂格式）*/
    private static void makeExcelFile() throws IOException, WriteException {
        //ByteArrayOutputStream os = new ByteArrayOutputStream();
        //WritableWorkbook workbook = Workbook.createWorkbook(os);
        //创建工作薄
        WritableWorkbook workbook = Workbook.createWorkbook(new File("D:\\ExcelDemo.xls"));
        //创建sheet,设置第二三四..个sheet，依次类推即可
        WritableSheet sheet = workbook.createSheet("First Sheet", 0);
        //设置行高
        sheet.setRowView(0, 1000, false);
        sheet.setRowView(1, 600, false);
        sheet.setRowView(2, 700, false);
        sheet.setRowView(3, 700, false);
        sheet.setRowView(4, 700, false);
        //设置列宽
        for (int i = 1; i < 5; i ++){
            sheet.setColumnView(i, 25);
        }
        for (int i = 7; i < 11; i ++){
            sheet.setColumnView(i, 25);
        }

        //构造表头
        //添加合并单元格，第一个参数是起始列，第二个参数是起始行，第三个参数是终止列，第四个参数是终止行
        sheet.mergeCells(1, 1, 4, 1);
        WritableCellFormat titleFormate = getWritableCellFormat();
        Label title = new Label(1,1,"极融-还款信息汇总",titleFormate);
        //设置背景颜色
        titleFormate.setBackground(Colour.PALE_BLUE);
        sheet.addCell(title);

        //添加合并单元格，第一个参数是起始列，第二个参数是起始行，第三个参数是终止列，第四个参数是终止行
        sheet.mergeCells(7, 1, 10, 1);
        //设置字体种类和黑体显示,字体为Arial,字号大小为10,采用黑体显示
        WritableCellFormat titleFormate1 = getWritableCellFormat();
        Label title1 = new Label(7,1,"米么-还款信息汇总",titleFormate1);
        //设置背景颜色
        titleFormate1.setBackground(Colour.IVORY);
        //设置第一行的高度
        sheet.addCell(title1);

        WritableCellFormat titleFormate2 = getWritableCellFormat();
        //创建要显示的具体内容
        Label total1 = new Label(1,2,"当日应还总额/元", titleFormate2);
        sheet.addCell(total1);
        Label total2 = new Label(1,3,"当日应还本息合计/元", titleFormate2);
        sheet.addCell(total2);
        Label total3 = new Label(1,4,"当日应还服务费合计/元", titleFormate2);
        sheet.addCell(total3);
        Label total4 = new Label(3,2,"当日已还款金额合计/元", titleFormate2);
        sheet.addCell(total4);
        Label total5 = new Label(3,3,"当日已还本息合计/元", titleFormate2);
        sheet.addCell(total5);
        Label total6 = new Label(3,4,"当日已还服务费合计/元", titleFormate2);
        sheet.addCell(total6);

        Label total7 = new Label(7,2,"当日应还总额/元", titleFormate2);
        sheet.addCell(total7);
        Label total8 = new Label(7,3,"当日应还本息合计/元", titleFormate2);
        sheet.addCell(total8);
        Label total9 = new Label(7,4,"当日应还服务费合计/元", titleFormate2);
        sheet.addCell(total9);
        Label total10 = new Label(9,2,"当日已还款金额合计/元", titleFormate2);
        sheet.addCell(total10);
        Label total = new Label(9,3,"当日已还本息合计/元", titleFormate2);
        sheet.addCell(total);
        Label total11 = new Label(9,4,"当日已还服务费合计/元", titleFormate2);
        sheet.addCell(total11);

        Label total12 = new Label(2,2,"10", titleFormate2);
        sheet.addCell(total12);
        Label total13 = new Label(2,3,"10", titleFormate2);
        sheet.addCell(total13);
        Label total14 = new Label(2,4,"10", titleFormate2);
        sheet.addCell(total14);
        Label total15 = new Label(4,2,"20", titleFormate2);
        sheet.addCell(total15);
        Label total16 = new Label(4,3,"20", titleFormate2);
        sheet.addCell(total16);
        Label total17 = new Label(4,4,"20", titleFormate2);
        sheet.addCell(total17);

        Label total18 = new Label(8,2,"30", titleFormate2);
        sheet.addCell(total18);
        Label total19 = new Label(8,3,"30", titleFormate2);
        sheet.addCell(total19);
        Label total20 = new Label(8,4,"30", titleFormate2);
        sheet.addCell(total20);
        Label total21 = new Label(10,2,"40", titleFormate2);
        sheet.addCell(total21);
        Label total22 = new Label(10,3,"40", titleFormate2);
        sheet.addCell(total22);
        Label total23 = new Label(10,4,"40", titleFormate2);
        sheet.addCell(total23);

        //把创建的内容写入到输出流中，并关闭输出流
        workbook.write();
        workbook.close();
       /* os.close();
       //添加bom头，此语句解决office打开excel文件乱码问题
        byte[] uft8bom={(byte)0xef,(byte)0xbb,(byte)0xbf};
        byte[] data = os.toByteArray();
        byte[] bytes = Bytes.concat(uft8bom,data);*/
        //return new ByteArrayInputStream(bytes);
    }

    private static WritableCellFormat getWritableCellFormat() throws WriteException {
        //设置字体种类和黑体显示,字体为Arial,字号大小为10,采用黑体显示(WritableFont.BOLD加粗，不想加粗去掉即可)
        WritableFont bold = new WritableFont(WritableFont.createFont("微软雅黑"), 10, WritableFont.BOLD);
        //生成一个单元格样式控制对象
        WritableCellFormat titleFormate = new WritableCellFormat(bold);
        //单元格中的内容水平方向居中
        titleFormate.setAlignment(jxl.format.Alignment.CENTRE);
        //设置边框(MEDIUM实线 DASH_DOT虚线 Border.TOP顶部线 left。。。)
        titleFormate.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
        //单元格的内容垂直方向居中
        titleFormate.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
        return titleFormate;
    }


    /**创建简单的excel*/
    public static class CreateExcel {

        public static void main(String[] args)
                throws IOException, RowsExceededException, WriteException {
            //1:创建excel文件
            File file=new File("test.xls");
            file.createNewFile();

            //2:创建工作簿
            WritableWorkbook workbook=Workbook.createWorkbook(file);
            //3:创建sheet,设置第二三四..个sheet，依次类推即可
            WritableSheet sheet=workbook.createSheet("用户管理", 0);
            //4：设置titles
            String[] titles={"编号","账号","密码"};
            //5:单元格
            Label label=null;
            //6:给第一行设置列名
            for(int i=0;i<titles.length;i++){
                //x,y,第一行的列名
                label=new Label(i,0,titles[i]);
                //7：添加单元格
                sheet.addCell(label);
            }
            //8：模拟数据库导入数据
            for(int i=1;i<10;i++){
                //添加编号，第二行第一列
                label=new Label(0,i,i+"");
                sheet.addCell(label);

                //添加账号
                label=new Label(1,i,"10010"+i);
                sheet.addCell(label);

                //添加密码
                label=new Label(2,i,"123456");
                sheet.addCell(label);
            }

            //写入数据，一定记得写入数据，不然你都开始怀疑世界了，excel里面啥都没有
            workbook.write();
            //最后一步，关闭工作簿
            workbook.close();
        }
    }

    /**读取excel*/
    public static class ReadExcel {

        public static void main(String[] args) throws Exception{
            //1:创建workbook
            Workbook workbook=Workbook.getWorkbook(new File("test.xls"));
            //2:获取第一个工作表sheet
            Sheet sheet=workbook.getSheet(0);
            //3:获取数据
            System.out.println("行："+sheet.getRows());
            System.out.println("列："+sheet.getColumns());
            for(int i=0;i<sheet.getRows();i++){
                for(int j=0;j<sheet.getColumns();j++){
                    Cell cell=sheet.getCell(j,i);
                    System.out.print(cell.getContents()+" ");
                }
                System.out.println();
            }

            //最后一步：关闭资源
            workbook.close();
        }


    }
}
