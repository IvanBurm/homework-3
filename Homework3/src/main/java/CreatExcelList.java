import java.io.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.text.*;
import org.apache.poi.openxml4j.exceptions.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreatExcelList {

    private static String[] columns = {"Имя", "Фамилия", "Отчество", "Возраст", "Пол", "Дата рождения", "Инн", "Почтовый индекс" , "Страна", "Область", "Город", "Улица", "Дом", "Квартира"};


    public static void main(String[] args) throws IOException,
            InvalidFormatException {
        int numberofrec,gender,rowNum = 1;;
        long age;
         Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("ExcelList");

        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        Row headerRow = sheet.createRow(0);

        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        for(numberofrec=1+(int)(Math.random()*30); numberofrec !=0; numberofrec --)
         {
             Row row = sheet.createRow(rowNum++);

             gender = 1 + (int)(Math.random()*2);
             if (gender == 1)
             {row.createCell(4).setCellValue("М");}
             else{ row.createCell(4).setCellValue("Ж");}

             Date dateofbirth = new Date(70 + (int)(Math.random()*48),1+(int)(Math.random()*12),1+(int)(Math.random()*31));
             Date datenow = new Date();
             age = (datenow.getTime()-dateofbirth.getTime())/(1000*60*60*24);
             age = age/365;
             SimpleDateFormat format = new SimpleDateFormat("dd.MM.yyyy");

             row.createCell(0).setCellValue(GetData("Имя.txt",gender));
             row.createCell(1).setCellValue(GetData("Фамилия.txt",gender));
             row.createCell(2).setCellValue(GetData("Отчество.txt",gender));
             row.createCell(5).setCellValue(format.format(dateofbirth).toString());
             row.createCell(3).setCellValue((int)age);
             row.createCell(6).setCellValue(GenerateINN());
             row.createCell(7).setCellValue(100000+(int)(Math.random()*900000));
             row.createCell(8).setCellValue(GetData("Страна.txt"));
             row.createCell(9).setCellValue(GetData("Область.txt"));
             row.createCell(10).setCellValue(GetData("Город.txt"));
             row.createCell(11).setCellValue(GetData("Улица.txt"));
             row.createCell(12).setCellValue(1+(int)(Math.random()*60));
             row.createCell(13).setCellValue(1+(int)(Math.random()*100));
         }

        for (int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        FileOutputStream fileOut = new FileOutputStream("src/main/resources/ExcelList.xlsx");
        workbook.write(fileOut);
        fileOut.close();
        System.out.println("Файл создан. Путь: src/main/resources/ExcelList.xlsx");

    }

    static String GenerateINN()
     {   int i;
        String innnumber="";
        int[] INN = new int[12];
        int[] coefficient = {3,7,2,4,10,3,5,9,4,6,8};
        INN[0]=7;
        INN[1]=7;
        INN[2]=(int)(Math.random()*6);
        if (INN[2] == 5)
          {INN[3]=(int)(Math.random()*2);}
        else
          {INN[3]=(int)(Math.random()*10);}
        for (i=4;i<10;i++)
        {
            INN[i]=(int)(Math.random()*10);
        }
        INN[10]=0;
        for (i=0;i<10;i++)
        {
            INN[10]=INN[10]+INN[i]*coefficient[i+1];
        }
        INN[10]=((INN[10]%11)%10);
         INN[11]=0;
         for (i=0;i<11;i++)
         {
             INN[11]=INN[11]+INN[i]*coefficient[i];
         }
         INN[11]=((INN[11]%11)%10);
         for (i=0;i<12;i++)
         {innnumber=innnumber+Integer.toString(INN[i]);}
        return innnumber;
     }
    static String GetData(String s, int gender)
    {
        try {
            FileInputStream fstream = new FileInputStream("src/main/resources/"+s);
            BufferedReader br = new BufferedReader(new InputStreamReader(fstream, "Cp1251"));
            String stringofdata, familiyending = "", otchestvoending = "";
            int count;
            if ((s == "Имя.txt") & (gender == 1)) {
                count = 1 + (int) (Math.random() * 24);
            }else
                {count = 24 + (int) (Math.random() * 24);}

            if ((s == "Фамилия.txt") & (gender == 2) & (count <= 36))
            {familiyending = familiyending + "а";}

            if ((gender == 1) & (s == "Отчество.txt"))
            { otchestvoending = otchestvoending+"ич";}
            if ((gender == 2) & (s == "Отчество.txt"))
            {otchestvoending = otchestvoending+"на";}

            while ((count != 0) & ((stringofdata = br.readLine()) != null)) {

                count --;

            }
              return stringofdata+familiyending+otchestvoending;

        } catch (IOException e)
        {
            System.out.println("Ошибка");
        }
        return "";
    }
    static String GetData(String s)
    {
        try {
            FileInputStream fstream = new FileInputStream("src/main/resources/"+s);
            BufferedReader br = new BufferedReader(new InputStreamReader(fstream, "Cp1251"));
            String stringofdata;
            int count;
                count = 1 + (int)(Math.random()*47);
            while ((count != 0) & ((stringofdata = br.readLine()) != null)) {
                count --;
            }
            return stringofdata;

        } catch (IOException e)
        {
            System.out.println("Ошибка");
        }
        return "";
    }
}
