import java.io.*;
        import org.apache.poi.ss.usermodel.*;

import java.net.URL;
import java.text.*;

public class Main {
    static DecimalFormat df = new DecimalFormat("#####0");

    public static void main(String[] args) {
        FileWriter fostream;
        PrintWriter out = null;
        String strOutputPath = "/Users/mykhailobozhedomov/IdeaProjects/extract-xlsx/exportxlsx";
        String strFilePrefix = "Watsons-Actualy";
        URL url = Main.class.getResource("Watsons-Actualy.xlsx");
        System.out.println(strOutputPath);

        try {
            InputStream inputStream = new FileInputStream(new File(url.getPath()));
            Workbook wb = WorkbookFactory.create(inputStream);
            Sheet sheet = wb.getSheet("Sheet1");
            System.out.println(sheet);

            fostream = new FileWriter(strOutputPath + "/" + strFilePrefix+ ".xml");
            System.out.println(strOutputPath);
            out = new PrintWriter(new BufferedWriter(fostream));

            out.println("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            out.println("<companies>");

            boolean firstRow = true;
            for (Row row : sheet) {
                System.out.println(row.getCell(0));
                if (firstRow == true) {
                    firstRow = false;
                    continue;
                }
                out.println("\t<company>");
                out.println(formatElement("\t\t", "company-id", "", formatCell(row.getCell(0))));
                out.println(formatElement("\t\t", "name", "lang=\"ru\"", formatCell(row.getCell(1))));
                out.println(formatElement("\t\t", "name-other", "lang=\"ru\"", formatCell(row.getCell(2))));
                out.println(formatElement("\t\t", "address", "lang=\"ru\"", formatCell(row.getCell(3))));
                out.println(formatElement("\t\t", "country", "lang=\"ru\"", formatCell(row.getCell(4))));
                out.println("\t\t<coordinates>");
                out.println(formatElement("\t\t\t", "lon", "", formatCell(row.getCell(6))));
                out.println(formatElement("\t\t\t", "lat", "", formatCell(row.getCell(7))));
                out.println("\t\t</coordinates>");
                out.println("\t\t<phone>");
                out.println(formatElement("\t\t\t", "number", "", formatCell(row.getCell(8))));
                out.println("\t\t</phone>");
                out.println("\t\t<email>ayupova@watsons.ru</email>");
                out.println("\t\t<url>https://watsons.com.ru/</url>");
                out.println("\t\t<add-url>https://watsons.com.ru/</add-url>");
                out.println(formatElement("\t\t\t", "working-time", "lang=\"ru\"", formatCell(row.getCell(13))));
                out.println("\t\t<rubric-id>184105798</rubric-id>");
                out.println("\t\t<photos>");
                String Str = new String(row.getCell(17).getStringCellValue());

                for (String retval : Str.split(",")) {
                    System.out.println(retval);
                    out.println("\t\t\t<photo>" + retval.trim() + "</photo>");
                }
                out.println("\t\t</photos>");
                out.println("\t</company>");
            }
            out.write("</companies>");
            out.flush();
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static String formatCell(Cell cell)
    {
        if (cell == null) {
            return "";
        }
        switch(cell.getCellType()) {
            case Cell.CELL_TYPE_BLANK:
                return "";
            case Cell.CELL_TYPE_BOOLEAN:
                return Boolean.toString(cell.getBooleanCellValue());
            case Cell.CELL_TYPE_ERROR:
                return "*error*";
            case Cell.CELL_TYPE_NUMERIC:
                return Main.df.format(cell.getNumericCellValue());
            case Cell.CELL_TYPE_STRING:
                return cell.getStringCellValue();
            default:
                return "<unknown value>";
        }
    }

    private static String formatElement(String prefix, String tag, String attr, String value) {
        StringBuilder sb = new StringBuilder(prefix);
        sb.append("<");
        sb.append(tag);
        if (attr.length() > 0) {
            sb.append(" ");
            sb.append(attr);
        }
        if (value != null && value.length() > 0) {
            sb.append(">");
            sb.append(value);
            sb.append("</");
            sb.append(tag);
            sb.append(">");
        } else {
            sb.append("/>");
        }
        return sb.toString();
    }
}
