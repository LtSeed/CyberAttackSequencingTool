import org.apache.commons.cli.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.*;
import java.util.stream.Stream;

public class Main {
    static String sorting_category_name;
    static Category sorting_category;
    static boolean fold;
    static boolean omit;
    public static void main(String[] args) {

        Date timer = new Date();

        Options options = new Options();
        options.addOption("p","path",true,"-p <path> Path to the excel file.");
        options.addOption("s","sheet",true,"-s <sheet name> Name of the excel sheet.");
        options.addOption("c","classification",true,"-c <class name> Set sorting categories, which default in 源IP.");
        options.addOption("f","fold",true,"-f <y/n> Collapse items with the same sorting category.");
        options.addOption("o","Omit",true,"-o <y/n> Omit the same information after folding.");

        CommandLine cli;
        CommandLineParser cliParser = new DefaultParser();
        HelpFormatter helpFormatter = new HelpFormatter();

        try {
            cli = cliParser.parse(options, args);
        } catch (ParseException e) {
            helpFormatter.printHelp(">>>>>> CAST-CLI options", options);
            e.printStackTrace();
            return;
        }


        sorting_category_name = cli.getOptionValue("c","源IP");
        String fold_str = cli.getOptionValue("f","y");
        fold = Stream.of("Y", "true", "t").anyMatch(fold_str::equalsIgnoreCase);
        String omit_str = cli.getOptionValue("o","y");
        omit = Stream.of("Y", "true", "t").anyMatch(omit_str::equalsIgnoreCase);
        String path = cli.getOptionValue("p");
        String sheet_name = cli.getOptionValue("s",null);

        if(!fold && omit){
            System.out.println("Error: Cannot omit when not fold.");
            fold = true;
        }

        XSSFWorkbook workbook;
        try {
            workbook = new XSSFWorkbook(new File(path));
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Error: Can not read the excel file at the path.");
            return;
        } catch (InvalidFormatException e) {
            e.printStackTrace();
            System.out.println("Error: The file are in a invalid format. Please check the file format of excel.");
            return;
        }
        XSSFSheet sheet;
        if(sheet_name != null) sheet = workbook.getSheet(sheet_name);
        else sheet = workbook.getSheetAt(0);

        XSSFRow categories = sheet.getRow(0);
        for (Cell cell : categories) {
            new Category(cell.getStringCellValue());
        }

        sorting_category = Category.valueOf(sorting_category_name);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Attack.addAttack(sheet.getRow(i));
        }

        XSSFSheet result = workbook.createSheet("result");
        XSSFRow row_0 = result.createRow(0);
        int i = 0;
        for (String s : Category.name_map.keySet()) {
            row_0.createCell(i++).setCellValue(s);
        }
        if(fold) {
            i = 1;
            for (String s : Attack.attacks.keySet()) {
                Attack attack = Attack.attacks.get(s);
                XSSFRow row = result.createRow(i++);
                for (Category category : attack.getAllInfo().keySet()) {
                    row.getCell(category.getNumber()).setCellValue(printList(attack.getAllInfo().get(category)));
                }
            }
        } else {
            XSSFCellStyle split = workbook.createCellStyle();
            split.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            split.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIGHT_BLUE.getIndex());
            boolean f = false;
            i = 1;
            for (String s : Attack.attacks.keySet()) {
                Attack attack = Attack.attacks.get(s);

                String[][] strings = attack.getAllInfoInStringArray();
                for (String[] stringArray : strings) {
                    XSSFRow row = result.createRow(i++);
                    for (int i1 = 0; i1 < stringArray.length; i1++) {
                        row.createCell(i1).setCellValue(stringArray[i1]);
                    }
                    if(f) row.setRowStyle(split);
                }
                f = !f;
            }
        }
        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Error occurred when saving the excel file.");
        }
        System.out.println("Done.("+(new Date().getTime() - timer.getTime())+"ms)");
    }

    private static String printList(List<String> strings) {
        StringBuilder sb = new StringBuilder();
        for (String string : strings) {
            sb.append(string).append('|');
        }
        return sb.toString();
    }

    private static class Attack {
        private static final TreeMap<String, Attack> attacks = new TreeMap<>();
        private final AttackInformation info;

        public static void addAttack(XSSFRow row){
            XSSFCell id_cell = row.getCell(sorting_category.getNumber());
            String id_cell_value = id_cell.getStringCellValue();
            Attack attack = attacks.getOrDefault(id_cell_value, new Attack());
            attack.info.addInfo(row);
            attacks.put(id_cell_value,attack);
        }
        public HashMap<Category, List<String>> getAllInfo(){
            return info.getAllInfo();
        }
        public String[][] getAllInfoInStringArray(){
            HashMap<Category, List<String>> map = info.getAllInfo();
            String[][] temp =
                    new String[map.size()][map.get(Category.getByNumber(0)).size()];
            for (Map.Entry<Category, List<String>> infos : map.entrySet()) {
                for (int i = 0; i < infos.getValue().size(); i++) {
                    temp[infos.getKey().number][i] = infos.getValue().get(i);
                }
            }
            return temp;
        }
        Attack(){
            info = new AttackInformation();
        }
    }

    private static class AttackInformation {
        private final HashMap<Category, List<String>> info = new HashMap<>(25);
        private void addInfo(XSSFRow row){
            int i = 0;
            for (Cell cell : row) {
                String value = cell.getStringCellValue();
                List<String> temp = info.getOrDefault(Category.getByNumber(i),new ArrayList<>());
                if(omit){
                    if(!temp.contains(value)) temp.add(value);
                } else temp.add(value);
                info.put(Category.getByNumber(i++), temp);
            }
        }
        public HashMap<Category, List<String>> getAllInfo(){
            return info;
        }

        AttackInformation (){
        }
    }

    private static class Category {
        private static int cnt = 0;
        private static final Map<String, Category> name_map = new HashMap<>();
        private static final Map<Integer, Category> number_map = new HashMap<>();
        private final int number;

        public static Category valueOf(String name){
            return name_map.getOrDefault(name,null);
        }

        public static Category getByNumber(int number){
            return number_map.getOrDefault(number,null);
        }

        Category (String name){
            name_map.put(name,this);
            number_map.put(cnt,this);
            this.number = cnt++;
        }

        public int getNumber() {
            return number;
        }
    }
}
