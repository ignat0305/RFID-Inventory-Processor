package RFIDreader;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

enum Categories{AMERICANFOOTBAL, OUTDOOR, SOCCER, SPORTWEAR, BASEBALL}
enum TagType{STICKER, HARDTAG}
enum Status{STOCKIN, STOCKOUT, STOCK}
enum Stage{CR0, PULLOVER, CR1, CR2, SMU, CS1, CS2, CFM, PROMO, LOOKSEE, MFA, OTH, PULLOVERMFA, PRECR0}
enum Team{L4A, L4MA, L4B, L4MB}
public class RFIDReader {
    class Record{
        TagType tagType;
        Stage stage;
        Categories categories;
        Status status;
        String code;
        String modelName;
        String article;
        Team team;
        LocalDateTime time;
        int row;
        public Record(LocalDateTime time, Team team, String article, String modelName,Categories categories, TagType tagType, Stage stage, Status status, String code, int row){
            this.article = article;
            this.modelName = modelName;
            this.categories = categories;
            this.tagType = tagType;
            this.stage = stage;
            this.status = status;
            this.team = team;
            this.code = code;
            this.row = row;
            this.time = time;
        }

        @Override
        public String toString() {
            return time.toString()+" "+team.toString()+" "+categories.toString()+" "+modelName+" "+article+" "+stage+" "+stage.toString()+" "+code+" "+tagType.toString()+" "+status.toString()+" "+row;
        }
    }
    String path;
    TreeMap<String, List<Record>> overall;
    TreeMap<String, Map<Status, Integer>> datCount;
    TreeMap<String, List<Record>> closed;
    Map<String, String> inHouse;
    List<int[]> bugs = new ArrayList<>();
    public RFIDReader(String path) throws IOException {
        Set<String> bug = new HashSet<>();
        overall = new TreeMap<>();
        datCount = new TreeMap<>();
        closed = new TreeMap<>();
        inHouse = new HashMap<>();
        this.path = path;
        FileInputStream file = new FileInputStream(path);
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);
        for(Row row: sheet){
            if(row.getRowNum()>=1 &&  row.getRowNum()!= 21654 && (row.getRowNum()< 23206 || row.getRowNum()>23225) && row.getCell(0).getStringCellValue().trim().equals("Sample Warehouse")){
//                System.out.println(row.getRowNum());
                DateTimeFormatter timeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
                LocalDateTime key = LocalDateTime.parse(row.getCell(2).getStringCellValue(), timeFormatter);
                String time = row.getCell(2).getStringCellValue().substring(0, 10);
//                System.out.println(time);
                overall.putIfAbsent(time, new ArrayList<>());
                datCount.putIfAbsent(time, new HashMap<>());
//                System.out.println(key);
                int cursor = 1;
                TagType tagType = TagType.HARDTAG;
                Stage stage = Stage.CR0;
                Categories categories = Categories.SPORTWEAR;
                Status status = Status.STOCKOUT;
                String code = "";
                String modelName = "";
                String article = "";
                Team team = Team.L4MB;

                try {
                    while (cursor<=14){
                            if(cursor!=13){
//                                System.out.print(cursor+" "+row.getCell(cursor).getStringCellValue()+" ");
                                if(cursor==1){
                                    switch (row.getCell(cursor).getStringCellValue().trim()){
                                        case "L4A"-> team = Team.L4A;
                                        case "L4MA"-> team = Team.L4MA;
                                        case "L4B"-> team = Team.L4B;
                                        case "L4MB"-> team = Team.L4MB;
                                    }
                                }
                                if(cursor ==3){
                                    code = row.getCell(cursor).getStringCellValue().trim();
                                }
                                if(cursor==4){
                                    switch (row.getCell(cursor).getStringCellValue().trim()){
                                        case "Stock"->{
                                            status = Status.STOCK;
                                            datCount.get(time).putIfAbsent(Status.STOCK, 0);
                                            datCount.get(time).compute(Status.STOCK, (a,b)->b+1);
                                        }
                                        case "Stock in"->{
                                            status = Status.STOCKIN;

                                            datCount.get(time).putIfAbsent(Status.STOCKIN, 0);
                                            datCount.get(time).compute(Status.STOCKIN, (a,b)->b+1);
                                        }
                                        case "Stock out"->{
                                            status = Status.STOCKOUT;
                                            datCount.get(time).putIfAbsent(Status.STOCKOUT, 0);
                                            datCount.get(time).compute(Status.STOCKOUT, (a,b)->b+1);

                                        }
                                    }
                                }
                                if(cursor==6){
                                    switch (row.getCell(cursor).getStringCellValue().trim()){
                                        case "Outdoor"-> categories = Categories.OUTDOOR;
                                        case "Soccer"-> categories = Categories.SOCCER;
                                        case "Sportswear" ->categories = Categories.SPORTWEAR;
                                        case "American Football"->categories = Categories.AMERICANFOOTBAL;
                                    }
                                }
                                if(cursor==8){
                                    switch (row.getCell(cursor).getStringCellValue().trim()){
                                        case "CR0" -> stage = Stage.CR0;
                                        case "CR1" -> stage = Stage.CR1;
                                        case "CR2" -> stage = Stage.CR2;
                                        case "CFM" -> stage = Stage.CFM;
                                        case "Pullover" -> stage = Stage.PULLOVER;
                                        case "Pre-CR0" -> stage = Stage.PRECR0;
                                        case "CS2" -> stage = Stage.CS2;
                                        case "CS1" -> stage = Stage.CS1;
                                        case "Looksee" -> stage = Stage.LOOKSEE;
                                        case "OTH" -> stage = Stage.OTH;
                                        case "MFA" -> stage = Stage.MFA;
                                        case "Promo" -> stage = Stage.PROMO;
                                        case "PulloverMFA" -> stage = Stage.PULLOVERMFA;
                                        case "SMU" -> stage = Stage.SMU;
                                    }
                                }
                                if(cursor==9){
                                    modelName = row.getCell(cursor).getStringCellValue().trim();
                                }
                                if(cursor==10){
                                    article = row.getCell(cursor).getStringCellValue().trim();
                                }
                                if(cursor == 14 && !row.getCell(cursor).getStringCellValue().isEmpty()){
                                    switch (row.getCell(cursor).getStringCellValue().trim()){
                                        case "Sticker"-> tagType = TagType.STICKER;
                                        case "HardTag"-> tagType = TagType.HARDTAG;
                                    }
                                }
                            }
                        cursor++;
                    }
                    Record item  = new Record(key,team, article, modelName, categories, tagType, stage, status, code, row.getRowNum());
                    if(item.status.name().equals("STOCKIN")){
                        inHouse.put(item.code, time);
                    }
                    if(item.status.name().equals("STOCKOUT")){
                        closed.putIfAbsent(time, new ArrayList<>());
                        closed.get(time).add(item);
                    }
                    overall.get(time).add(item);
//                    System.out.println("\n");
                }catch (NullPointerException e ){
                    break;
                }
            }
        }
    }

    public void printTheMap(){
        for(var x: overall.keySet()){
            System.out.println(x);
            for(var y: overall.get(x)){
                System.out.println(y);
            }
            System.out.print(x+" ");
            for(var y: datCount.get(x).keySet()){
                System.out.print(y.toString()+" "+datCount.get(x).get(y)+" ");
            }
            System.out.print("\n");
        }
    }
    public void create(String outputPath) throws IOException {

        Workbook workbook = new XSSFWorkbook();

        Font boldFont = workbook.createFont();
        boldFont.setBold(true);

        CellStyle boldStyle = workbook.createCellStyle();
        boldStyle.setFont(boldFont);

        DateTimeFormatter timeFormat =
                DateTimeFormatter.ofPattern("HH:mm:ss");

        for (String date : overall.keySet()) {

            Sheet sheet = workbook.createSheet(date);
            int rowNum = 0;

            // =========================
            // A. Status Count
            // =========================
            Row title = sheet.createRow(rowNum++);
            title.createCell(0).setCellValue(date + " Status Count");
            title.getCell(0).setCellStyle(boldStyle);

            Map<Status, Integer> count = datCount.get(date);

            Row countRow = sheet.createRow(rowNum++);
            countRow.createCell(0).setCellValue(
                    "STOCK : " + count.getOrDefault(Status.STOCK, 0));
            countRow.createCell(1).setCellValue(
                    "STOCKIN : " + count.getOrDefault(Status.STOCKIN, 0));
            countRow.createCell(2).setCellValue(
                    "STOCKOUT : " + count.getOrDefault(Status.STOCKOUT, 0));

            rowNum++;

            // =========================
            // B. Daily Records
            // =========================
            Row header = sheet.createRow(rowNum++);

            String[] columns = {
                    "Time", "Team", "Category", "Model",
                    "Article", "Stage", "Code",
                    "TagType", "Status", "Row"
            };

            for (int i = 0; i < columns.length; i++) {
                Cell cell = header.createCell(i);
                cell.setCellValue(columns[i]);
                cell.setCellStyle(boldStyle);
            }

            for (Record r : overall.get(date)) {

                Row row = sheet.createRow(rowNum++);

                row.createCell(0).setCellValue(
                        r.time.format(timeFormat));
                row.createCell(1).setCellValue(r.team.toString());
                row.createCell(2).setCellValue(r.categories.toString());
                row.createCell(3).setCellValue(r.modelName);
                row.createCell(4).setCellValue(r.article);
                row.createCell(5).setCellValue(r.stage.toString());
                row.createCell(6).setCellValue(r.code);
                row.createCell(7).setCellValue(r.tagType.toString());
                row.createCell(8).setCellValue(r.status.toString());
                row.createCell(9).setCellValue(r.row);
            }

            rowNum += 2;

            // =========================
            // C. Closed RFID
            // =========================
            Row closeTitle = sheet.createRow(rowNum++);
            closeTitle.createCell(0).setCellValue("Closed RFID");
            closeTitle.getCell(0).setCellStyle(boldStyle);

            Row closeHeader = sheet.createRow(rowNum++);

            String[] closeColumns = {
                    "Code", "Article", "Model",
                    "StockIn Date", "StockOut Time"
            };

            for (int i = 0; i < closeColumns.length; i++) {
                Cell cell = closeHeader.createCell(i);
                cell.setCellValue(closeColumns[i]);
                cell.setCellStyle(boldStyle);
            }

            List<Record> closeList =
                    closed.getOrDefault(date, new ArrayList<>());

            for (Record out : closeList) {

                Row row = sheet.createRow(rowNum++);

                row.createCell(0).setCellValue(out.code);
                row.createCell(1).setCellValue(out.article);
                row.createCell(2).setCellValue(out.modelName);

                // ★ StockIn Date：使用 RFID 真正的入庫日期
                String stockInDate =
                        inHouse.getOrDefault(out.code, "NOT FOUND");

                row.createCell(3).setCellValue(stockInDate);

                // ★ StockOut Time：使用真正的出庫時間
                row.createCell(4).setCellValue(
                        out.time.format(timeFormat));
            }

            // Auto size
            for (int i = 0; i < 10; i++) {
                sheet.autoSizeColumn(i);
            }
        }

        try (FileOutputStream output =
                     new FileOutputStream(outputPath)) {

            workbook.write(output);
        }
        System.out.print("Done");
        workbook.close();
    }



}
