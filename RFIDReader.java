package RFIDreader;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

enum Categories{AMERICANFOOTBAL, OUTDOOR, SOCCER, SPORTWEAR, BASEBALL, INNOVATION}
enum TagType{STICKER, HARDTAG}
enum Status{STOCKIN, STOCKOUT, STOCK}
enum Stage{CR0, PULLOVER, CR1, CR2, SMU, CS1, CS2, CFM, PROMO, LOOKSEE, MFA, OTH, PULLOVERMFA, PRECR0}
enum Team{L4A, L4AMA, L4B, L4BMB}
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
        LocalDateTime inTime;
        LocalDateTime outTime;
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
            this.inTime = time;
        }
        public Record(LocalDateTime inTime, LocalDateTime outTime, Team team, String article, String modelName,Categories categories, TagType tagType, Stage stage, Status status, String code, int row){
            this.article = article;
            this.modelName = modelName;
            this.categories = categories;
            this.tagType = tagType;
            this.stage = stage;
            this.status = status;
            this.team = team;
            this.code = code;
            this.row = row;
            this.inTime = inTime;
            this.outTime = outTime;
        }

        @Override
        public String toString() {
            return inTime.toString()+" "+outTime+" "+team.toString()+" "+categories.toString()+" "+modelName+" "+article+" "+stage+" "+stage.toString()+" "+code+" "+tagType.toString()+" "+status.toString()+" "+row;
        }
    }
    String path;
    TreeMap<String, List<Record>> overall;
    TreeMap<String, Map<Status, Integer>> datCount;
    TreeMap<String, List<Record>> closed;
    Map<String, Record> inHouse;
    List<String> bugs = new ArrayList<>();
    public RFIDReader(String path) throws IOException {
        Set<String> bug = new HashSet<>();
        overall = new TreeMap<>();
        datCount = new TreeMap<>();
        closed = new TreeMap<>();
        inHouse = new HashMap<>();
        this.path = path;
        int timeCell = 2;
        int modelCell = -1;
        int artCell = -1;
        int tagCell = -1;
        int statusCell = -1;
        int seasonCell = -1;
        int stageCell = -1;
        int categoryCell = -1;
        int idCell = -1;
        int teamCell= -1;
        FileInputStream file = new FileInputStream(path);
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

//        System.out.println(sheet.getRow(2).getCell(15).getNumericCellValue());
//        System.out.println(sheet.getRow(2).getCell(15).getDateCellValue());
//        System.out.println(sheet.getRow(2).getCell(15).getLocalDateTimeCellValue());
        rowLoop:
        for(Row row: sheet){
            if(row.getRowNum()==0){
                int cursor = 0;
                while (cursor<row.getLastCellNum()+1){
                    switch (row.getCell(cursor).getStringCellValue().trim().toUpperCase()){
                        case "TEAM" ->teamCell = cursor;
                        case "SEASON"->seasonCell = cursor;
                        case "STAGE"->stageCell = cursor;
                        case "ID" -> idCell = cursor;
                        case "CATEGORY"-> categoryCell = cursor;
                        case "TAG" -> tagCell = cursor;
                        case "TIME" -> timeCell = cursor;
                        case "MODEL" -> modelCell = cursor;
                        case "STATUS" -> statusCell = cursor;
                        case "ARTICLE"-> artCell = cursor;
                    }
                    cursor++;
                }
            } else if(row.getRowNum()>=1 && row.getCell(0).getStringCellValue().trim().equals("Sample Warehouse")){
                String time = "";
                String season = row.getCell(seasonCell).getStringCellValue().trim();
                Categories categories;
                Team team;
                TagType tag;
                Status status;
                Stage stage;
                try {
                    categories = Categories.valueOf(row.getCell(categoryCell).getStringCellValue().trim().toUpperCase());
                    team = Team.valueOf(row.getCell(teamCell).getStringCellValue().trim().toUpperCase());
                    tag = TagType.valueOf(row.getCell(tagCell).getStringCellValue().trim().toUpperCase());
                    status = Status.valueOf(row.getCell(statusCell).getStringCellValue().trim().toUpperCase());
                    stage = Stage.valueOf(row.getCell(stageCell).getStringCellValue().trim().toUpperCase());
                }catch (IllegalArgumentException e){
                    bugs.add(row.getRowNum()+" incorrect argument");
                    continue rowLoop;
                }
                String model = row.getCell(modelCell).toString().trim();
                String article = row.getCell(artCell).getStringCellValue().trim();
                String id = row.getCell(idCell).getStringCellValue().trim().toUpperCase();
                LocalDateTime key;
                try{
                    DateTimeFormatter timeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
                    key = LocalDateTime.parse(row.getCell(timeCell).getStringCellValue().trim(), timeFormatter);
                }catch (IllegalStateException e){
                    bugs.add(row.getRowNum()+" Time format issue");
                    continue rowLoop;
                }
                Record item = new Record(key, team, article, model, categories, tag, stage, status, id, row.getRowNum());
                String date = key.toString().substring(0, 10);
                if(status.equals(Status.STOCKIN)){
                    if(inHouse.containsKey(id)){
                        bugs.add(item.toString());
                    }else{
                        inHouse.put(id, item);
                        datCount.putIfAbsent(date, new HashMap<>());
                        datCount.get(date).putIfAbsent(item.status, 0);
                        datCount.get(date).compute(item.status, (a,b)->b+1);
                        overall.putIfAbsent(date, new ArrayList<>());
                        overall.get(date).add(item);
                    }
                }else if(status.equals(Status.STOCK) || status.equals(Status.STOCKOUT)){
                    if(!inHouse.containsKey(id) || !inHouse.get(id).team.equals(item.team)){
                        bugs.add(item.toString());
                    }else{
                        overall.putIfAbsent(date, new ArrayList<>());
                        if(status.equals(Status.STOCK)){
                            overall.get(date).add(item);
                            datCount.putIfAbsent(date, new HashMap<>());
                            datCount.get(date).putIfAbsent(item.status, 0);
                            datCount.get(date).compute(item.status, (a,b)->b+1);
                        }else{
                            overall.get(date).add(item);
                            Record update = new Record(inHouse.get(id).inTime, key, item.team, item.article, item.modelName, item.categories, item.tagType, item.stage, item.status, item.code, item.row);
                            closed.putIfAbsent(date, new ArrayList<>());
                            closed.get(date).add(update);
                            datCount.putIfAbsent(date, new HashMap<>());
                            datCount.get(date).putIfAbsent(item.status, 0);
                            datCount.get(date).compute(item.status, (a,b)->b+1);
                            inHouse.remove(id);
                        }
                    }
                }
            }
        }
        file.close();
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

    public void bugLogs(){
        if(bugs.isEmpty()){
            System.out.print("None");
        }else{
            for(var x: bugs){
                System.out.println(x);
            }
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
        Sheet sheet1 = workbook.createSheet("Bug_Logs");
        int rowC = 0;
        Row head = sheet1.createRow(rowC++);
        head.createCell(0).setCellValue("Title");
        for(var x: bugs){
            Row current = sheet1.createRow(rowC++);
            current.createCell(0).setCellValue(x);
        }

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
                        r.inTime.format(timeFormat));
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
                String stockInDate = out.inTime.toString();

                row.createCell(3).setCellValue(stockInDate);

                // ★ StockOut Time：使用真正的出庫時間
                row.createCell(4).setCellValue(
                        out.outTime.format(timeFormat));
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
        System.out.println("Done");
        workbook.close();
    }



}
