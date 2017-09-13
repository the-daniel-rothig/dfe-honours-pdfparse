package uk.gov.education.honours;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;

import javax.net.ssl.HttpsURLConnection;
import java.awt.Color;
import java.io.*;
import java.net.URL;
import java.net.URLEncoder;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;

class KissflowApi {

    private final String kissflowApiKey;

    KissflowApi(String kissflowApiKey) {

        this.kissflowApiKey = kissflowApiKey;
    }

    void importShortlist(File shortlistFile) throws IOException, InvalidFormatException, org.json.simple.parser.ParseException {
        Sheet sheet = WorkbookFactory.create(shortlistFile).getSheetAt(0);
        validateWorksheetIsShortlist(sheet);

        Iterator<Row> rowIterator = sheet.rowIterator();
        rowIterator.next();
        JSONArray records = (JSONArray) callJsonEndpoint("https://kf-0000580.appspot.com/api/1/Honours/list/p1/99999", null, "GET", null);

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            String[] split = row.getCell(7).getStringCellValue().split("/");
            String id = split[split.length-1];

            Optional recordOrNone = records.stream().filter(x -> Objects.equals(((JSONObject) x).get("Id"), id)).findFirst();
            if (!recordOrNone.isPresent()) continue;
            JSONObject record = (JSONObject) recordOrNone.get();

            boolean isProgressNotUpdate =
                    (   Objects.equals(record.getOrDefault("Directorate_shortlist", ""), "")
                     && !Objects.equals(row.getCell(1).getStringCellValue(), ""))
                  ||(   Objects.equals(record.getOrDefault("Departmental_shortlist", ""), "")
                     && !Objects.equals(row.getCell(0).getStringCellValue(), ""));

            Map<String,String> body = new HashMap<>();

            body.put("Departmental_shortlist", URLEncoder.encode(row.getCell(0).getStringCellValue(), "UTF-8"));
            body.put("Directorate_shortlist", URLEncoder.encode(Integer.valueOf(Integer.parseInt(row.getCell(1).getStringCellValue())).toString(), "UTF-8"));
            body.put("Round", URLEncoder.encode(row.getCell(2).getStringCellValue(), "UTF-8"));
            body.put("Proposed_Award", URLEncoder.encode(row.getCell(3).getStringCellValue(), "UTF-8"));
            body.put("Proposed_Committee", URLEncoder.encode(row.getCell(4).getStringCellValue(), "UTF-8"));
            body.put("Proposed_Category", URLEncoder.encode(row.getCell(5).getStringCellValue(), "UTF-8"));

            String fullbody = body.entrySet().stream().map(x -> String.format("%s=%s", x.getKey(), x.getValue())).collect(Collectors.joining("&"));

            if (isProgressNotUpdate) {
                callJsonEndpoint("https://kf-0000580.appspot.com/api/1/Honours/" + id + "/done", fullbody, "POST", "application/x-www-form-urlencoded");
            } else {
                callJsonEndpoint("https://kf-0000580.appspot.com/api/1/Honours/" + id + "/update", fullbody, "PUT", "application/x-www-form-urlencoded");
            }
        }
    }

    private static void validateWorksheetIsShortlist(Sheet sheet) {
        //some basic validation
        boolean goodFormat =
                Objects.equals(sheet.getRow(0).getCell(0).getStringCellValue(), "Departmental rank")
            &&  Objects.equals(sheet.getRow(0).getCell(1).getStringCellValue(), "Directorate rank")
            &&  Objects.equals(sheet.getRow(0).getCell(2).getStringCellValue(), "Round")
            &&  Objects.equals(sheet.getRow(0).getCell(3).getStringCellValue(), "Proposed award")
            &&  Objects.equals(sheet.getRow(0).getCell(4).getStringCellValue(), "Proposed committee")
            &&  Objects.equals(sheet.getRow(0).getCell(5).getStringCellValue(), "Proposed category");

        if (!goodFormat) {
            throw new UnsupportedOperationException("Wrong table format: Columns should be Departmental rank, Directorate rank, Round, Proposed award, Proposed committee, Proposed category");
        }
    }

    XSSFWorkbook getShortlist(String directorate, String round) throws IOException, org.json.simple.parser.ParseException {
        boolean filterToDirectorate = !Objects.equals(directorate, null) && !Objects.equals(directorate, "");
        boolean filterToRound = !Objects.equals(round, null) && !Objects.equals(round, "");

        JSONArray results = (JSONArray) callJsonEndpoint("https://kf-0000580.appspot.com/api/1/Honours/list/p1/99999", null, "GET", null);
        int currentRow = 0;

        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet();
        XSSFRow header = appendRow(sheet, currentRow++, Arrays.asList(
            "Departmental rank",
            "Directorate rank",
            "Round",
            "Proposed award",
            "Proposed committee",
            "Proposed category",
            "Directorate",
            "Case link",
            "Forenames",
            "Surnames",
            "Short citation",
            "Region",
            "Gender",
            "Ethnic group"
        ));

        XSSFCellStyle headerStyle = wb.createCellStyle();
        XSSFFont headerFont = wb.createFont();
        headerFont.setBold(true);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(new XSSFColor(new Color(25, 129, 183)));
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Iterator<Cell> cellIterator = header.cellIterator();
        while(cellIterator.hasNext()) cellIterator.next().setCellStyle(headerStyle);

        XSSFCellStyle depStyle = wb.createCellStyle();
        depStyle.setFillForegroundColor(new XSSFColor(new Color(249, 203, 156)));
        depStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        XSSFCellStyle dirStyle = wb.createCellStyle();
        dirStyle.setFillForegroundColor(new XSSFColor(new Color(200, 231, 247)));
        dirStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        for(Object o : results) {
            JSONObject item = (JSONObject) o;

            if (filterToDirectorate && (!Objects.equals(directorate, item.getOrDefault("Directorate","")) || !item.containsKey("Assigned To-Directorate Shortlist"))) continue;
            if (filterToRound && ((Double) item.getOrDefault("Directorate_shortlist", 0.0) < 0.5 || !Objects.equals(directorate, item.getOrDefault("Round","")) || !item.containsKey("Assigned To-Department Shortlist"))) continue;

            XSSFRow thisRow = appendRow(sheet, currentRow, Arrays.asList(
                    (String) item.getOrDefault("Departmental_shortlist", ""),
                    Integer.valueOf(Double.valueOf((double) item.getOrDefault("Directorate_shortlist", 0.0)).intValue()).toString(),
                    (String) item.getOrDefault("Round", ""),
                    (String) item.getOrDefault("Proposed_Award", ""),
                    (String) item.getOrDefault("Proposed_Committee", ""),
                    (String) item.getOrDefault("Proposed_Category", ""),
                    (String) item.getOrDefault("Directorate", ""),
                    "https://kf-0000580.appspot.com/#/inbox/Provide%20Input/Sh25328874_8f06_11e7_addd_062ed84aadae/Acc70d887e_8f06_11e7_addd_062ed84aadae/" + item.getOrDefault("Id", ""),
                    (String) item.getOrDefault("First_Name", ""),
                    (String) item.getOrDefault("Last_Name", ""),
                    (String) item.getOrDefault("Short_citation", ""),
                    (String) item.getOrDefault("Region", ""),
                    (String) item.getOrDefault("Gender", ""),
                    (String) item.getOrDefault("Ethnic_Group", "")
            ));

            thisRow.getCell(0).setCellStyle(depStyle);
            for (int i = 1; i<6; i++) thisRow.getCell(i).setCellStyle(dirStyle);
        }

        return wb;
    }

    XSSFWorkbook getFinalShortlist(String round) throws IOException, org.json.simple.parser.ParseException {
        boolean filterToRound = false;

        if (round.matches("[0-9]{4} [A-Z]{2}")) {
            filterToRound = true;
        }

        XSSFWorkbook wb = new XSSFWorkbook();

        XSSFSheet sheet = wb.createSheet("Sheet 1");
        int currentRow = 0;
        XSSFRow header = appendRow(sheet, currentRow++, Arrays.asList(
                "Department",
                "Hon List",
                "Year",
                "Surname",
                "Forename(s)",
                "AKA",
                "Preferred Name",
                "Title",
                "Post-Nominals",
                "Award",
                "Original Award",
                "DOB",
                "Approx DoB",
                "Leaving Current Post",
                "Total Length of Service",
                "Length of Service in Post",
                "Length of Service in Grade",
                "Nationality",
                "Foreign National?",
                "Short Citation",
                "Long Citation",
                "Edited Long Citation",
                "Address 1",
                "Address 2",
                "Address 3",
                "Town",
                "County",
                "Country",
                "Postcode",
                "SecureAddress",
                "Telephone Number",
                "Rating",
                "Public",
                "NomineesOrigin",
                "Committee",
                "Original Committee",
                "Category",
                "Original Category",
                "Gender",
                "Voluntary Work",
                "Previous Honours & Dates",
                "Previous Recommendations",
                "NominatorsOrigin"
        ));

        XSSFCellStyle headerStyle = wb.createCellStyle();
        XSSFFont headerFont = wb.createFont();
        headerFont.setBold(true);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(new XSSFColor(new Color(25, 129, 183)));
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Iterator<Cell> cellIterator = header.cellIterator();
        while(cellIterator.hasNext()) cellIterator.next().setCellStyle(headerStyle);

        JSONArray results = (JSONArray) callJsonEndpoint("https://kf-0000580.appspot.com/api/1/Honours/list/p1/99999", null, "GET", null);
        for (Object o : results) {
            JSONObject item = (JSONObject) o;

            if(filterToRound && !Objects.equals(round, item.getOrDefault("Round", "")) && (Objects.equals(item.getOrDefault("Departmental_shortlist", ""),"") || item.containsKey("Assigned To"))) {
                continue;
            }

            Matcher year = Pattern.compile("[0-9]{4}").matcher((String) item.getOrDefault("Round", ""));
            boolean dobApproximate = Objects.equals(item.getOrDefault("Is_Date_of_Birth_approximate", ""), "Yes");
            String[] addressLines = ((String) item.getOrDefault("Address", "")).split("\n");

            appendRow(sheet, currentRow++, Arrays.asList(
                    "DfE",
                    (String) item.getOrDefault("Round", ""),
                    year.find() ? year.group() : "",
                    (String) item.getOrDefault("Last_Name", ""),
                    (String) item.getOrDefault("First_Name", ""),
                    "",
                    (String) item.getOrDefault("Preferrred_Names", ""),
                    (String) item.getOrDefault("Title", ""),
                    (String) item.getOrDefault("PostNominals", ""),
                    (String) item.getOrDefault("Proposed_Award", ""),
                    (String) item.getOrDefault("Proposed_Award", ""),
                    dobApproximate ? "" : (String) item.getOrDefault("Date_of_Birth", ""),
                    dobApproximate ? (String) item.getOrDefault("Date_of_Birth", "") : "",
                    (String) item.getOrDefault("Date_Leaving_Post", ""),
                    (String) item.getOrDefault("Total_Length_Of_Service", ""),
                    Objects.equals(item.getOrDefault("Is_this_a_state_nomination", ""), "Yes")? "" : item.get("Years").toString(),
                    Objects.equals(item.getOrDefault("Is_this_a_state_nomination", ""), "Yes") ? item.get("Years").toString() : "",
                    (String) item.getOrDefault("Nationality",""),
                    ((String) item.getOrDefault("Nationality","")).contains("British") ? "N" : "Y",
                    (String) item.getOrDefault("Short_citation", ""),
                    (String) item.getOrDefault("Long_citation", ""),
                    (String) item.getOrDefault("Long_citation", ""),
                    addressLines[0],
                    addressLines.length > 5 ? addressLines[1] : "",
                    addressLines.length > 6 ? addressLines[2] : "",
                    addressLines.length > 4 ? addressLines[addressLines.length-4] : "",
                    addressLines.length > 3 ? addressLines[addressLines.length-3] : "",
                    addressLines.length > 1 ? addressLines[addressLines.length-1] : "",
                    addressLines.length > 2 ? addressLines[addressLines.length-2] : "",
                    Objects.equals(item.getOrDefault("Secure_Address", ""), "Yes") ? "Y" : "N",
                    (String) item.getOrDefault("Telephone", ""),
                    (String) item.getOrDefault("Departmental_shortlist", ""),
                    ((String) item.getOrDefault("Nomination_Type", "")).contains("Public") ? (String) item.getOrDefault("Nomination_Type", "") : "",
                    (String) item.getOrDefault("Ethnic_Group", ""),
                    (String) item.getOrDefault("Proposed_Committee", ""),
                    (String) item.getOrDefault("Proposed_Committee", ""),
                    (String) item.getOrDefault("Proposed_Category", ""),
                    (String) item.getOrDefault("Proposed_Category", ""),
                    (String) item.getOrDefault("Gender", ""),
                    Objects.equals(item.getOrDefault("Voluntary_work", ""), "Yes") ? "Y" : "N",
                    (String) item.getOrDefault("Previous_HonoursRecommendations", ""),
                    (String) item.getOrDefault("Previous_HonoursRecommendations", ""),
                    (String) item.getOrDefault("Ethnic_Group_", "")));
        }

        return wb;
    }

    private static XSSFRow appendRow(XSSFSheet sheet, int currentRow, List<String> strings) {
        XSSFRow row = sheet.createRow(currentRow);
        int cellNum = 0;
        for (String string : strings) {
            row.createCell(cellNum++).setCellValue(string);
        }
        return row;
    }


    String sendToKissflow(List<Section> res, String fileName, String nominationFileLocation, List<NominationUploader.FileNameAndPath> uploadedFiles) throws IOException, ParseException, org.json.simple.parser.ParseException {
        String httpsURL = "https://kf-0000580.appspot.com/api/1/Honours/submit";
        ResParser d = new ResParser(res);
        JSONObject body = new JSONObject();

        JSONObject nomFile = new JSONObject();
        nomFile.put(fileName, nominationFileLocation);
        body.put("Nomination_doc", nomFile);

        JSONObject evidence = new JSONObject();
        for(NominationUploader.FileNameAndPath fnap : uploadedFiles) evidence.put(fnap.name, fnap.path);
        body.put("Evidence", evidence);

        body.put("Nomination_Type", "Public - Central");

        // Map nominee details
        body.put("Title", d.getProperty(1, "Title"));
        body.put("First_Name", d.getProperty(1, "Forename"));
        body.put("Last_Name", d.getProperty(1, "Surname"));
        body.put("Telephone", d.getProperty(1, "Telephone number"));
        body.put("Address", Stream.of("Street", "Town or City", "County", "Postcode", "Country").map(x -> d.getProperty(1, x)).filter(x -> x != null).collect(Collectors.joining("\n")));
        body.put("Date_of_Birth", d.getSimpleAnswer(1, "What is your nominee's date of birth?"));
        body.put("Age", Double.valueOf((new Date().getTime() - d.getSimpleAnswerAsDate(1, "What is your nominee's date of birth?").getTime()) / (1000*3600*24*365.24) + 1).intValue());
        body.put("Disability", Objects.equals(d.getProperty(1, "Equality monitoring", "Disability"),"Yes") ? "Yes" : "No");
        body.put("Nationality", d.getSimpleAnswer(1, "What is your nominee’s nationality?"));
        body.put("Ethnic_Group", d.getProperty(1, "Equality monitoring", "Ethnic Origin Group"));


        // Map last place of work
        if (!d.getRows(2, "List the posts your nominee has excelled in").isEmpty()) {
            Map<String, String> row = d.getRows(2, "List the posts your nominee has excelled in").get(0);
            body.put("Description", row.get("Name"));


            String s = row.get("Start Date to End Date");
            Matcher fromToPattern = Pattern.compile("([A-Za-z]+ [0-9]{4}) to ([A-Za-z]+ [0-9]{4})").matcher(s);
            Matcher fromPattern = Pattern.compile("([A-Za-z]+ [0-9]{4}).*").matcher(s);
            SimpleDateFormat sdf = new SimpleDateFormat("MMM yyyy");

            if (fromToPattern.find()) {
                int years = Double.valueOf(Math.ceil((sdf.parse(fromToPattern.group(2)).getTime() - sdf.parse(fromToPattern.group(1)).getTime()) / 3600000 / 24 / 365.25)).intValue();
                body.put("Years", years);
                body.put("Date_Leaving_Post", fromToPattern.group(2));
            }
            else if (fromPattern.find()) {
                int years = Double.valueOf(Math.ceil((new Date().getTime() - sdf.parse(fromPattern.group(1)).getTime())/3600000/24/365.25)).intValue();
                body.put("Years", years);
            }
        }

        // Map nominator details
        body.put("Title_", d.getProperty(0, "Title"));
        body.put("First_Name_", d.getProperty(0, "Forename"));
        body.put("Last_Name_", d.getProperty(0, "Surname"));
        body.put("Telephone_", d.getProperty(0, "Telephone number"));
        body.put("Email", d.getProperty(0, "Email"));
        body.put("Address_", Stream.of("Street", "Town or City", "County", "Postcode", "Country").map(x -> d.getProperty(0, x)).filter(x -> x != null).collect(Collectors.joining("\n")));
        body.put("Ethnic_Group_", d.getProperty(0, "Equality monitoring", "Ethnic Origin Group"));

        body.put("Relationship_to_nominee", d.getSimpleAnswer(1, "What is your relationship to the nominee?"));


        JSONObject submit = (JSONObject) callJsonEndpoint(httpsURL, body.toJSONString(), "POST", "application/json");

        return String.format("https://kf-0000580.appspot.com/#/inbox/Provide Input/Sh25328874_8f06_11e7_addd_062ed84aadae/Ac56fe5508_8f07_11e7_addd_062ed84aadae/%s", (String) submit.get("Id"));
    }

    Object callJsonEndpoint(String httpsURL, String body, String method, String contentType) throws IOException, org.json.simple.parser.ParseException {
        URL myurl = new URL(httpsURL);
        HttpsURLConnection con = (HttpsURLConnection)myurl.openConnection();
        con.setRequestMethod(method);

        if (contentType != null) con.setRequestProperty("Content-Type", contentType);
        //con.setRequestProperty("User-Agent", "Mozilla/4.0 (compatible; MSIE 5.0;Windows98;DigExt)");
        con.setRequestProperty("api-key",kissflowApiKey);
        con.setDoInput(true);

        if (body != null) {
            con.setRequestProperty("Content-Length", String.valueOf(body.length()));

            con.setDoOutput(true);

            DataOutputStream output = new DataOutputStream(con.getOutputStream());

            output.writeBytes(body);
            output.close();
        }

        System.out.println("Resp Code:"+con .getResponseCode());

        return new JSONParser().parse(new InputStreamReader(con.getInputStream()));
    }


    static private class ResParser {
        final private List<Section> res;

        ResParser(List<Section> res) {
            this.res = res;
        }

        String getProperty(int section, String property) {
            return getProperty(section, "Details", property);
        }

        String getProperty(int section, String subsection, String property) {
            try {
                return res.get(section).questions.get(subsection).answers.get(0).get(property);
            } catch (IndexOutOfBoundsException ignored) {
                return null;
            } catch (NullPointerException ignored) {
                return null;
            }
        }

        String getSimpleAnswer(int section, String question) {
            try {
                return res.get(section).questions.get(question).simpleAnswer;
            } catch (IndexOutOfBoundsException ignored) {
                return null;
            } catch (NullPointerException ignored) {
                return null;
            }
        }

        Date getSimpleAnswerAsDate(int section, String question) throws UnsupportedEncodingException, ParseException {
            String dateString = res.get(section).questions.get(question).simpleAnswer;
            return new SimpleDateFormat("dd/MM/yyyy").parse(dateString);
        }

        List<Map<String,String>> getRows (int section, String question) {
            return res.get(section).questions.get(question).answers;
        }
    }

}
