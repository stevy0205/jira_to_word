import java.io.*;
import java.math.BigInteger;
import java.net.URL;
import java.net.HttpURLConnection;
import java.util.HashMap;
import java.util.LinkedHashMap;

import org.apache.poi.xwpf.usermodel.*;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;


public class JiraToWord2 {

    public static String loginResponse = "";
    public static String jSessionID = "";
    public static String[] jsonData = new String[100];
    public static String baseURL = "https://jira.ppal.directory/rest/";
    public static String loginURL = "auth/1/session";
    public static String biExportURL = "atm/1.0/testcase/";
    public static String testCaseURL = "atm/1.0/testcase/search?query=folder%20=%20/Sandbox";

    public static String testResultUrl = "atm/1.0/testcase/CFRVIA-T59/testresult/latest";
    public static String loginUserName = "uwagstev";
    public static String loginPassWord = "Uwag!Stev!0205";
    public static boolean errorsOccurred = false;
    private static String[] testExpectedResults;
    private static String[] testDescriptions;
    private static String[] testcases;

    public static void main(String[] args) throws Exception {


        if(!errorsOccurred)
        {
            loginResponse = loginToJira(baseURL, loginURL, loginUserName, loginPassWord);
            if(loginResponse == "ERROR") { errorsOccurred = true; }
        } else {
            System.out.println("Login Failed");
        }
        if(!errorsOccurred)
        {
            jSessionID = parseJSessionID(loginResponse);
            if(jSessionID == "ERROR") { errorsOccurred = true; }
        } else {
            System.out.println("Jsession could not be parsed");
        }
        if(!errorsOccurred)
        {
            testcases = new JiraToWord2().getTestCases();
            for (int i = 0; i < testcases.length; i++) {

                System.out.println(getJsonData(baseURL, testResultUrl, jSessionID));
                jsonData[i] = getJsonData(baseURL, biExportURL+testcases[i], jSessionID);
                System.out.println(i);

                if(jsonData[i] == "ERROR") { errorsOccurred = true; }
            }

        } else {
            System.out.println("JSON Data unavailable");
        }

        if(!errorsOccurred){

            new JiraToWord2().createDoc();

        }

        if(!errorsOccurred)
        {
            System.out.println("SUCCESS");
        } else {
            System.out.println("FAILURE");
        }
    }


    public static String loginToJira(String baseURL, String loginURL, String loginUserName, String loginPassWord) {
        String loginResponse = "";
        URL url = null;
        HttpURLConnection conn = null;
        String input = "";
        OutputStream os = null;
        BufferedReader br = null;
        String output = null;

        try {
            // Create URL object
            url = new URL(baseURL + loginURL);
            // Use the URL to create connection
            conn = (HttpURLConnection) url.openConnection();
            // Set properties
            conn.setDoOutput(true);
            conn.setRequestMethod("GET");
            conn.setRequestProperty("Content-Type", "application/json");
            conn.setRequestProperty("Authorization", "Basic dXdhZ3N0ZXY6VXdhZyFTdGV2ITAyMDU=");



            // Create JSON post data

            input = "{\"username\":\"" + loginUserName + "\", \"password\":\"" + loginPassWord + "\"}";
            System.out.println("Input:"+ input);

            // Send our request
            os = conn.getOutputStream();
            os.write(input.getBytes());
            os.flush();


            // Handle our response
            if(conn.getResponseCode() == 200){
                br = new BufferedReader(new InputStreamReader((conn.getInputStream())));
                while((output = br.readLine()) != null){
                    loginResponse += output;
                }
                System.out.println(conn.getResponseMessage());
                System.out.println(conn.getContentEncoding());
                conn.disconnect();
            } else {
                System.out.println(conn.getResponseCode()+conn.getResponseMessage());
                System.out.println("Response failed");
            }
        } catch (Exception ex) {
            System.out.println("Error in loginToJira: " + ex.getMessage());
            loginResponse = "ERROR";
        }
        System.out.println("\nloginResponse:");
        System.out.println(loginResponse);
        return loginResponse;
    }


    public static String parseJSessionID(String input){
        String jSessionID = "";
        try {
            JSONParser parser = new JSONParser();
            Object obj = parser.parse(input);
            JSONObject jsonObject = (JSONObject) obj;
            JSONObject sessionJsonObj = (JSONObject) jsonObject.get("session");
            jSessionID = (String) sessionJsonObj.get("value");
        } catch (Exception ex) {
            System.out.println("Error in parseJSessionID: " + ex.getMessage());
            ex.printStackTrace();
            jSessionID = "ERROR";
        }
        System.out.println("\njSessionID:");
        System.out.println(jSessionID);
        return jSessionID;
    }

    public String[] getTestCases() throws ParseException {
        String jsonTestCases = "";
        try {
            URL url = new URL(baseURL + testCaseURL);
            String cookie = "JSESSIONID=" + jSessionID;
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();
            conn.setRequestMethod("GET");
            conn.setRequestProperty("Content-Type", "application/json");
            conn.setRequestProperty("Autorization","Basic dXdhZ3N0ZXY6VXdhZyFTdGV2ITAyMDU=");
            conn.setRequestProperty("Cookie", cookie);
            if(conn.getResponseCode() == 200)
            {
                BufferedReader br = new BufferedReader(new InputStreamReader(conn.getInputStream()));
                String output = "";
                while((output = br.readLine()) != null){
                    jsonTestCases+=output;
                }
                conn.disconnect();
            }
        } catch (Exception ex){
            System.out.println("Error in getJsonData: " + ex.getMessage());
            jsonTestCases = "ERROR";
        }

        // Make JSON Object from JSON String
        JSONParser parser = new JSONParser();
        JSONArray json = (JSONArray) parser.parse(jsonTestCases);
        String[] testCases = new String[json.size()];
        for (int i = 0; i < json.size(); i++) {
            JSONObject jsonObject = (JSONObject)json.get(i);
            testCases[i] = (String) jsonObject.get("key");
            System.out.println("Testcase Nr." + i+1 +":" + testCases[i]);
        }

        return testCases;
    }

    public static String getJsonData(String baseURL, String biExportURL, String jSessionID) {
        String jsonData = "";
        try {
            URL url = new URL(baseURL + biExportURL);
            String cookie = "JSESSIONID=" + jSessionID;
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();
            conn.setRequestMethod("GET");
            conn.setRequestProperty("Content-Type", "application/json");
            conn.setRequestProperty("Autorization","Basic dXdhZ3N0ZXY6VXdhZyFTdGV2ITAyMDU=");
            conn.setRequestProperty("Cookie", cookie);
            if(conn.getResponseCode() == 200)
            {
                BufferedReader br = new BufferedReader(new InputStreamReader(conn.getInputStream()));
                String output = "";
                while((output = br.readLine()) != null){
                    jsonData += output;
                }
                conn.disconnect();
            }
        } catch (Exception ex){
            System.out.println("Error in getJsonData: " + ex.getMessage());
            jsonData = "ERROR";
        }
        System.out.println("\njsonData:");
        System.out.println(jsonData);

        //deal with Special Characters
        jsonData = jsonData.replaceAll("amp;","");
        jsonData = jsonData.replaceAll("&gt;",">");
        jsonData = jsonData.replaceAll("&lt;","<");
        jsonData = jsonData.replaceAll("\\.",". ");
        jsonData = jsonData.replaceAll("<[a-z]([^>])*>|</[a-z]([^>])*>", "");

        System.out.println(jsonData);
        return jsonData;
    }

    public LinkedHashMap<String,String> jsonToTxt(String jsonData) throws IOException, ParseException {


        System.out.println("Jsondata is "+jsonData);

        // Make JSON Object from JSON String
        JSONParser parser = new JSONParser();
        JSONObject json = (JSONObject) parser.parse(jsonData);

        // Get Values From Keys
        LinkedHashMap<String,String> map = new LinkedHashMap<>(1000);

        String name = json.get("name").toString();
        map.put("projectName",name);

        String createdOn = json.get("createdOn").toString();
        map.put("createdOn",createdOn);

        String objective = json.get("objective").toString();
        map.put("objective",objective);

        String priority = json.get("priority").toString();
        map.put("priority",priority);

        String precondition = json.get("precondition").toString();
        map.put("precondition",precondition);

        String owner = json.get("owner").toString();
        map.put("owner",owner);

        String updatedBy = json.get("updatedBy").toString();
        map.put("updatedBy",updatedBy);

        String status = json.get("status").toString();
        map.put("status",status);

        String folder = json.get("folder").toString();
        map.put("folder",folder);

        String estTime = json.get("estimatedTime").toString();
        map.put("estimatedTime",estTime);

        //Get Values from JSONArrays
        JSONObject testScript = (JSONObject) json.get("testScript");
        JSONArray steps = (JSONArray) testScript.get("steps");
        JSONObject[] stepsArray = new JSONObject[1000];

        testExpectedResults = new String[steps.size()];
        testDescriptions = new String[steps.size()];

        for (int i = steps.size()-1; i >= 0; i--) {
            stepsArray[i] = (JSONObject) steps.get(i);
            testExpectedResults[i] = stepsArray[i].get("expectedResult").toString();
            map.put("Erwartetes Testergebnis",testExpectedResults[i]);
            testDescriptions[i] = stepsArray[i].get("description").toString();
            map.put("Testbeschreibung",testDescriptions[i]);
        }



        // Write to txt.file
        FileWriter file = new FileWriter("/Users/uwagstev/Documents/json.txt");
        for (String line : map.keySet()) {
            line += "\n";
            file.write(line);
        }

        file.flush();
        file.close();

        return map;
    }



    public void createDoc() throws IOException, ParseException {
        String fileName = "c:/Users/uwagstev/Documents/word.docx";
        try (XWPFDocument doc = new XWPFDocument()) {
            for (int i = 0; i < testcases.length; i++) {
                HashMap<String, String> map = new JiraToWord2().jsonToTxt(jsonData[i]);

                createTitle(map, doc);

                createSubtitleWithName(map, doc);

                createTestGoals(map, doc);

                createEstimatedTime(map, doc);

                createPreconditions(map, doc);

                //Create enumeration for Preconditions

                createTable(doc);

                doc.createParagraph().createRun().addBreak();

                createStatusTable(doc);

                doc.createParagraph().createRun().addBreak(BreakType.PAGE);

            }

            // save it to .docx file
            try (FileOutputStream out = new FileOutputStream(fileName)) {
                doc.write(out);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }

        }

    }

    private static void createPreconditions(HashMap<String, String> map, XWPFDocument doc) {
        //create subtitle preconditions
        XWPFParagraph p6 = doc.createParagraph();
        XWPFRun precondition = p6.createRun();

        precondition.setBold(true);
        precondition.setFontSize(12);
        precondition.setFontFamily("Arial");
        precondition.setText("Voraussetzungen/Bedingungen: " + "\r");

        //create text for precondition
        XWPFParagraph p7 = doc.createParagraph();
        XWPFRun preconditionText = p7.createRun();

        preconditionText.setFontSize(11);
        preconditionText.setFontFamily("Arial");
        preconditionText.setText(map.get("precondition").replaceAll("<.*?>", ""));
    }

    private static void createEstimatedTime(HashMap<String, String> map, XWPFDocument doc) {
        //create text for estimated time
        XWPFParagraph p5 = doc.createParagraph();
        XWPFRun estTimeText = p5.createRun();
        Double estTime = Double.parseDouble(map.get("estimatedTime"))*0.00000027777777777778;
        String estTimeStr = String.format("%.2f",estTime);

        estTimeText.setFontSize(11);
        estTimeText.setFontFamily("Arial");

        estTimeText.setText("Voraussichtliche Dauer: " + estTimeStr + "h");
    }

    private static void createTestGoals(HashMap<String, String> map, XWPFDocument doc) {
        //create subtitle testgoals
        XWPFParagraph p3 = doc.createParagraph();
        XWPFRun testgoal = p3.createRun();

        testgoal.setBold(true);
        testgoal.setFontSize(12);
        testgoal.setFontFamily("Arial");
        testgoal.setText("Testziele");

        //create text for testgoals
        XWPFParagraph p4 = doc.createParagraph();
        XWPFRun testgoalText = p4.createRun();

        testgoalText.setFontSize(11);
        testgoalText.setFontFamily("Arial");
        testgoalText.setText(map.get("objective").replaceAll("<.*?>", ""));
    }

    private static void createSubtitleWithName(HashMap<String, String> map, XWPFDocument doc) {
        //create a subtitle with name
        XWPFParagraph p2 = doc.createParagraph();
        XWPFRun name = p2.createRun();

        name.setBold(true);
        name.setFontSize(14);
        name.setFontFamily("Arial");
        name.setText(map.get("projectName").replaceAll("\\[.*?\\]","").replaceFirst(" ","") + "\r");
    }

    private static void createTitle(HashMap<String, String> map, XWPFDocument doc) {
        // create a title
        XWPFParagraph p1 = doc.createParagraph();
        XWPFRun title = p1.createRun();

        title.setBold(true);
        title.setFontSize(16);
        title.setFontFamily("Arial");
        title.setText(map.get("folder").replaceAll("/", "") + "\r");
    }

    private static void createTable(XWPFDocument doc) {
        //Create Title for Table
        XWPFParagraph p8 = doc.createParagraph();
        XWPFRun tableTitle = p8.createRun();

        tableTitle.setFontFamily("Arial");
        tableTitle.setFontSize(12);
        tableTitle.setBold(true);
        tableTitle.setText("Prozedur");

        //create table with steps (expected and description of teststep)
        XWPFTable table = doc.createTable();

        CTTblLayoutType type = table.getCTTbl().getTblPr().addNewTblLayout();
        type.setType(STTblLayoutType.FIXED);

        //Creating first Row
        XWPFTableRow row1 = table.getRow(0);
        table.setWidth("100%");
        table.setTableAlignment(TableRowAlign.CENTER);

        row1.getCell(0).setText("ID");
        row1.getCell(0).setColor("F2F2F2");

        row1.addNewTableCell().setText("Aktion");
        row1.getCell(1).setColor("F2F2F2");

        row1.addNewTableCell().setText("Reaktion");
        row1.getCell(2).setColor("F2F2F2");

        row1.addNewTableCell().setText("ja");
        row1.getCell(3).setColor("F2F2F2");


        row1.addNewTableCell().setText("nein");
        row1.getCell(4).setColor("F2F2F2");

        for (int i = 0; i < testExpectedResults.length; i++) {
            XWPFTableRow row = table.createRow();
            row.getCell(0).setText(String.valueOf(i+1));
            row.getCell(0).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
            row.getCell(0).getParagraphs().get(0).getRuns().get(0).setFontSize(9);
            row.getCell(0).getParagraphs().get(0).getRuns().get(0).setFontFamily("Arial");

            row.getCell(1).setText(testDescriptions[i]);
            row.getCell(1).getParagraphs().get(0).getRuns().get(0).setFontSize(9);
            row.getCell(1).getParagraphs().get(0).getRuns().get(0).setFontFamily("Arial");

            row.getCell(2).setText(testExpectedResults[i]);
            row.getCell(2).getParagraphs().get(0).getRuns().get(0).setFontSize(9);
            row.getCell(2).getParagraphs().get(0).getRuns().get(0).setFontFamily("Arial");
        }


        table.getRow(0).getCell(0).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(400));
        table.getRow(0).getCell(1).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(4500));
        table.getRow(0).getCell(2).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(3800));
        table.getRow(0).getCell(3).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(600));
        table.getRow(0).getCell(4).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(600));


        for (int i= 0; i < 5 ; i ++) {
            table.getRow(0).getCell(i).getParagraphs().get(0).setSpacingAfter(10);
            table.getRow(0).getCell(i).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
            table.getRow(0).getCell(i).getParagraphs().get(0).setVerticalAlignment(TextAlignment.BOTTOM);
            table.getRow(0).getCell(i).getParagraphs().get(0).getRuns().get(0).setBold(true);
            table.getRow(0).getCell(i).getParagraphs().get(0).getRuns().get(0).setFontSize(9);
            table.getRow(0).getCell(i).getParagraphs().get(0).getRuns().get(0).setFontFamily("Arial");
            table.getRow(0).getCell(i).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.BOTTOM);
        }
    }

    private static void createStatusTable(XWPFDocument doc) {
        //create table with status (hardcoded)

        XWPFTable statusTable = doc.createTable();
        statusTable.setWidth("100%");

        XWPFTableRow row = statusTable.getRow(0);

        row.getCell(0).setText("  Status:");

        row.addNewTableCell().setText("Erfüllt");

        row.addNewTableCell().setText("□");

        row.addNewTableCell().setText("Teilweise erfüllt");

        row.addNewTableCell().setText("□");

        row.addNewTableCell().setText("Nicht erfüllt");

        row.addNewTableCell().setText("□");

        //Set Column Widths
        statusTable.getRow(0).getCell(0).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1500));
        statusTable.getRow(0).getCell(1).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1500));
        statusTable.getRow(0).getCell(2).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(400));
        statusTable.getRow(0).getCell(3).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1500));
        statusTable.getRow(0).getCell(4).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(400));
        statusTable.getRow(0).getCell(5).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1500));
        statusTable.getRow(0).getCell(6).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(400));

        statusTable.setTopBorder(XWPFTable.XWPFBorderType.THICK, 14, 0, "000000");
        statusTable.setBottomBorder(XWPFTable.XWPFBorderType.THICK, 15, 0, "000000");
        statusTable.setLeftBorder(XWPFTable.XWPFBorderType.THICK, 14, 0, "000000");
        statusTable.setRightBorder(XWPFTable.XWPFBorderType.THICK, 14, 0, "000000");
        statusTable.removeInsideVBorder();


        XWPFTableRow sRow2 = statusTable.createRow();
        sRow2.getCell(0).setText("  Datum:");
        sRow2.getCell(0).getParagraphs().get(0).getRuns().get(0).setBold(true);
        sRow2.getCell(0).getParagraphs().get(0).getRuns().get(0).setFontSize(10);
        sRow2.getCell(0).getParagraphs().get(0).getRuns().get(0).setFontFamily("Arial");
        sRow2.getCell(0).getParagraphs().get(0).setSpacingAfter(0);
        sRow2.setHeightRule(TableRowHeightRule.AT_LEAST);

        XWPFTableRow sRow3 = statusTable.createRow();
        sRow3.getCell(0).setText("  Getestet von:" );
        sRow3.getCell(0).getParagraphs().get(0).getRuns().get(0).setFontSize(10);
        sRow3.getCell(0).getParagraphs().get(0).getRuns().get(0).setFontFamily("Arial");
        sRow3.getCell(0).getParagraphs().get(0).getRuns().get(0).setBold(true);
        sRow3.getCell(1).setText("DPAG:");
        sRow3.getCell(1).getParagraphs().get(0).getRuns().get(0).addBreak();
        sRow3.getCell(1).setText("SPL:");
        sRow3.getCell(0).getParagraphs().get(0).setSpacingBefore(20);
        sRow3.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.TOP);

        XWPFTableRow sRow4 = statusTable.createRow();
        sRow4.getCell(0).setText("  Bemerkungen:");
        sRow4.getCell(0).getParagraphs().get(0).getRuns().get(0).setFontSize(10);
        sRow4.getCell(0).getParagraphs().get(0).getRuns().get(0).setFontFamily("Arial");
        sRow4.getCell(0).getParagraphs().get(0).getRuns().get(0).setBold(true);
        sRow4.getCell(0).getParagraphs().get(0).setSpacingBefore(20);
        sRow4.setHeight(1000);

        for (int i = 0; i < 7; i++) {
            if(i%2==0&&i!=0) {
                row.getCell(i).getParagraphs().get(0).getRuns().get(0).setFontSize(14);
                row.getCell(i).getParagraphs().get(0).getRuns().get(0).setFontFamily("Arial");
                row.getCell(i).getParagraphs().get(0).setAlignment(ParagraphAlignment.LEFT);
                row.getCell(i).getCTTc().addNewTcPr().addNewTextDirection().setVal(STTextDirection.TB_V);
            } else {
                row.getCell(i).getParagraphs().get(0).setAlignment(ParagraphAlignment.RIGHT);
                row.getCell(i).getParagraphs().get(0).getRuns().get(0).setFontFamily("Arial");
                row.getCell(i).getParagraphs().get(0).getRuns().get(0).setFontSize(10);
                row.getCell(i).getParagraphs().get(0).getRuns().get(0).setBold(true);
            }
        }

        for (int i=0; i <7 ; i ++) {
            statusTable.getRow(0).getCell(i).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
            statusTable.getRow(0).getCell(i).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
            statusTable.getRow(0).getCell(i).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
            statusTable.getRow(0).getCell(i).getParagraphs().get(0).setSpacingAfter(0);
        }

        for (int i = 0; i < 3 ; i++) {
            if(i!=2) {
                statusTable.getRow(i).getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.BOTTOM);
                statusTable.getRow(i).getCell(0).getParagraphs().get(0).setAlignment(ParagraphAlignment.LEFT);
            }
            if(i==0&&i==1){
                statusTable.getRow(i).getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                statusTable.getRow(i).getCell(0).getParagraphs().get(0).setAlignment(ParagraphAlignment.LEFT);
            }
            statusTable.getRow(i).getCell(0).getParagraphs().get(0).setSpacingAfter(0);
        }
        sRow2.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        row.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        statusTable.setInsideHBorder(XWPFTable.XWPFBorderType.OUTSET,14,0,"000000");
    }

}



