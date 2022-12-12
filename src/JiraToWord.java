import java.net.URL;
import java.net.HttpURLConnection;
import java.io.OutputStream;
import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.util.Base64;

import org.json.simple.JSONObject;
import org.json.simple.parser.*;


public class JiraToWord {

    public static String loginResponse = "";
    public static String jSessionID = "";
    public static String jsonData = "";

    public static String baseURL = "https://jira.ppal.directory/rest/";
    public static String loginURL = "auth/1/session";
    public static String biExportURL = "atm/1.0/testcase/CFRVIA-T58";

    public static String loginUserName = "uwagstev";
    public static String loginPassWord = "Uwag!Stev!0205";
    public static boolean errorsOccurred = false;



    public static void main(String[] args){


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
            jsonData = getJsonData(baseURL, biExportURL, jSessionID);
            if(jsonData == "ERROR") { errorsOccurred = true; }
        } else {
            System.out.println("JSON Data unavailable");
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

    public static String getJsonData(String baseURL, String biExportURL, String jSessionID){
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
        return jsonData;
    }


}