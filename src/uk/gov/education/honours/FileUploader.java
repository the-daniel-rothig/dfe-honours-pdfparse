package uk.gov.education.honours;

import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.nio.file.Files;

class FileUploader {
    private final String fileApiKey;

    FileUploader(String fileApiKey) {
        this.fileApiKey = fileApiKey;
    }

    String sendFile(String pathToFile, String filename) throws Exception {
        String fileId = uploadFile(pathToFile, filename);
        return getFileUrl(fileId);
    }

    private String getFileUrl(String fileId) throws IOException, ParseException {
        String url = "https://upload.kissflow.com/group/";
        HttpURLConnection connection = (HttpURLConnection) new URL(url).openConnection();

        connection.setRequestProperty("User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64; rv:55.0) Gecko/20100101 Firefox/55.0");

        // Indicate a POST request
        connection.setRequestMethod("POST");
        connection.setDoOutput(true);
        connection.setRequestProperty("Content-Type", "application/x-www-form-urlencoded; charset=UTF-8");

        DataOutputStream output = new DataOutputStream(connection.getOutputStream());
        output.writeBytes(
                String.format("pub_key=%s&%s=%s", URLEncoder.encode(fileApiKey, "UTF-8"), URLEncoder.encode("files[]", "UTF-8"), URLEncoder.encode(fileId, "UTF-8"))
        );
        output.close();

        JSONObject parse = (JSONObject) new JSONParser().parse(new InputStreamReader(connection.getInputStream()));

        System.out.println(parse.toJSONString());

        return parse.get("cdn_url") + "/nth/0/";
    }

    private String uploadFile(String pathToFile, String filename) throws IOException, ParseException {
        // Task attachments endpoint
        String url = "https://upload.kissflow.com/base/?jsonerrors=1";
        File theFile = new File(pathToFile);

        HttpURLConnection connection = (HttpURLConnection) new URL(url).openConnection();

        connection.setRequestProperty("User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64; rv:55.0) Gecko/20100101 Firefox/55.0");

        // Indicate a POST request
        connection.setDoOutput(true);

        String boundary = Long.toHexString(System.currentTimeMillis());
        connection.setRequestProperty("Content-Type", "multipart/form-data; boundary=" + boundary);

        PrintWriter writer = null;
        try {
            OutputStream outputStream = connection.getOutputStream();
            writer = new PrintWriter(new OutputStreamWriter(outputStream, "UTF-8"));
            //String fileName = theFile.getName();

            String LINE_FEED = "\r\n";
            writer.append("--").append(boundary).append(LINE_FEED);
            writer.append("Content-Disposition: form-data; name=\"UPLOADCARE_PUB_KEY\"").append(LINE_FEED).append(LINE_FEED);
            writer.append(fileApiKey).append(LINE_FEED);

            writer.append("--").append(boundary).append(LINE_FEED);
            writer.append("Content-Disposition: form-data; name=\"UPLOADCARE_STORE\"").append(LINE_FEED).append(LINE_FEED);
            writer.append("auto").append(LINE_FEED);

            writer.append("--").append(boundary).append(LINE_FEED);
            writer.append("Content-Disposition: form-data; name=\"file\"; filename=\"").append(filename).append("\"").append(LINE_FEED);
            writer.append("Content-Type: ").append(Files.probeContentType(theFile.toPath())).append(LINE_FEED);
            writer.append(LINE_FEED);
            writer.flush();

            FileInputStream inputStream = new FileInputStream(theFile);
            byte[] buffer = new byte[4096];
            int bytesRead;
            while ((bytesRead = inputStream.read(buffer)) != -1) {
                outputStream.write(buffer, 0, bytesRead);
            }
            outputStream.flush();
            inputStream.close();


            writer.append(LINE_FEED);
            writer.append("--").append(boundary).append("--").append(LINE_FEED);
            writer.append(LINE_FEED);
            writer.flush();
            writer.close();
        } catch (Exception ignored) {

        } finally {
            if (writer != null) writer.close();
        }

        JSONObject parse = (JSONObject) new JSONParser().parse(new InputStreamReader(connection.getInputStream()));

        return (String) parse.get("file");
    }
}
