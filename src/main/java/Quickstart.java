import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.ValueRange;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.Arrays;
import java.util.List;

/**
 * Google Sheet API를 OAuth 2.0과 Service Account로 연동하는 예제코드
 *
 * @author lks21c
 */
public class Quickstart {
    /**
     * OAUTH 2.0 연동시 지정한 OAuth 2.0 클라이언트 이름
     */
    private static final String APPLICATION_NAME =
            "Google Sheets API Java Quickstart";

    /**
     * OAUTH 2.0 연동시 credential을 디스크에 저장할 위치
     */
    private static final java.io.File DATA_STORE_DIR = new java.io.File(
            System.getProperty("user.home"), ".credentials/sheets.googleapis.com-java-quickstart");

    /**
     * Global instance of the {@link FileDataStoreFactory}.
     */
    private static FileDataStoreFactory DATA_STORE_FACTORY;

    /**
     * Global instance of the JSON factory.
     */
    private static final JsonFactory JSON_FACTORY =
            JacksonFactory.getDefaultInstance();

    /**
     * Global instance of the HTTP transport.
     */
    private static HttpTransport HTTP_TRANSPORT;

    /**
     * Google Sheet API 권한을 SCOPE로 지정
     */
    private static final List<String> SCOPES =
            Arrays.asList(SheetsScopes.SPREADSHEETS);

    /**
     * HTTP_TRANSPORT, DATA_STORE_FACTORY 초기화
     */
    static {
        try {
            HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
            DATA_STORE_FACTORY = new FileDataStoreFactory(DATA_STORE_DIR);
        } catch (Throwable t) {
            t.printStackTrace();
            System.exit(1);
        }
    }

    /**
     * OAUTH 2.0 연동시 사용될 callback용 local receiver 포트 지정
     */
    private static final int LOCAL_SERVER_RECEIVER_PORT = 8080;

    /**
     * 인증 모드 2개
     */
    private enum AuthMode {
        OAUTH20, SERVICE_ACCOUNT
    }

    /**
     * OAUTH 2.0용 credential 생성
     *
     * @return Credential object.
     * @throws IOException
     */
    public static Credential getOauth2Authorize() throws IOException {
        // OAUTH 2.0용 secret josn 로드
        InputStream in =
                Quickstart.class.getResourceAsStream("/client_secret.json");
        GoogleClientSecrets clientSecrets =
                GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        GoogleAuthorizationCodeFlow flow =
                new GoogleAuthorizationCodeFlow.Builder(
                        HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                        .setDataStoreFactory(DATA_STORE_FACTORY)
                        .setAccessType("offline")
                        .build();

        LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(LOCAL_SERVER_RECEIVER_PORT).build();

        Credential credential = new AuthorizationCodeInstalledApp(
                flow, receiver).authorize("user");
        System.out.println(
                "Credentials saved to " + DATA_STORE_DIR.getAbsolutePath());
        return credential;
    }

    /**
     * Service Account용 credentail 생성
     * @return Credential object.
     * @throws IOException
     */
    public static Credential getServiceAccountAuthorize() throws IOException {

        InputStream in =
                Quickstart.class.getResourceAsStream("/service.json");
        Credential credential = GoogleCredential.fromStream(in)
                .createScoped(SCOPES);
        return credential;
    }

    /**
     * Google Credential 정보를 가지고 Google Sheet서비스를 초기화 한다.
     *
     * @return 인증이 통과된 Sheets API client service
     * @throws IOException
     */
    public static Sheets getSheetsService(AuthMode authMode) throws IOException {
        Credential credential = null;
        if (authMode == AuthMode.OAUTH20) {
            credential = getOauth2Authorize();
        } else if (authMode == AuthMode.SERVICE_ACCOUNT) {
            credential = getServiceAccountAuthorize();
        }
        return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }

    public static void main(String[] args) throws IOException {
        // 기호에 따라 OAUTH2.0용 인증이나 서비스 계정으로 인증을 수행 후 Sheet Service 객체를 불러온다.
        // Sheets service = getSheetsService(AuthMode.OAUTH20);
        Sheets service = getSheetsService(AuthMode.SERVICE_ACCOUNT);

        // 아래의 샘플 구글 시트 URL에서 중간의 문자열이 spreed sheet id에 해당한다.
        // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
        String spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";

        String range = "A3:A11";
        ValueRange response = service.spreadsheets().values()
                .get(spreadsheetId, range)
                .execute();
        List<List<Object>> values = response.getValues();
        if (values == null || values.size() == 0) {
            System.out.println("No data found.");
        } else {
            for (List row : values) {
                if (row.size() > 0) {
                    System.out.println(row.get(0).toString());
                }
            }
        }
    }
}
