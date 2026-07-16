import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.GridLayout;
import java.io.FileOutputStream;
import java.net.CookieManager;
import java.net.CookiePolicy;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * NewNetOne DirectEvent Eventlog importer  (standalone Java, JDK 17).
 *
 * Logs in to the portal, submits the Search form for a date range / filters,
 * follows every "Next >>" page, and writes all rows to an .xlsx file.
 *
 * ---- BUILD (Windows, plain javac + POI jars in a lib\ folder) ----
 *   javac -cp "lib/*" NewNetOneEventlog.java
 *
 * ---- RUN ----
 *   java  -cp "lib/*;."  NewNetOneEventlog
 *   (on Linux/Mac use a colon:   -cp "lib/*:.")
 *
 * You need these POI jars in lib\  (download from the Apache POI binary release):
 *   poi-5.x.x.jar
 *   poi-ooxml-5.x.x.jar
 *   poi-ooxml-lite-5.x.x.jar
 *   xmlbeans-5.x.x.jar
 *   commons-collections4-4.x.jar
 *   commons-compress-1.x.jar
 *   commons-io-2.x.jar
 *   log4j-api-2.x.x.jar            (POI logging; program still runs if warnings appear)
 *   SparseBitSet-1.x.jar
 */
public class NewNetOneEventlog {

    // ------------------------- USER SETTINGS -------------------------
    static String USERNAME = "";            // filled in via the login window
    static String PASSWORD = "";

    static String START_DATE = "16-7-2026";   // d-M-yyyy (site format)
    static String END_DATE   = "16-7-2026";
    static String START_TIME = "00:00";
    static String END_TIME   = "23:59";

    static final String SERVER  = "live";   // "live" or "test" (ddlServer)
    static final String SERVICE = "";       // "" = all services (ddlService)

    static final String FILTER_ERROR       = "";
    static final String FILTER_USERNAME    = "";
    static final String FILTER_EVENTLOGID  = "";
    static final String FILTER_SEARCHVALUE = "";

    static final boolean SHOW_INFORMATION = false;
    static final boolean SHOW_WARNING     = true;
    static final boolean SHOW_ERROR       = true;
    static final boolean SHOW_PENDING     = true;

    static final int MAX_PAGES = 20;
    static final String OUTPUT_FILENAME = "Eventlog.xlsx";
    // -----------------------------------------------------------------

    static final String LOGIN_URL    =
        "https://documentation.newnetone.com/Login.aspx?ReturnUrl=%2fEventlog%2fDefault.aspx";
    static final String EVENTLOG_URL =
        "https://documentation.newnetone.com/Eventlog/Default.aspx";

    HttpClient client;

    public static void main(String[] args) throws Exception {
        if (!showLoginDialog()) {
            System.out.println("Cancelled by user.");
            return;
        }
        new NewNetOneEventlog().run();
    }

    /** Shows a GUI window asking for username, password, date range and time range. */
    static boolean showLoginDialog() {
        JTextField tfUser      = new JTextField(USERNAME, 20);
        JPasswordField tfPass  = new JPasswordField(20);
        JTextField tfStartDate = new JTextField(START_DATE, 20);
        JTextField tfEndDate   = new JTextField(END_DATE, 20);
        JTextField tfStartTime = new JTextField(START_TIME, 20);
        JTextField tfEndTime   = new JTextField(END_TIME, 20);

        JPanel panel = new JPanel(new GridLayout(0, 2, 8, 8));
        panel.add(new JLabel("Username:"));               panel.add(tfUser);
        panel.add(new JLabel("Password:"));                panel.add(tfPass);
        panel.add(new JLabel("Start Date (d-M-yyyy):"));   panel.add(tfStartDate);
        panel.add(new JLabel("End Date (d-M-yyyy):"));     panel.add(tfEndDate);
        panel.add(new JLabel("Start Time (HH:mm):"));      panel.add(tfStartTime);
        panel.add(new JLabel("End Time (HH:mm):"));        panel.add(tfEndTime);

        int result = JOptionPane.showConfirmDialog(null, panel,
                "NewNetOne Eventlog - Login & Search Range",
                JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
        if (result != JOptionPane.OK_OPTION) return false;

        USERNAME   = tfUser.getText().trim();
        PASSWORD   = new String(tfPass.getPassword());
        START_DATE = tfStartDate.getText().trim();
        END_DATE   = tfEndDate.getText().trim();
        START_TIME = tfStartTime.getText().trim();
        END_TIME   = tfEndTime.getText().trim();

        if (USERNAME.isEmpty() || PASSWORD.isEmpty()) {
            JOptionPane.showMessageDialog(null, "Username and password are required.",
                    "Missing credentials", JOptionPane.WARNING_MESSAGE);
            return false;
        }
        return true;
    }

    void run() throws Exception {
        // Java's HttpClient carries cookies automatically via a CookieManager
        CookieManager cm = new CookieManager();
        cm.setCookiePolicy(CookiePolicy.ACCEPT_ALL);
        client = HttpClient.newBuilder()
                .cookieHandler(cm)
                .followRedirects(HttpClient.Redirect.NORMAL)
                .connectTimeout(Duration.ofSeconds(30))
                .build();

        // 1. GET login page -> ViewState tokens
        System.out.println("Loading login page...");
        String html = get(LOGIN_URL);
        String vs  = field(html, "__VIEWSTATE");
        String vsg = field(html, "__VIEWSTATEGENERATOR");
        String ev  = field(html, "__EVENTVALIDATION");

        // 2. POST credentials
        System.out.println("Logging in...");
        String body = "__EVENTTARGET=&__EVENTARGUMENT="
                + "&__VIEWSTATE=" + enc(vs)
                + "&__VIEWSTATEGENERATOR=" + enc(vsg)
                + "&__EVENTVALIDATION=" + enc(ev)
                + "&tbUsername=" + enc(USERNAME)
                + "&tbPassword=" + enc(PASSWORD)
                + "&btnLogin=" + enc("Login");
        html = post(LOGIN_URL, body);

        if (html.contains("id=\"tbPassword\"") && !html.contains("gvEventlog")) {
            System.out.println("LOGIN FAILED - check username / password.");
            return;
        }
        System.out.println("Login OK.");

        // 3. GET eventlog page -> fresh tokens
        html = get(EVENTLOG_URL);
        vs  = field(html, "__VIEWSTATE");
        vsg = field(html, "__VIEWSTATEGENERATOR");
        ev  = field(html, "__EVENTVALIDATION");
        try (java.io.FileWriter fw = new java.io.FileWriter("debug_eventlog_form.html", StandardCharsets.UTF_8)) {
            fw.write(html);
        }

        // 4. POST search
        System.out.println("Searching " + START_DATE + " " + START_TIME
                + "  ->  " + END_DATE + " " + END_TIME + " ...");
        body = buildSearchBody(vs, vsg, ev);
        html = post(EVENTLOG_URL, body);

        // 5. parse pages, follow "Next >>"
        showProgressDialog();
        List<String[]> rows = new ArrayList<>();
        int page = 1;
        try {
            while (true) {
                int before = rows.size();
                parseTable(html, rows);
                System.out.println("Page " + page + ": total rows so far = " + rows.size());
                updateProgress(page, rows.size());

                vs  = field(html, "__VIEWSTATE");
                vsg = field(html, "__VIEWSTATEGENERATOR");
                ev  = field(html, "__EVENTVALIDATION");

                if (!html.contains("id=\"ctl00_ContentPlaceHolder1_lbVolgende\"")) break;
                if (page >= MAX_PAGES) break;
                if (rows.size() == before && page > 1) break; // safety: no progress

                body = buildPostbackBody(vs, vsg, ev, "ctl00$ContentPlaceHolder1$lbVolgende");
                html = post(EVENTLOG_URL, body);
                page++;
            }
        } finally {
            closeProgressDialog();
        }

        String outputPath = resolveOutputPath();
        writeXlsx(rows, outputPath);
        System.out.println("Done. Imported " + rows.size()
                + " rows across " + page + " page(s) -> " + outputPath);

        int finalPage = page;
        SwingUtilities.invokeLater(() -> JOptionPane.showMessageDialog(null,
                "Scraped " + rows.size() + " row(s) across " + finalPage + " page(s)!\n"
                        + "Saved to " + outputPath + "\n\n"
                        + "The hamsters have earned a nap. 🐹💤",
                "All done! 🎉", JOptionPane.INFORMATION_MESSAGE));
    }

    /** Resolves ~/Desktop/Eventlog/Eventlog.xlsx, creating the folder if needed. */
    static String resolveOutputPath() {
        java.io.File dir = new java.io.File(System.getProperty("user.home"), "Desktop" + java.io.File.separator + "Eventlog");
        if (!dir.exists()) dir.mkdirs();
        return new java.io.File(dir, OUTPUT_FILENAME).getAbsolutePath();
    }

    // ------------------------- funny progress dialog -------------------------
    JDialog progressDialog;
    JLabel funnyLabel;
    JLabel statsLabel;

    static final String[] FUNNY_MESSAGES = {
        "Herding eventlogs into neat little rows... 🐑",
        "Politely asking the server for more cookies... 🍪",
        "Convincing ASP.NET this is a legit human... 🤖",
        "Clicking 'Next >>' faster than humanly possible... ⚡",
        "Summoning rows from the database void... 👻",
        "Counting bytes like sheep before bedtime... 💤",
        "Bribing the GridView with kindness... 🙏",
        "Making Excel sweat a little... 📊",
        "Reticulating splines... 📐",
        "Almost there, don't jinx it... 🤞",
    };

    void showProgressDialog() {
        SwingUtilities.invokeLater(() -> {
            progressDialog = new JDialog((java.awt.Frame) null, "Scraping in progress...", false);
            JPanel panel = new JPanel(new java.awt.BorderLayout(10, 10));
            panel.setBorder(BorderFactory.createEmptyBorder(15, 20, 15, 20));

            funnyLabel = new JLabel(FUNNY_MESSAGES[0], SwingConstants.CENTER);
            JProgressBar bar = new JProgressBar();
            bar.setIndeterminate(true);
            statsLabel = new JLabel("Page 0 · 0 rows found", SwingConstants.CENTER);

            panel.add(funnyLabel, java.awt.BorderLayout.NORTH);
            panel.add(bar, java.awt.BorderLayout.CENTER);
            panel.add(statsLabel, java.awt.BorderLayout.SOUTH);

            progressDialog.setContentPane(panel);
            progressDialog.setSize(380, 130);
            progressDialog.setLocationRelativeTo(null);
            progressDialog.setDefaultCloseOperation(JDialog.DO_NOTHING_ON_CLOSE);
            progressDialog.setVisible(true);
        });
    }

    void updateProgress(int page, int totalRows) {
        SwingUtilities.invokeLater(() -> {
            if (funnyLabel != null) funnyLabel.setText(FUNNY_MESSAGES[(page - 1) % FUNNY_MESSAGES.length]);
            if (statsLabel != null) statsLabel.setText("Page " + page + " · " + totalRows + " rows found");
        });
    }

    void closeProgressDialog() {
        SwingUtilities.invokeLater(() -> {
            if (progressDialog != null) progressDialog.dispose();
        });
    }

    // ------------------------- request bodies -------------------------
    String commonFilters() {
        StringBuilder s = new StringBuilder();
        s.append("&ctl00%24ContentPlaceHolder1%24ddlServer=").append(enc(SERVER));
        s.append("&ctl00%24ContentPlaceHolder1%24ddlService=").append(enc(SERVICE));
        s.append("&ctl00%24ContentPlaceHolder1%24tbEventlogID=").append(enc(FILTER_EVENTLOGID));
        s.append("&ctl00%24ContentPlaceHolder1%24tbUsername=").append(enc(FILTER_USERNAME));
        s.append("&ctl00%24ContentPlaceHolder1%24tbError=").append(enc(FILTER_ERROR));
        s.append("&ctl00%24ContentPlaceHolder1%24tbWaarde=").append(enc(FILTER_SEARCHVALUE));
        s.append("&ctl00%24ContentPlaceHolder1%24calDatumVan%24Text=").append(enc(START_DATE));
        s.append("&ctl00%24ContentPlaceHolder1%24tbStartTime=").append(enc(START_TIME));
        s.append("&ctl00%24ContentPlaceHolder1%24calDatumTot%24Text=").append(enc(END_DATE));
        s.append("&ctl00%24ContentPlaceHolder1%24tbEndTime=").append(enc(END_TIME));
        if (SHOW_INFORMATION) s.append("&ctl00%24ContentPlaceHolder1%24cblEventtype%240=on");
        if (SHOW_WARNING)     s.append("&ctl00%24ContentPlaceHolder1%24cblEventtype%241=on");
        if (SHOW_ERROR)       s.append("&ctl00%24ContentPlaceHolder1%24cblEventtype%242=on");
        if (SHOW_PENDING)     s.append("&ctl00%24ContentPlaceHolder1%24cblEventtype%243=on");
        return s.toString();
    }

    String buildSearchBody(String vs, String vsg, String ev) {
        return "ctl00_toolkitScriptMaster_HiddenField="
                + "&__EVENTTARGET=&__EVENTARGUMENT=&__LASTFOCUS="
                + "&__VIEWSTATE=" + enc(vs)
                + "&__VIEWSTATEGENERATOR=" + enc(vsg)
                + "&__VIEWSTATEENCRYPTED="
                + "&__EVENTVALIDATION=" + enc(ev)
                + commonFilters()
                + "&ctl00%24ContentPlaceHolder1%24btnSearch=Search";
    }

    String buildPostbackBody(String vs, String vsg, String ev, String target) {
        return "ctl00_toolkitScriptMaster_HiddenField="
                + "&__EVENTTARGET=" + enc(target)
                + "&__EVENTARGUMENT=&__LASTFOCUS="
                + "&__VIEWSTATE=" + enc(vs)
                + "&__VIEWSTATEGENERATOR=" + enc(vsg)
                + "&__VIEWSTATEENCRYPTED="
                + "&__EVENTVALIDATION=" + enc(ev)
                + commonFilters();
    }

    // ------------------------- HTML table parsing -------------------------
    static int debugDumpCount = 0;

    void parseTable(String html, List<String[]> rows) {
        String lower = html.toLowerCase();
        int marker = lower.indexOf("gveventlog");
        if (marker < 0) {
            System.out.println("WARNING: 'gvEventlog' not found in server response - can't locate results table.");
            dumpDebugHtml(html);
            return;
        }
        int tableStart = lower.lastIndexOf("<table", marker);
        if (tableStart < 0) {
            System.out.println("WARNING: found 'gvEventlog' but no enclosing <table> before it.");
            dumpDebugHtml(html);
            return;
        }
        int tableEnd = findMatchingTableEnd(html, tableStart);
        String tbl = html.substring(tableStart, tableEnd);

        Matcher trM = Pattern.compile("<tr[^>]*>(.*?)</tr>",
                Pattern.DOTALL | Pattern.CASE_INSENSITIVE).matcher(tbl);
        Map<String, Integer> col = null;
        int rowsFromThisPage = 0;

        while (trM.find()) {
            String tr = trM.group(1);
            List<String> cells = new ArrayList<>();
            Matcher cellM = Pattern.compile("<t[dh][^>]*>(.*?)</t[dh]>",
                    Pattern.DOTALL | Pattern.CASE_INSENSITIVE).matcher(tr);
            while (cellM.find()) cells.add(clean(cellM.group(1)));
            if (cells.isEmpty()) continue;

            if (col == null) {
                // first non-empty row = header; map column names instead of assuming fixed positions
                col = new HashMap<>();
                for (int i = 0; i < cells.size(); i++) {
                    String h = cells.get(i).toLowerCase();
                    if (h.contains("vendor"))        col.put("vendor", i);
                    else if (h.contains("service"))   col.put("service", i);
                    else if (h.contains("error"))      col.put("error", i);
                    else if (h.contains("username"))  col.put("username", i);
                    else if (h.contains("response"))  col.put("response", i);
                    else if (h.contains("created"))   col.put("created", i);
                }
                continue;
            }

            if (col.size() < 6) continue;          // header wasn't recognized - bail out on data rows
            if (cells.size() < col.size()) continue; // pager/footer row, too few cells

            rows.add(new String[]{
                    cell(cells, col.get("vendor")), cell(cells, col.get("service")),
                    cell(cells, col.get("error")), cell(cells, col.get("username")),
                    cell(cells, col.get("response")), cell(cells, col.get("created"))
            });
            rowsFromThisPage++;
        }

        if (col == null || col.size() < 6) {
            System.out.println("WARNING: results table found but header columns didn't match "
                    + "(vendor/service/error/username/response/created).");
            dumpDebugHtml(html);
        } else if (rowsFromThisPage == 0) {
            System.out.println("NOTE: header recognized but 0 data rows on this page.");
        }
    }

    static String cell(List<String> cells, Integer idx) {
        return (idx == null || idx >= cells.size()) ? "" : cells.get(idx);
    }

    /** Finds the end of the <table> that starts at tableStart, accounting for nested tables. */
    static int findMatchingTableEnd(String html, int tableStart) {
        Matcher tm = Pattern.compile("<table[^>]*>|</table>", Pattern.CASE_INSENSITIVE).matcher(html);
        tm.region(tableStart, html.length());
        int depth = 0;
        while (tm.find()) {
            if (tm.group().charAt(1) == '/') {
                depth--;
                if (depth == 0) return tm.end();
            } else {
                depth++;
            }
        }
        return html.length();
    }

    /** Saves the raw server response to disk so parsing failures can be diagnosed. */
    static void dumpDebugHtml(String html) {
        if (debugDumpCount >= 3) return; // don't spam the folder
        debugDumpCount++;
        String name = "debug_page" + debugDumpCount + ".html";
        try (java.io.FileWriter fw = new java.io.FileWriter(name, StandardCharsets.UTF_8)) {
            fw.write(html);
            System.out.println("Saved raw response to " + name + " for inspection.");
        } catch (Exception ignored) {}
    }

    static String clean(String s) {
        s = s.replaceAll("(?s)<[^>]*>", " ");        // strip tags
        s = s.replace("&nbsp;", " ").replace("&amp;", "&")
             .replace("&lt;", "<").replace("&gt;", ">")
             .replace("&quot;", "\"").replace("&#39;", "'");
        s = s.replaceAll("\\s+", " ").trim();
        return s;
    }

    // ------------------------- xlsx output -------------------------
    static final String[] HEADERS = {"Vendor", "Service", "Error", "Username",
                                      "ResponseTimeMs", "Created"};

    void writeXlsx(List<String[]> rows, String outputPath) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
            CellStyle bold = wb.createCellStyle();
            Font f = wb.createFont(); f.setBold(true); bold.setFont(f);

            writeSheet(wb, "Eventlog", rows, bold);

            // group rows by the first few words of the Error text, so similar
            // errors (same prefix, different trailing detail) share a tab
            Map<String, List<String[]>> byError = new LinkedHashMap<>();
            for (String[] row : rows) {
                String key = errorGroupKey(row[2]);
                byError.computeIfAbsent(key, k -> new ArrayList<>()).add(row);
            }

            // groups that only occur once are one-off errors - collect them
            // together on a single "Unique Errors" tab instead of one tab each
            List<String[]> uniqueErrors = new ArrayList<>();
            List<Map.Entry<String, List<String[]>>> repeatedGroups = new ArrayList<>();
            for (Map.Entry<String, List<String[]>> e : byError.entrySet()) {
                if (e.getKey().equals("No Error") || e.getValue().size() > 1) {
                    repeatedGroups.add(e);
                } else {
                    uniqueErrors.addAll(e.getValue());
                }
            }

            Set<String> usedNames = new HashSet<>();
            usedNames.add("Eventlog");
            if (!byError.isEmpty()) {
                Sheet sh = wb.createSheet(uniqueSheetName("Unique Errors", usedNames));
                int r = writeErrorCountSummary(sh, byError, bold);
                r++; // blank spacer row
                writeSheetBody(sh, r, uniqueErrors, bold);
            }
            for (Map.Entry<String, List<String[]>> e : repeatedGroups) {
                writeSheet(wb, uniqueSheetName(e.getKey(), usedNames), e.getValue(), bold);
            }

            try (FileOutputStream out = new FileOutputStream(outputPath)) {
                wb.write(out);
            }
        }
    }

    static final int ERROR_GROUP_WORDS = 3;

    /** Groups similar errors by their first few words (e.g. "[Push] InputInvalid:"). */
    static String errorGroupKey(String err) {
        if (err.isEmpty()) return "No Error";
        String[] words = err.trim().split("\\s+");
        int n = Math.min(ERROR_GROUP_WORDS, words.length);
        return String.join(" ", Arrays.copyOfRange(words, 0, n));
    }

    void writeSheet(Workbook wb, String name, List<String[]> rows, CellStyle bold) {
        Sheet sh = wb.createSheet(name);
        writeSheetBody(sh, 0, rows, bold);
    }

    /** Writes the header row + data rows starting at sheet row startRow. */
    void writeSheetBody(Sheet sh, int startRow, List<String[]> rows, CellStyle bold) {
        Row hr = sh.createRow(startRow);
        for (int c = 0; c < HEADERS.length; c++) {
            Cell cell = hr.createCell(c);
            cell.setCellValue(HEADERS[c]);
            cell.setCellStyle(bold);
        }
        int r = startRow + 1;
        for (String[] row : rows) {
            Row xr = sh.createRow(r++);
            for (int c = 0; c < row.length; c++) xr.createCell(c).setCellValue(row[c]);
        }
        for (int c = 0; c < HEADERS.length; c++) sh.autoSizeColumn(c);
    }

    /** Writes an "Error Type | Count" summary table at the top of the sheet. Returns next free row. */
    int writeErrorCountSummary(Sheet sh, Map<String, List<String[]>> byError, CellStyle bold) {
        Row hr = sh.createRow(0);
        Cell c0 = hr.createCell(0); c0.setCellValue("Error Type"); c0.setCellStyle(bold);
        Cell c1 = hr.createCell(1); c1.setCellValue("Count");     c1.setCellStyle(bold);
        int r = 1;
        for (Map.Entry<String, List<String[]>> e : byError.entrySet()) {
            Row row = sh.createRow(r++);
            row.createCell(0).setCellValue(e.getKey());
            row.createCell(1).setCellValue(e.getValue().size());
        }
        sh.autoSizeColumn(0);
        sh.autoSizeColumn(1);
        return r;
    }

    /** Excel sheet names: <=31 chars, no \/?*[]:, unique within the workbook. */
    static String uniqueSheetName(String raw, Set<String> used) {
        String name = raw.replaceAll("[\\\\/?*\\[\\]:]", " ").trim();
        if (name.isEmpty()) name = "Error";
        if (name.length() > 31) name = name.substring(0, 31);
        String base = name;
        int n = 2;
        while (used.contains(name)) {
            String suffix = " (" + n + ")";
            name = base.substring(0, Math.min(base.length(), 31 - suffix.length())) + suffix;
            n++;
        }
        used.add(name);
        return name;
    }

    // ------------------------- http + helpers -------------------------
    String get(String url) throws Exception {
        HttpRequest req = HttpRequest.newBuilder(URI.create(url))
                .header("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
                .GET().build();
        return client.send(req, HttpResponse.BodyHandlers.ofString(StandardCharsets.UTF_8)).body();
    }

    String post(String url, String body) throws Exception {
        HttpRequest req = HttpRequest.newBuilder(URI.create(url))
                .header("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
                .header("Content-Type", "application/x-www-form-urlencoded")
                .POST(HttpRequest.BodyPublishers.ofString(body, StandardCharsets.UTF_8))
                .build();
        return client.send(req, HttpResponse.BodyHandlers.ofString(StandardCharsets.UTF_8)).body();
    }

    /** Extract the value="" of an <input> identified by id (or name). */
    static String field(String html, String id) {
        Matcher m = Pattern.compile(
                "(?:id|name)=\"" + Pattern.quote(id) + "\"[^>]*?value=\"([^\"]*)\"",
                Pattern.CASE_INSENSITIVE | Pattern.DOTALL).matcher(html);
        if (m.find()) return m.group(1);
        // value attribute sometimes precedes id
        m = Pattern.compile(
                "value=\"([^\"]*)\"[^>]*?(?:id|name)=\"" + Pattern.quote(id) + "\"",
                Pattern.CASE_INSENSITIVE | Pattern.DOTALL).matcher(html);
        return m.find() ? m.group(1) : "";
    }

    static String enc(String s) {
        return java.net.URLEncoder.encode(s == null ? "" : s, StandardCharsets.UTF_8);
    }
}
