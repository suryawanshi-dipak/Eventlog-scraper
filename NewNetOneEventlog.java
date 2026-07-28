import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.BasicStroke;
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Frame;
import java.awt.GradientPaint;
import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.RenderingHints;
import java.awt.event.KeyEvent;
import java.awt.geom.Arc2D;
import java.awt.geom.Ellipse2D;
import java.awt.geom.QuadCurve2D;
import java.awt.geom.RoundRectangle2D;
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

    static final java.time.format.DateTimeFormatter SITE_DATE_FMT =
        java.time.format.DateTimeFormatter.ofPattern("d-M-yyyy");
    static String START_DATE = java.time.LocalDate.now().format(SITE_DATE_FMT);   // d-M-yyyy (site format)
    static String END_DATE   = java.time.LocalDate.now().format(SITE_DATE_FMT);
    static String START_TIME = "00:00";
    static String END_TIME   = "23:59";

    static final String SERVER  = "live";   // "live" or "test" (ddlServer)
    static final String SERVICE = "";       // "" = all services (ddlService)

    static String FILTER_ERROR             = "";   // filled in via the login window
    static final String FILTER_USERNAME    = "";
    static final String FILTER_EVENTLOGID  = "";
    static String FILTER_SEARCHVALUE       = "";   // filled in via the login window

    static boolean SHOW_INFORMATION = false;   // filled in via the login window
    static boolean SHOW_WARNING     = true;
    static boolean SHOW_ERROR       = true;
    static boolean SHOW_PENDING     = true;

    static final int MAX_PAGES = Integer.MAX_VALUE; // no cap - follow "Next >>" until it disappears
    static final String OUTPUT_FILENAME = "Eventlog.xlsx";
    // -----------------------------------------------------------------

    static final String LOGIN_URL    =
        "https://documentation.newnetone.com/Login.aspx?ReturnUrl=%2fEventlog%2fDefault.aspx";
    static final String EVENTLOG_URL =
        "https://documentation.newnetone.com/Eventlog/Default.aspx";
    static final String EVENTLOG_DETAIL_URL_BASE =
        "https://documentation.newnetone.com/Eventlog/Eventlog.aspx?ID=";

    HttpClient client;

    public static void main(String[] args) throws Exception {
        initLogger();
        log("=== NewNetOneEventlog starting (pid=" + ProcessHandle.current().pid() + ") ===");

        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception e) {
            log("Look and feel setup failed (non-fatal): " + e);
        }

        try {
            if (!showLoginDialog()) {
                log("Cancelled by user at login dialog.");
                return;
            }
            log("Login dialog OK. user=" + USERNAME + " range=" + START_DATE + " " + START_TIME
                    + " -> " + END_DATE + " " + END_TIME);
            new NewNetOneEventlog().run();
            log("=== Finished normally ===");
        } catch (Throwable t) {
            java.io.StringWriter sw = new java.io.StringWriter();
            t.printStackTrace(new java.io.PrintWriter(sw));
            log("FATAL ERROR: " + t);
            log(sw.toString());
            final String msg = t.getClass().getSimpleName() + ": " + t.getMessage();
            final String logPath = LOG_PATH;
            SwingUtilities.invokeLater(() -> JOptionPane.showMessageDialog(null,
                    "Something went wrong:\n" + msg + "\n\nFull details were saved to:\n" + logPath,
                    "Error", JOptionPane.ERROR_MESSAGE));
        } finally {
            closeLogger();
        }
    }

    // ------------------------- logging -------------------------
    static String LOG_PATH;
    static java.io.PrintWriter logWriter;

    static java.io.File eventlogDir() {
        java.io.File dir = new java.io.File(System.getProperty("user.home"),
                "Desktop" + java.io.File.separator + "Eventlog");
        if (!dir.exists()) dir.mkdirs();
        return dir;
    }

    static void initLogger() {
        try {
            LOG_PATH = new java.io.File(eventlogDir(), "app_log.txt").getAbsolutePath();
            logWriter = new java.io.PrintWriter(new java.io.FileWriter(LOG_PATH, false), true);
        } catch (Exception e) {
            logWriter = null; // fall back to console-only logging
        }
    }

    static void log(String msg) {
        String line = "[" + java.time.LocalTime.now().withNano(0) + "] " + msg;
        System.out.println(line);
        if (logWriter != null) logWriter.println(line);
    }

    static void closeLogger() {
        if (logWriter != null) logWriter.close();
    }

    static final Color BRAND_ORANGE     = new Color(230, 126, 34);
    static final Color BRAND_BLUE_LIGHT = new Color(74, 122, 255);
    static final Color BRAND_BLUE_DARK  = new Color(20, 40, 200);
    static final Color BRAND_DARK   = new Color(60, 60, 60);
    static final Color BRAND_MUTED  = new Color(120, 120, 120);
    static final Color PANEL_WHITE  = Color.WHITE;

    static java.awt.Font uiFont(int style, int size) {
        return new java.awt.Font("Segoe UI", style, size);
    }

    /** Shows a styled GUI window asking for username, password, date range and time range. */
    static boolean showLoginDialog() {
        JTextField tfUser      = new JTextField(USERNAME, 18);
        JPasswordField tfPass  = new JPasswordField(18);
        JTextField tfStartDate = new JTextField(START_DATE, 18);
        JTextField tfEndDate   = new JTextField(END_DATE, 18);
        JTextField tfStartTime = new JTextField(START_TIME, 18);
        JTextField tfEndTime   = new JTextField(END_TIME, 18);
        JTextField tfErrorFilter = new JTextField(FILTER_ERROR, 18);
        JTextField tfSearchValue = new JTextField(FILTER_SEARCHVALUE, 18);
        JCheckBox cbInformation = new JCheckBox("Information", SHOW_INFORMATION);
        JCheckBox cbWarning     = new JCheckBox("Warning", SHOW_WARNING);
        JCheckBox cbError       = new JCheckBox("Error", SHOW_ERROR);
        JCheckBox cbPending     = new JCheckBox("Pending", SHOW_PENDING);

        JDialog dialog = new JDialog((Frame) null, "NewNetOne Eventlog", true);

        JPanel content = new JPanel(new BorderLayout(0, 14));
        content.setBackground(PANEL_WHITE);
        content.setBorder(BorderFactory.createEmptyBorder(20, 24, 16, 24));

        JLabel title = new JLabel("🐯  NewNetOne Eventlog Scraper");
        title.setFont(uiFont(java.awt.Font.BOLD, 18));
        title.setForeground(BRAND_DARK);
        JLabel subtitle = new JLabel("Sign in and choose a search range");
        subtitle.setFont(uiFont(java.awt.Font.PLAIN, 12));
        subtitle.setForeground(BRAND_MUTED);

        JPanel header = new JPanel();
        header.setLayout(new BoxLayout(header, BoxLayout.Y_AXIS));
        header.setBackground(PANEL_WHITE);
        header.add(title);
        header.add(Box.createVerticalStrut(4));
        header.add(subtitle);
        header.add(Box.createVerticalStrut(6));
        header.add(new JSeparator());

        JPanel form = new JPanel(new GridBagLayout());
        form.setBackground(PANEL_WHITE);
        GridBagConstraints gc = new GridBagConstraints();
        gc.insets = new Insets(6, 4, 6, 4);
        gc.fill = GridBagConstraints.HORIZONTAL;

        int row = 0;
        addFormRow(form, gc, row++, "Username", tfUser);
        addFormRow(form, gc, row++, "Password", tfPass);
        addFormRow(form, gc, row++, "Start Date (d-M-yyyy)", tfStartDate);
        addFormRow(form, gc, row++, "End Date (d-M-yyyy)", tfEndDate);
        addFormRow(form, gc, row++, "Start Time (HH:mm)", tfStartTime);
        addFormRow(form, gc, row++, "End Time (HH:mm)", tfEndTime);
        addFormRow(form, gc, row++, "Error", tfErrorFilter);
        addFormRow(form, gc, row++, "Search value", tfSearchValue);

        JPanel eventTypes = new JPanel(new FlowLayout(FlowLayout.LEFT, 12, 0));
        eventTypes.setBackground(PANEL_WHITE);
        for (JCheckBox cb : new JCheckBox[]{cbInformation, cbWarning, cbError, cbPending}) {
            cb.setFont(uiFont(java.awt.Font.PLAIN, 13));
            cb.setBackground(PANEL_WHITE);
            eventTypes.add(cb);
        }
        gc.gridx = 0; gc.gridy = row++; gc.gridwidth = 2;
        form.add(eventTypes, gc);
        gc.gridwidth = 1;

        JButton cancelBtn = new JButton("Cancel");
        styleButton(cancelBtn, new Color(235, 235, 235), BRAND_DARK);

        JButton okBtn = new GradientButton("Start Scraping", BRAND_BLUE_LIGHT, BRAND_BLUE_DARK);

        JPanel buttons = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 0));
        buttons.setBackground(PANEL_WHITE);
        buttons.add(cancelBtn);
        buttons.add(okBtn);

        content.add(header, BorderLayout.NORTH);
        content.add(form, BorderLayout.CENTER);
        content.add(buttons, BorderLayout.SOUTH);

        dialog.setContentPane(content);
        dialog.getRootPane().setDefaultButton(okBtn);
        dialog.setResizable(false);

        boolean[] okPressed = {false};
        okBtn.addActionListener(e -> { okPressed[0] = true; dialog.dispose(); });
        cancelBtn.addActionListener(e -> dialog.dispose());
        dialog.getRootPane().registerKeyboardAction(e -> dialog.dispose(),
                KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0), JComponent.WHEN_IN_FOCUSED_WINDOW);

        dialog.pack();
        dialog.setLocationRelativeTo(null);
        dialog.setVisible(true); // blocks (modal) until dispose()

        if (!okPressed[0]) return false;

        USERNAME     = tfUser.getText().trim();
        PASSWORD     = new String(tfPass.getPassword());
        START_DATE   = tfStartDate.getText().trim();
        END_DATE     = tfEndDate.getText().trim();
        START_TIME   = tfStartTime.getText().trim();
        END_TIME     = tfEndTime.getText().trim();
        FILTER_ERROR       = tfErrorFilter.getText().trim();
        FILTER_SEARCHVALUE = tfSearchValue.getText().trim();
        SHOW_INFORMATION = cbInformation.isSelected();
        SHOW_WARNING     = cbWarning.isSelected();
        SHOW_ERROR       = cbError.isSelected();
        SHOW_PENDING     = cbPending.isSelected();

        if (USERNAME.isEmpty() || PASSWORD.isEmpty()) {
            JOptionPane.showMessageDialog(null, "Username and password are required.",
                    "Missing credentials", JOptionPane.WARNING_MESSAGE);
            return false;
        }
        return true;
    }

    static void addFormRow(JPanel form, GridBagConstraints gc, int row, String label, JComponent field) {
        gc.gridx = 0; gc.gridy = row; gc.weightx = 0;
        JLabel l = new JLabel(label);
        l.setFont(uiFont(java.awt.Font.PLAIN, 13));
        l.setForeground(BRAND_DARK);
        form.add(l, gc);

        gc.gridx = 1; gc.weightx = 1;
        field.setFont(uiFont(java.awt.Font.PLAIN, 13));
        form.add(field, gc);
    }

    static void styleButton(JButton btn, Color bg, Color fg) {
        btn.setFont(uiFont(java.awt.Font.BOLD, 13));
        btn.setBackground(bg);
        btn.setForeground(fg);
        btn.setFocusPainted(false);
        btn.setOpaque(true);
        btn.setBorder(BorderFactory.createEmptyBorder(8, 18, 8, 18));
    }

    /** A button painted with a diagonal color gradient instead of a flat background. */
    static class GradientButton extends JButton {
        private final Color from, to;

        GradientButton(String text, Color from, Color to) {
            super(text);
            this.from = from;
            this.to = to;
            setContentAreaFilled(false);
            setBorderPainted(false);
            setFocusPainted(false);
            setOpaque(false);
            setForeground(Color.WHITE);
            setFont(uiFont(java.awt.Font.BOLD, 13));
            setBorder(BorderFactory.createEmptyBorder(8, 22, 8, 22));
        }

        @Override
        protected void paintComponent(Graphics g) {
            Graphics2D g2 = (Graphics2D) g.create();
            g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            g2.setPaint(new GradientPaint(0, 0, from, getWidth(), getHeight(), to));
            g2.fill(new RoundRectangle2D.Float(0, 0, getWidth(), getHeight(), 10, 10));
            g2.dispose();
            super.paintComponent(g);
        }
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
        log("Loading login page...");
        String html = get(LOGIN_URL);
        String vs  = field(html, "__VIEWSTATE");
        String vsg = field(html, "__VIEWSTATEGENERATOR");
        String ev  = field(html, "__EVENTVALIDATION");
        log("Login page tokens: VIEWSTATE=" + vs.length() + " chars, GENERATOR=" + vsg.length()
                + " chars, EVENTVALIDATION=" + ev.length() + " chars");

        // 2. POST credentials
        log("Logging in as " + USERNAME + " ...");
        String body = "__EVENTTARGET=&__EVENTARGUMENT="
                + "&__VIEWSTATE=" + enc(vs)
                + "&__VIEWSTATEGENERATOR=" + enc(vsg)
                + "&__EVENTVALIDATION=" + enc(ev)
                + "&tbUsername=" + enc(USERNAME)
                + "&tbPassword=" + enc(PASSWORD)
                + "&btnLogin=" + enc("Login");
        html = post(LOGIN_URL, body);

        if (html.contains("id=\"tbPassword\"") && !html.contains("gvEventlog")) {
            log("LOGIN FAILED - server returned the login form again (check username / password).");
            SwingUtilities.invokeLater(() -> JOptionPane.showMessageDialog(null,
                    "Login failed. Please check your username and password and try again.",
                    "Login Failed", JOptionPane.ERROR_MESSAGE));
            return;
        }
        log("Login OK.");

        // 3. GET eventlog page -> fresh tokens
        html = get(EVENTLOG_URL);
        vs  = field(html, "__VIEWSTATE");
        vsg = field(html, "__VIEWSTATEGENERATOR");
        ev  = field(html, "__EVENTVALIDATION");
        saveDebugFile("debug_eventlog_form.html", html);

        // 4. POST search
        log("Searching " + START_DATE + " " + START_TIME + "  ->  " + END_DATE + " " + END_TIME + " ...");
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
                log("Page " + page + ": total rows so far = " + rows.size());
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
        log("Done. Imported " + rows.size() + " rows across " + page + " page(s) -> " + outputPath);

        int finalPage = page;
        SwingUtilities.invokeLater(() -> JOptionPane.showMessageDialog(null,
                "Scraped " + rows.size() + " row(s) across " + finalPage + " page(s)!\n"
                        + "Saved to " + outputPath + "\n\n"
                        + "The tiger is full and taking a nap. 🐯💤",
                "All done! 🎉", JOptionPane.INFORMATION_MESSAGE));
    }

    /** Resolves ~/Desktop/Eventlog/Eventlog_<today's date>.xlsx, creating the folder if needed. */
    static String resolveOutputPath() {
        java.io.File dir = new java.io.File(System.getProperty("user.home"), "Desktop" + java.io.File.separator + "Eventlog");
        if (!dir.exists()) dir.mkdirs();
        String today = java.time.LocalDate.now().format(SITE_DATE_FMT);
        String dated = OUTPUT_FILENAME.replaceFirst("\\.xlsx$", "_" + today + ".xlsx");
        return new java.io.File(dir, dated).getAbsolutePath();
    }

    // ------------------------- tiger progress dialog -------------------------
    JDialog progressDialog;
    JLabel funnyLabel;
    JLabel statsLabel;
    TigerEatingPanel tigerPanel;

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
            progressDialog = new JDialog((Frame) null, "Scraping in progress...", false);

            JPanel panel = new JPanel(new BorderLayout(10, 10));
            panel.setBackground(PANEL_WHITE);
            panel.setBorder(BorderFactory.createEmptyBorder(16, 20, 16, 20));

            JLabel title = new JLabel("Feeding the tiger your eventlog rows...", SwingConstants.CENTER);
            title.setFont(uiFont(java.awt.Font.BOLD, 14));
            title.setForeground(BRAND_DARK);

            tigerPanel = new TigerEatingPanel();

            funnyLabel = new JLabel(FUNNY_MESSAGES[0], SwingConstants.CENTER);
            funnyLabel.setFont(uiFont(java.awt.Font.PLAIN, 12));
            funnyLabel.setForeground(BRAND_MUTED);

            statsLabel = new JLabel("Page 0 · 0 rows found", SwingConstants.CENTER);
            statsLabel.setFont(uiFont(java.awt.Font.BOLD, 12));
            statsLabel.setForeground(BRAND_DARK);

            JPanel south = new JPanel();
            south.setLayout(new BoxLayout(south, BoxLayout.Y_AXIS));
            south.setBackground(PANEL_WHITE);
            funnyLabel.setAlignmentX(0.5f);
            statsLabel.setAlignmentX(0.5f);
            south.add(funnyLabel);
            south.add(Box.createVerticalStrut(4));
            south.add(statsLabel);

            panel.add(title, BorderLayout.NORTH);
            panel.add(tigerPanel, BorderLayout.CENTER);
            panel.add(south, BorderLayout.SOUTH);

            progressDialog.setContentPane(panel);
            progressDialog.setSize(440, 250);
            progressDialog.setLocationRelativeTo(null);
            progressDialog.setResizable(false);
            progressDialog.setDefaultCloseOperation(JDialog.DO_NOTHING_ON_CLOSE);
            progressDialog.setVisible(true);
        });
    }

    static final int TIGER_CYCLE_PAGES = 5; // pages per lap - total page count is unbounded now

    void updateProgress(int page, int totalRows) {
        SwingUtilities.invokeLater(() -> {
            if (funnyLabel != null) funnyLabel.setText(FUNNY_MESSAGES[(page - 1) % FUNNY_MESSAGES.length]);
            if (statsLabel != null) statsLabel.setText("Page " + page + " · " + totalRows + " rows found");
            if (tigerPanel != null) {
                double lap = ((page - 1) % TIGER_CYCLE_PAGES + 1) / (double) TIGER_CYCLE_PAGES;
                tigerPanel.setProgress(lap);
            }
        });
    }

    void closeProgressDialog() {
        SwingUtilities.invokeLater(() -> {
            if (tigerPanel != null) tigerPanel.setProgress(1.0);
            if (progressDialog != null) progressDialog.dispose();
            if (tigerPanel != null) tigerPanel.stopAnimation();
        });
    }

    /** A tiger that strolls across the panel eating a row of meat as scraping progresses. */
    static class TigerEatingPanel extends JPanel {
        private static final int NUM_MEATS = 8;
        private volatile double targetFraction = 0.0;
        private double currentFraction = 0.0;
        private int frame = 0;
        private final Timer timer;

        TigerEatingPanel() {
            setPreferredSize(new Dimension(400, 120));
            setOpaque(true);
            setBackground(new Color(235, 245, 232));
            timer = new Timer(45, e -> {
                currentFraction += (targetFraction - currentFraction) * 0.12;
                frame++;
                repaint();
            });
            timer.start();
        }

        void setProgress(double fraction) {
            targetFraction = Math.max(0.0, Math.min(1.0, fraction));
        }

        void stopAnimation() {
            timer.stop();
        }

        @Override
        protected void paintComponent(Graphics g) {
            super.paintComponent(g);
            Graphics2D g2 = (Graphics2D) g;
            g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);

            int w = getWidth(), h = getHeight();
            int margin = 34;
            int baseline = h - 34;

            g2.setPaint(new GradientPaint(0, 0, new Color(214, 235, 250), 0, h, new Color(235, 248, 230)));
            g2.fillRect(0, 0, w, h);

            g2.setColor(new Color(196, 168, 118));
            g2.fillRect(0, baseline + 20, w, Math.max(0, h - (baseline + 20)));

            int usableW = Math.max(1, w - 2 * margin);
            for (int i = 0; i < NUM_MEATS; i++) {
                double frac = (i + 0.5) / NUM_MEATS;
                int x = margin + (int) (frac * usableW);
                boolean eaten = frac < currentFraction - 0.02;
                drawMeat(g2, x, baseline, eaten);
            }

            int tigerX = margin + (int) (currentFraction * usableW);
            drawTiger(g2, tigerX, baseline, frame);
        }

        private void drawMeat(Graphics2D g2, int cx, int baseline, boolean eaten) {
            int y = baseline - 6;
            if (eaten) {
                g2.setColor(new Color(160, 160, 160));
                g2.drawOval(cx - 5, y + 2, 10, 5);
                return;
            }
            g2.setColor(new Color(235, 235, 235));
            g2.fillRoundRect(cx - 3, y + 6, 6, 10, 4, 4);
            g2.setColor(new Color(178, 92, 48));
            g2.fillOval(cx - 11, y - 8, 22, 18);
            g2.setColor(new Color(140, 62, 30));
            g2.drawOval(cx - 11, y - 8, 22, 18);
        }

        private void drawTiger(Graphics2D g2, int cx, int baseline, int frame) {
            int bodyY = baseline - 22;
            double bounce = Math.sin(frame * 0.35) * 2;
            int by = (int) (bodyY + bounce);
            Color furOrange = new Color(240, 150, 40);
            Color stripe = new Color(60, 40, 20);

            // tail
            g2.setStroke(new BasicStroke(4, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));
            g2.setColor(furOrange);
            double tailWag = Math.sin(frame * 0.4) * 10;
            g2.draw(new QuadCurve2D.Double(cx - 18, by + 6, cx - 30, by - 10 + tailWag, cx - 34, by - 22 + tailWag));

            // legs (alternate for a walking effect)
            g2.setStroke(new BasicStroke(5, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));
            int legOffset = (frame / 3) % 2 == 0 ? 2 : -2;
            g2.drawLine(cx - 10, by + 14, cx - 10 + legOffset, by + 24);
            g2.drawLine(cx + 8, by + 14, cx + 8 - legOffset, by + 24);

            // body
            g2.setColor(furOrange);
            g2.fill(new RoundRectangle2D.Double(cx - 16, by - 6, 34, 22, 16, 16));
            g2.setColor(stripe);
            g2.setStroke(new BasicStroke(2));
            for (int i = 0; i < 3; i++) g2.drawLine(cx - 8 + i * 8, by - 4, cx - 6 + i * 8, by + 12);

            // head
            int hx = cx + 16, hy = by - 6;
            g2.setColor(furOrange);
            g2.fill(new Ellipse2D.Double(hx - 12, hy - 12, 24, 22));

            // ears
            g2.fillOval(hx - 11, hy - 18, 9, 9);
            g2.fillOval(hx + 2, hy - 18, 9, 9);
            g2.setColor(stripe);
            g2.fillOval(hx - 8, hy - 15, 4, 4);
            g2.fillOval(hx + 5, hy - 15, 4, 4);

            // eyes
            g2.setColor(Color.WHITE);
            g2.fillOval(hx - 6, hy - 4, 5, 5);
            g2.fillOval(hx + 3, hy - 4, 5, 5);
            g2.setColor(Color.BLACK);
            g2.fillOval(hx - 5, hy - 3, 2, 2);
            g2.fillOval(hx + 4, hy - 3, 2, 2);

            // chewing mouth
            double open = (Math.sin(frame * 0.6) + 1) / 2.0 * 8 + 2;
            g2.setColor(new Color(120, 20, 20));
            g2.fill(new Arc2D.Double(hx - 2, hy + 2, 16, open, 200, 140, Arc2D.CHORD));

            // face stripes
            g2.setColor(stripe);
            g2.setStroke(new BasicStroke(1.5f));
            g2.drawLine(hx - 8, hy - 8, hx - 4, hy - 2);
            g2.drawLine(hx + 2, hy - 8, hx + 6, hy - 2);
        }
    }

    // ------------------------- request bodies -------------------------
    String commonFilters() {
        StringBuilder s = new StringBuilder();
        s.append("&ctl00%24ContentPlaceHolder1%24ddlServer=").append(enc(SERVER));
        s.append("&ctl00%24ContentPlaceHolder1%24ddlService=").append(enc(SERVICE));
        s.append("&ctl00%24ContentPlaceHolder1%24tbEventlogID=").append(enc(FILTER_EVENTLOGID));
        s.append("&ctl00%24ContentPlaceHolder1%24tbUsername=").append(enc(FILTER_USERNAME));
        s.append("&ctl00%24ContentPlaceHolder1%24tbError=").append(enc(FILTER_ERROR));
        if (!FILTER_SEARCHVALUE.isEmpty()) {
            s.append("&ctl00%24ContentPlaceHolder1%24tbWaarde=").append(enc(FILTER_SEARCHVALUE));
        }
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
        return buildPostbackBody(vs, vsg, ev, target, "");
    }

    String buildPostbackBody(String vs, String vsg, String ev, String target, String argument) {
        return "ctl00_toolkitScriptMaster_HiddenField="
                + "&__EVENTTARGET=" + enc(target)
                + "&__EVENTARGUMENT=" + enc(argument) + "&__LASTFOCUS="
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
            log("WARNING: 'gvEventlog' not found in server response - can't locate results table.");
            dumpDebugHtml(html);
            return;
        }
        int tableStart = lower.lastIndexOf("<table", marker);
        if (tableStart < 0) {
            log("WARNING: found 'gvEventlog' but no enclosing <table> before it.");
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
                    cell(cells, 0), // leftmost, unlabeled column holds the row's numeric event ID
                    cell(cells, col.get("vendor")), cell(cells, col.get("service")),
                    cell(cells, col.get("error")), cell(cells, col.get("username")),
                    cell(cells, col.get("response")), cell(cells, col.get("created"))
            });
            rowsFromThisPage++;
        }

        if (col == null || col.size() < 6) {
            log("WARNING: results table found but header columns didn't match "
                    + "(vendor/service/error/username/response/created).");
            dumpDebugHtml(html);
        } else if (rowsFromThisPage == 0) {
            log("NOTE: header recognized but 0 data rows on this page.");
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
        saveDebugFile("debug_page" + debugDumpCount + ".html", html);
    }

    /** Writes a diagnostic file into the (always-writable) Eventlog output folder - never fatal. */
    static void saveDebugFile(String name, String content) {
        java.io.File f = new java.io.File(eventlogDir(), name);
        try (java.io.FileWriter fw = new java.io.FileWriter(f, StandardCharsets.UTF_8)) {
            fw.write(content);
            log("Saved " + f.getAbsolutePath() + " for inspection.");
        } catch (Exception e) {
            log("Could not write debug file " + f.getAbsolutePath() + ": " + e);
        }
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
    static final String[] HEADERS = {"ID", "Vendor", "Service", "Error", "Username",
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
                String key = errorGroupKey(row[3]);
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

    static final int MAX_IDS_IN_SUMMARY = 5;

    /** Writes an "Error Type | Count | IDs" summary table at the top of the sheet. Returns next free row. */
    int writeErrorCountSummary(Sheet sh, Map<String, List<String[]>> byError, CellStyle bold) {
        Row hr = sh.createRow(0);
        Cell c0 = hr.createCell(0); c0.setCellValue("Error Type"); c0.setCellStyle(bold);
        Cell c1 = hr.createCell(1); c1.setCellValue("Count");     c1.setCellStyle(bold);
        Cell c2 = hr.createCell(2); c2.setCellValue("IDs");       c2.setCellStyle(bold);
        int r = 1;
        for (Map.Entry<String, List<String[]>> e : byError.entrySet()) {
            Row row = sh.createRow(r++);
            row.createCell(0).setCellValue(e.getKey());
            row.createCell(1).setCellValue(e.getValue().size());
            List<String[]> groupRows = e.getValue();
            String ids = groupRows.stream().limit(MAX_IDS_IN_SUMMARY).map(row2 -> row2[0])
                    .collect(java.util.stream.Collectors.joining(", "));
            if (groupRows.size() > MAX_IDS_IN_SUMMARY) ids += ", ...";
            row.createCell(2).setCellValue(ids);
        }
        sh.autoSizeColumn(0);
        sh.autoSizeColumn(1);
        sh.autoSizeColumn(2);
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
        HttpResponse<String> resp = client.send(req, HttpResponse.BodyHandlers.ofString(StandardCharsets.UTF_8));
        log("GET  " + url + " -> HTTP " + resp.statusCode() + " (" + resp.body().length() + " chars)");
        return resp.body();
    }

    String post(String url, String body) throws Exception {
        HttpRequest req = HttpRequest.newBuilder(URI.create(url))
                .header("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
                .header("Content-Type", "application/x-www-form-urlencoded")
                .POST(HttpRequest.BodyPublishers.ofString(body, StandardCharsets.UTF_8))
                .build();
        HttpResponse<String> resp = client.send(req, HttpResponse.BodyHandlers.ofString(StandardCharsets.UTF_8));
        log("POST " + url + " -> HTTP " + resp.statusCode() + " (" + resp.body().length() + " chars)");
        return resp.body();
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
