package com.sd.discovery.single;

import cn.hutool.core.date.DateUtil;
import cn.hutool.core.thread.ThreadUtil;
import cn.hutool.core.util.NumberUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import cn.hutool.poi.excel.cell.CellUtil;
import com.google.common.collect.Lists;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.input.DragEvent;
import javafx.scene.input.Dragboard;
import javafx.scene.input.TransferMode;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import java.io.*;
import java.util.*;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;

public class MainController {

    // ===== FXML 注入 =====
    @FXML private StackPane dropZone;
    @FXML private Label dropIcon, dropFileName, dropFileDetail;
    @FXML private TextField outputDirField;
    @FXML private TextField usernameField;
    @FXML private PasswordField passwordField;
    @FXML private CheckBox rememberCheck;
    @FXML private Button startBtn;
    @FXML private Label statGroups, statHits, statBest, statTime;
    @FXML private ProgressBar progressBar;
    @FXML private Label progressLabel, progressPercent;
    @FXML private ScrollPane logScrollPane;
    @FXML private VBox logContainer;

    // ===== 状态 =====
    private File selectedFile;
    private String workDir;
    private final AtomicBoolean running = new AtomicBoolean(false);
    private Thread workerThread;
    private long startTime;
    private String sessionStartTime; // 本次会话时间戳 yyyyMMddHHmmss
    private int writerRowIndex = 1;  // 当前写入行索引
    private final AtomicInteger groupCount = new AtomicInteger(0);
    private final AtomicInteger hitCount = new AtomicInteger(0);
    private final AtomicInteger bestCount = new AtomicInteger(0);
    private int totalGroups = 0;

    private static final String CRED_FILE = System.getProperty("user.home") + "/procalc5/.credentials";
    private static final String DEFAULT_URL = "https://procalc5.proflute.se/rotor";
    private static final int MAX_EMPTY = 5;

    // ===== 初始化 =====
    @FXML
    public void initialize() {
        // 默认工作目录
        boolean isWindows = System.getProperty("os.name").toLowerCase().contains("win");
        workDir = isWindows ? "C:\\procalc5" : System.getProperty("user.home") + "/procalc5";
        outputDirField.setText(workDir);

        // 加载保存的凭证
        loadCredentials();

        // 自动查找 procalc5.proflute.xlsx
        File defaultFile = new File(workDir + (isWindows ? "\\" : "/") + "procalc5.proflute.xlsx");
        if (defaultFile.exists()) {
            setFileSelected(defaultFile);
        }
    }

    // ===== 凭证管理 =====
    private void loadCredentials() {
        try {
            File f = new File(CRED_FILE);
            if (f.exists()) {
                Properties props = new Properties();
                props.load(new FileInputStream(f));
                usernameField.setText(props.getProperty("username", ""));
                passwordField.setText(props.getProperty("password", ""));
                rememberCheck.setSelected(true);
            }
        } catch (Exception e) {
            // ignore
        }
    }

    private void saveCredentials() {
        if (!rememberCheck.isSelected()) {
            new File(CRED_FILE).delete();
            return;
        }
        try {
            new File(CRED_FILE).getParentFile().mkdirs();
            Properties props = new Properties();
            props.setProperty("username", usernameField.getText().trim());
            props.setProperty("password", passwordField.getText().trim());
            props.store(new FileOutputStream(CRED_FILE), "Procalc5 credentials");
        } catch (Exception e) {
            // ignore
        }
    }

    // ===== 文件拖拽/选择 =====
    @FXML
    public void onDragOver(DragEvent event) {
        if (event.getDragboard().hasFiles()) {
            event.acceptTransferModes(TransferMode.COPY);
        }
        event.consume();
    }

    @FXML
    public void onDragDropped(DragEvent event) {
        Dragboard db = event.getDragboard();
        if (db.hasFiles()) {
            File file = db.getFiles().get(0);
            if (file.getName().endsWith(".xlsx")) {
                setFileSelected(file);
            } else {
                appendLog("⚠ 请选择 .xlsx 文件", "warn");
            }
        }
        event.setDropCompleted(true);
        event.consume();
    }

    @FXML
    public void onSelectFile() {
        FileChooser fc = new FileChooser();
        fc.setTitle("选择 Excel 输入文件");
        fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel 文件", "*.xlsx"));
        File dir = new File(workDir);
        if (dir.exists()) fc.setInitialDirectory(dir);
        File file = fc.showOpenDialog(dropZone.getScene().getWindow());
        if (file != null) {
            setFileSelected(file);
        }
    }

    private void setFileSelected(File file) {
        selectedFile = file;
        dropIcon.setVisible(true);
        dropIcon.setManaged(true);
        dropFileName.setText(file.getName());
        try {
            int rows = ExcelUtil.getReader(file).read().size() - 1;
            long sizeKb = file.length() / 1024;
            dropFileDetail.setText(rows + " 组数据 · " + sizeKb + "KB");
            totalGroups = rows;
            statGroups.setText("0/" + totalGroups);
        } catch (Exception e) {
            dropFileDetail.setText(file.getAbsolutePath());
        }
        dropZone.setStyle(dropZone.getStyle().replace("#c7c7cc", "#34c759").replace("dashed", "solid"));
    }

    // ===== 目录选择 =====
    @FXML
    public void onSelectDir() {
        DirectoryChooser dc = new DirectoryChooser();
        dc.setTitle("选择输出目录");
        File dir = new File(outputDirField.getText());
        if (dir.exists()) dc.setInitialDirectory(dir);
        File selected = dc.showDialog(dropZone.getScene().getWindow());
        if (selected != null) {
            outputDirField.setText(selected.getAbsolutePath());
            workDir = selected.getAbsolutePath();
        }
    }

    @FXML
    public void onOpenDir() {
        String dir = outputDirField.getText();
        if (dir != null && !dir.isEmpty()) {
            new File(dir).mkdirs();
            try {
                if (System.getProperty("os.name").toLowerCase().contains("mac")) {
                    Runtime.getRuntime().exec(new String[]{"open", dir});
                } else if (System.getProperty("os.name").toLowerCase().contains("win")) {
                    Runtime.getRuntime().exec(new String[]{"explorer", dir});
                } else {
                    Runtime.getRuntime().exec(new String[]{"xdg-open", dir});
                }
            } catch (IOException e) {
                appendLog("打开目录失败: " + e.getMessage(), "error");
            }
        }
    }

    // ===== 运行控制 =====
    @FXML
    public void onStart() {
        if (running.get()) {
            // 停止
            running.set(false);
            if (workerThread != null) workerThread.interrupt();
            startBtn.setText("▶  开始运行");
            startBtn.setStyle(startBtn.getStyle().replace("#ff3b30", "#0071e3"));
            appendLog("⏹ 用户手动停止", "warn");
            return;
        }

        // 校验
        if (selectedFile == null || !selectedFile.exists()) {
            appendLog("⚠ 请先选择输入文件", "warn");
            return;
        }
        if (StrUtil.isBlank(usernameField.getText()) || StrUtil.isBlank(passwordField.getText())) {
            appendLog("⚠ 请输入用户名和密码", "warn");
            return;
        }

        // 保存凭证
        saveCredentials();

        // 启动
        running.set(true);
        groupCount.set(0);
        hitCount.set(0);
        bestCount.set(0);
        startTime = System.currentTimeMillis();
        startBtn.setText("⏹  停止运行");
        startBtn.setStyle(startBtn.getStyle().replace("#0071e3", "#ff3b30"));
        logContainer.getChildren().clear();
        appendLog("程序启动，工作目录: " + workDir, "info");

        workerThread = new Thread(() -> {
            try {
                runAutomation();
            } catch (Exception e) {
                Platform.runLater(() -> appendLog("❌ 异常: " + e.getMessage(), "error"));
            } finally {
                Platform.runLater(() -> {
                    running.set(false);
                    startBtn.setText("▶  开始运行");
                    startBtn.setStyle(startBtn.getStyle().replace("#ff3b30", "#0071e3"));
                    updateTimer();
                });
            }
        });
        workerThread.setDaemon(true);
        workerThread.start();

        // 启动计时器
        startTimer();
    }

    private void startTimer() {
        Thread timerThread = new Thread(() -> {
            while (running.get()) {
                ThreadUtil.safeSleep(1000);
                Platform.runLater(this::updateTimer);
            }
        });
        timerThread.setDaemon(true);
        timerThread.start();
    }

    private void updateTimer() {
        long elapsed = (System.currentTimeMillis() - startTime) / 1000;
        int min = (int) (elapsed / 60);
        int sec = (int) (elapsed % 60);
        statTime.setText(String.format("%02d:%02d", min, sec));
    }

    // ===== 核心自动化逻辑 =====
    private void runAutomation() {
        String sep = System.getProperty("os.name").toLowerCase().contains("win") ? "\\" : "/";
        String username = usernameField.getText().trim();
        String password = passwordField.getText().trim();

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--no-sandbox");
        options.addArguments("--disable-dev-shm-usage");
        WebDriver driver = new ChromeDriver(options);

        try {
            // 登录
            Platform.runLater(() -> appendLog("正在打开浏览器...", "info"));
            driver.get(DEFAULT_URL);
            ThreadUtil.safeSleep(8000);
            driver.findElement(By.id("userNameInput")).sendKeys(username);
            driver.findElement(By.id("passwordInput")).sendKeys(password);
            driver.findElement(By.xpath("//*[@id=\"submitButton\"]")).click();
            ThreadUtil.safeSleep(5000);
            Platform.runLater(() -> appendLog("登录成功", "info"));

            // Excel 初始化
            sessionStartTime = DateUtil.format(DateUtil.date(), "yyyyMMddHHmmss");
            String allFilePath = workDir + sep + "calculate_results_all.xlsx";
            File allFile = new File(allFilePath);
            List<String> header = Lists.newArrayList("序号", "计算时间", " Wet Air:", "", "", "", "", "Process left",
                "", "", "", "", "", "", "Process Right", "", "", "", "", "", "Reactivation",
                "", "", "", "", "", "", "RPH");
            ExcelWriter excelWriter;
            if (allFile.exists()) {
                int existingCount = ExcelUtil.getReader(allFilePath).read().size();
                excelWriter = ExcelUtil.getWriter(allFilePath);
                writerRowIndex = existingCount + 1; // +1 空行分隔
            } else {
                excelWriter = ExcelUtil.getWriter(allFilePath);
                excelWriter.writeHeadRow(header);
                writerRowIndex = 1;
            }

            List<List<Object>> paraList = ExcelUtil.getReader(selectedFile.getAbsolutePath()).read();
            Double lastFoundTemp = null;
            List<List<Object>> bestResults = new ArrayList<>();
            int groupIdx = 0;

            for (List<Object> list : paraList) {
                if (!running.get()) break;
                if (paraList.indexOf(list) == 0) continue;

                groupIdx++;
                final int gIdx = groupIdx;
                Platform.runLater(() -> {
                    statGroups.setText(gIdx + "/" + totalGroups);
                    progressLabel.setText("正在处理第 " + gIdx + " 组...");
                    progressBar.setProgress((double) gIdx / totalGroups);
                    progressPercent.setText((int)((double)gIdx / totalGroups * 100) + "%");
                });

                // 解析参数 (与原代码一致)
                String linesNumber = StrUtil.toString(list.get(0));
                String UnitsofMeasure = StrUtil.toString(list.get(1));
                String RelativeHumidity = StrUtil.toString(list.get(2));
                String WetBulb = StrUtil.toString(list.get(3));
                String Pressurealtitud = StrUtil.toString(list.get(4));
                String PressurealtitudV = StrUtil.toString(list.get(5));
                String Showbypass = StrUtil.toString(list.get(6));
                String Reactivationinputtype = StrUtil.toString(list.get(7));
                String AirflowRange = StrUtil.toString(list.get(8));
                String Dewpointrange = StrUtil.toString(list.get(9));
                String Performancesafetyfactor = StrUtil.toString(list.get(10));
                String PerformancesafetyfactorV = StrUtil.toString(list.get(11));
                String ProcessAirflow = StrUtil.toString(list.get(12));
                String DesiccantNedia = StrUtil.toString(list.get(13));
                String SectorLayout = StrUtil.toString(list.get(14));
                String RotorDiameter = StrUtil.toString(list.get(16));
                String RotorDepth = StrUtil.toString(list.get(17));
                String NetFaceAreaCalculation = StrUtil.toString(list.get(18));
                String SealingArea = StrUtil.toString(list.get(19));
                String ProcessStrC = StrUtil.toString(list.get(20));
                String ProcessStrGKG = StrUtil.toString(list.get(21));
                String Rph = StrUtil.toString(list.get(22));
                Double ReactivationStart = Double.parseDouble(list.get(23).toString());
                Double ReactivationEnd = Double.parseDouble(list.get(24).toString());
                Double Reactivationbc = Double.parseDouble(list.get(25).toString());
                Double fanweiStart = Double.parseDouble(StrUtil.split(((String) list.get(26)), "~").get(0));
                Double fanweiEnd = Double.parseDouble(StrUtil.split(((String) list.get(26)), "~").get(1));
                String Reactivation1 = StrUtil.toString(list.get(27));
                String Reactivation2 = StrUtil.toString(list.get(28));
                String Reactivation3 = StrUtil.toString(list.get(29));

                // 设置网页参数（复用原有逻辑）
                setWebParams(driver, UnitsofMeasure, RelativeHumidity, WetBulb, Pressurealtitud,
                    PressurealtitudV, Showbypass, Reactivationinputtype, AirflowRange, Dewpointrange,
                    Performancesafetyfactor, PerformancesafetyfactorV, ProcessAirflow, DesiccantNedia,
                    SectorLayout, RotorDiameter, RotorDepth, NetFaceAreaCalculation, SealingArea,
                    ProcessStrC, ProcessStrGKG, Rph, Reactivation1, Reactivation2, Reactivation3);

                if (!running.get()) break;

                // 倒序查找
                WebElement Reactivation = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[3]/div/div[2]/div/div[1]/div/div/input"));
                WebElement gkg = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[2]/div/div/input"));
                WebElement button = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[9]/div[3]/button"));

                // 边界值检查
                fillInput(driver, Reactivation, ReactivationStart.toString());
                safeClick(button);
                ThreadUtil.safeSleep(1500);
                if (StrUtil.isEmpty(gkg.getAttribute("value"))) continue;
                Double gkgLeft = Double.parseDouble(gkg.getAttribute("value"));
                fillInput(driver, Reactivation, ReactivationEnd.toString());
                safeClick(button);
                ThreadUtil.safeSleep(1500);
                if (StrUtil.isEmpty(gkg.getAttribute("value"))) continue;

                boolean qk1 = (fanweiStart <= gkgLeft && gkgLeft <= fanweiEnd);
                Double ReactivationStartReal = ReactivationStart;
                Double ReactivationEndReal = ReactivationEnd;

                if (!qk1) {
                    fillInput(driver, Reactivation, StrUtil.toString(NumberUtil.add(ReactivationStart, Reactivationbc)));
                    safeClick(button);
                    ThreadUtil.safeSleep(1500);
                    if (StrUtil.isEmpty(gkg.getAttribute("value"))) continue;
                    Double gkgTemp = Double.parseDouble(gkg.getAttribute("value"));
                    if (NumberUtil.compare(gkgTemp, gkgLeft) > 0) {
                        if (NumberUtil.compare(gkgLeft, fanweiStart) < 0 && NumberUtil.compare(gkgTemp, gkgLeft) != 0) {
                            Double wendus = NumberUtil.mul((Double) NumberUtil.div(NumberUtil.sub(fanweiStart, gkgLeft), NumberUtil.sub(gkgTemp, gkgLeft)), Reactivationbc);
                            ReactivationStartReal = NumberUtil.add(ReactivationStart.doubleValue(), NumberUtils.safeMultiply(wendus, 1.0, 0).doubleValue());
                        } else { ReactivationStartReal = null; }
                    } else {
                        if (NumberUtil.compare(gkgLeft, fanweiEnd) > 0 && NumberUtil.compare(gkgLeft, gkgTemp) != 0) {
                            Double wendus = NumberUtil.mul((Double) NumberUtil.div(NumberUtil.sub(gkgLeft, fanweiEnd), NumberUtil.sub(gkgLeft, gkgTemp)), Reactivationbc);
                            ReactivationStartReal = NumberUtil.add(ReactivationStart.doubleValue(), NumberUtils.safeMultiply(wendus, 1.0, 0).doubleValue());
                        } else { ReactivationStartReal = null; }
                    }
                }
                if (ReactivationEndReal == null || ReactivationStartReal == null) {
                    ReactivationStartReal = ReactivationStart;
                    ReactivationEndReal = ReactivationEnd;
                }

                Double searchUpper = (lastFoundTemp != null) ? Math.min(lastFoundTemp, ReactivationEndReal) : ReactivationEndReal;
                Double searchLower = ReactivationStartReal;
                Platform.runLater(() -> appendLog("[倒序查找] 第" + gIdx + "组: " + searchUpper + "°C → " + searchLower + "°C", ""));

                if (NumberUtil.compare(searchUpper, searchLower) < 0) continue;

                boolean flag = false;
                Double tempCurrent = searchUpper;
                int emptyCount = 0;
                StringBuilder ss = new StringBuilder();
                List<double[]> groupHits = new ArrayList<>();

                while (NumberUtil.compare(tempCurrent, searchLower) >= 0 && running.get()) {
                    try {
                        ((org.openqa.selenium.JavascriptExecutor) driver).executeScript(
                            "var backdrop = document.querySelector('.MuiBackdrop-root');" +
                            "if(backdrop && getComputedStyle(backdrop).opacity > 0) { backdrop.click(); }");
                        ThreadUtil.safeSleep(300);
                        fillInput(driver, Reactivation, StrUtil.toString(tempCurrent));
                    } catch (Exception e) {
                        Platform.runLater(() -> appendLog("填充温度失败: " + e.getMessage(), "warn"));
                        tempCurrent = NumberUtil.sub(tempCurrent, Reactivationbc);
                        continue;
                    }
                    safeClick(button);
                    ThreadUtil.safeSleep(1500);

                    String gkgValue;
                    try { gkgValue = gkg.getAttribute("value"); } catch (Exception e) {
                        emptyCount++;
                        if (emptyCount >= MAX_EMPTY) break;
                        tempCurrent = NumberUtil.sub(tempCurrent, Reactivationbc);
                        continue;
                    }
                    if (StrUtil.isEmpty(gkgValue)) {
                        emptyCount++;
                        if (emptyCount >= MAX_EMPTY) break;
                        tempCurrent = NumberUtil.sub(tempCurrent, Reactivationbc);
                        continue;
                    }
                    emptyCount = 0;
                    final Double gkgVal = Double.parseDouble(gkgValue);

                    if (fanweiStart <= gkgVal && gkgVal <= fanweiEnd) {
                        BaiscApplication.toList(driver, ss, linesNumber, excelWriter, sessionStartTime, true);
                        groupHits.add(new double[]{tempCurrent, gkgVal});
                        lastFoundTemp = tempCurrent;
                        flag = true;
                        hitCount.incrementAndGet();
                        final Double curTemp = tempCurrent;
                        Platform.runLater(() -> {
                            statHits.setText(String.valueOf(hitCount.get()));
                            appendLog("★ 命中! 温度=" + curTemp + ", g/kg=" + gkgVal, "hit");
                        });
                    }
                    tempCurrent = NumberUtil.sub(tempCurrent, Reactivationbc);
                    if (flag && !(fanweiStart <= gkgVal && gkgVal <= fanweiEnd)) break;
                }

                // 最优解
                if (!groupHits.isEmpty()) {
                    double midValue = (fanweiStart + fanweiEnd) / 2.0;
                    double[] best = groupHits.get(0);
                    double bestDiff = Math.abs(best[1] - midValue);
                    for (int i = 1; i < groupHits.size(); i++) {
                        double diff = Math.abs(groupHits.get(i)[1] - midValue);
                        if (diff < bestDiff || (diff == bestDiff && groupHits.get(i)[0] > best[0])) {
                            best = groupHits.get(i);
                            bestDiff = diff;
                        }
                    }
                    final double fMid = midValue, fBestGkg = best[1], fBestTemp = best[0], fBestDiff = bestDiff;
                    Platform.runLater(() -> appendLog("[最优解] 第" + gIdx + "组: 中间值=" + NumberUtil.round(fMid, 4)
                        + ", 最优: g/kg=" + fBestGkg + "(" + fBestTemp + "°C), 差值=" + NumberUtil.round(fBestDiff, 4), "hit"));

                    fillInput(driver, Reactivation, StrUtil.toString(best[0]));
                    safeClick(button);
                    ThreadUtil.safeSleep(1500);
                    List<Object> bestRow = BaiscApplication.collectRowData(driver, linesNumber);
                    bestRow.add(1, sessionStartTime); // 插入时间戳到B列
                    bestRow.add(NumberUtil.round(midValue, 4)); // 额外列：范围中间值
                    // 写入全量文件（红色+加粗标记最优解）
                    int bestRowIdx = writerRowIndex++;
                    excelWriter.writeRow(bestRow);
                    applyRowStyle(excelWriter, bestRowIdx, bestRow.size(), true, true);
                    bestResults.add(bestRow);
                    bestCount.incrementAndGet();
                    Platform.runLater(() -> statBest.setText(String.valueOf(bestCount.get())));
                }
            }

            // 写入文件
            excelWriter.flush();
            // 输出最优解汇总到 {时间戳}_result_02.xlsx
            String bestFilePath = workDir + sep + sessionStartTime + "_result_02.xlsx";
            if (!bestResults.isEmpty()) {
                ExcelWriter bestWriter = ExcelUtil.getWriter(bestFilePath);
                List<String> bestHeader = Lists.newArrayList("序号", "计算时间", " Wet Air:", "", "", "", "", "Process left",
                    "", "", "", "", "", "", "Process Right", "", "", "", "", "", "Reactivation",
                    "", "", "", "", "", "", "RPH", "范围中间值");
                bestWriter.writeHeadRow(bestHeader);
                int bestIdx = 1;
                for (List<Object> row : bestResults) {
                    bestWriter.writeRow(row);
                    applyRowStyle(bestWriter, bestIdx, row.size(), true, true);
                    bestIdx++;
                }
                bestWriter.flush();
            }
            final String finalBestPath = bestFilePath;
            final int finalGroupCount = groupIdx;
            Platform.runLater(() -> {
                appendLog("===== 运行完成 =====", "info");
                appendLog("共 " + finalGroupCount + " 组, " + hitCount.get() + " 条有效数据, " + bestCount.get() + " 个最优解", "info");
                appendLog("全量数据已追加到: " + allFilePath, "info");
                appendLog("最优解已写入: " + finalBestPath, "info");
                progressBar.setProgress(1.0);
                progressLabel.setText("运行完成");
                progressPercent.setText("100%");
            });

        } catch (Exception e) {
            Platform.runLater(() -> appendLog("❌ " + e.getMessage(), "error"));
        } finally {
            try { driver.quit(); } catch (Exception e) {}
        }
    }

    // ===== 网页参数设置（从 BaiscApplication 提取）=====
    private void setWebParams(WebDriver driver, String UnitsofMeasure, String RelativeHumidity, String WetBulb,
        String Pressurealtitud, String PressurealtitudV, String Showbypass, String Reactivationinputtype,
        String AirflowRange, String Dewpointrange, String Performancesafetyfactor, String PerformancesafetyfactorV,
        String ProcessAirflow, String DesiccantNedia, String SectorLayout, String RotorDiameter, String RotorDepth,
        String NetFaceAreaCalculation, String SealingArea, String ProcessStrC, String ProcessStrGKG,
        String Rph, String Reactivation1, String Reactivation2, String Reactivation3) throws InterruptedException {

        // 范围1
        if (StrUtil.equalsIgnoreCase(UnitsofMeasure, "si")) {
            driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[1]/div/label[1]/span/input")).click();
        } else {
            driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[1]/div/label[2]/span/input")).click();
        }
        if (!StrUtil.equalsIgnoreCase(RelativeHumidity, "勾选")) {
            driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[2]/div/div/label[1]/span/input")).click();
        }
        if (!StrUtil.equalsIgnoreCase(WetBulb, "勾选")) {
            driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[2]/div/div/label[2]/span/input")).click();
        }
        if (StrUtil.equalsIgnoreCase(Pressurealtitud, "Altitude")) {
            driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[3]/div/div[1]/label[1]/span/input")).click();
        } else {
            driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[3]/div/div[1]/label[2]/span/input")).click();
        }
        fillInput(driver, "//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[3]/div/div[2]/input", PressurealtitudV);
        if (StrUtil.equalsIgnoreCase(Showbypass, "No")) {
            driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[4]/div/div/label[1]/span/input")).click();
        } else {
            driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[4]/div/div/label[2]/span/input")).click();
        }
        if (StrUtil.equalsIgnoreCase(Reactivationinputtype, "Temp")) {
            driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[5]/div/div/label[1]/span/input")).click();
        } else {
            driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[5]/div/div/label[2]/span/input")).click();
        }
        if (StrUtil.equalsIgnoreCase(AirflowRange, "Default")) {
            driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[6]/div/div/label[1]/span/input")).click();
        } else {
            driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[6]/div/div/label[2]/span/input")).click();
        }
        if (StrUtil.equalsIgnoreCase(Dewpointrange, "Default")) {
            driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[7]/div/div/label[1]/span/input")).click();
        } else {
            driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[7]/div/div/label[2]/span/input")).click();
        }
        // Performance safety factor
        driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[8]/div/div/div/div")).click();
        ThreadUtil.safeSleep(500);
        List<WebElement> psfOptions = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
        boolean flag = true;
        if (StrUtil.equalsIgnoreCase(Performancesafetyfactor, "None")) { psfOptions.get(0).click(); flag = false; }
        else if (StrUtil.equalsIgnoreCase(Performancesafetyfactor, "+Δ% Moisture")) { psfOptions.get(1).click(); flag = false; }
        else if (StrUtil.equalsIgnoreCase(Performancesafetyfactor, "x Moisture")) { psfOptions.get(2).click(); flag = false; }
        if (flag) psfOptions.get(0).click();
        if (!StrUtil.equalsIgnoreCase(Performancesafetyfactor, "None")) {
            driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[8]/div/div/div[2]/div/div")).click();
            List<WebElement> opts = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
            for (WebElement s : opts) {
                if (StrUtil.equalsAnyIgnoreCase(PerformancesafetyfactorV, s.getAttribute("data-value"))) { s.click(); break; }
            }
        }
        // 范围2
        fillInput(driver, "//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[1]/div/div[2]/input", ProcessAirflow);
        // Media
        driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[3]/div/div[2]")).click();
        ThreadUtil.safeSleep(1000);
        List<WebElement> medias = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
        flag = true;
        for (WebElement s : medias) {
            String sTemp = StrUtil.equalsIgnoreCase(s.getAttribute("data-value"), "1") ? "PPS" : "PPP";
            if (StrUtil.equalsIgnoreCase(DesiccantNedia, sTemp)) { s.click(); flag = false; break; }
        }
        if (flag) medias.get(0).click();
        // Sector Layout
        Platform.runLater(() -> appendLog("[Sector Layout] " + SectorLayout, ""));
        driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[4]/div/div[2]")).click();
        ThreadUtil.safeSleep(1500);
        flag = true;
        for (int retry = 0; retry < 3 && flag; retry++) {
            try {
                List<WebElement> slOptions = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
                for (WebElement s : slOptions) {
                    if (StrUtil.equalsIgnoreCase(s.getText().trim(), SectorLayout)) { s.click(); flag = false; break; }
                }
                if (flag && slOptions.size() > 0) { slOptions.get(0).click(); flag = false; }
            } catch (org.openqa.selenium.StaleElementReferenceException e) {
                ThreadUtil.safeSleep(1000);
                driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[4]/div/div[2]")).click();
                ThreadUtil.safeSleep(1500);
            }
        }
        // Rotor diameter/depth
        boolean isCustomSector = StrUtil.equalsAnyIgnoreCase(SectorLayout, "Custom 2-sector", "Custom 3 sector");
        if (isCustomSector) {
            fillInput(driver, driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[6]/div/div[2]/input")), RotorDiameter);
        } else {
            for (int retry = 0; retry < 3; retry++) {
                try {
                    driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[6]/div/div[2]")).click();
                    ThreadUtil.safeSleep(1000);
                    List<WebElement> opts = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
                    flag = true;
                    for (WebElement s : opts) { if (StrUtil.equalsIgnoreCase(RotorDiameter, s.getAttribute("data-value"))) { s.click(); flag = false; break; } }
                    if (flag) opts.get(0).click();
                    break;
                } catch (org.openqa.selenium.StaleElementReferenceException e) { ThreadUtil.safeSleep(1000); }
            }
        }
        if (isCustomSector) {
            fillInput(driver, driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[7]/div/div[2]/input")), RotorDepth);
        } else {
            for (int retry = 0; retry < 3; retry++) {
                try {
                    driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[7]/div/div[2]")).click();
                    ThreadUtil.safeSleep(1000);
                    List<WebElement> opts = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
                    flag = true;
                    for (WebElement s : opts) { if (StrUtil.equalsIgnoreCase(RotorDepth, s.getAttribute("data-value"))) { s.click(); flag = false; break; } }
                    if (flag) opts.get(0).click();
                    break;
                } catch (org.openqa.selenium.StaleElementReferenceException e) { ThreadUtil.safeSleep(1000); }
            }
        }
        // Net face area
        driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[8]/div/div[2]")).click();
        ThreadUtil.safeSleep(1000);
        List<WebElement> nfaOptions = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
        flag = true;
        for (WebElement s : nfaOptions) {
            String sTemp = StrUtil.equalsIgnoreCase(s.getAttribute("data-value"), "0") ? "Sealing area" : "Active area";
            if (StrUtil.equalsIgnoreCase(NetFaceAreaCalculation, sTemp)) { s.click(); flag = false; break; }
        }
        if (flag) nfaOptions.get(0).click();
        fillInput(driver, "//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[9]/div/div[2]/input", SealingArea);
        // 底部参数
        Platform.runLater(() -> appendLog("[填充值] C=" + ProcessStrC + ", GKG=" + ProcessStrGKG + ", Rph=" + Rph, ""));
        fillInput(driver, "//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[1]/div/div/input", ProcessStrC);
        fillInput(driver, "//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[2]/div/div/input", ProcessStrGKG);
        fillInput(driver, "//*[@id=\"root\"]/div/div/div[1]/div[2]/div[8]/div/div[2]/div/div/div/div/input", Rph);
        fillInput(driver, "//*[@id=\"root\"]/div/div/div[1]/div[2]/div[5]/div/div[2]/div/div[1]/div/div/input", Reactivation1);
        fillInput(driver, "//*[@id=\"root\"]/div/div/div[1]/div[2]/div[5]/div/div[2]/div/div[2]/div/div/input", Reactivation2);
        fillInput(driver, "//*[@id=\"root\"]/div/div/div[1]/div[2]/div[5]/div/div[2]/div/div[5]/div/div/input", Reactivation3);
    }

    // ===== 工具方法 =====
    private static void fillInput(WebDriver driver, String xpath, String value) {
        WebElement el = driver.findElement(By.xpath(xpath));
        fillInput(driver, el, value);
    }

    private static void fillInput(WebDriver driver, WebElement el, String value) {
        ((org.openqa.selenium.JavascriptExecutor) driver).executeScript(
            "var el = arguments[0];" +
            "el.click(); el.focus();" +
            "var nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;" +
            "nativeInputValueSetter.call(el, arguments[1]);" +
            "el.dispatchEvent(new Event('input', { bubbles: true }));" +
            "el.dispatchEvent(new Event('change', { bubbles: true }));",
            el, value);
    }

    private void safeClick(WebElement button) {
        try { BaiscApplication.click(button); } catch (Exception e) {}
    }

    // ===== 日志 =====
    @FXML
    public void onClearLog() {
        logContainer.getChildren().clear();
    }

    /**
     * 对指定行应用样式：加粗和/或红色字体
     */
    private void applyRowStyle(ExcelWriter writer, int rowIndex, int colCount, boolean bold, boolean red) {
        Sheet sheet = writer.getSheet();
        Row row = sheet.getRow(rowIndex);
        if (row == null) return;
        Workbook wb = writer.getWorkbook();
        org.apache.poi.ss.usermodel.CellStyle style = wb.createCellStyle();
        org.apache.poi.ss.usermodel.Font font = wb.createFont();
        if (bold) font.setBold(true);
        if (red) font.setColor(org.apache.poi.ss.usermodel.IndexedColors.RED.getIndex());
        style.setFont(font);
        for (int i = 0; i < colCount; i++) {
            Cell cell = row.getCell(i);
            if (cell != null) cell.setCellStyle(style);
        }
    }

    private void appendLog(String msg, String type) {
        Platform.runLater(() -> {
            Label label = new Label(msg);
            label.setStyle("-fx-font-family: monospace; -fx-font-size: 11; -fx-wrap-text: false;");
            if ("hit".equals(type)) label.setStyle(label.getStyle() + " -fx-text-fill: #34c759;");
            else if ("warn".equals(type)) label.setStyle(label.getStyle() + " -fx-text-fill: #ff9500;");
            else if ("error".equals(type)) label.setStyle(label.getStyle() + " -fx-text-fill: #ff3b30;");
            else if ("info".equals(type)) label.setStyle(label.getStyle() + " -fx-text-fill: #007aff;");
            logContainer.getChildren().add(label);
            // 自动滚动到底部
            logScrollPane.layout();
            logScrollPane.setVvalue(1.0);
        });
    }
}
