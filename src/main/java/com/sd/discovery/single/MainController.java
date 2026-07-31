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
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
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
    @FXML private TextArea logTextArea;

    // ===== 状态 =====
    private File selectedFile;
    private String workDir;
    private final AtomicBoolean running = new AtomicBoolean(false);
    private Thread workerThread;
    private long startTime;
    private String sessionStartTime;
    private String sessionTimeDisplay;
    private int writerRowIndex = 1;
    private final AtomicInteger groupCount = new AtomicInteger(0);
    private final AtomicInteger hitCount = new AtomicInteger(0);
    private final AtomicInteger bestCount = new AtomicInteger(0);
    private int totalGroups = 0;
    // 停止时保存数据用
    private ExcelWriter currentExcelWriter;
    private List<List<Object>> currentBestResults;
    private String currentAllFilePath;
    private String currentBestFilePath;
    private int currentGroupIdx = 0;

    private static final String CRED_FILE = System.getProperty("user.home") + "/procalc5/.credentials";
    private static final String DEFAULT_URL = "https://procalc5.proflute.se/rotor";
    private static final int MAX_EMPTY = 5;

    // ===== 初始化 =====
    @FXML
    public void initialize() {
        boolean isWindows = System.getProperty("os.name").toLowerCase().contains("win");
        workDir = isWindows ? "C:\\procalc5" : System.getProperty("user.home") + "/procalc5";
        outputDirField.setText(workDir);
        loadCredentials();
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
        } catch (Exception e) { }
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
        } catch (Exception e) { }
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
            // 弹出确认对话框
            Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
            alert.setTitle("停止运行");
            alert.setHeaderText("确定要停止运行吗？");
            alert.setContentText(String.format("已完成 %d 组，命中 %d 条，最优解 %d 个。\n是否保存已采集的数据？",
                currentGroupIdx, hitCount.get(), bestCount.get()));
            ButtonType saveBtn = new ButtonType("保存并停止");
            ButtonType discardBtn = new ButtonType("直接停止");
            ButtonType cancelBtn = new ButtonType("取消", ButtonBar.ButtonData.CANCEL_CLOSE);
            alert.getButtonTypes().setAll(saveBtn, discardBtn, cancelBtn);
            alert.showAndWait().ifPresent(response -> {
                if (response == cancelBtn) return;
                running.set(false);
                if (response == saveBtn) saveDataOnStop();
                if (workerThread != null) workerThread.interrupt();
                startBtn.setText("▶  开始运行");
                startBtn.setStyle(startBtn.getStyle().replace("#ff3b30", "#0071e3"));
                appendLog("⏹ 用户手动停止", "warn");
            });
            return;
        }

        if (selectedFile == null || !selectedFile.exists()) {
            appendLog("⚠ 请先选择输入文件", "warn");
            return;
        }
        if (StrUtil.isBlank(usernameField.getText()) || StrUtil.isBlank(passwordField.getText())) {
            appendLog("⚠ 请输入用户名和密码", "warn");
            return;
        }

        saveCredentials();
        running.set(true);
        groupCount.set(0);
        hitCount.set(0);
        bestCount.set(0);
        startTime = System.currentTimeMillis();
        startBtn.setText("⏹  停止运行");
        startBtn.setStyle(startBtn.getStyle().replace("#0071e3", "#ff3b30"));
        logTextArea.clear();
        startOutputRedirect();
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

    // ===== 核心自动化逻辑 (Selenium，与 start.sh 使用同一浏览器驱动) =====
    private void runAutomation() {
        String sep = System.getProperty("os.name").toLowerCase().contains("win") ? "\\" : "/";
        String username = usernameField.getText().trim();
        String password = passwordField.getText().trim();

        WebDriver driver = null;
        try {
            appendLog("正在启动 Chrome...", "info");
            ChromeOptions options = new ChromeOptions();
            options.addArguments("--start-maximized");
            driver = new ChromeDriver(options);
            appendLog("浏览器启动成功", "info");

            // 登录
            appendLog("正在打开网页...", "info");
            driver.get(DEFAULT_URL);
            ThreadUtil.safeSleep(8000);
            driver.findElement(By.id("userNameInput")).sendKeys(username);
            driver.findElement(By.id("passwordInput")).sendKeys(password);
            driver.findElement(By.id("submitButton")).click();
            ThreadUtil.safeSleep(5000);
            appendLog("登录成功", "info");

            // 等待表单加载
            try {
                long deadline = System.currentTimeMillis() + 30000;
                while (driver.findElements(By.cssSelector("input[type='radio']")).isEmpty()
                    && System.currentTimeMillis() < deadline) ThreadUtil.safeSleep(200);
                if (driver.findElements(By.cssSelector("input[type='radio']")).isEmpty()) throw new RuntimeException("form not ready");
            } catch (Exception e) {
                appendLog("等待表单超时，继续...", "warn");
            }
            ThreadUtil.safeSleep(2000);

            // Excel 初始化
            sessionStartTime = DateUtil.format(DateUtil.date(), "yyyyMMdd_HHmmss");
            sessionTimeDisplay = DateUtil.format(DateUtil.date(), "yyyy年M月d日H时m分");
            currentAllFilePath = workDir + sep + "calculate_results_all.xlsx";
            currentBestFilePath = workDir + sep + sessionStartTime + "_result_02.xlsx";
            File allFile = new File(currentAllFilePath);
            List<String> header = Lists.newArrayList("序号", "计算时间", " Wet Air:", "", "", "", "", "Process left",
                "", "", "", "", "", "", "Process Right", "", "", "", "", "", "Reactivation",
                "", "", "", "", "", "", "RPH");
            if (allFile.exists()) {
                int existingCount = ExcelUtil.getReader(currentAllFilePath).read().size();
                currentExcelWriter = ExcelUtil.getWriter(currentAllFilePath);
                writerRowIndex = existingCount + 1;
            } else {
                currentExcelWriter = ExcelUtil.getWriter(currentAllFilePath);
                currentExcelWriter.writeHeadRow(header);
                writerRowIndex = 1;
            }

            List<List<Object>> paraList = ExcelUtil.getReader(selectedFile.getAbsolutePath()).read();
            Double lastFoundTemp = null;
            currentBestResults = new ArrayList<>();
            currentGroupIdx = 0;

            for (List<Object> list : paraList) {
                if (!running.get()) break;
                if (paraList.indexOf(list) == 0) continue;

                currentGroupIdx++;
                final int gIdx = currentGroupIdx;
                Platform.runLater(() -> {
                    statGroups.setText(gIdx + "/" + totalGroups);
                    progressLabel.setText("正在处理第 " + gIdx + " 组...");
                    progressBar.setProgress((double) gIdx / totalGroups);
                    progressPercent.setText((int)((double)gIdx / totalGroups * 100) + "%");
                });

                // 解析参数
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

                // 设置网页参数
                setWebParams(driver, UnitsofMeasure, RelativeHumidity, WetBulb, Pressurealtitud,
                    PressurealtitudV, Showbypass, Reactivationinputtype, AirflowRange, Dewpointrange,
                    Performancesafetyfactor, PerformancesafetyfactorV, ProcessAirflow, DesiccantNedia,
                    SectorLayout, RotorDiameter, RotorDepth, NetFaceAreaCalculation, SealingArea,
                    ProcessStrC, ProcessStrGKG, Rph, Reactivation1, Reactivation2, Reactivation3);

                if (!running.get()) break;

                // 倒序查找
                String reactXpath = "//*[@id=\"root\"]/div/div/div[1]/div[2]/div[3]/div/div[2]/div/div[1]/div/div/input";
                String gkgXpath = "//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[2]/div/div/input";
                String buttonXpath = "//*[@id=\"root\"]/div/div/div[1]/div[2]/div[9]/div[3]/button";

                // 边界值检查
                appendLog("[边界检查] 第" + gIdx + "组 起始: React=" + ReactivationStart, "");
                fillInputByXpath(driver, reactXpath, ReactivationStart.toString());
                clickByXpath(driver, buttonXpath);
                ThreadUtil.safeSleep(1500);
                String gkgStartVal = getInputValueByXpath(driver, gkgXpath);
                appendLog("[边界检查] React=" + ReactivationStart + " → g/kg=" + gkgStartVal, "");
                if (StrUtil.isEmpty(gkgStartVal)) { appendLog("[边界检查] 起始值为空，跳过本组", "warn"); continue; }
                Double gkgLeft = Double.parseDouble(gkgStartVal);

                fillInputByXpath(driver, reactXpath, ReactivationEnd.toString());
                clickByXpath(driver, buttonXpath);
                ThreadUtil.safeSleep(1500);
                String gkgEndVal = getInputValueByXpath(driver, gkgXpath);
                appendLog("[边界检查] React=" + ReactivationEnd + " → g/kg=" + gkgEndVal, "");
                if (StrUtil.isEmpty(gkgEndVal)) { appendLog("[边界检查] 结束值为空，跳过本组", "warn"); continue; }

                boolean qk1 = (fanweiStart <= gkgLeft && gkgLeft <= fanweiEnd);
                Double ReactivationStartReal = ReactivationStart;
                Double ReactivationEndReal = ReactivationEnd;

                if (!qk1) {
                    fillInputByXpath(driver, reactXpath, StrUtil.toString(NumberUtil.add(ReactivationStart, Reactivationbc)));
                    clickByXpath(driver, buttonXpath);
                    ThreadUtil.safeSleep(1500);
                    String gkgTempStr = getInputValueByXpath(driver, gkgXpath);
                    if (StrUtil.isEmpty(gkgTempStr)) continue;
                    Double gkgTemp = Double.parseDouble(gkgTempStr);
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
                        ((JavascriptExecutor) driver).executeScript(
                            "var backdrop = document.querySelector('.MuiBackdrop-root'); if(backdrop && getComputedStyle(backdrop).opacity > 0) { backdrop.click(); }");
                        ThreadUtil.safeSleep(300);
                        fillInputByXpath(driver, reactXpath, StrUtil.toString(tempCurrent));
                    } catch (Exception e) {
                        Platform.runLater(() -> appendLog("填充温度失败: " + e.getMessage(), "warn"));
                        tempCurrent = NumberUtil.sub(tempCurrent, Reactivationbc);
                        continue;
                    }
                    clickByXpath(driver, buttonXpath);
                    ThreadUtil.safeSleep(1500);

                    String gkgValue;
                    try {
                        gkgValue = getInputValueByXpath(driver, gkgXpath);
                    } catch (Exception e) {
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
                    appendLog("[迭代] 温度=" + tempCurrent + "°C → g/kg=" + gkgVal + " (范围:" + fanweiStart + "~" + fanweiEnd + ")", "");

                    boolean isHit = (fanweiStart <= gkgVal && gkgVal <= fanweiEnd);
                    toListSelenium(driver, ss, linesNumber, currentExcelWriter, sessionTimeDisplay, isHit);

                    if (isHit) {
                        appendLog("[数据采集] 命中! 已写入第" + writerRowIndex + "行", "hit");
                        groupHits.add(new double[]{tempCurrent, gkgVal});
                        lastFoundTemp = tempCurrent;
                        flag = true;
                        hitCount.incrementAndGet();
                        final Double curTemp = tempCurrent;
                        Platform.runLater(() -> statHits.setText(String.valueOf(hitCount.get())));
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

                    fillInputByXpath(driver, reactXpath, StrUtil.toString(best[0]));
                    clickByXpath(driver, buttonXpath);
                    ThreadUtil.safeSleep(1500);
                    List<Object> bestRow = collectRowDataSelenium(driver, linesNumber);
                    appendLog("[最优解写入] 采集到" + bestRow.size() + "个字段", "");
                    bestRow.add(1, sessionTimeDisplay);
                    bestRow.add(NumberUtil.round(midValue, 4));
                    int bestRowIdx = writerRowIndex++;
                    currentExcelWriter.writeRow(bestRow);
                    applyRowStyle(currentExcelWriter, bestRowIdx, bestRow.size(), false, true);
                    currentBestResults.add(bestRow);
                    bestCount.incrementAndGet();
                    Platform.runLater(() -> statBest.setText(String.valueOf(bestCount.get())));
                }
            }

            // 写入文件
            saveResults();
            final int finalGroupCount = currentGroupIdx;
            appendLog("===== 运行完成 =====", "info");
            appendLog("共 " + finalGroupCount + " 组, " + hitCount.get() + " 条有效数据, " + bestCount.get() + " 个最优解", "info");
            appendLog("全量数据已追加到: " + currentAllFilePath, "info");
            appendLog("最优解已写入: " + currentBestFilePath, "info");
            Platform.runLater(() -> {
                progressBar.setProgress(1.0);
                progressLabel.setText("运行完成");
                progressPercent.setText("100%");
            });

        } catch (Exception e) {
            Platform.runLater(() -> appendLog("❌ " + e.getMessage(), "error"));
        } finally {
            if (driver != null) try { driver.quit(); } catch (Exception e) {}
        }
    }

    // ===== 数据保存 =====
    private void saveResults() {
        try {
            if (currentExcelWriter != null) currentExcelWriter.flush();
            if (currentBestResults != null && !currentBestResults.isEmpty()) {
                String sep = System.getProperty("os.name").toLowerCase().contains("win") ? "\\" : "/";
                ExcelWriter bestWriter = ExcelUtil.getWriter(currentBestFilePath);
                List<String> bestHeader = Lists.newArrayList("序号", "计算时间", " Wet Air:", "", "", "", "", "Process left",
                    "", "", "", "", "", "", "Process Right", "", "", "", "", "", "Reactivation",
                    "", "", "", "", "", "", "RPH", "范围中间值");
                bestWriter.writeHeadRow(bestHeader);
                int bestIdx = 1;
                for (List<Object> row : currentBestResults) {
                    bestWriter.writeRow(row);
                    applyRowStyle(bestWriter, bestIdx, row.size(), true, false);
                    bestIdx++;
                }
                bestWriter.flush();
            }
        } catch (Exception e) {
            appendLog("保存结果失败: " + e.getMessage(), "error");
        }
    }

    private void saveDataOnStop() {
        try {
            if (currentExcelWriter != null) {
                currentExcelWriter.flush();
                appendLog("全量数据已保存到: " + currentAllFilePath, "info");
            }
            if (currentBestResults != null && !currentBestResults.isEmpty()) {
                ExcelWriter bestWriter = ExcelUtil.getWriter(currentBestFilePath);
                List<String> bestHeader = Lists.newArrayList("序号", "计算时间", " Wet Air:", "", "", "", "", "Process left",
                    "", "", "", "", "", "", "Process Right", "", "", "", "", "", "Reactivation",
                    "", "", "", "", "", "", "RPH", "范围中间值");
                bestWriter.writeHeadRow(bestHeader);
                int bestIdx = 1;
                for (List<Object> row : currentBestResults) {
                    bestWriter.writeRow(row);
                    applyRowStyle(bestWriter, bestIdx, row.size(), true, false);
                    bestIdx++;
                }
                bestWriter.flush();
                appendLog("最优解已保存到: " + currentBestFilePath, "info");
            }
            appendLog("数据保存完成 (" + hitCount.get() + "条数据, " + bestCount.get() + "个最优解)", "info");
        } catch (Exception e) {
            appendLog("保存数据失败: " + e.getMessage(), "error");
        }
    }

    // ===== 网页参数设置 (Selenium版) =====
    private void setWebParams(WebDriver driver, String UnitsofMeasure, String RelativeHumidity, String WetBulb,
        String Pressurealtitud, String PressurealtitudV, String Showbypass, String Reactivationinputtype,
        String AirflowRange, String Dewpointrange, String Performancesafetyfactor, String PerformancesafetyfactorV,
        String ProcessAirflow, String DesiccantNedia, String SectorLayout, String RotorDiameter, String RotorDepth,
        String NetFaceAreaCalculation, String SealingArea, String ProcessStrC, String ProcessStrGKG,
        String Rph, String Reactivation1, String Reactivation2, String Reactivation3) throws InterruptedException {
        WebDriver page = driver;

        String base = "//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div";

        // 范围1
        appendLog("[范围1] Units=" + UnitsofMeasure + ", Humidity=" + RelativeHumidity + ", WetBulb=" + WetBulb
            + ", Pressure=" + Pressurealtitud + "(" + PressurealtitudV + ")", "");
        if (StrUtil.equalsIgnoreCase(UnitsofMeasure, "si")) {
            clickByXpath(page, base + "/div[1]/div/label[1]/span/input");
        } else {
            clickByXpath(page, base + "/div[1]/div/label[2]/span/input");
        }
        if (!StrUtil.equalsIgnoreCase(RelativeHumidity, "勾选")) {
            clickByXpath(page, base + "/div[2]/div/div/label[1]/span/input");
        }
        if (!StrUtil.equalsIgnoreCase(WetBulb, "勾选")) {
            clickByXpath(page, base + "/div[2]/div/div/label[2]/span/input");
        }
        if (StrUtil.equalsIgnoreCase(Pressurealtitud, "Altitude")) {
            clickByXpath(page, base + "/div[3]/div/div[1]/label[1]/span/input");
        } else {
            clickByXpath(page, base + "/div[3]/div/div[1]/label[2]/span/input");
        }
        fillInputByXpath(page, base + "/div[3]/div/div[2]/input", PressurealtitudV);
        appendLog("[范围1] Bypass=" + Showbypass + ", ReactType=" + Reactivationinputtype
            + ", AirflowRange=" + AirflowRange + ", DewpointRange=" + Dewpointrange, "");
        if (StrUtil.equalsIgnoreCase(Showbypass, "No")) {
            clickByXpath(page, base + "/div[4]/div/div/label[1]/span/input");
        } else {
            clickByXpath(page, base + "/div[4]/div/div/label[2]/span/input");
        }
        if (StrUtil.equalsIgnoreCase(Reactivationinputtype, "Temp")) {
            clickByXpath(page, base + "/div[5]/div/div/label[1]/span/input");
        } else {
            clickByXpath(page, base + "/div[5]/div/div/label[2]/span/input");
        }
        if (StrUtil.equalsIgnoreCase(AirflowRange, "Default")) {
            clickByXpath(page, base + "/div[6]/div/div/label[1]/span/input");
        } else {
            clickByXpath(page, base + "/div[6]/div/div/label[2]/span/input");
        }
        if (StrUtil.equalsIgnoreCase(Dewpointrange, "Default")) {
            clickByXpath(page, base + "/div[7]/div/div/label[1]/span/input");
        } else {
            clickByXpath(page, base + "/div[7]/div/div/label[2]/span/input");
        }
        // Performance safety factor
        clickByXpath(page, base + "/div[8]/div/div/div/div");
        ThreadUtil.safeSleep(500);
        // 获取下拉选项
        List<String> psfOptions = getOptionTexts(page);
        appendLog("[下拉] Performance Safety Factor: " + Performancesafetyfactor + " (选项:" + psfOptions + ")", "");
        boolean selected = false;
        if (StrUtil.equalsIgnoreCase(Performancesafetyfactor, "None")) { selectOptionByIndex(page, 0); selected = true; }
        else if (StrUtil.equalsIgnoreCase(Performancesafetyfactor, "+Δ% Moisture")) { selectOptionByIndex(page, 1); selected = true; }
        else if (StrUtil.equalsIgnoreCase(Performancesafetyfactor, "x Moisture")) { selectOptionByIndex(page, 2); selected = true; }
        if (!selected) selectOptionByIndex(page, 0);
        if (!StrUtil.equalsIgnoreCase(Performancesafetyfactor, "None")) {
            clickByXpath(page, base + "/div[8]/div/div/div[2]/div/div");
            ThreadUtil.safeSleep(500);
            selectOptionByDataValue(page, PerformancesafetyfactorV);
            appendLog("[下拉] Safety Factor Value=" + PerformancesafetyfactorV, "");
        }

        // 范围2
        String base2 = "//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div";
        fillInputByXpath(page, base2 + "/div[1]/div/div[2]/input", ProcessAirflow);
        appendLog("[范围2] ProcessAirflow=" + ProcessAirflow, "");
        // Media
        clickByXpath(page, base2 + "/div[3]/div/div[2]");
        ThreadUtil.safeSleep(1000);
        selectMediaOption(page, DesiccantNedia);
        appendLog("[下拉] Media=" + DesiccantNedia, "");
        // Sector Layout
        Platform.runLater(() -> appendLog("[Sector Layout] " + SectorLayout, ""));
        clickByXpath(page, base2 + "/div[4]/div/div[2]");
        ThreadUtil.safeSleep(1500);
        selectOptionByText(page, SectorLayout, 3);
        appendLog("[下拉] Sector Layout 已选择: " + SectorLayout, "");
        verifySelectedText(driver, "Sector Layout", base2 + "/div[4]/div/div[2]", SectorLayout);
        // Rotor diameter/depth
        boolean isCustomSector = StrUtil.equalsAnyIgnoreCase(SectorLayout, "Custom 2-sector", "Custom 3 sector");
        if (isCustomSector) {
            fillInputByXpath(page, base2 + "/div[6]/div/div[2]/input", RotorDiameter);
        } else {
            for (int retry = 0; retry < 3; retry++) {
                try {
                    clickByXpath(page, base2 + "/div[6]/div/div[2]");
                    ThreadUtil.safeSleep(1000);
                    selectOptionByDataValue(page, RotorDiameter);
                    appendLog("[下拉] Rotor Diameter=" + RotorDiameter, "");
                    break;
                } catch (Exception e) { ThreadUtil.safeSleep(1000); }
            }
        }
        if (isCustomSector) {
            fillInputByXpath(page, base2 + "/div[7]/div/div[2]/input", RotorDepth);
        } else {
            for (int retry = 0; retry < 3; retry++) {
                try {
                    clickByXpath(page, base2 + "/div[7]/div/div[2]");
                    ThreadUtil.safeSleep(1000);
                    selectOptionByDataValue(page, RotorDepth);
                    appendLog("[下拉] Rotor Depth=" + RotorDepth, "");
                    break;
                } catch (Exception e) { ThreadUtil.safeSleep(1000); }
            }
        }
        // Net face area
        clickByXpath(page, base2 + "/div[8]/div/div[2]");
        ThreadUtil.safeSleep(1000);
        selectNetFaceArea(page, NetFaceAreaCalculation);
        appendLog("[下拉] Net Face Area=" + NetFaceAreaCalculation + ", SealingArea=" + SealingArea, "");
        fillInputByXpath(page, base2 + "/div[9]/div/div[2]/input", SealingArea);
        // 底部参数
        Platform.runLater(() -> appendLog("[填充值] C=" + ProcessStrC + ", GKG=" + ProcessStrGKG + ", Rph=" + Rph, ""));
        String processCXpath = "//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[1]/div/div/input";
        String processGkgXpath = "//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[2]/div/div/input";
        String rphXpath = "//*[@id=\"root\"]/div/div/div[1]/div[2]/div[8]/div/div[2]/div/div/div/div/input";
        fillInputByXpath(page, processCXpath, ProcessStrC);
        verifyInput(driver, "C", processCXpath, ProcessStrC);
        fillInputByXpath(page, processGkgXpath, ProcessStrGKG);
        verifyInput(driver, "GKG", processGkgXpath, ProcessStrGKG);
        fillInputByXpath(page, rphXpath, Rph);
        verifyInput(driver, "RPH", rphXpath, Rph);
        fillInputByXpath(page, "//*[@id=\"root\"]/div/div/div[1]/div[2]/div[5]/div/div[2]/div/div[1]/div/div/input", Reactivation1);
        fillInputByXpath(page, "//*[@id=\"root\"]/div/div/div[1]/div[2]/div[5]/div/div[2]/div/div[2]/div/div/input", Reactivation2);
        fillInputByXpath(page, "//*[@id=\"root\"]/div/div/div[1]/div[2]/div[5]/div/div[2]/div/div[5]/div/div/input", Reactivation3);
        appendLog("[填充] Reactivation: " + Reactivation1 + ", " + Reactivation2 + ", " + Reactivation3, "");
        appendLog("[参数设置完成] 所有参数已填充到网页", "info");
    }

    // ===== Selenium 辅助方法 =====

    private String getInputValueByXpath(WebDriver driver, String xpath) {
        return driver.findElement(By.xpath(xpath)).getAttribute("value");
    }

    private void fillInputByXpath(WebDriver driver, String xpath, String value) {
        WebElement element = driver.findElement(By.xpath(xpath));
        ((JavascriptExecutor) driver).executeScript(
            "var el=arguments[0], value=arguments[1]; el.focus();"
                + "var setter=Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype,'value').set;"
                + "setter.call(el,value); el.dispatchEvent(new Event('input',{bubbles:true}));"
                + "el.dispatchEvent(new Event('change',{bubbles:true}));",
            element, value);
    }

    private void clickByXpath(WebDriver driver, String xpath) {
        WebElement element = driver.findElement(By.xpath(xpath));
        try {
            element.click();
        } catch (Exception e) {
            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", element);
        }
    }

    private List<WebElement> visibleOptions(WebDriver driver) {
        List<WebElement> visible = new ArrayList<>();
        for (WebElement option : driver.findElements(By.cssSelector("li[role='option']"))) {
            if (option.isDisplayed()) visible.add(option);
        }
        return visible;
    }

    private List<String> getOptionTexts(WebDriver driver) {
        List<String> values = new ArrayList<>();
        for (WebElement option : visibleOptions(driver)) values.add(option.getAttribute("data-value"));
        return values;
    }

    private void selectOptionByIndex(WebDriver driver, int index) {
        List<WebElement> options = visibleOptions(driver);
        if (options.size() <= index) throw new IllegalStateException("下拉选项未就绪");
        options.get(index).click();
    }

    private void selectOptionByDataValue(WebDriver driver, String value) {
        for (WebElement option : visibleOptions(driver)) {
            if (StrUtil.equals(option.getAttribute("data-value"), value)) {
                option.click();
                return;
            }
        }
        throw new IllegalStateException("未找到下拉选项: " + value);
    }

    private void selectOptionByText(WebDriver driver, String value, int maxRetry) {
        for (int retry = 0; retry < maxRetry; retry++) {
            for (WebElement option : visibleOptions(driver)) {
                if (StrUtil.equalsIgnoreCase(option.getText().trim(), value)) {
                    option.click();
                    return;
                }
            }
            ThreadUtil.safeSleep(1000);
        }
        throw new IllegalStateException("未找到下拉选项: " + value);
    }

    private void selectMediaOption(WebDriver driver, String media) {
        for (WebElement option : visibleOptions(driver)) {
            String label = StrUtil.equals(option.getAttribute("data-value"), "1") ? "PPS" : "PPP";
            if (StrUtil.equalsIgnoreCase(label, media)) {
                option.click();
                return;
            }
        }
        throw new IllegalStateException("未找到干燥剂选项: " + media);
    }

    private void selectNetFaceArea(WebDriver driver, String area) {
        for (WebElement option : visibleOptions(driver)) {
            String label = StrUtil.equals(option.getAttribute("data-value"), "0") ? "Sealing area" : "Active area";
            if (StrUtil.equalsIgnoreCase(label, area)) {
                option.click();
                return;
            }
        }
        throw new IllegalStateException("未找到净面面积选项: " + area);
    }

    private void verifyInput(WebDriver driver, String name, String xpath, String expected) {
        String actual = getInputValueByXpath(driver, xpath);
        if (!StrUtil.equals(actual == null ? "" : actual.trim(), expected == null ? "" : expected.trim())) {
            throw new IllegalStateException(name + " 写入失败，期望: " + expected + "，实际: " + actual);
        }
        Platform.runLater(() -> appendLog("[已确认] " + name + " = " + actual, ""));
    }

    private void verifySelectedText(WebDriver driver, String name, String xpath, String expected) {
        String actual = driver.findElement(By.xpath(xpath)).getText().trim();
        if (!actual.contains(expected)) {
            throw new IllegalStateException(name + " 选择失败，期望: " + expected + "，实际: " + actual);
        }
        Platform.runLater(() -> appendLog("[已确认] " + name + " = " + actual, ""));
    }


    private List<Object> collectRowDataSelenium(WebDriver driver, String lineNumber) {
        return BaiscApplication.collectRowData(driver, lineNumber);
    }

    private void toListSelenium(WebDriver driver, StringBuilder ss, String lineNumber,
        ExcelWriter excelWriter, String timestamp, boolean isHit) {
        List<Object> row = collectRowDataSelenium(driver, lineNumber);
        row.add(1, timestamp);
        ss.append(row).append("\r\n");
        int rowIdx = writerRowIndex++;
        excelWriter.writeRow(row);
        if (isHit) applyRowStyle(excelWriter, rowIdx, row.size(), true, false);
    }

    // ===== 样式和日志 =====

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

    @FXML
    public void onClearLog() {
        logTextArea.clear();
    }

    @FXML
    public void onCopyLog() {
        String text = logTextArea.getSelectedText();
        if (text != null && !text.isEmpty()) {
            javafx.scene.input.Clipboard clipboard = javafx.scene.input.Clipboard.getSystemClipboard();
            javafx.scene.input.ClipboardContent content = new javafx.scene.input.ClipboardContent();
            content.putString(text);
            clipboard.setContent(content);
            return;
        }
        logTextArea.selectAll();
        logTextArea.copy();
        logTextArea.deselect();
    }

    private void startOutputRedirect() {
        final OutputStream logStream = new OutputStream() {
            private final StringBuilder lineBuffer = new StringBuilder();
            @Override
            public void write(int b) {
                if (b == '\n') {
                    final String line = lineBuffer.toString().trim();
                    lineBuffer.setLength(0);
                    if (!line.isEmpty()) {
                        Platform.runLater(() -> {
                            logTextArea.appendText(line + "\n");
                            logTextArea.setScrollTop(Double.MAX_VALUE);
                        });
                    }
                } else if (b != '\r') {
                    lineBuffer.append((char) b);
                }
            }
        };
        System.setOut(new PrintStream(logStream, true));
        System.setErr(new PrintStream(logStream, true));
    }

    private void appendLog(String msg, String type) {
        Platform.runLater(() -> {
            String prefix;
            if ("hit".equals(type)) prefix = "✓ ";
            else if ("warn".equals(type)) prefix = "⚠ ";
            else if ("error".equals(type)) prefix = "✗ ";
            else if ("info".equals(type)) prefix = "● ";
            else prefix = "  ";
            String time = String.format("[%02d:%02d:%02d] ",
                (int)((System.currentTimeMillis() - startTime) / 3600000),
                (int)(((System.currentTimeMillis() - startTime) / 60000) % 60),
                (int)(((System.currentTimeMillis() - startTime) / 1000) % 60));
            logTextArea.appendText(time + prefix + msg + "\n");
            logTextArea.setScrollTop(Double.MAX_VALUE);
        });
    }
}
