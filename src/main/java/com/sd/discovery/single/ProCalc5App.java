package com.sd.discovery.single;

import cn.hutool.core.date.DateUtil;
import cn.hutool.core.thread.ThreadUtil;
import cn.hutool.core.util.NumberUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import com.google.common.collect.Lists;
import com.microsoft.playwright.*;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.PrintStream;
import java.util.Arrays;
import java.util.List;

/**
 * ProCalc5 自动化计算工具 - macOS GUI 启动器 (Playwright版)
 */
public class ProCalc5App {

    private JFrame frame;
    private JTextField inputFilePath;
    private JTextField outputDirPath;
    private JTextArea logArea;
    private JButton runButton;
    private JButton browseInputButton;
    private JButton browseOutputButton;
    private volatile boolean running = false;

    public static void main(String[] args) {
        if (args.length >= 2) {
            String inputFile = args[0];
            String outputDir = args[1];
            System.out.println("命令行模式启动");
            System.out.println("输入文件: " + inputFile);
            System.out.println("输出目录: " + outputDir);
            ProCalc5App app = new ProCalc5App();
            try {
                app.runProCalc5(inputFile, outputDir);
                System.out.println("\n✅ 任务执行完成！");
            } catch (Exception ex) {
                System.out.println("\n❌ 任务执行出错: " + ex.getMessage());
                ex.printStackTrace();
                System.exit(1);
            }
            System.exit(0);
            return;
        }
        SwingUtilities.invokeLater(() -> {
            try { UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName()); } catch (Exception ignored) {}
            new ProCalc5App().createAndShowGUI();
        });
    }

    // ========== GUI 部分 (不变) ==========

    private void createAndShowGUI() {
        frame = new JFrame("ProCalc5 自动化计算工具");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(700, 550);
        frame.setLocationRelativeTo(null);
        frame.setResizable(true);

        JPanel mainPanel = new JPanel(new BorderLayout(10, 10));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(15, 15, 15, 15));

        JPanel paramPanel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(5, 5, 5, 5);

        gbc.gridx = 0; gbc.gridy = 0;
        paramPanel.add(new JLabel("输入 Excel 文件:"), gbc);
        inputFilePath = new JTextField(30);
        inputFilePath.setText(getDefaultInputPath());
        gbc.gridx = 1; gbc.weightx = 1.0;
        paramPanel.add(inputFilePath, gbc);
        browseInputButton = new JButton("浏览...");
        browseInputButton.addActionListener(e -> browseInputFile());
        gbc.gridx = 2; gbc.weightx = 0;
        paramPanel.add(browseInputButton, gbc);

        gbc.gridx = 0; gbc.gridy = 1; gbc.weightx = 0;
        paramPanel.add(new JLabel("输出目录:"), gbc);
        outputDirPath = new JTextField(30);
        outputDirPath.setText(getDefaultOutputDir());
        gbc.gridx = 1; gbc.weightx = 1.0;
        paramPanel.add(outputDirPath, gbc);
        browseOutputButton = new JButton("浏览...");
        browseOutputButton.addActionListener(e -> browseOutputDir());
        gbc.gridx = 2; gbc.weightx = 0;
        paramPanel.add(browseOutputButton, gbc);

        mainPanel.add(paramPanel, BorderLayout.NORTH);

        logArea = new JTextArea();
        logArea.setEditable(false);
        logArea.setFont(new Font("Menlo", Font.PLAIN, 12));
        logArea.setBackground(new Color(30, 30, 30));
        logArea.setForeground(new Color(0, 255, 0));
        logArea.setCaretColor(Color.WHITE);
        JScrollPane scrollPane = new JScrollPane(logArea);
        scrollPane.setBorder(BorderFactory.createTitledBorder("运行日志"));
        mainPanel.add(scrollPane, BorderLayout.CENTER);

        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.CENTER, 15, 5));
        runButton = new JButton("▶ 开始运行");
        runButton.setFont(new Font("Dialog", Font.BOLD, 14));
        runButton.setPreferredSize(new java.awt.Dimension(150, 38));
        runButton.addActionListener(e -> startTask());
        buttonPanel.add(runButton);
        JButton clearButton = new JButton("清空日志");
        clearButton.addActionListener(e -> logArea.setText(""));
        buttonPanel.add(clearButton);
        mainPanel.add(buttonPanel, BorderLayout.SOUTH);

        frame.setContentPane(mainPanel);
        redirectOutput();

        frame.addWindowListener(new WindowAdapter() {
            @Override
            public void windowClosing(WindowEvent e) {
                if (running) {
                    int result = JOptionPane.showConfirmDialog(frame, "任务正在运行中，确定退出吗？", "确认退出", JOptionPane.YES_NO_OPTION);
                    if (result != JOptionPane.YES_OPTION) return;
                }
                System.exit(0);
            }
        });

        frame.setVisible(true);
        log("ProCalc5 自动化计算工具已启动 (Playwright版)");
        log("请选择输入 Excel 文件后点击「开始运行」");
    }

    private String getDefaultInputPath() {
        String userDir = System.getProperty("user.dir");
        File defaultFile = new File(userDir, "procalc5/procalc5.proflute.xlsx");
        return defaultFile.exists() ? defaultFile.getAbsolutePath() : "";
    }

    private String getDefaultOutputDir() {
        String userDir = System.getProperty("user.dir");
        return new File(userDir, "procalc5").getAbsolutePath();
    }

    private void browseInputFile() {
        JFileChooser chooser = new JFileChooser();
        chooser.setFileFilter(new FileNameExtensionFilter("Excel 文件 (*.xlsx)", "xlsx"));
        if (!inputFilePath.getText().isEmpty()) chooser.setCurrentDirectory(new File(inputFilePath.getText()).getParentFile());
        int result = chooser.showOpenDialog(frame);
        if (result == JFileChooser.APPROVE_OPTION) inputFilePath.setText(chooser.getSelectedFile().getAbsolutePath());
    }

    private void browseOutputDir() {
        JFileChooser chooser = new JFileChooser();
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        if (!outputDirPath.getText().isEmpty()) chooser.setCurrentDirectory(new File(outputDirPath.getText()));
        int result = chooser.showOpenDialog(frame);
        if (result == JFileChooser.APPROVE_OPTION) outputDirPath.setText(chooser.getSelectedFile().getAbsolutePath());
    }

    private void startTask() {
        String inputFile = inputFilePath.getText().trim();
        String outputDir = outputDirPath.getText().trim();
        if (inputFile.isEmpty() || !new File(inputFile).exists()) {
            JOptionPane.showMessageDialog(frame, "请选择有效的输入 Excel 文件！", "错误", JOptionPane.ERROR_MESSAGE);
            return;
        }
        if (outputDir.isEmpty()) {
            outputDir = new File(inputFile).getParent();
            outputDirPath.setText(outputDir);
        }
        final String finalInputFile = inputFile;
        final String finalOutputDir = outputDir;
        running = true;
        setUIEnabled(false);
        logArea.setText("");
        log("========================================");
        log("开始执行自动化计算任务...");
        log("输入文件: " + inputFile);
        log("输出目录: " + outputDir);
        log("========================================");

        Thread taskThread = new Thread(() -> {
            try {
                runProCalc5(finalInputFile, finalOutputDir);
                log("\n✅ 任务执行完成！");
                SwingUtilities.invokeLater(() -> JOptionPane.showMessageDialog(frame,
                        "计算任务已完成！\n结果保存在: " + finalOutputDir + "/result.xlsx", "完成", JOptionPane.INFORMATION_MESSAGE));
            } catch (Exception ex) {
                log("\n❌ 任务执行出错: " + ex.getMessage());
                ex.printStackTrace();
                SwingUtilities.invokeLater(() -> JOptionPane.showMessageDialog(frame,
                        "执行出错: " + ex.getMessage(), "错误", JOptionPane.ERROR_MESSAGE));
            } finally {
                running = false;
                SwingUtilities.invokeLater(() -> setUIEnabled(true));
            }
        });
        taskThread.setDaemon(true);
        taskThread.start();
    }

    private void setUIEnabled(boolean enabled) {
        runButton.setEnabled(enabled);
        browseInputButton.setEnabled(enabled);
        browseOutputButton.setEnabled(enabled);
        inputFilePath.setEnabled(enabled);
        outputDirPath.setEnabled(enabled);
        runButton.setText(enabled ? "▶ 开始运行" : "⏳ 运行中...");
    }

    private void redirectOutput() {
        PrintStream printStream = new PrintStream(new java.io.OutputStream() {
            private final StringBuilder buffer = new StringBuilder();
            @Override
            public void write(int b) {
                buffer.append((char) b);
                if (b == '\n' || buffer.length() > 200) {
                    final String text = buffer.toString();
                    buffer.setLength(0);
                    SwingUtilities.invokeLater(() -> { logArea.append(text); logArea.setCaretPosition(logArea.getDocument().getLength()); });
                }
            }
            @Override
            public void flush() {
                if (buffer.length() > 0) {
                    final String text = buffer.toString();
                    buffer.setLength(0);
                    SwingUtilities.invokeLater(() -> { logArea.append(text); logArea.setCaretPosition(logArea.getDocument().getLength()); });
                }
            }
        }, true);
        System.setOut(printStream);
        System.setErr(printStream);
    }

    private void log(String msg) {
        if (logArea != null) {
            logArea.append(msg + "\n");
            logArea.setCaretPosition(logArea.getDocument().getLength());
        } else {
            System.out.println(msg);
        }
    }

    // ========== Playwright 辅助方法 ==========

    private String getInputValueByIndex(Page page, int index) {
        Object result = page.evaluate("() => { var el = document.querySelectorAll('input[type=text]')[" + index + "]; return el ? el.value : ''; }");
        return result != null ? result.toString() : "";
    }

    private String getInputValueByXpath(Page page, String xpath) {
        String jsXpath = xpath.replace("'", "\\'");
        Object result = page.evaluate(
            "() => { var r = document.evaluate('" + jsXpath + "', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null); var el = r.singleNodeValue; return el ? el.value : ''; }");
        return result != null ? result.toString() : "";
    }

    private void fillInputByIndex(Page page, int index, String value) {
        page.evaluate("(val) => { var el = document.querySelectorAll('input[type=text]')[" + index + "]; if(el) { el.focus(); el.value=''; } }", value);
        page.locator("input[type='text']").nth(index).fill(value);
        ThreadUtil.safeSleep(500);
    }

    /**
     * 选择 MUI Select 下拉框选项
     * 策略: Playwright click 打开下拉框 → 用 evaluate 找到目标选项坐标 → 用 Playwright mouse.click 点击选项中心
     */
    private void selectComboboxByDataValue(Page page, int comboboxIndex, String dataValue) {
        page.locator("div[role='combobox']").nth(comboboxIndex).click();
        ThreadUtil.safeSleep(1500);
        try {
            Locator matched = page.locator("li[role='option'][data-value='" + dataValue + "']");
            if (matched.count() > 0) {
                matched.first().click();
                log("  [DEBUG] combo[" + comboboxIndex + "] dataValue=" + dataValue + " locator.click OK");
            } else {
                page.locator("li[role='option']").first().click();
                log("  [DEBUG] combo[" + comboboxIndex + "] dataValue=" + dataValue + " fallback first");
            }
        } catch (Exception e) {
            log("  [WARN] combo[" + comboboxIndex + "] click 失败: " + e.getMessage());
        }
        ThreadUtil.safeSleep(500);
    }

    private void selectComboboxByText(Page page, int comboboxIndex, String textContains) {
        // 先打开下拉框
        page.locator("div[role='combobox']").nth(comboboxIndex).click();
        ThreadUtil.safeSleep(1500);
        // 调试: 打印下拉选项信息
        String debugInfo = (String) page.evaluate(
            "() => {" +
            "  var all = document.querySelectorAll('li[role=option]');" +
            "  var vis = [];" +
            "  for (var i = 0; i < all.length; i++) { if (all[i].offsetParent !== null) vis.push(all[i]); }" +
            "  var r = 'total=' + all.length + ' vis=' + vis.length;" +
            "  for (var i = 0; i < Math.min(5, vis.length); i++) {" +
            "    var rect = vis[i].getBoundingClientRect();" +
            "    r += ' [' + vis[i].getAttribute('data-value') + '|' + vis[i].textContent.trim().substring(0,15) + '|y=' + Math.round(rect.top) + ']';" +
            "  }" +
            "  return r;" +
            "}");
        log("  [DEBUG-DROPDOWN] combo[" + comboboxIndex + "] " + debugInfo);
        // 方法: 先尝试 Playwright locator.click (non-force)
        try {
            Locator matched = page.locator("li[role='option']").filter(new Locator.FilterOptions().setHasText(textContains));
            int matchCount = matched.count();
            log("  [DEBUG] combo[" + comboboxIndex + "] text='" + textContains + "' matched=" + matchCount);
            if (matchCount > 0) {
                matched.first().click();
                log("  [DEBUG] combo[" + comboboxIndex + "] locator.click OK");
            } else {
                page.locator("li[role='option']").first().click();
                log("  [DEBUG] combo[" + comboboxIndex + "] fallback first option click");
            }
        } catch (Exception e) {
            log("  [WARN] combo[" + comboboxIndex + "] locator.click 失败: " + e.getMessage());
        }
        ThreadUtil.safeSleep(500);
        // 验证选择是否成功: 检查combobox文本内容
        String currentText = (String) page.evaluate(
            "(idx) => {" +
            "  var combos = document.querySelectorAll('div[role=combobox]');" +
            "  return combos[idx] ? combos[idx].textContent.trim() : '';" +
            "}", comboboxIndex);
        log("  [DEBUG-VERIFY] combo[" + comboboxIndex + "] after click text='" + currentText + "'");
        // 如果选择失败(文本为空)，使用键盘导航重试
        if (currentText == null || currentText.isEmpty() || currentText.equals("\u200B") || currentText.equals("\u200C")) {
            log("  [DEBUG] combo[" + comboboxIndex + "] click未生效，尝试键盘导航...");
            // 重新打开下拉框
            page.locator("div[role='combobox']").nth(comboboxIndex).click();
            ThreadUtil.safeSleep(1000);
            // 获取选项列表和索引
            String optionInfo = (String) page.evaluate(
                "(searchText) => {" +
                "  var options = document.querySelectorAll('li[role=option]');" +
                "  var vis = [];" +
                "  for (var i = 0; i < options.length; i++) {" +
                "    if (options[i].offsetParent !== null) {" +
                "      vis.push({text: options[i].textContent.trim(), idx: i});" +
                "    }" +
                "  }" +
                "  var targetIdx = -1;" +
                "  for (var i = 0; i < vis.length; i++) {" +
                "    if (vis[i].text.indexOf(searchText) >= 0) { targetIdx = i; break; }" +
                "  }" +
                "  if (targetIdx < 0) targetIdx = 0;" +
                "  return targetIdx + ':' + vis.length;" +
                "}", textContains);
            log("  [DEBUG] combo[" + comboboxIndex + "] keyboard nav info='" + optionInfo + "'");
            if (optionInfo != null && optionInfo.contains(":")) {
                int targetIdx = Integer.parseInt(optionInfo.split(":")[0]);
                int totalVis = Integer.parseInt(optionInfo.split(":")[1]);
                // 用键盘下箭头导航到目标选项
                for (int i = 0; i < targetIdx; i++) {
                    page.keyboard().press("ArrowDown");
                    ThreadUtil.safeSleep(50);
                }
                ThreadUtil.safeSleep(200);
                page.keyboard().press("Enter");
                log("  [DEBUG] combo[" + comboboxIndex + "] keyboard nav: ArrowDown x" + targetIdx + " + Enter");
            }
            ThreadUtil.safeSleep(500);
        }
    }

    private void selectComboboxByDataValueWithFallback(Page page, int comboboxIndex, String dataValue) {
        page.locator("div[role='combobox']").nth(comboboxIndex).click();
        ThreadUtil.safeSleep(1000);
        String clickResult = (String) page.evaluate(
            "(dv) => {" +
            "  var allOpts = document.querySelectorAll('li[role=option]');" +
            "  var options = [];" +
            "  for (var i = 0; i < allOpts.length; i++) { if (allOpts[i].offsetParent !== null) options.push(allOpts[i]); }" +
            "  if (options.length === 0) return 'no-options';" +
            "  for (var i = 0; i < options.length; i++) {" +
            "    if (options[i].getAttribute('data-value') === dv) {" +
            "      var r = options[i].getBoundingClientRect();" +
            "      return 'click:' + (r.left + r.width/2) + ',' + (r.top + r.height/2);" +
            "    }" +
            "  }" +
            "  var target = parseFloat(dv);" +
            "  if (!isNaN(target)) {" +
            "    var bestIdx = 0, bestVal = -1;" +
            "    for (var i = 0; i < options.length; i++) {" +
            "      var v = parseFloat(options[i].getAttribute('data-value'));" +
            "      if (!isNaN(v) && v <= target && v > bestVal) { bestVal = v; bestIdx = i; }" +
            "    }" +
            "    var r = options[bestIdx].getBoundingClientRect();" +
            "    return 'click:' + (r.left + r.width/2) + ',' + (r.top + r.height/2);" +
            "  }" +
            "  var r = options[0].getBoundingClientRect();" +
            "  return 'click:' + (r.left + r.width/2) + ',' + (r.top + r.height/2);" +
            "}", dataValue);
        if (clickResult != null && clickResult.startsWith("click:")) {
            String[] coords = clickResult.substring(6).split(",");
            page.mouse().click(Double.parseDouble(coords[0]), Double.parseDouble(coords[1]));
            log("  [DEBUG] combo[" + comboboxIndex + "] target=" + dataValue + " mouse.click");
        } else {
            log("  [WARN] combo[" + comboboxIndex + "] 未找到选项: " + clickResult);
        }
        ThreadUtil.safeSleep(500);
    }

    private void clickRadioBySelector(Page page, String cssSelector) {
        try {
            // 先等待元素出现
            page.waitForSelector(cssSelector, new Page.WaitForSelectorOptions().setTimeout(15000));
            page.locator(cssSelector).first().click();
            log("  [DEBUG] radio click OK: " + cssSelector);
        } catch (Exception e) {
            log("  [WARN] radio locator.click 失败，尝试JS点击: " + e.getMessage().split("\n")[0]);
            try {
                page.evaluate(
                    "(selector) => {" +
                    "  var el = document.querySelector(selector);" +
                    "  if (el) { el.click(); return 'ok'; }" +
                    "  return 'not found';" +
                    "}", cssSelector);
                log("  [DEBUG] radio JS click done: " + cssSelector);
            } catch (Exception e2) {
                log("  [ERROR] radio JS click 也失败: " + e2.getMessage().split("\n")[0]);
            }
        }
        ThreadUtil.safeSleep(300);
    }

    private void fillInputByLabel(Page page, String labelText, String value) {
        // MUI label 没有 for 属性，输入框在 label 的祖父级元素中
        String inputSelector = (String) page.evaluate(
            "(searchText) => {" +
            "  var labels = document.querySelectorAll('label, div');" +
            "  var lower = searchText.toLowerCase();" +
            "  for (var i = 0; i < labels.length; i++) {" +
            "    if (labels[i].textContent.trim().toLowerCase() === lower) {" +
            "      var el = labels[i];" +
            "      for (var level = 0; level < 4; level++) {" +
            "        el = el.parentElement;" +
            "        if (!el) break;" +
            "        var input = el.querySelector('input[type=text]:not([readonly])');" +
            "        if (input) {" +
            "          if (input.id) return '#id:' + input.id;" +
            "          input.setAttribute('data-pw-found', searchText);" +
            "          return '#attr:' + searchText;" +
            "        }" +
            "      }" +
            "    }" +
            "  }" +
            "  return '';" +
            "}", labelText);
        if (inputSelector != null && !inputSelector.isEmpty()) {
            if (inputSelector.startsWith("#id:")) {
                String id = inputSelector.substring(4);
                page.locator("#" + id).fill(value);
            } else if (inputSelector.startsWith("#attr:")) {
                String attr = inputSelector.substring(6);
                page.locator("input[data-pw-found='" + attr + "']").fill(value);
            }
        } else {
            log("  [WARN] 找不到 '" + labelText + "' 输入框");
        }
        ThreadUtil.safeSleep(300);
    }

    private void fillInputByAdornment(Page page, String adornment, String value) {
        page.evaluate(
            "(args) => {" +
            "  var adornments = document.querySelectorAll('.MuiInputAdornment-root');" +
            "  for (var i = 0; i < adornments.length; i++) {" +
            "    if (adornments[i].textContent.trim() === args[0]) {" +
            "      var container = adornments[i].closest('.MuiInputBase-root') || adornments[i].parentElement.parentElement;" +
            "      var input = container.querySelector('input');" +
            "      if (input && !input.readOnly) { input.focus(); return input; }" +
            "    }" +
            "  }" +
            "  return null;" +
            "}", java.util.Arrays.asList(adornment, value));
        // Use JS to find and fill
        String evalResult = (String) page.evaluate(
            "(args) => {" +
            "  var adornments = document.querySelectorAll('.MuiInputAdornment-root');" +
            "  for (var i = 0; i < adornments.length; i++) {" +
            "    if (adornments[i].textContent.trim() === args[0]) {" +
            "      var container = adornments[i].closest('.MuiInputBase-root') || adornments[i].parentElement.parentElement;" +
            "      var input = container.querySelector('input');" +
            "      if (input && !input.readOnly) { return 'found'; }" +
            "    }" +
            "  }" +
            "  return 'not-found';" +
            "}", java.util.Arrays.asList(adornment));
        if ("found".equals(evalResult)) {
            // Click the adornment's input container, then fill
            page.locator(".MuiInputAdornment-root").locator("text=" + adornment).locator("..").locator("..").locator("input").fill(value);
        }
    }

    /**
     * 设置 Reactivation 输入值并等待自动计算
     */
    private String readGkgOutput(Page page) {
        // 从 Process Right card (div[7]) 读取 g/kg 输出 (index 1)
        // Process Right g/kg 随 Reactivation 温度变化而变化
        Object result = page.evaluate(
            "() => {" +
            "  var card = document.evaluate(\"//*[@id='root']/div/div/div[1]/div[2]/div[7]\", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;" +
            "  if (!card) return '';" +
            "  var inputs = card.querySelectorAll('input[type=text]');" +
            "  return inputs.length > 1 ? inputs[1].value : '';" +
            "}");
        return result != null ? result.toString() : "";
    }

    private void setReactivationValue(Page page, int inputIndex, String value) {
        fillInputByIndex(page, inputIndex, value);
        page.keyboard().press("Tab");
        ThreadUtil.safeSleep(500);
        // 点击 Calculate 按钮触发计算
        try {
            page.locator("button:has-text('Calculate')").click();
        } catch (Exception e) {
            log("  [WARN] Calculate按钮点击失败: " + e.getMessage());
        }
        ThreadUtil.safeSleep(3000);
        String newGkg = readGkgOutput(page);
        log("  [DEBUG] setReactivity=" + value + " gkg=" + newGkg);
    }

    // ========== ProCalc5 自动化核心逻辑 (Playwright版) ==========

    private void runProCalc5(String inputExcelPath, String outputDir) {
        StringBuilder ss = new StringBuilder();
        String sheetName = DateUtil.format(DateUtil.date(), "yyyyMMddHHmmss");

        // 确保输出目录存在且可写
        File outDirFile = new File(outputDir);
        if (!outDirFile.exists()) {
            outDirFile.mkdirs();
        }
        // 检查是否可写（DMG挂载卷为只读）
        if (!outDirFile.canWrite()) {
            String fallback = new File(inputExcelPath).getParent();
            if (fallback == null) fallback = System.getProperty("user.home") + "/Desktop";
            log("[WARN] 输出目录不可写: " + outputDir);
            log("[WARN] 结果将保存到: " + fallback);
            outputDir = fallback;
            outDirFile = new File(outputDir);
            if (!outDirFile.exists()) outDirFile.mkdirs();
        }
        String resultFile = outputDir + File.separator + "result.xlsx";
        log("输出目录: " + outputDir);
        log("结果文件: " + resultFile);

        // 启动 Playwright 浏览器
        Playwright playwright = Playwright.create();
        Browser browser;
        try {
            browser = playwright.chromium().launch(
                new BrowserType.LaunchOptions()
                    .setHeadless(false)
                    .setChannel("chrome")
                    .setSlowMo(100));
        } catch (Exception e) {
            log("Chrome channel 启动失败，尝试默认 Chromium: " + e.getMessage());
            browser = playwright.chromium().launch(
                new BrowserType.LaunchOptions().setHeadless(false).setSlowMo(100));
        }
        BrowserContext context = browser.newContext(new Browser.NewContextOptions().setViewportSize(1920, 1080));
        Page page = context.newPage();
        page.setDefaultTimeout(30000);

        try {
            page.navigate("https://procalc5.proflute.se/rotor");
            ThreadUtil.safeSleep(8000);

            // 登录
            page.fill("#userNameInput", "EXTCNJANZHA");
            page.fill("#passwordInput", "a[d?>9v78ugU");
            page.click("#submitButton");
            ThreadUtil.safeSleep(8000);
            log("已登录 ProCalc5");
            // 等待表单加载（radio按钮出现）
            try {
                page.waitForSelector("input[type='radio']", new Page.WaitForSelectorOptions().setTimeout(30000));
                log("表单已加载");
            } catch (Exception e) {
                log("[WARN] 等待表单超时，继续执行...");
            }
            ThreadUtil.safeSleep(2000);

            ExcelWriter excelWriter = ExcelUtil.getWriter(resultFile, sheetName);
            excelWriter.writeHeadRow(Lists.newArrayList("序号", " Wet Air:", "", "", "", "",
                    "Process left", "", "", "", "", "", "",
                    "Process Right", "", "", "", "", "",
                    "Reactivation", "", "", "", "", "", "", "RPH"));

            List<List<Object>> paraList = ExcelUtil.getReader(inputExcelPath).read();
            int totalRows = paraList.size() - 1;
            log("共 " + totalRows + " 组参数待计算");

            for (List<Object> list : paraList) {
                if (paraList.indexOf(list) == 0) continue;
                int currentRow = paraList.indexOf(list);
                log(String.format("[%d/%d] 正在处理第 %s 组...", currentRow, totalRows, list.get(0)));

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

                try {
                // === 范围1: 基本设置 ===
                // Units of measure
                if (StrUtil.equalsIgnoreCase(UnitsofMeasure, "si")) {
                    clickRadioBySelector(page, "input[type='radio'][value='SI']");
                } else {
                    clickRadioBySelector(page, "input[type='radio'][value='IP']");
                }
                // Relative humidity checkbox
                if (!StrUtil.equalsIgnoreCase(RelativeHumidity, "勾选")) {
                    page.locator("input[name='relativehumidity']").click();
                    ThreadUtil.safeSleep(200);
                }
                // Wet bulb checkbox
                if (!StrUtil.equalsIgnoreCase(WetBulb, "勾选")) {
                    page.locator("input[name='wetbulb']").click();
                    ThreadUtil.safeSleep(200);
                }
                // Pressure/Altitude
                if (StrUtil.equalsIgnoreCase(Pressurealtitud, "Altitude")) {
                    clickRadioBySelector(page, "input[type='radio'][value='Altitude']");
                } else {
                    clickRadioBySelector(page, "input[type='radio'][value='Pressure']");
                }
                // Pressure/Altitude value
                ThreadUtil.safeSleep(500);
                // 使用 JS 查找 Pressure/Altitude 输入框
                page.evaluate(
                    "(val) => {" +
                    "  var inputs = document.querySelectorAll('input[type=text]');" +
                    "  for (var i = 0; i < inputs.length; i++) {" +
                    "    if (!inputs[i].id) continue;" +
                    "    var label = document.querySelector('label[for=\"' + inputs[i].id + '\"]');" +
                    "    if (label) {" +
                    "      var lt = label.textContent.toLowerCase();" +
                    "      if (lt.includes('pressure') || lt.includes('altitude')) {" +
                    "        var nativeSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;" +
                    "        inputs[i].focus();" +
                    "        nativeSetter.call(inputs[i], val);" +
                    "        inputs[i].dispatchEvent(new Event('input', {bubbles:true}));" +
                    "        inputs[i].dispatchEvent(new Event('change', {bubbles:true}));" +
                    "        return 'set';" +
                    "      }" +
                    "    }" +
                    "  }" +
                    "  return 'not-found';" +
                    "}", PressurealtitudV);
                ThreadUtil.safeSleep(200);

                // Show bypass
                if (StrUtil.equalsIgnoreCase(Showbypass, "No")) {
                    clickRadioBySelector(page, "input[type='radio'][value='0'][name=':r6:']");
                } else {
                    clickRadioBySelector(page, "input[type='radio'][value='1'][name=':r6:']");
                }
                // Reactivation input type
                if (StrUtil.equalsIgnoreCase(Reactivationinputtype, "Temp")) {
                    clickRadioBySelector(page, "input[type='radio'][value='Temp']");
                } else {
                    clickRadioBySelector(page, "input[type='radio'][value='Power']");
                }
                // Airflow range
                if (StrUtil.equalsIgnoreCase(AirflowRange, "Default")) {
                    clickRadioBySelector(page, "input[type='radio'][value='0'][name=':r8:']");
                } else {
                    clickRadioBySelector(page, "input[type='radio'][value='1'][name=':r8:']");
                }
                // Dew point range
                if (StrUtil.equalsIgnoreCase(Dewpointrange, "Default")) {
                    clickRadioBySelector(page, "input[type='radio'][value='0'][name=':r9:']");
                } else {
                    clickRadioBySelector(page, "input[type='radio'][value='1'][name=':r9:']");
                }

                // ======================================
                // 策略: 先填充所有文本输入框，最后设置下拉框
                // (因为输入框填充会触发React重新验证并清除下拉框值)
                // ======================================

                // --- 阶段1: 填充所有文本输入框 ---

                // Process Airflow
                fillInputByLabel(page, "Process airflow", ProcessAirflow);
                ThreadUtil.safeSleep(300);

                // Sealing/Active Area input
                if (StrUtil.equalsIgnoreCase(NetFaceAreaCalculation, "Sealing area")) {
                    fillInputByLabel(page, "Sealing area", SealingArea);
                } else {
                    fillInputByLabel(page, "Active area", SealingArea);
                }
                ThreadUtil.safeSleep(300);

                // Process Left 输入
                log("  [DEBUG] Process Left 输入: ProcessStrC=" + ProcessStrC + " ProcessStrGKG=" + ProcessStrGKG);
                page.evaluate(
                    "(args) => {" +
                    "  var card = document.evaluate(\"//*[@id='root']/div/div/div[1]/div[2]/div[6]\", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;" +
                    "  if (!card) return;" +
                    "  var inputs = card.querySelectorAll('input[type=text]:not([readonly])');" +
                    "  var setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;" +
                    "  if (inputs.length >= 2) {" +
                    "    inputs[0].focus(); setter.call(inputs[0], args[0]);" +
                    "    inputs[0].dispatchEvent(new Event('input', {bubbles:true}));" +
                    "    inputs[0].dispatchEvent(new Event('change', {bubbles:true}));" +
                    "    inputs[0].blur();" +
                    "    inputs[1].focus(); setter.call(inputs[1], args[1]);" +
                    "    inputs[1].dispatchEvent(new Event('input', {bubbles:true}));" +
                    "    inputs[1].dispatchEvent(new Event('change', {bubbles:true}));" +
                    "    inputs[1].blur();" +
                    "  } else if (inputs.length >= 1) {" +
                    "    inputs[0].focus(); setter.call(inputs[0], args[0]);" +
                    "    inputs[0].dispatchEvent(new Event('input', {bubbles:true}));" +
                    "    inputs[0].dispatchEvent(new Event('change', {bubbles:true}));" +
                    "    inputs[0].blur();" +
                    "  }" +
                    "}", Arrays.asList(ProcessStrC, ProcessStrGKG));
                ThreadUtil.safeSleep(500);

                // RPH 输入
                page.evaluate(
                    "(val) => {" +
                    "  var adornments = document.querySelectorAll('.MuiInputAdornment-root');" +
                    "  for (var i = 0; i < adornments.length; i++) {" +
                    "    if (adornments[i].textContent.trim() === 'rph') {" +
                    "      var container = adornments[i].closest('.MuiInputBase-root') || adornments[i].parentElement.parentElement;" +
                    "      var input = container.querySelector('input');" +
                    "      if (input && !input.readOnly) {" +
                    "        input.focus();" +
                    "        var nativeSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;" +
                    "        nativeSetter.call(input, val);" +
                    "        input.dispatchEvent(new Event('input', {bubbles:true}));" +
                    "        input.dispatchEvent(new Event('change', {bubbles:true}));" +
                    "        input.blur();" +
                    "      }" +
                    "    }" +
                    "  }" +
                    "}", Rph);
                ThreadUtil.safeSleep(300);

                // Reactivation Ambient SCMH 填充
                String reactAmbientXpath = "//*[@id='root']/div/div/div[1]/div[2]/div[5]";
                Locator reactAmbientCard = page.locator("xpath=" + reactAmbientXpath);
                Locator raInputs = reactAmbientCard.locator("input[type='text']");
                int raCount = raInputs.count();
                log("  [DEBUG] Reactivation Ambient inputs count=" + raCount);
                if (raCount >= 5) {
                    boolean isEditable = (boolean) raInputs.nth(4).evaluate("el => !el.readOnly");
                    if (isEditable) {
                        raInputs.nth(4).fill("2000");
                        log("  [DEBUG] 填充 Reactivation Ambient SCMH=2000");
                    }
                }
                ThreadUtil.safeSleep(500);

                // Reactivation 卡片输入
                String reactCardXpath = "//*[@id='root']/div/div/div[1]/div[2]/div[3]";
                Locator reactCard = page.locator("xpath=" + reactCardXpath);
                Locator reactInputs = reactCard.locator("input[type='text']:not([readonly])");
                int reactInputCount = reactInputs.count();
                log("  [DEBUG] Reactivation inputs count=" + reactInputCount);
                if (reactInputCount >= 3) {
                    reactInputs.nth(0).fill(Reactivation1);
                    ThreadUtil.safeSleep(300);
                    reactInputs.nth(1).fill(Reactivation2);
                    ThreadUtil.safeSleep(300);
                    reactInputs.nth(2).fill(Reactivation3);
                    ThreadUtil.safeSleep(300);
                }

                // --- 阶段2: 设置所有下拉框 (必须在输入框填充之后!) ---
                log("  [DEBUG] 开始设置下拉框...");

                // Performance safety factor dropdown [0]
                selectComboboxByText(page, 0, Performancesafetyfactor);
                ThreadUtil.safeSleep(300);

                // Performance safety factor value (if not None)
                if (!StrUtil.equalsIgnoreCase(Performancesafetyfactor, "None")) {
                    selectComboboxByDataValue(page, 1, PerformancesafetyfactorV);
                    ThreadUtil.safeSleep(300);
                }

                // Desiccant media dropdown [1]
                selectComboboxByText(page, 1, DesiccantNedia);

                // Sector layout dropdown [2]
                selectComboboxByText(page, 2, SectorLayout);

                // Rotor diameter dropdown [4]
                selectComboboxByText(page, 4, RotorDiameter);

                // Rotor depth dropdown [5]
                selectComboboxByText(page, 5, RotorDepth);

                // Net face area calculation dropdown [6]
                if (StrUtil.equalsIgnoreCase(NetFaceAreaCalculation, "Sealing area")) {
                    selectComboboxByText(page, 6, "Sealing");
                } else {
                    selectComboboxByText(page, 6, "Active");
                }

                // 点击Calculate触发初始计算
                try { page.locator("button:has-text('Calculate')").click(); } catch (Exception ignored) {}
                ThreadUtil.safeSleep(3000);

                // 检查页面错误
                String pageErrors = (String) page.evaluate(
                    "() => {" +
                    "  var alerts = document.querySelectorAll('.MuiAlert-root');" +
                    "  var r = '';" +
                    "  for (var i = 0; i < alerts.length; i++) r += alerts[i].textContent.trim() + '; ';" +
                    "  return r;" +
                    "}");
                if (pageErrors != null && !pageErrors.isEmpty()) {
                    log("  [ERROR] 页面错误: " + pageErrors);
                }
                // 检查combobox实际选中值
                String comboValues = (String) page.evaluate(
                    "() => {" +
                    "  var combos = document.querySelectorAll('div[role=combobox]');" +
                    "  var r = '';" +
                    "  for (var i = 0; i < combos.length; i++) r += '[' + i + ']=' + combos[i].textContent.trim() + '; ';" +
                    "  return r;" +
                    "}");
                log("  [DEBUG] Combo选中: " + comboValues);

                // 查找 g/kg 输出 (Reactivation card index 1)
                String gkgValue = readGkgOutput(page);
                log("  [DEBUG] 初始 g/kg='" + gkgValue + "'");

                // Reactivation 输入框 (第1个可编辑的 Reactivation input)
                // 全局索引 18 通常是 Reactivation 温度输入
                int reactInputGlobalIndex = 18;

                // === 步长计算 ===
                setReactivationValue(page, reactInputGlobalIndex, ReactivationStart.toString());
                ThreadUtil.safeSleep(1000);

                // 重新读取 g/kg
                gkgValue = readGkgOutput(page);
                log("  [DEBUG] ReactivationStart=" + ReactivationStart + " gkg='" + gkgValue + "' fanwei=[" + fanweiStart + "," + fanweiEnd + "]");
                if (StrUtil.isEmpty(gkgValue)) { log("  gkg为空，跳过第" + linesNumber + "组"); continue; }
                Double gkgLeft = Double.parseDouble(gkgValue);
                log("  [DEBUG] gkgLeft=" + gkgLeft);

                // ReactivationEnd 计算
                setReactivationValue(page, reactInputGlobalIndex, ReactivationEnd.toString());
                ThreadUtil.safeSleep(1000);
                String gkgEndValue = readGkgOutput(page);
                log("  [DEBUG] ReactivationEnd=" + ReactivationEnd + " gkg='" + gkgEndValue + "'");
                if (StrUtil.isEmpty(gkgEndValue)) { log("  gkg为空(End)，跳过第" + linesNumber + "组"); continue; }

                boolean qk1 = (fanweiStart <= gkgLeft && gkgLeft <= fanweiEnd);
                Double ReactivationStartReal = ReactivationStart;
                Double ReactivationEndReal = ReactivationEnd;

                if (!qk1) {
                    setReactivationValue(page, reactInputGlobalIndex, StrUtil.toString(NumberUtil.add(ReactivationStart, Reactivationbc)));
                    ThreadUtil.safeSleep(1000);
                    String gkgNext = readGkgOutput(page);
                    if (StrUtil.isEmpty(gkgNext)) continue;
                    Double gkgTemp = Double.parseDouble(gkgNext);
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

                boolean flag = false;
                Double templeft = ReactivationStartReal.doubleValue();
                while (templeft <= ReactivationEndReal.doubleValue()) {
                    setReactivationValue(page, reactInputGlobalIndex, StrUtil.toString(templeft));
                    ThreadUtil.safeSleep(1000);
                    String gkgTempStr = readGkgOutput(page);
                    if (StrUtil.isEmpty(gkgTempStr)) break;
                    Double gkgTemp = Double.parseDouble(gkgTempStr);
                    if (fanweiStart <= gkgTemp && gkgTemp <= fanweiEnd) {
                        log("  [匹配] gkgTemp=" + gkgTemp + " 在范围[" + fanweiStart + "," + fanweiEnd + "]内，写入结果");
                        writeResult(page, ss, linesNumber, excelWriter);
                        flag = true;
                    } else {
                        log("  [不匹配] gkgTemp=" + gkgTemp + " 不在范围[" + fanweiStart + "," + fanweiEnd + "]");
                    }
                    templeft = NumberUtil.add(templeft, Reactivationbc);
                    if (flag && !(fanweiStart <= gkgTemp && gkgTemp <= fanweiEnd)) break;
                }
                log("第 " + linesNumber + " 组完成");
                ThreadUtil.safeSleep(1000);
                } catch (Exception groupEx) {
                    log("[ERROR] 第 " + linesNumber + " 组失败: " + groupEx.getMessage().split("\n")[0]);
                    // 尝试刷新页面恢复状态
                    try {
                        page.navigate("https://procalc5.proflute.se/rotor");
                        ThreadUtil.safeSleep(5000);
                    } catch (Exception navEx) {
                        log("[ERROR] 页面恢复失败: " + navEx.getMessage().split("\n")[0]);
                    }
                }
            }
            excelWriter.flush();
            log("所有参数计算完成，结果已保存到: " + resultFile);
        } catch (Exception e) {
            log("执行异常: " + e.getMessage());
            e.printStackTrace();
        } finally {
            try { context.close(); } catch (Exception ignored) {}
            try { browser.close(); } catch (Exception ignored) {}
            try { playwright.close(); } catch (Exception ignored) {}
        }
    }

    private void writeResult(Page page, StringBuilder ss, String lineNumber, ExcelWriter excelWriter) {
        List<String> list = Lists.newArrayList();

        // Wet Air (div[2])
        String wetAirXpath = "//*[@id='root']/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div";
        String v1 = getInputValueByXpath(page, wetAirXpath + "[1]/div/div/input");
        String v2 = getInputValueByXpath(page, wetAirXpath + "[2]/div/div/input");
        String v3 = getInputValueByXpath(page, wetAirXpath + "[3]/div/div/input");
        String v4 = getInputValueByXpath(page, wetAirXpath + "[4]/div/div/input");
        String v5 = getInputValueByXpath(page, wetAirXpath + "[5]/div/div/input");
        ss.append(lineNumber).append(" Wet Air:");
        ss.append(" ").append(v1).append(" ").append(v2).append(" ").append(v3).append(" ").append(v4).append(" ").append(v5);
        list.addAll(Arrays.asList(lineNumber, v1, v2, v3, v4, v5));

        // Process left (div[6])
        String plXpath = "//*[@id='root']/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div";
        String vv1 = getInputValueByXpath(page, plXpath + "[1]/div/div/input");
        String vv2 = getInputValueByXpath(page, plXpath + "[2]/div/div/input");
        String vv3 = getInputValueByXpath(page, plXpath + "[3]/div/div/input");
        String vv4 = getInputValueByXpath(page, plXpath + "[4]/div/div/input");
        String vv5 = getInputValueByXpath(page, plXpath + "[5]/div/div/input");
        String vv6 = getInputValueByXpath(page, plXpath + "[6]/div/div/input");
        String vv7 = getInputValueByXpath(page, plXpath + "[7]/div/div/input");
        ss.append(" ").append(vv1).append(" ").append(vv2).append(" ").append(vv3).append(" ").append(vv4).append(" ").append(vv5).append(" ").append(vv6).append(" ").append(vv7);
        list.addAll(Arrays.asList(vv1, vv2, vv3, vv4, vv5, vv6, vv7));

        // Process Right (div[7])
        String prXpath = "//*[@id='root']/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div";
        vv1 = getInputValueByXpath(page, prXpath + "[1]/div/div/input");
        vv2 = getInputValueByXpath(page, prXpath + "[2]/div/div/input");
        vv3 = getInputValueByXpath(page, prXpath + "[3]/div/div/input");
        vv4 = getInputValueByXpath(page, prXpath + "[4]/div/div/input");
        vv5 = getInputValueByXpath(page, prXpath + "[5]/div/div/input");
        vv6 = getInputValueByXpath(page, prXpath + "[6]/div/div/input");
        ss.append(" process right:");
        ss.append(" ").append(vv1).append(" ").append(vv2).append(" ").append(vv3).append(" ").append(vv4).append(" ").append(vv5).append(" ").append(vv6);
        list.addAll(Arrays.asList(vv1, vv2, vv3, vv4, vv5, vv6));

        // Reactivation (div[3])
        String reactXpath = "//*[@id='root']/div/div/div[1]/div[2]/div[3]/div/div[2]/div/div";
        vv1 = getInputValueByXpath(page, reactXpath + "[1]/div/div/input");
        vv2 = getInputValueByXpath(page, reactXpath + "[2]/div/div/input");
        vv3 = getInputValueByXpath(page, reactXpath + "[3]/div/div/input");
        vv4 = getInputValueByXpath(page, reactXpath + "[4]/div/div/input");
        vv5 = getInputValueByXpath(page, reactXpath + "[5]/div/div/input");
        vv6 = getInputValueByXpath(page, reactXpath + "[6]/div/div/input");
        vv7 = getInputValueByXpath(page, reactXpath + "[7]/div/div/input");
        ss.append(" Reactivation:");
        ss.append(" ").append(vv1).append(" ").append(vv2).append(" ").append(vv3).append(" ").append(vv4).append(" ").append(vv5).append(" ").append(vv6).append(" ").append(vv7);
        list.addAll(Arrays.asList(vv1, vv2, vv3, vv4, vv5, vv6, vv7));

        // RPH (div[8])
        String rphXpath = "//*[@id='root']/div/div/div[1]/div[2]/div[8]/div/div[2]/div/div/div/div/input";
        vv1 = getInputValueByXpath(page, rphXpath);
        ss.append(" RPH: ").append(vv1).append("\r\n");
        list.addAll(Arrays.asList(vv1));

        excelWriter.writeRow(list);
        log("  写入结果行: " + lineNumber);
    }
}
