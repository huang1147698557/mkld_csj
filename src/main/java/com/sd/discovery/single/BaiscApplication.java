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
import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

/**
 * BaiscApplication
 *
 * @Author: gaoweiqi
 * @CreateDate: 2021/2/3 16:26
 */
@SpringBootApplication(scanBasePackages = {"com.sd.discovery.*"})
public class BaiscApplication {

  // 默认工作目录：Windows=C:\procalc5，Mac/Linux=用户目录下的procalc5
  private static String workDir;
  private static String inputFile; // 自定义输入文件路径（--input=），为空时使用默认
  private static String sessionStartTime; // 本次会话时间戳，用于文件名
  private static String sessionTimeDisplay; // 可读时间格式，用于Excel单元格
  private static int writerRowIndex = 1;  // 当前写入行索引（用于追加写入）

  static {
    boolean isWindows = System.getProperty("os.name").toLowerCase().contains("win");
    if (isWindows) {
      workDir = "C:\\procalc5";
    } else {
      workDir = System.getProperty("user.home") + "/procalc5";
    }
  }

  public static void main(String[] args) {
    SpringApplication.run(BaiscApplication.class, args);

    // 支持命令行参数覆盖：--workdir=/path/to/dir  --input=/path/to/file.xlsx  --chromedriver=/path/to/chromedriver
    for (String arg : args) {
      if (arg.startsWith("--workdir=")) {
        workDir = arg.substring("--workdir=".length());
      } else if (arg.startsWith("--input=")) {
        inputFile = arg.substring("--input=".length());
      } else if (arg.startsWith("--chromedriver=")) {
        // 仅当用户显式指定时才设置，否则由 Selenium Manager 自动管理
        System.setProperty("webdriver.chrome.driver", arg.substring("--chromedriver=".length()));
      }
    }

    System.out.println("工作目录: " + workDir);
    if (inputFile != null) {
      System.out.println("输入文件: " + inputFile);
    }
    test();
  }

  // 跨平台输入：用JavaScript设置React受控输入框的值（解决clear/sendKeys不触发React状态更新的问题）
  private static void fillInput(WebDriver driver, String xpath, String value) {
    WebElement el = driver.findElement(By.xpath(xpath));
    fillInput(driver, el, value);
  }

  private static void fillInput(WebDriver driver, WebElement el, String value) {
    ((org.openqa.selenium.JavascriptExecutor) driver).executeScript(
      "var el = arguments[0];" +
      "el.click();" +
      "el.focus();" +
      "var nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;" +
      "nativeInputValueSetter.call(el, arguments[1]);" +
      "el.dispatchEvent(new Event('input', { bubbles: true }));" +
      "el.dispatchEvent(new Event('change', { bubbles: true }));",
      el, value
    );
  }

  private static void test() {
    String usernameValue = System.getenv("PROCALC_USERNAME");
    String passwordValue = System.getenv("PROCALC_PASSWORD");
    if (StrUtil.isBlank(usernameValue) || StrUtil.isBlank(passwordValue)) {
      throw new IllegalStateException("请设置 PROCALC_USERNAME 和 PROCALC_PASSWORD 环境变量");
    }
    StringBuilder ss = new StringBuilder();
    ChromeOptions options = new ChromeOptions();
    options.addArguments("--no-sandbox");
    options.addArguments("--disable-dev-shm-usage");
    WebDriver driver = new ChromeDriver(options);
    sessionStartTime = DateUtil.format(DateUtil.date(), "yyyyMMdd_HHmmss");
    sessionTimeDisplay = DateUtil.format(DateUtil.date(), "yyyy年M月d日H时m分");
    driver.get("https://procalc5.proflute.se/rotor");
    WebElement username = waitForVisibleElement(driver, "用户名输入框",
        By.id("userNameInput"), By.id("username"), By.name("username"),
        By.name("UserName"), By.cssSelector("input[type='email']"));
    WebElement password = waitForVisibleElement(driver, "密码输入框",
        By.id("passwordInput"), By.id("password"), By.name("password"), By.name("Password"));
    username.sendKeys(usernameValue);
    password.sendKeys(passwordValue);
    WebElement login = waitForVisibleElement(driver, "登录按钮",
        By.id("submitButton"), By.id("kc-login"), By.cssSelector("button[type='submit']"),
        By.cssSelector("input[type='submit']"));
    login.click();
    waitForVisibleElement(driver, "转子计算表单",
        By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[9]/div[3]/button"));
    String sep = System.getProperty("os.name").toLowerCase().contains("win") ? "\\" : "/";
    // 全量累积文件：追加模式
    String allFilePath = workDir + sep + "calculate_results_all.xlsx";
    File allFile = new File(allFilePath);
    List<String> header = Lists.newArrayList("序号", "计算时间", " Wet Air:", "", "", "", "", "Process left",
        "", "", "", "", "", "", "Process Right", "", "", "", "", "", "Reactivation",
        "", "", "", "", "", "", "RPH");
    ExcelWriter excelWriter;
    if (allFile.exists()) {
      // 追加：读取已有行数，跳过一行空白后继续写入
      int existingCount = ExcelUtil.getReader(allFilePath).read().size();
      excelWriter = ExcelUtil.getWriter(allFilePath);
      writerRowIndex = existingCount + 1; // +1 空行分隔
    } else {
      // 新建文件并写表头
      excelWriter = ExcelUtil.getWriter(allFilePath);
      excelWriter.writeHeadRow(header);
      writerRowIndex = 1;
    }
    //登陆成功
    List<List<Object>> paraList = ExcelUtil.getReader(inputFile != null ? inputFile : workDir + sep + "procalc5.proflute.xlsx").read();
    Double lastFoundTemp = null; // 跨行保持上一轮找到的温度值，用于下一轮倒序查找上限
    List<List<Object>> bestResults = new ArrayList<>(); // 各组最优解汇总
    int groupIndex = 0; // 组号计数器
    for (List<Object> list : paraList) {
      if (paraList.indexOf(list) == 0) {
        continue;
      }
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
      //元素赋值
      //范围1赋值
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
      //Performance safety factor None DeltaPercent Multiplier
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[8]/div/div/div/div")).click();
      ThreadUtil.safeSleep(500);
      List<WebElement> DesiccantMedias1 = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
      boolean flag = true;
      if (StrUtil.equalsIgnoreCase(Performancesafetyfactor, "None")) {
        DesiccantMedias1.get(0).click();
        flag = false;
      } else if (StrUtil.equalsIgnoreCase(Performancesafetyfactor, "+Δ% Moisture")) {
        DesiccantMedias1.get(1).click();
        flag = false;
      } else if (StrUtil.equalsIgnoreCase(Performancesafetyfactor, "x Moisture")) {
        DesiccantMedias1.get(2).click();
        flag = false;
      }
      if (flag) {
        DesiccantMedias1.get(0).click();
      }
      if (!StrUtil.equalsIgnoreCase(Performancesafetyfactor, "None")) {
        driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[8]/div/div/div[2]/div/div")).click();
        List<WebElement> RotorDepths = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
        for (WebElement s : RotorDepths) {
          String sTemp = s.getAttribute("data-value");
          if (StrUtil.equalsAnyIgnoreCase(PerformancesafetyfactorV, sTemp)) {
            s.click();
            break;
          }
        }
      }
      //范围2赋值
      fillInput(driver, "//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[1]/div/div[2]/input", ProcessAirflow);
      //media
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[3]/div/div[2]")).click();
      ThreadUtil.safeSleep(1000);
      List<WebElement> DesiccantMedias = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
      flag = true;
      for (WebElement s : DesiccantMedias) {
        String sTemp = StrUtil.equalsIgnoreCase(s.getAttribute("data-value"), "1") ? "PPS" : "PPP";
        if (StrUtil.equalsIgnoreCase(DesiccantNedia, sTemp)) {
          s.click();
          flag = false;
          break;
        }
      }
      if (flag) {
        DesiccantMedias.get(0).click();
      }
      //Sector layout - 用文本内容匹配，防止StaleElementReferenceException
      System.out.println("[Sector Layout] 期望值: " + SectorLayout);
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[4]/div/div[2]")).click();
      ThreadUtil.safeSleep(1500);
      flag = true;
      for (int retry = 0; retry < 3 && flag; retry++) {
        try {
          List<WebElement> SectorLayouts = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
          System.out.println("[Sector Layout] 找到 " + SectorLayouts.size() + " 个选项");
          for (WebElement s : SectorLayouts) {
            String text = s.getText().trim();
            System.out.println("  选项: '" + text + "' (data-value=" + s.getAttribute("data-value") + ")");
            if (StrUtil.equalsIgnoreCase(text, SectorLayout)) {
              System.out.println("[Sector Layout] 匹配成功，选择: " + text);
              s.click();
              flag = false;
              break;
            }
          }
          if (flag && SectorLayouts.size() > 0) {
            System.out.println("[Sector Layout] 未匹配到，使用默认: " + SectorLayouts.get(0).getText());
            SectorLayouts.get(0).click();
            flag = false;
          }
        } catch (org.openqa.selenium.StaleElementReferenceException e) {
          System.out.println("[Sector Layout] 元素失效，重试第" + (retry + 1) + "次");
          ThreadUtil.safeSleep(1000);
          driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[4]/div/div[2]")).click();
          ThreadUtil.safeSleep(1500);
        }
      }
/////
      //Rotor diameter - Custom 2-sector/Custom 3 sector时为输入框，其他为下拉
      boolean isCustomSector = StrUtil.equalsAnyIgnoreCase(SectorLayout, "Custom 2-sector", "Custom 3 sector");
      if (isCustomSector) {
        WebElement rotorDiameterInput = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[6]/div/div[2]/input"));
        fillInput(driver, rotorDiameterInput, RotorDiameter);
      } else {
        for (int retry = 0; retry < 3; retry++) {
          try {
            driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[6]/div/div[2]")).click();
            ThreadUtil.safeSleep(1000);
            List<WebElement> RotorDiameters = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
            flag = true;
            for (WebElement s : RotorDiameters) {
              if (StrUtil.equalsIgnoreCase(RotorDiameter, s.getAttribute("data-value"))) {
                s.click();
                flag = false;
                break;
              }
            }
            if (flag) {
              RotorDiameters.get(0).click();
            }
            break;
          } catch (org.openqa.selenium.StaleElementReferenceException e) {
            System.out.println("Rotor diameter下拉元素失效，重试第" + (retry + 1) + "次");
            ThreadUtil.safeSleep(1000);
          }
        }
      }
      //Rotor depth - Custom 2-sector/Custom 3 sector时为输入框，其他为下拉
      if (isCustomSector) {
        WebElement rotorDepthInput = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[7]/div/div[2]/input"));
        fillInput(driver, rotorDepthInput, RotorDepth);
      } else {
        for (int retry = 0; retry < 3; retry++) {
          try {
            driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[7]/div/div[2]")).click();
            ThreadUtil.safeSleep(1000);
            List<WebElement> RotorDepths = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
            flag = true;
            for (WebElement s : RotorDepths) {
              if (StrUtil.equalsIgnoreCase(RotorDepth, s.getAttribute("data-value"))) {
                s.click();
                flag = false;
                break;
              }
            }
            if (flag) {
              RotorDepths.get(0).click();
            }
            break;
          } catch (org.openqa.selenium.StaleElementReferenceException e) {
            System.out.println("Rotor depth下拉元素失效，重试第" + (retry + 1) + "次");
            ThreadUtil.safeSleep(1000);
          }
        }
      }
      //Net face area calculation
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[8]/div/div[2]")).click();
      ThreadUtil.safeSleep(1000);
      List<WebElement> Netfaceareacalculations = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
      flag = true;
      for (WebElement s : Netfaceareacalculations) {
        String sTemp = StrUtil.equalsIgnoreCase(s.getAttribute("data-value"), "0") ? "Sealing area" : "Active area";
        if (StrUtil.equalsIgnoreCase(NetFaceAreaCalculation, sTemp)) {
          s.click();
          flag = false;
          break;
        }
      }
      if (flag) {
        Netfaceareacalculations.get(0).click();
      }
      fillInput(driver, "//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[9]/div/div[2]/input", SealingArea);
      //按顺序放值
      System.out.println("[填充值] ProcessStrC=" + ProcessStrC + ", ProcessStrGKG=" + ProcessStrGKG + ", Rph=" + Rph);
      System.out.println("[填充值] Reactivation1=" + Reactivation1 + ", Reactivation2=" + Reactivation2 + ", Reactivation3=" + Reactivation3);
      fillInput(driver, "//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[1]/div/div/input", ProcessStrC);
      fillInput(driver, "//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[2]/div/div/input", ProcessStrGKG);
      fillInput(driver, "//*[@id=\"root\"]/div/div/div[1]/div[2]/div[8]/div/div[2]/div/div/div/div/input", Rph);
      fillInput(driver, "//*[@id=\"root\"]/div/div/div[1]/div[2]/div[5]/div/div[2]/div/div[1]/div/div/input", Reactivation1);
      fillInput(driver, "//*[@id=\"root\"]/div/div/div[1]/div[2]/div[5]/div/div[2]/div/div[2]/div/div/input", Reactivation2);
      fillInput(driver, "//*[@id=\"root\"]/div/div/div[1]/div[2]/div[5]/div/div[2]/div/div[5]/div/div/input", Reactivation3);
      WebElement button = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[9]/div[3]/button"));
      //按照步长处理数据
      WebElement Reactivation = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[3]/div/div[2]/div/div[1]/div/div/input"));
      WebElement gkg = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[2]/div/div/input"));
      //先算范围的边界值
      fillInput(driver, Reactivation, ReactivationStart.toString());
      click(button);
      ThreadUtil.safeSleep(1500);
      if (StrUtil.isEmpty(gkg.getAttribute("value"))) {
        continue;
      }
      Double gkgLeft = Double.parseDouble(gkg.getAttribute("value"));
      fillInput(driver, Reactivation, ReactivationEnd.toString());
      click(button);
      ThreadUtil.safeSleep(1500);
      if (StrUtil.isEmpty(gkg.getAttribute("value"))) {
        continue;
      }
      boolean qk1 = (fanweiStart <= gkgLeft && gkgLeft <= fanweiEnd);
      Double ReactivationStartReal = ReactivationStart;
      Double ReactivationEndReal = ReactivationEnd;
      //左侧不在范围内
      if (!qk1) {
        fillInput(driver, Reactivation, StrUtil.toString(NumberUtil.add(ReactivationStart, Reactivationbc)));
        click(button);
        ThreadUtil.safeSleep(1500);
        if (StrUtil.isEmpty(gkg.getAttribute("value"))) {
          continue;
        }
        Double gkgTemp = Double.parseDouble(gkg.getAttribute("value"));
        //递增模式
        if (NumberUtil.compare(gkgTemp, gkgLeft) > 0) {
          if (NumberUtil.compare(gkgLeft, fanweiStart) < 0 && NumberUtil.compare(gkgTemp, gkgLeft) != 0) {
            Double wendus = NumberUtil.mul((Double) NumberUtil.div(NumberUtil.sub(fanweiStart, gkgLeft),
                NumberUtil.sub(gkgTemp, gkgLeft)), Reactivationbc);
            ReactivationStartReal = NumberUtil.add(ReactivationStart.doubleValue(), NumberUtils.safeMultiply(wendus, 1.0, 0).doubleValue());
          } else {
            ReactivationStartReal = null;
          }
        } else { //递减模式
          if (NumberUtil.compare(gkgLeft, fanweiEnd) > 0 && NumberUtil.compare(gkgLeft, gkgTemp) != 0) {
            Double wendus = NumberUtil.mul((Double) NumberUtil.div(NumberUtil.sub(gkgLeft, fanweiEnd),
                NumberUtil.sub(gkgLeft, gkgTemp)), Reactivationbc);
            ReactivationStartReal = NumberUtil.add(ReactivationStart.doubleValue(), NumberUtils.safeMultiply(wendus, 1.0, 0).doubleValue());
          } else {
            ReactivationStartReal = null;
          }
        }
      }
      //右侧不在范围内
      if (ReactivationEndReal == null || ReactivationStartReal == null) {
        ReactivationStartReal = ReactivationStart;
        ReactivationEndReal = ReactivationEnd;
      }
      // ===== 倒序查找 Reactivation 温度 =====
      // 确定本轮查找上限：优先使用上一行找到的温度，否则使用 ReactivationEndReal
      Double searchUpper = (lastFoundTemp != null) ? Math.min(lastFoundTemp, ReactivationEndReal) : ReactivationEndReal;
      Double searchLower = ReactivationStartReal;
      System.out.println("[倒序查找] searchLower=" + searchLower + ", searchUpper=" + searchUpper + ", 步长=" + Reactivationbc);

      if (NumberUtil.compare(searchUpper, searchLower) < 0) {
        System.out.println("[倒序查找] 上限小于下限，跳过本行");
        continue;
      }

      flag = false;
      Double tempCurrent = searchUpper; // 从上限开始倒序
      int iterCount = 0;
      int emptyCount = 0; // 连续空值计数器，容忍异常空值
      final int MAX_EMPTY = 5; // 连续5次空值才真正中断
      // ===== 本轮命中记录收集（用于最优解筛选）=====
      List<double[]> groupHits = new ArrayList<>(); // [温度, g/kg]
      while (NumberUtil.compare(tempCurrent, searchLower) >= 0) {
        iterCount++;
        try {
          // 尝试关闭可能存在的遮罩层(MUI Backdrop/Dialog)
          ((org.openqa.selenium.JavascriptExecutor) driver).executeScript(
            "var backdrop = document.querySelector('.MuiBackdrop-root');" +
            "if(backdrop && getComputedStyle(backdrop).opacity > 0) { backdrop.click(); }"
          );
          ThreadUtil.safeSleep(300);
          fillInput(driver, Reactivation, StrUtil.toString(tempCurrent));
        } catch (Exception e) {
          System.out.println("[倒序查找] 填充温度失败: " + e.getMessage());
          tempCurrent = NumberUtil.sub(tempCurrent, Reactivationbc);
          continue;
        }
        try {
          click(button);
        } catch (Exception e) {
          System.out.println(e.getMessage());
        }
        ThreadUtil.safeSleep(1500);
        String gkgValue;
        try {
          gkgValue = gkg.getAttribute("value");
        } catch (Exception e) {
          System.out.println("[倒序查找] 获取g/kg值失败: " + e.getMessage());
          emptyCount++;
          if (emptyCount >= MAX_EMPTY) break;
          tempCurrent = NumberUtil.sub(tempCurrent, Reactivationbc);
          continue;
        }
        if (StrUtil.isEmpty(gkgValue)) {
          emptyCount++;
          System.out.println("[倒序查找] gkg为空(连续第" + emptyCount + "次)，继续查找...");
          if (emptyCount >= MAX_EMPTY) {
            System.out.println("[倒序查找] 连续" + MAX_EMPTY + "次空值，中断查找");
            break;
          }
          tempCurrent = NumberUtil.sub(tempCurrent, Reactivationbc);
          continue;
        }
        emptyCount = 0; // 有值则重置空值计数
        Double gkgTemp = Double.parseDouble(gkgValue);
        System.out.println("[倒序查找] 温度=" + tempCurrent + ", g/kg=" + gkgTemp + ", 范围=[" + fanweiStart + "~" + fanweiEnd + "]");
        boolean isHit = (fanweiStart <= gkgTemp && gkgTemp <= fanweiEnd);
        toList(driver, ss, linesNumber, excelWriter, sessionTimeDisplay, isHit);
        if (isHit) {
          groupHits.add(new double[]{tempCurrent, gkgTemp}); // 记录命中
          lastFoundTemp = tempCurrent; // 记录满足条件的温度
          flag = true;
          System.out.println("[倒序查找] ★ 命中! 温度=" + tempCurrent + ", 更新lastFoundTemp");
        }
        tempCurrent = NumberUtil.sub(tempCurrent, Reactivationbc); // 倒序递减
        if (flag && !(fanweiStart <= gkgTemp && gkgTemp <= fanweiEnd)) {
          System.out.println("[倒序查找] 已离开范围，结束查找，共迭代" + iterCount + "次");
          break;
        }
      }
      if (!flag) {
        System.out.println("[倒序查找] 未找到满足条件的温度，共迭代" + iterCount + "次");
      }
      // ===== 最优解筛选 =====
      groupIndex++;
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
        System.out.println("[最优解] 第" + groupIndex + "组: 范围中间值=" + NumberUtil.round(midValue, 4)
            + ", 最接近命中: g/kg=" + best[1] + "(温度=" + best[0] + "°C), 差值=" + NumberUtil.round(bestDiff, 4));
        // 将最优解温度重新填入网页，采集完整数据
        fillInput(driver, Reactivation, StrUtil.toString(best[0]));
        click(button);
        ThreadUtil.safeSleep(1500);
        List<Object> bestRow = collectRowData(driver, linesNumber);
        bestRow.add(1, sessionTimeDisplay); // 插入可读时间到第2列
        bestRow.add(NumberUtil.round(midValue, 4)); // 额外列：范围中间值
        // 写入全量文件（红色标记最优解）
        int bestRowIdx = writerRowIndex++;
        excelWriter.writeRow(bestRow);
        applyRowStyle(excelWriter, bestRowIdx, bestRow.size(), false, true);
        bestResults.add(bestRow);
      } else {
        System.out.println("[最优解] 第" + groupIndex + "组: 无命中记录，跳过");
      }
      System.out.println("Datas:" + ss);
      ThreadUtil.safeSleep(1000);
    }
    excelWriter.flush();
    // ===== 输出最优解汇总到 {时间戳}_result_02.xlsx =====
    String bestFilePath = workDir + sep + sessionStartTime + "_result_02.xlsx";
    ExcelWriter bestWriter = ExcelUtil.getWriter(bestFilePath);
    List<String> bestHeader = Lists.newArrayList("序号", "计算时间", " Wet Air:", "", "", "", "", "Process left",
        "", "", "", "", "", "", "Process Right", "", "", "", "", "", "Reactivation",
        "", "", "", "", "", "", "RPH", "范围中间值");
    bestWriter.writeHeadRow(bestHeader);
    int bestIdx = 1;
    for (List<Object> row : bestResults) {
      bestWriter.writeRow(row);
      applyRowStyle(bestWriter, bestIdx, row.size(), true, false);
      bestIdx++;
    }
    bestWriter.flush();
    System.out.println("\n===== 数据汇总完成 =====");
    System.out.println("全量数据已追加到: " + allFilePath);
    System.out.println("共" + bestResults.size() + "组最优解，已写入: " + bestFilePath);
    System.out.println("总数据:" + ss);
  }

  private static WebElement waitForVisibleElement(WebDriver driver, String description, By... selectors) {
    long deadline = System.currentTimeMillis() + 60000;
    while (System.currentTimeMillis() < deadline) {
      for (By selector : selectors) {
        for (WebElement element : driver.findElements(selector)) {
          if (element.isDisplayed() && element.isEnabled()) {
            return element;
          }
        }
      }
      ThreadUtil.safeSleep(250);
    }
    throw new IllegalStateException("等待" + description + "超时，当前页面: " + driver.getCurrentUrl()
        + "，标题: " + driver.getTitle());
  }

  public static void click(WebElement button) {
    click(button, 3);
  }

  public static void click(WebElement button, int time) {
    try {
      button.click();
    } catch (Exception e) {
      System.out.println("按钮点击异常重新 点击");
      if (time > 0) {
        ThreadUtil.safeSleep(1000);
        click(button, --time);
      }

    }
  }

  /**
   * 采集当前网页上所有结果字段，返回一行数据（与 toList 列顺序一致，不含中间值）
   */
  public static List<Object> collectRowData(WebDriver driver, String lineNumber) {
    List<Object> row = new ArrayList<>();
    row.add(lineNumber);
    // Wet Air (5 fields)
    for (int i = 1; i <= 5; i++) {
      row.add(driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div[" + i + "]/div/div/input")).getAttribute("value"));
    }
    // Process left (7 fields)
    for (int i = 1; i <= 7; i++) {
      row.add(driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[" + i + "]/div/div/input")).getAttribute("value"));
    }
    // Process right (6 fields)
    for (int i = 1; i <= 6; i++) {
      row.add(driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[" + i + "]/div/div/input")).getAttribute("value"));
    }
    // Reactivation (7 fields)
    for (int i = 1; i <= 7; i++) {
      row.add(driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[3]/div/div[2]/div/div[" + i + "]/div/div/input")).getAttribute("value"));
    }
    // RPH (1 field)
    row.add(driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[8]/div/div[2]/div/div/div/div/input")).getAttribute("value"));
    return row;
  }

  public static void toList(WebDriver driver, StringBuilder ss, String lineNumber,
      ExcelWriter excelWriter, String timestamp, boolean isHit) {
    List<String> list = Lists.newArrayList();
    list.add(lineNumber);
    list.add(timestamp); // B列：计算时间
    String v1 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div[1]/div/div/input")).getAttribute("value");
    String v2 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div[2]/div/div/input")).getAttribute("value");
    String v3 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div[3]/div/div/input")).getAttribute("value");
    String v4 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div[4]/div/div/input")).getAttribute("value");
    String v5 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div[5]/div/div/input")).getAttribute("value");
    ss.append(lineNumber);
    ss.append(" Wet Air:");
    ss.append(" " + v1 + " " + v2 + " " + v3 + " " + v4 + " " + v5);
    list.addAll(Arrays.asList(v1, v2, v3, v4, v5));
    String vv1 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[1]/div/div/input")).getAttribute("value");
    String vv2 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[2]/div/div/input")).getAttribute("value");
    String vv3 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[3]/div/div/input")).getAttribute("value");
    String vv4 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[4]/div/div/input")).getAttribute("value");
    String vv5 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[5]/div/div/input")).getAttribute("value");
    String vv6 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[6]/div/div/input")).getAttribute("value");
    String vv7 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[7]/div/div/input")).getAttribute("value");
//    ss.append(" Process left:");
    ss.append(" " + vv1 + " " + vv2 + " " + vv3 + " " + vv4 + " " + vv5 + " " + vv6 + " " + vv7);
    list.addAll(Arrays.asList(vv1, vv2, vv3, vv4, vv5, vv6, vv7));
    vv1 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[1]/div/div/input")).getAttribute("value");
    vv2 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[2]/div/div/input")).getAttribute("value");
    vv3 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[3]/div/div/input")).getAttribute("value");
    vv4 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[4]/div/div/input")).getAttribute("value");
    vv5 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[5]/div/div/input")).getAttribute("value");
    vv6 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[6]/div/div/input")).getAttribute("value");
    ss.append(" process right:");
    ss.append(" " + vv1 + " " + vv2 + " " + vv3 + " " + vv4 + " " + vv5 + " " + vv6);
    list.addAll(Arrays.asList(vv1, vv2, vv3, vv4, vv5, vv6));
    vv1 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[3]/div/div[2]/div/div[1]/div/div/input")).getAttribute("value");
    vv2 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[3]/div/div[2]/div/div[2]/div/div/input")).getAttribute("value");
    vv3 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[3]/div/div[2]/div/div[3]/div/div/input")).getAttribute("value");
    vv4 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[3]/div/div[2]/div/div[4]/div/div/input")).getAttribute("value");
    vv5 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[3]/div/div[2]/div/div[5]/div/div/input")).getAttribute("value");
    vv6 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[3]/div/div[2]/div/div[6]/div/div/input")).getAttribute("value");
    vv7 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[3]/div/div[2]/div/div[7]/div/div/input")).getAttribute("value");
    ss.append(" Reactivation:");
    ss.append(" " + vv1 + " " + vv2 + " " + vv3 + " " + vv4 + " " + vv5 + " " + vv6 + " " + vv7);
    list.addAll(Arrays.asList(vv1, vv2, vv3, vv4, vv5, vv6, vv7));
    vv1 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[8]/div/div[2]/div/div/div/div/input")).getAttribute("value");
    ss.append(" RPH:");
    ss.append(" " + vv1);
    ss.append("\r\n");
    list.addAll(Arrays.asList(vv1));
    int rowIdx = writerRowIndex++;
    excelWriter.writeRow(list);
    if (isHit) {
      applyRowStyle(excelWriter, rowIdx, list.size(), true, false);
    }
  }

  /**
   * 对指定行应用样式：加粗和/或红色字体
   */
  private static void applyRowStyle(ExcelWriter writer, int rowIndex, int colCount, boolean bold, boolean red) {
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
}
