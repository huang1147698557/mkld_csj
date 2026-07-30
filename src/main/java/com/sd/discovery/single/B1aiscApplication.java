package com.sd.discovery.single;

import cn.hutool.core.thread.ThreadUtil;
import cn.hutool.core.util.NumberUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelUtil;
import java.util.List;
import java.util.Scanner;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

/**
 * BaiscApplication
 *
 * @Author: gaoweiqi
 * @CreateDate: 2021/2/3 16:26
 */
public class B1aiscApplication {

  private static void test1() {
    StringBuilder ss = new StringBuilder();
    WebDriver driver = new ChromeDriver();
    driver.get("https://procalc5.proflute.se/rotor");
    ThreadUtil.safeSleep(8000);
    WebElement username = driver.findElement(By.id("userNameInput"));
    WebElement password = driver.findElement(By.id("passwordInput"));
    username.sendKeys("extcnlilzhe");
    password.sendKeys("018anhjtmgh$");
    WebElement login = driver.findElement(By.xpath("//*[@id=\"submitButton\"]"));
    login.click();
    //登陆成功
    Scanner scanner = new Scanner(System.in);
    //EXTCNLILZHE 374mucycdaw* procalc5.proflute.xlsx
    System.out.print("登录成功后，请输入要执行的文件名(包含文件后缀):");
    String temp = scanner.next();
    List<List<Object>> paraList = ExcelUtil.getReader("C:\\procalc5\\" + temp).read();
    for (List<Object> list : paraList) {
      if (paraList.indexOf(list) == 0) {
        continue;
      }
      String linesNumber = StrUtil.toString(list.get(0));
      String ProcessAirflow = StrUtil.toString(list.get(1));
      String DesiccantNedia = StrUtil.toString(list.get(2));
      String SectorLayout = StrUtil.toString(list.get(3));
      String RotorDiameter = StrUtil.toString(list.get(4));
      String RotorDepth = StrUtil.toString(list.get(5));
      String NetFaceAreaCalculation = StrUtil.toString(list.get(6));
      String SealingArea = StrUtil.toString(list.get(7));
      String ProcessStr1 = StrUtil.toString(list.get(8));
      String ProcessStr2 = StrUtil.toString(list.get(9));
      String Power = StrUtil.toString(list.get(10));
      Double ReactivationStart = Double.parseDouble(list.get(11).toString());
      Double ReactivationEnd = Double.parseDouble(list.get(12).toString());
      Double Reactivationbc = Double.parseDouble(list.get(13).toString());
      Double fanweiStart = Double.parseDouble(StrUtil.split(((String) list.get(14)), "~").get(0));
      Double fanweiEnd = Double.parseDouble(StrUtil.split(((String) list.get(14)), "~").get(1));

      String Reactivation1 = StrUtil.toString(list.get(15));
      String Reactivation2 = StrUtil.toString(list.get(16));
      String Reactivation3 = StrUtil.toString(list.get(17));
      //元素赋值
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[1]/div/div[2]/input"))
          .sendKeys(Keys.chord(Keys.CONTROL, "a"), ProcessAirflow);
      //media
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[3]/div/div[2]")).click();
      ThreadUtil.safeSleep(1000);
      List<WebElement> DesiccantMedias = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
      boolean flag = true;
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
      //Sector layout
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[4]/div/div[2]")).click();
      ThreadUtil.safeSleep(1000);
      List<WebElement> SectorLayouts = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
      flag = true;
      for (WebElement s : SectorLayouts) {
        switch (s.getAttribute("data-value")) {
          case "1":
            if (StrUtil.equalsAnyIgnoreCase(SectorLayout, "L270/90")) {
              s.click();
              flag = false;
              break;
            }
            ;
          case "2":
            if (StrUtil.equalsAnyIgnoreCase(SectorLayout, "L180/180")) {
              s.click();
              flag = false;
              break;
            }
            ;
          case "3":
            if (StrUtil.equalsAnyIgnoreCase(SectorLayout, "L240/120")) {
              s.click();
              flag = false;
              break;
            }
            ;
          case "4":
            if (StrUtil.equalsAnyIgnoreCase(SectorLayout, "L240/90/30")) {
              s.click();
              flag = false;
              break;
            }
            ;
          case "5":
            if (StrUtil.equalsAnyIgnoreCase(SectorLayout, "L180/90/90")) {
              s.click();
              flag = false;
              break;
            }
            ;
          case "6":
            if (StrUtil.equalsAnyIgnoreCase(SectorLayout, "L240/60/60")) {
              s.click();
              flag = false;
              break;
            }
            ;
          case "7":
            if (StrUtil.equalsAnyIgnoreCase(SectorLayout, "L230/65/65")) {
              s.click();
              flag = false;
              break;
            }
            ;
        }
      }
      if (flag) {
        SectorLayouts.get(0).click();
      }
/////
      WebElement RotorDiameterE = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[6]/div/div[2]"));
      RotorDiameterE.click();
      //Rotor diameter
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
      //Rotor depth
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[7]/div/div[2]")).click();
      ThreadUtil.safeSleep(1000);
      List<WebElement> RotorDepths = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
      flag = true;//*[@id="menu-"]/div[3]/ul
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
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[9]/div/div[2]/input")).sendKeys(Keys.chord(Keys.CONTROL, "a"), SealingArea);
      //按顺序放值
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[1]/div/div/input")).sendKeys(Keys.chord(Keys.CONTROL, "a"), ProcessStr1);
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[2]/div/div/input")).sendKeys(Keys.chord(Keys.CONTROL, "a"), ProcessStr2);
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[8]/div/div[2]/div/div/div/div/input")).sendKeys(Keys.chord(Keys.CONTROL, "a"), Power);
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[5]/div/div[2]/div/div[1]/div/div/input")).sendKeys(Keys.chord(Keys.CONTROL, "a"), Reactivation1);
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[5]/div/div[2]/div/div[2]/div/div/input")).sendKeys(Keys.chord(Keys.CONTROL, "a"), Reactivation2);
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[5]/div/div[2]/div/div[5]/div/div/input")).sendKeys(Keys.chord(Keys.CONTROL, "a"), Reactivation3);
      WebElement button = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[9]/button"));
      //按照步长处理数据
      WebElement Reactivation = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[3]/div/div[2]/div/div[1]/div/div/input"));
      WebElement gkg = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[2]/div/div/input"));
      //先算范围的边界值
      flag = false;
      Reactivation.sendKeys(Keys.chord(Keys.CONTROL, "a"), ReactivationStart.toString());
      button.click();
      ThreadUtil.safeSleep(1500);
      if(StrUtil.isEmpty(gkg.getAttribute("value"))){
        continue;
      }
      Double gkgLeft = Double.parseDouble(gkg.getAttribute("value"));
      Reactivation.sendKeys(Keys.chord(Keys.CONTROL, "a"), ReactivationEnd.toString());
      button.click();
      ThreadUtil.safeSleep(1500);
      if(StrUtil.isEmpty(gkg.getAttribute("value"))){
        continue;
      }
      Double gkgRight = Double.parseDouble(gkg.getAttribute("value"));
      boolean qk1 = (fanweiStart <= gkgLeft && gkgLeft <= fanweiEnd);
      boolean qk2 = (fanweiStart <= gkgRight && gkgRight <= fanweiEnd);
      Double ReactivationStartReal = ReactivationStart;
      Double ReactivationEndReal = ReactivationEnd;
      //左侧不在范围内
      if (!qk1) {
        Reactivation.sendKeys(Keys.chord(Keys.CONTROL, "a"), StrUtil.toString(NumberUtil.add(ReactivationStart, Reactivationbc)));
        button.click();
        ThreadUtil.safeSleep(1500);
        if(StrUtil.isEmpty(gkg.getAttribute("value"))){
          continue;
        }
        Double gkgTemp = Double.parseDouble(gkg.getAttribute("value"));
        //递增模式
        if (gkgTemp > gkgLeft) {
          if (gkgLeft < fanweiStart) {
            Double wendus = NumberUtil.mul((Double) NumberUtil.div(NumberUtil.sub(fanweiStart, gkgLeft),
                NumberUtil.sub(gkgTemp, gkgLeft)), Reactivationbc);
            ReactivationStartReal = NumberUtil.add(ReactivationStart.doubleValue(), NumberUtils.safeMultiply(wendus, 1.0, 0).doubleValue());
          } else {
            ReactivationStartReal = null;
          }
        } else { //递减模式
          if (gkgLeft > fanweiEnd) {
            Double wendus = NumberUtil.mul((Double) NumberUtil.div(NumberUtil.sub(gkgLeft, fanweiEnd),
                NumberUtil.sub(gkgLeft, gkgTemp)), Reactivationbc);
            ReactivationStartReal = NumberUtil.add(ReactivationStart.doubleValue(), NumberUtils.safeMultiply(wendus, 1.0, 0).doubleValue());
          } else {
            ReactivationStartReal = null;
          }
        }
      }
      //右侧不在范围内
    /*  if (!qk2) {
        Reactivation.sendKeys(Keys.chord(Keys.CONTROL, "a"), StrUtil.toString(NumberUtil.sub(ReactivationEnd, Reactivationbc)));
        button.click();
        ThreadUtil.safeSleep(1500);
        Double gkgTemp = Double.parseDouble(gkg.getAttribute("value"));
        //递增模式
        if (gkgTemp > gkgRight) {
          if (gkgRight < fanweiStart) {
            Double wendus = NumberUtil.mul((Double) NumberUtil.div(NumberUtil.sub(fanweiEnd, gkgRight),
                NumberUtil.sub(gkgTemp, gkgRight)), Reactivationbc);
            ReactivationEndReal = NumberUtil.sub(ReactivationEnd.doubleValue(), NumberUtils.safeMultiply(wendus, 1.0, 0).doubleValue());
          } else {
            ReactivationEndReal = null;
          }
        } else {//递减模式
          if (gkgRight > fanweiEnd) {
            Double wendus = NumberUtil.mul((Double) NumberUtil.div(NumberUtil.sub(gkgRight, fanweiEnd),
                NumberUtil.sub(gkgRight, gkgTemp)), Reactivationbc);
            ReactivationEndReal = NumberUtil.sub(ReactivationEnd.doubleValue(), NumberUtils.safeMultiply(wendus, 1.0, 0).doubleValue());
          } else {
            ReactivationEndReal = null;
          }
        }
      }*/

      if (ReactivationEndReal == null || ReactivationStartReal == null) {
        ReactivationStartReal = ReactivationStart;
        ReactivationEndReal = ReactivationEnd;
      }
      flag = false;
      Double templeft = ReactivationStartReal.doubleValue();
      while (templeft <= ReactivationEndReal.doubleValue()) {
        Reactivation.sendKeys(Keys.chord(Keys.CONTROL, "a"), StrUtil.toString(templeft));
        button.click();
        ThreadUtil.safeSleep(1500);
        if (StrUtil.isEmpty(gkg.getAttribute("value"))) {
          break;
        }
        if(StrUtil.isEmpty(gkg.getAttribute("value"))){
          continue;
        }
        Double gkgTemp = Double.parseDouble(gkg.getAttribute("value"));
        if (fanweiStart <= gkgTemp && gkgTemp <= fanweiEnd) {
          toList(driver, ss, linesNumber);
          flag = true;
        }
        templeft = NumberUtil.add(templeft, Reactivationbc);
        if (flag && !(fanweiStart <= gkgTemp && gkgTemp <= fanweiEnd)) {
          break;
        }

      }
      System.out.println("Datas:" + ss);
    }
    System.out.println("总数据:" + ss);
  }

  public static void toList(WebDriver driver, StringBuilder ss, String lineNumber) {
    String v1 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div[1]/div/div/input")).getAttribute("value");
    String v2 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div[2]/div/div/input")).getAttribute("value");
    String v3 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div[3]/div/div/input")).getAttribute("value");
    String v4 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div[4]/div/div/input")).getAttribute("value");
    String v5 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div[5]/div/div/input")).getAttribute("value");
    String vv1 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[1]/div/div/input")).getAttribute("value");
    String vv2 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[2]/div/div/input")).getAttribute("value");
    String vv3 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[3]/div/div/input")).getAttribute("value");
    String vv4 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[4]/div/div/input")).getAttribute("value");
    String vv5 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[5]/div/div/input")).getAttribute("value");
    String vv6 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[6]/div/div/input")).getAttribute("value");
    String vv7 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[3]/div/div[2]/div/div[1]/div/div/input")).getAttribute("value");
    ss.append(lineNumber + " " + v1 + " " + v2 + " " + v3 + " " + v4 + " " + v5 + " " + vv1 + " " + vv2 + " " + vv3 + " " + vv4 + " " + vv5 + " " + vv6+" "+vv7).append("\r\n");

  }
}