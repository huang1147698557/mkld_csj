package com.base.demo;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
//import com.osisoft.jdbc.Driver;

public class test {

  public static void main(String[] args) {
    // 定义数据库连接参数
    String jdbcUrl = "jdbc:adonet://192.168.121.153;Provider=iHOLEDB.iHistorian.1;Persist Security Info=False;USER ID=administrator;Password=A1.;Mode=Read";
    String username = "administrator";
    String password = "A1.";

    // 定义 SQL 查询语句
    String sqlQuery = "SELECT * FROM ihtags";

    // 创建数据库连接
    Connection connection = null;
    Statement statement = null;
    ResultSet resultSet = null;

    try {
      // 显式加载 JDBC 驱动程序
      Class.forName("com.osisoft.jdbc.Driver");

      // 建立数据库连接
      connection = DriverManager.getConnection(jdbcUrl, username, password);

      // 设置连接为只读模式
      if (connection != null) {
        connection.setReadOnly(true);
      }

      // 创建 SQL 语句对象
      statement = connection.createStatement();

      // 执行 SQL 查询
      resultSet = statement.executeQuery(sqlQuery);

      // 处理查询结果
      while (resultSet.next()) {
        // 假设表中有 TagID, TagName, Description 字段
        int tagId = resultSet.getInt("TagID");
        String tagName = resultSet.getString("TagName");
        String description = resultSet.getString("Description");

        System.out.println("TagID: " + tagId + ", TagName: " + tagName + ", Description: " + description);
      }

    } catch (ClassNotFoundException e) {
      System.err.println("JDBC 驱动程序未找到: " + e.getMessage());
    } catch (SQLException e) {
      System.err.println("数据库错误: " + e.getMessage());
    } finally {
      // 关闭资源
      try {
        if (resultSet != null) {
          resultSet.close();
        }
        if (statement != null) {
          statement.close();
        }
        if (connection != null) {
          connection.close();
        }
      } catch (SQLException e) {
        System.err.println("关闭资源时出错: " + e.getMessage());
      }
    }
  }
}
