import java.sql.Connection;
import java.sql.DriverManager;

public class test1 {
  public static void main(String[] args) {
    Connection conn = null;
    try {
      // 加载驱动程序
      Class.forName("com.osisoft.jdbc.Driver");
      // 建立连接
      conn = DriverManager.getConnection(
          "jdbc:osisoft:pi://192.168.121.153:5460/MyDatabase",
          "administrator", "A1.");
      // 检查连接状态
      if (conn != null) {
        System.out.println("连接成功！");
      }
    } catch (Exception e) {
      e.printStackTrace();
    } finally {
      // 关闭连接
      if (conn != null) {
        try {
          conn.close();
        } catch (Exception e) {
          e.printStackTrace();
        }
      }
    }
  }
}
