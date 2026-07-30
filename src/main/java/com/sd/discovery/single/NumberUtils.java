package com.sd.discovery.single;

import com.google.common.base.Optional;
import org.apache.commons.lang3.ObjectUtils;
import java.math.BigDecimal;
import java.text.DecimalFormat;

/**
 * 数字计算比较工具类
 *
 * @Author: gaoweiqi
 * @CreateDate: 2022/12/12 13:15
 */
public class NumberUtils {

  public static DecimalFormat DF_FOUR = new DecimalFormat("#0.0000");

  private static final double EPSILON = 1e-6;

  /**
   * @param b1
   * @param b2
   * @return
   */
  public static int safeCompare(BigDecimal b1, BigDecimal b2) {
    if (null == b1) {
      b1 = BigDecimal.ZERO;
    }
    if (null == b2) {
      b2 = BigDecimal.ZERO;
    }
    return b1.compareTo(b2);
  }

  /**
   * @param b1
   * @param bn
   * @return
   */
  public static BigDecimal safeAdd(BigDecimal b1, BigDecimal... bn) {
    if (null == b1) {
      b1 = BigDecimal.ZERO;
    }
    if (null != bn) {
      for (BigDecimal b : bn) {
        b1 = b1.add(null == b ? BigDecimal.ZERO : b);
      }
    }
    return b1;
  }

  public static Integer safeAdd(Integer b1, Integer... bn) {
    if (null == b1) {
      b1 = 0;
    }
    Integer r = b1;
    if (null != bn) {
      for (Integer b : bn) {
        r += Optional.fromNullable(b).or(0);
      }
    }
    return r > 0 ? r : 0;
  }

  public static Long safeAdd(Long b1, Long... bn) {
    if (null == b1) {
      b1 = 0L;
    }
    Long r = b1;
    if (null != bn) {
      for (Long b : bn) {
        r += Optional.fromNullable(b).or(0L);
      }
    }
    return r > 0 ? r : 0;
  }

  /**
   * 计算金额
   *
   * @param b1
   * @param bn
   * @return
   */
  public static BigDecimal safeSubtract(BigDecimal b1, BigDecimal... bn) {
    return safeSubtract(true, b1, bn);
  }

  public static BigDecimal safeSubtract(Boolean isZero, BigDecimal b1, BigDecimal... bn) {
    if (null == b1) {
      b1 = BigDecimal.ZERO;
    }
    BigDecimal r = b1;
    if (null != bn) {
      for (BigDecimal b : bn) {
        r = r.subtract((null == b ? BigDecimal.ZERO : b));
      }
    }
    return isZero ? (r.compareTo(BigDecimal.ZERO) == -1 ? BigDecimal.ZERO : r) : r;
  }

  public static Integer safeSubtract(Integer b1, Integer... bn) {
    return safeSubtract(true, b1, bn);
  }

  public static Integer noSafeSubtract(Integer b1, Integer b2) {
    return safeSubtract(false, b1, b2);
  }

  public static Integer safeSubtract(Boolean isZero, Integer b1, Integer... bn) {
    if (null == b1) {
      b1 = 0;
    }
    Integer r = b1;
    if (null != bn) {
      for (Integer b : bn) {
        r -= Optional.fromNullable(b).or(0);
      }
    }
    return isZero ? (null != r && r > 0 ? r : 0) : r;
  }

  public static <T extends Number> BigDecimal safeDivide(T b1, T b2) {
    return safeDivide(b1, b2, BigDecimal.ZERO);
  }

  public static <T extends Number> BigDecimal safeDivide(T b1, T b2, BigDecimal defaultValue) {
    if (null == b1 || null == b2) {
      return defaultValue;
    }
    try {
      return BigDecimal.valueOf(b1.doubleValue())
          .divide(BigDecimal.valueOf(b2.doubleValue()), 2, BigDecimal.ROUND_HALF_EVEN);
    } catch (Exception e) {
      return defaultValue;
    }
  }

  public static <T extends Number> BigDecimal safeDivideDown(T b1, T b2) {
    return safeDivideDown(b1, b2, BigDecimal.ZERO);
  }

  public static <T extends Number> BigDecimal safeDivideDown(T b1, T b2, BigDecimal defaultValue) {
    if (null == b1 || null == b2) {
      return defaultValue;
    }
    try {
      return BigDecimal.valueOf(b1.doubleValue())
          .divide(BigDecimal.valueOf(b2.doubleValue()), 2, BigDecimal.ROUND_DOWN);
    } catch (Exception e) {
      return defaultValue;
    }
  }

  public static <T extends Number> BigDecimal safeDivideUp(T b1, T b2) {
    return safeDivideUp(b1, b2, BigDecimal.ZERO);
  }

  public static <T extends Number> BigDecimal safeDivideUp(T b1, T b2, BigDecimal defaultValue) {
    if (null == b1 || null == b2) {
      return defaultValue;
    }
    try {
      return BigDecimal.valueOf(b1.doubleValue()).divide(BigDecimal.valueOf(b2.doubleValue()), 2, BigDecimal.ROUND_UP);
    } catch (Exception e) {
      return defaultValue;
    }
  }

  public static <T extends Number> BigDecimal safeMultiply(T b1, T b2, Integer scale) {
    if (null == b1 || null == b2) {
      return BigDecimal.ZERO;
    }
    if (scale != null) {
      return BigDecimal.valueOf(b1.doubleValue()).multiply(BigDecimal.valueOf(b2.doubleValue()))
          .setScale(scale, BigDecimal.ROUND_HALF_UP);
    }
    return BigDecimal.valueOf(b1.doubleValue()).multiply(BigDecimal.valueOf(b2.doubleValue()));
  }

  public interface Filter<TElement> {

    boolean apply(TElement element);
  }

  /**
   * 四舍五入保留小数点后两位
   *
   * @param d
   * @return
   */
  public static BigDecimal safeSetScale(BigDecimal d) {
    return null == d ? BigDecimal.ZERO : d.setScale(2, BigDecimal.ROUND_HALF_UP);
  }

  /**
   * 取几个数的最小值
   *
   * @param bn
   * @return
   */
  public static Integer getMinValue(Integer... bn) {
    Integer minValue = null;
    for (Integer b : bn) {
      if (minValue == null) {
        minValue = b;
      }
      if (b == null) {
        continue;
      }
      if (lt(b, minValue)) {
        minValue = b;
      }
    }
    return minValue;
  }

  /**
   * 取几个数的最小值
   *
   * @param bn
   * @return
   */
  public static Long getMinValue(Long... bn) {
    Long minValue = null;
    for (Long b : bn) {
      if (minValue == null) {
        minValue = b;
      }
      if (b == null) {
        continue;
      }
      if (lt(b, minValue)) {
        minValue = b;
      }
    }
    return minValue;
  }

  public static <T extends Comparable<? super T>> int compare(T o1, T o2) {
    return ObjectUtils.compare(o1, o2);
  }

  /**
   * 相等
   *
   * @param o1
   * @param o2
   * @return
   */
  public static <T extends Comparable<? super T>> boolean eq(T o1, T o2) {
    return isNotNull(o1, o2) && compare(o1, o2) == 0;
  }

  /**
   * 不等
   *
   * @param o1
   * @param o2
   * @return
   */
  public static <T extends Comparable<? super T>> boolean nq(T o1, T o2) {
    return !eq(o1, o2);
  }

  /**
   * 大于等于
   *
   * @param o1
   * @param o2
   * @return
   */
  public static <T extends Comparable<? super T>> boolean ge(T o1, T o2) {
    return isNotNull(o1, o2) && compare(o1, o2) >= 0;
  }

  /**
   * 大于
   *
   * @param o1
   * @param o2
   * @param <T>
   * @return
   */
  public static <T extends Comparable<? super T>> boolean gt(T o1, T o2) {
    return isNotNull(o1, o2) && compare(o1, o2) == 1;
  }

  /**
   * 小于等于
   *
   * @param o1
   * @param o2
   * @param <T>
   * @return
   */
  public static <T extends Comparable<? super T>> boolean le(T o1, T o2) {
    return isNotNull(o1, o2) && compare(o1, o2) <= 0;
  }

  /**
   * 小于
   *
   * @param o1
   * @param o2
   * @param <T>
   * @return
   */
  public static <T extends Comparable<? super T>> boolean lt(T o1, T o2) {
    return isNotNull(o1, o2) && compare(o1, o2) == -1;
  }

  public static <T extends Number> boolean isTrue(T object) {
    return isNotNull(object) && object.intValue() == 1;
  }

  public static <T extends Number> boolean isFalse(T object) {
    return isNotNull(object) && object.intValue() == 0;
  }

  /**
   * 不等于null
   *
   * @param objects
   * @param <T>
   * @return
   */
  public static <T> boolean isNotNull(T... objects) {
    if (null != objects && objects.length > 0) {
      for (T object : objects) {
        if (null == object) {
          return false;
        }
      }
      return true;
    }
    return false;
  }

  /**
   * 计算结果保留四位小数
   * @return
   */
  public static Double dataFormatFour(Double data) {
    return Double.valueOf(DF_FOUR.format(data));
  }

  /**
   * 判断double类型值是否为0
   * @param value 待计算的double类型值
   * @return true:为0 false:非0
   */
  public static boolean isZero(double value) {
    return Math.abs(value) < EPSILON;
  }

  public static void main(String[] args) {
    Integer integer = NumberUtils.safeSubtract(10, 11);
    Integer integer1 = NumberUtils.noSafeSubtract(10, 11);
    System.out.println("integer = " + integer);
    System.out.println("integer1 = " + integer1);

    System.out.println("glliu===========");
    boolean aFalse = isZero(0.00000005D);
    System.out.println("0.00000005D结果:" + aFalse);
    boolean zero = isZero(0.0001D);
    System.out.println("0.0001D结果:" + zero);

  }
}