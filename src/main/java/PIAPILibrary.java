import com.sun.jna.Library;
import com.sun.jna.Native;
import com.sun.jna.ptr.FloatByReference;
import com.sun.jna.ptr.IntByReference;

public interface PIAPILibrary extends Library {

  PIAPILibrary INSTANCE = (PIAPILibrary) Native.loadLibrary("piapi32", PIAPILibrary.class);
  // piapi32中的接口 见piapi.h
  int piut_setservernode(String servername);
  int pitm_systime();

  // 查找tagname指定的位号并存入pt中
  // PIINT32 pipt_findpoint( char PIPTR *tagname, int32 PIPTR *pt );
  int pipt_findpoint(String tagname, IntByReference pt);

  // 获取PI时间
  // 这个函数解析传入的时间字符串并返回pi本地时间。如果传递的字符串是相对时间，则使用reltime作为计算绝对时间的起点。
  // 如果时间字符串有效则返回0，如果无效则返回-1。有效的时间字符串是一种绝对格式，包含dd-mm -yy hh:mm:ss、+|- n和|h|m|s格式的相对时间、
  // 用单词指定的绝对时间（今天、昨天、星期天、星期一、…）、用星号表示当前时间，或者使用单词绝对时间和相对时间之一的组合时间。
  // PIINT32 pitm_parsetime( char PIPTR *str, int32 reltime, int32 PIPTR *timedate);
  int pitm_parsetime(String timestr, int reltime, IntByReference timedate);

  // 获取timedate时间位号pt的值并存入rval、istat中
  // 这个函数检索发送到pi系统的特定点的最新值。返回快照的工程单位值、状态码、时间和日期。
  // 如果点类型不是实数，则以istat形式返回值，而rval参数返回为0。
  // PIINT32 pisn_getsnapshot( int32 pt, float PIPTR *rval, int32 PIPTR *istat, int32 *timedate );
  int pisn_getsnapshot(int pt, FloatByReference rval, IntByReference istat, IntByReference timedate);

  // 写入位号pt在时间timedate的值
  // PIINT32 pisn_putsnapshot( int32 pt, float rval,  int32 istat, int32 timedate );
  int pisn_putsnapshot(int pt, float rval, int istat, int timedate);

  // 写入位号pt的值rval、istat、timedate、wait
  // PIINT32 piar_putvalue( int32 pt, float rval, int32 istat, int32 timedate, int32 wait );
  int piar_putvalue(int pt, float rval, int istat, int timedate, int wait);

  //这个函数检索服务器上定义的所有单位。第一次调用该函数时，应该将Index设置为0，随后的调用应该将Index增加到number-1，以检索剩余的单元名称。
  // int32 piba_getunit( char PIPTR * unit, int32 len, int32 index,  int32 PIPTR * number );
  int piba_getunit(String unit, int len, int index, IntByReference number);
}

