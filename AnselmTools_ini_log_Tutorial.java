import com.agile.api.IAgileSession;
import com.agile.api.IChange;
import com.agile.api.IDataObject;
import com.agile.api.INode;
import com.agile.px.ActionResult;
import com.agile.px.ICustomAction;
import com.anselm.tools.agile.AUtil;
import com.anselm.tools.record.Ini;
import com.anselm.tools.record.Log;

public class AnselmTools_ini_log_Tutorial implements ICustomAction{

	@Override
	public ActionResult doAction(IAgileSession arg0, INode arg1, IDataObject ido) {
		// 初始化
		Ini ini = new Ini("C:\\Agile\\Config.ini");
		Log log = new Log();
		try {
			// log 設定
			String strLogFilePath = ini.getValue("Program Use", "LogFile") + "AnselmTools_ini_log_Tutorial_.log";
			log.setLogFile(strLogFilePath, ini);
			log.logSeparatorBar();
			log.setTopic("AnselmTools_ini_log_Tutorial_");
			
			log.log("-------------------------程式開始-------------------------");
			
			// @Step 1. 取得 admin 權限的 AgileSession 連線			
			IAgileSession admin = AUtil.getAgileSession(ini, "AgileAP");
			
			// @Step 2. 檢查是否為可執行之流程及站別
			IChange ichange = (IChange) admin.getObject(IChange.OBJECT_TYPE, ido.getName());
			log.log(1, "當前表單名稱 : " + ichange.getName());
			
			log.log("-------------------------程式結束-------------------------");
			
		} catch (Exception e) {
			e.printStackTrace();
			log.logException(e);
		}finally {
			close(log, ini);
		}
		
		return new ActionResult(ActionResult.STRING, "done!");
	}
	
	
	/**************************************************************************
	 * 解構&釋放
	 *************************************************************************/
	public void close(Log log, Ini ini) {
		try {
			log.close();
			log = null;
			ini = null;
		} catch (Exception e) {
		}
	}

}
