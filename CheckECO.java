package prac0914;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

import com.agile.api.APIException;
import com.agile.api.ChangeConstants;
import com.agile.api.IAgileSession;
import com.agile.api.IChange;
import com.agile.api.IDataObject;
import com.agile.api.IItem;
import com.agile.api.IManufacturerPart;
import com.agile.api.INode;
import com.agile.api.IRow;
import com.agile.api.IStatus;
import com.agile.api.ITable;
import com.agile.api.ItemConstants;
import com.agile.api.ManufacturerPartConstants;
import com.agile.px.ActionResult;
import com.agile.px.EventActionResult;
import com.agile.px.IEventAction;
import com.agile.px.IEventInfo;
import com.agile.px.IObjectEventInfo;
import com.anselm.tools.record.Ini;
import com.anselm.tools.record.Log;

public class CheckECO implements IEventAction{

	private boolean fasle;

	@Override
	public EventActionResult doAction(IAgileSession session, INode actionNode, IEventInfo affectedObject) {
		
		// use log
		Ini ini = new Ini("C:\\Agile\\Config.ini");
		Log log = new Log();
		StringBuilder error = new StringBuilder();
		
		try {
			
			log.logSeparatorBar();
			log.setTopic("AnselmTools_ini_log_Tutorial_");
			log.log("-------------------------程式開始-------------------------");
				
			// init
			ArrayList<String> incompatibleList = new ArrayList<String>(5);
			ArrayList<String> noAttachments = new ArrayList<String>(5);
			ArrayList<String> errorIPN = new ArrayList<String>(5);
			
			String errorlog = "";
			HashMap params = new HashMap();  
			
			IObjectEventInfo info = (IObjectEventInfo) affectedObject;
			IDataObject obj = info.getDataObject();	
			IChange change = (IChange) session.getObject(ChangeConstants.CLASS_ECO, obj.getName());
			ITable table_eco = change.getTable(ChangeConstants.TABLE_AFFECTEDITEMS);
			Iterator it_eco = table_eco.iterator();
			
			while(it_eco.hasNext()){
				IRow row = (IRow) it_eco.next();
				log.log(row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_ITEM_NUMBER));	
				
				String affIPN = (String) row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_TEXT06);
				log.log(row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_TEXT06));
				
				IItem item = (IItem) session.getObject(ItemConstants.CLASS_PART, row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_ITEM_NUMBER));		
				
				IItem item1 = (IItem) row.getReferent();
				ITable table = (ITable) item1.getTable(ItemConstants.TABLE_MANUFACTURERS);
				Iterator<?> it = table.iterator();
				item1.refresh();
				while (it.hasNext()) {
					int attNum = 0;
					IRow row1 = (IRow) it.next();
					
					// 檢查 MPN Lifecycle 狀態				
					String manuName = row1.getValue(ItemConstants.ATT_MANUFACTURERS_MFR_NAME)+"";
					log.log("Mfr. Name: " + manuName);
					
					String manuNum = row1.getValue(ItemConstants.ATT_MANUFACTURERS_MFR_PART_NUMBER)+"";
					log.log("Mfr. Part Number: " + manuNum);
					
					String manuLifeCycle = row1.getValue(ItemConstants.ATT_MANUFACTURERS_MFR_TAB_LIST01)+"";
					log.log("目前狀態是: " + manuLifeCycle);
					
					// 檢查 Manufacturer LifeCycle 狀態
					checkLifeCycle(manuLifeCycle, manuName, incompatibleList, log);
					
					// 檢查 Manufacturer Part 是否有附件	
					checkMfrPart(params, manuName, manuNum, session, attNum, affIPN, errorIPN, row1, noAttachments, error, log);

				}
				
				log.log("錯誤資料確認");
				log.log(incompatibleList);
				log.log(noAttachments);
				log.log(errorIPN);
				log.log("----------------------------");
				
				if (incompatibleList.size() == 0 && noAttachments.size() == 0 && errorIPN.size() == 0) {
					IStatus nextStatus_eco = change.getDefaultNextStatus();
					log.log("Next default status = " + nextStatus_eco.getName());
					log.log("------------------------- 進站完成 -------------------------");
				} else {
					log.log("------------------------- 需要修正 -------------------------");
					
					if (incompatibleList.size() == 0)
						errorlog = errorlog + "";
					else
						errorlog = errorlog + "Mfr. Part Number " + incompatibleList + " MPN Lifecycle 狀態非 Active 請確認, ";
					
					if (noAttachments.size() == 0)
						errorlog = errorlog + "";
					else
						errorlog = errorlog + " Manufacturer Part " + noAttachments + " 無附件, ";
					
					if (errorIPN.size() == 0)
						errorlog = errorlog + "";
					else
						errorlog = errorlog + " Manufacturer Part " + errorIPN + " IPN 有誤 ";
										
//					errorlog = errorlog + "Mfr. Part Number " + incompatibleList + " MPN Lifecycle 狀態非 Active 請確認, Manufacturer Part " + noAttachments + " 無附件, Manufacturer Part " + errorIPN + "IPN 有誤";
					throw new Exception(errorlog);
				}
			}
			
			return new EventActionResult(info, new ActionResult(ActionResult.STRING, "finish"));
			
		} catch (Exception e) {
			e.printStackTrace();
			log.logException(e);
			return new EventActionResult(affectedObject, new ActionResult(ActionResult.EXCEPTION, e));
		} 
		

	}
	
	
	/************************************************************************
	 檢查是否有附件
	*************************************************************************/
	public void checkMfrPart(HashMap params, String manuName, String manuNum, IAgileSession session, int attNum, String affIPN, ArrayList errorIPN, IRow row1, ArrayList noAttachments, StringBuilder error, Log log) throws APIException {
		params.put(ManufacturerPartConstants.ATT_GENERAL_INFO_MANUFACTURER_NAME, manuName);
		params.put(ManufacturerPartConstants.ATT_GENERAL_INFO_MANUFACTURER_PART_NUMBER, manuNum);
		IManufacturerPart mfrPrat = (IManufacturerPart) session.getObject(IManufacturerPart.OBJECT_TYPE, params);
		log.log(mfrPrat.getName());
		ITable mfratt = mfrPrat.getTable(ItemConstants.TABLE_ATTACHMENTS);
		Iterator<?> it_mfratt = mfratt.iterator();
		
		while (it_mfratt.hasNext()) {
			IRow attrow = (IRow) it_mfratt.next();						
			attNum += 1;
			if(checkIPN(attrow, affIPN, error, log)) {
				errorIPN.add(row1.getValue(ItemConstants.ATT_MANUFACTURERS_MFR_PART_NUMBER)+"");
			}
		}
		
		if (attNum == 0) {
			noAttachments.add(mfrPrat.getName());
		}		
	}

	/************************************************************************
	 檢查 Lifecycle
	*************************************************************************/
	@SuppressWarnings({ "unchecked", "rawtypes" })
	public void checkLifeCycle(String manuLifeCycle, String manuName, ArrayList incompatibleList, Log log) throws APIException {
		if (manuLifeCycle.equals("Disqualified") || manuLifeCycle.equals("Obsolete")) {
			log.log("MPN Lifecycle 狀態不符合條件");
			incompatibleList.add(manuName);
		} 
	}

	/************************************************************************
	 檢查 IPN
	*************************************************************************/
	public boolean checkIPN(IRow attrow, String affIPN, StringBuilder error, Log log) throws APIException {
		System.out.println(attrow.getValue(ItemConstants.ATT_ATTACHMENTS_TEXT01));
		if (!attrow.getValue(ItemConstants.ATT_ATTACHMENTS_TEXT01).equals(affIPN)) {
			log.log("IPN 不同");
			return true;
		} else {
			return false;
		}
	}

}