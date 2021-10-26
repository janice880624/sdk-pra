package prac0914;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

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

	@Override
	public EventActionResult doAction(IAgileSession session, INode actionNode, IEventInfo affectedObject) {
		
		// use log
		Ini ini = new Ini("C:\\Agile\\Config.ini");
		Log log = new Log();
		
		try {
			
			log.logSeparatorBar();
			log.setTopic("AnselmTools_ini_log_Tutorial_");
			log.log("-------------------------�{���}�l-------------------------");
			
			/*****************************************************************************
			 * Step.1 MPN Lifecycle ���A
			 * Step.2 �ˬd Manufacturer Part �O�_������
			 ******************************************************************************/
					
			// init
			ArrayList<String> incompatibleList = new ArrayList<String>(5);
			ArrayList<String> noAttachments = new ArrayList<String>(5);
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
				
				IItem item1 = (IItem) row.getReferent();
				ITable table = (ITable) item1.getTable(ItemConstants.TABLE_MANUFACTURERS);
				Iterator<?> it = table.iterator();
				item1.refresh();
				while (it.hasNext()) {
					int attNum = 0;
					IRow row1 = (IRow) it.next();
					
					// �ˬd MPN Lifecycle ���A				
					String manuName = row1.getValue(ItemConstants.ATT_MANUFACTURERS_MFR_NAME)+"";
					log.log("Mfr. Name: " + manuName);
					String manuNum = row1.getValue(ItemConstants.ATT_MANUFACTURERS_MFR_PART_NUMBER)+"";
					log.log("Mfr. Part Number: " + manuNum);
					String manuLifeCycle = row1.getValue(ItemConstants.ATT_MANUFACTURERS_MFR_TAB_LIST01)+"";
					log.log("�ثe���A�O: " + manuLifeCycle);
					
					if (manuLifeCycle.equals("Disqualified") || manuLifeCycle.equals("Obsolete")) {
						log.log("MPN Lifecycle ���A���ŦX����");
						incompatibleList.add(row1.getValue(ItemConstants.ATT_MANUFACTURERS_MFR_NAME)+"");
					} 
					

					// �ˬd Manufacturer Part �O�_������					
					params.put(ManufacturerPartConstants.ATT_GENERAL_INFO_MANUFACTURER_NAME, manuName);
					params.put(ManufacturerPartConstants.ATT_GENERAL_INFO_MANUFACTURER_PART_NUMBER, manuNum);
					IManufacturerPart mfrPrat = (IManufacturerPart) session.getObject(IManufacturerPart.OBJECT_TYPE, params);
					log.log(mfrPrat.getName());
					ITable mfratt = mfrPrat.getTable(ItemConstants.TABLE_ATTACHMENTS);
					Iterator<?> it_mfratt = mfratt.iterator();
					
					while (it_mfratt.hasNext()) {
						IRow attrow = (IRow) it_mfratt.next();						
						attNum += 1;						
					}
					
					if (attNum == 0) {
						noAttachments.add(mfrPrat.getName());
					}
				}
				
				log.log(incompatibleList);
				
				if (incompatibleList.size() == 0 || noAttachments.size() == 0) {
					IStatus nextStatus_eco = change.getDefaultNextStatus();
					log.log("Next default status = " + nextStatus_eco.getName());
					log.log("------------------------- �i������ -------------------------");
				} else {
					log.log("------------------------- �ݭn�ץ� -------------------------");
					errorlog = errorlog + "Mfr. Part Number " + incompatibleList + " MPN Lifecycle ���A�D Active �нT�{, Manufacturer Part " + noAttachments + " �L���� ";
					throw new Exception(errorlog);
				}
			}
			
			return new EventActionResult(info, new ActionResult(ActionResult.STRING, "finish"));
			
		} catch (Exception e) {
			e.printStackTrace();
			return new EventActionResult(affectedObject, new ActionResult(ActionResult.EXCEPTION, e));
		}

	}

}

