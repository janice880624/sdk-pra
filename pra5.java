package prac0914;

import java.util.Iterator;

import com.agile.api.ChangeConstants;
import com.agile.api.IAgileSession;
import com.agile.api.IChange;
import com.agile.api.IDataObject;
import com.agile.api.INode;
import com.agile.api.IRow;
import com.agile.api.ITable;
import com.agile.px.ActionResult;
import com.agile.px.EventActionResult;
import com.agile.px.ICustomAction;
import com.agile.px.IEventAction;
import com.agile.px.IEventInfo;
import com.agile.px.IObjectEventInfo;

public class EventPra5 implements IEventAction{

	@Override
public EventActionResult doAction(IAgileSession session, INode actionNode, IEventInfo affectedObject) {

		try {
			// TODO Auto-generated method stub
			System.out.println("------------------------------------------------");
			System.out.println("程式開始");

			IObjectEventInfo info = (IObjectEventInfo) affectedObject;
			IDataObject obj = info.getDataObject();
			IChange change = (IChange)session.getObject(ChangeConstants.CLASS_ECO, obj.getName());

			ITable affectedItems = change.getTable(ChangeConstants.TABLE_AFFECTEDITEMS);

			Iterator it = affectedItems.iterator();

			while(it.hasNext()) {
				IRow row = (IRow) it.next();
				System.out.println(row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_ITEM_NUMBER));
				System.out.println("OLD_REV：" + row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV));
				System.out.println("NEW_REV：" + row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV));

				// 字串是否是數字
				boolean integerOrNot1 = ((String) row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV)).matches("[+-]?\\d*(\\.\\d+)?");

				// 是否為整數
				boolean integerOrNot2 = ((String) row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV)).matches("-?\\d+");

				if (row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV).equals("")) {
					if (row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV).equals("1") == false) {
						row.setValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV, "1");
						System.out.println("ex: 0 -> 1");
					}
				} else if (integerOrNot1 == true && integerOrNot2 == true){
					int i = Integer.valueOf((String) row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV)) + 1;
					if (row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV).equals(Integer.toString(i)) == false) {
						row.setValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV, Integer.toString(i));
						System.out.println("ex: 1 -> 2");
					}
				} else if (integerOrNot1 == true && integerOrNot2 == false){
					double num2 = Double.parseDouble((String) row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV));
					double j = num2 + 0.1;
					j = Math.round(j*100.0)/100.0;
					if (row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV).equals(Double.toString(j)) == false) {
						row.setValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV, Double.toString(j));
						System.out.println("ex: 0.1 -> 0.2");
					}
				} else if (integerOrNot1 == false && integerOrNot2 == false) {
					String first = ((String) row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV)).substring(0, 1);
					String x = ((String) row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV)).replaceAll("[^\\d]", "");
					int k = Integer.parseInt(x);
				    if (k < 10){
				    	k = Integer.parseInt(x) + 1;
				    	if (row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV).equals(first + "0" +Integer.toString(k)) == false) {
							row.setValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV, first + "0" +Integer.toString(k));
							System.out.println("ex: 0.1 -> 0.2");
						}
				    } else if(k == 99) {
				    	char result = first.charAt(0);
				        int asciiValue = result;
				        int asciiValue2 = asciiValue;
				        String second = Character.toString((char) (asciiValue2 + 1));
				        if (row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV).equals(second + "01") == false) {
							row.setValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV, second + "01");
							System.out.println("ex: 0.1 -> 0.2");
						}
				    }else{
				    	k = Integer.parseInt(x) + 1;
				    	if (row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV).equals(first + Integer.toString(k)) == false) {
							row.setValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV, first + Integer.toString(k));
							System.out.println("ex: 0.1 -> 0.2");
						}
				    }
				    System.out.println("ex: A01 -> A02");
				}



				System.out.println("-----------");

			}

			System.out.println("程式結束");
			return new EventActionResult(info, new ActionResult(ActionResult.EXCEPTION, "finish"));

		}catch (Exception e) {
			e.printStackTrace();
			return new EventActionResult(affectedObject, new ActionResult(ActionResult.EXCEPTION, e));
		}
	}

}
