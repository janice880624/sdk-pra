package prac0914;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;

import com.agile.api.ChangeConstants;
import com.agile.api.IAgileSession;
import com.agile.api.IChange;
import com.agile.api.IDataObject;
import com.agile.api.IItem;
import com.agile.api.INode;
import com.agile.api.IRow;
import com.agile.api.ITable;
import com.agile.api.ItemConstants;
import com.agile.px.ActionResult;
import com.agile.px.EventActionResult;
import com.agile.px.ICustomAction;
import com.agile.px.IEventAction;
import com.agile.px.IEventInfo;
import com.agile.px.IObjectEventInfo;

public class EventPra6 implements IEventAction{

	@Override
public EventActionResult doAction(IAgileSession session, INode actionNode, IEventInfo affectedObject) {

		try {
			// TODO Auto-generated method stub
			System.out.println("------------------------------------------------");
			System.out.println("程式開始");

			IObjectEventInfo info = (IObjectEventInfo) affectedObject;
			IDataObject obj = info.getDataObject();

			// init
			ArrayList<Integer> numList = new ArrayList<Integer>();
			ArrayList<String> nameList = new ArrayList<String>();

			IItem item = (IItem) session.getObject(ItemConstants.CLASS_PART, obj.getName());
			ITable bomTable = item.getTable(ItemConstants.TABLE_BOM);
			Iterator it = bomTable.iterator();
			Iterator it2 = bomTable.iterator();
			while(it.hasNext()) {
				IRow row = (IRow) it.next();
				System.out.println(row.getValue(ItemConstants.ATT_BOM_ITEM_NUMBER));
				System.out.println(row.getValue(ItemConstants.ATT_BOM_FIND_NUM));

				if (row.getValue(ItemConstants.ATT_BOM_FIND_NUM).equals("0")) {
					System.out.println("這個是新的");
				}

				int i = Integer.valueOf((String) row.getValue(ItemConstants.ATT_BOM_FIND_NUM));
				numList.add(i);
				nameList.add((String) row.getValue(ItemConstants.ATT_BOM_ITEM_NUMBER));
				System.out.println("--------");
			}


			while(it2.hasNext()) {
				IRow row2 = (IRow) it2.next();
				int maxElement = Collections.max(numList);
				if (row2.getValue(ItemConstants.ATT_BOM_FIND_NUM).equals("0")) {
					row2.setValue(ItemConstants.ATT_BOM_FIND_NUM, Integer.toString(maxElement+10));
					System.out.println("替換完成");
				}
			}

			System.out.println("程式結束");
			return new EventActionResult(info, new ActionResult(ActionResult.EXCEPTION, "finish"));

		}catch (Exception e) {
			e.printStackTrace();
			return new EventActionResult(affectedObject, new ActionResult(ActionResult.EXCEPTION, e));
		}
	}

}
