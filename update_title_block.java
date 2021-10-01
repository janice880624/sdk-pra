package prac0914;

import java.util.Iterator;

import com.agile.api.IAdmin;
import com.agile.api.IAgileClass;
import com.agile.api.IAgileSession;
import com.agile.api.IItem;
import com.agile.api.INode;
import com.agile.api.IQuery;
import com.agile.api.IRow;
import com.agile.api.ITable;
import com.agile.api.ItemConstants;
import com.agile.px.ActionResult;
import com.agile.px.EventActionResult;
import com.agile.px.IEventAction;
import com.agile.px.IEventInfo;

public class EventHelloWorld implements IEventAction{

	@Override
	public EventActionResult doAction(IAgileSession session, INode actionNode, IEventInfo req) {

		try {
			// TODO Auto-generated method stub
			System.out.println("連結成功");
			System.out.println("-----------");

			IItem item = (IItem) session.getObject(ItemConstants.CLASS_DOCUMENT, req.getEventName());
			System.out.println("表單名稱: " + item.getName());
			String ChangeName = item.getCell(1539) + " " + item.getCell(1540);


			String main = (String)item.getCell(1539).toString();
			String subsitude = (String)item.getCell(1540).toString();
//
			IAdmin admin = session.getAdminInstance();
			IAgileClass cls = admin.getAgileClass("Alternative BOM (全域替代)");
			IQuery query = (IQuery)session.createObject(IQuery.OBJECT_TYPE, cls);
			query.setCaseSensitive(false);
			query.setCriteria("[Title Block.Number (文件編號)] starts with 'AT'");
			ITable results = query.execute();
			Iterator it = results.iterator();

			if (main.equals(subsitude)) {
				System.out.println("請更改選單");
			} else {
				System.out.println("通過");
				while(it.hasNext()) {
					IRow row = (IRow) it.next();
					System.out.println(row.getValue(ItemConstants.ATT_TITLE_BLOCK_NUMBER));
					String OriginalName = row.getReferent().getCell(1539) + " " + row.getReferent().getCell(1540);

					if (ChangeName.equals(OriginalName)) {
						if (row.getValue(ItemConstants.ATT_TITLE_BLOCK_NUMBER).equals(item.getName())) {
							continue;
						}else {
							System.out.println("有相等的請更改");
						}
					} else {
						System.out.println("更改成功");
					}
				}
				System.out.println("xxxxxxxxxxxxxxxxxxxx");
			}
			return new EventActionResult(req, new ActionResult(ActionResult.EXCEPTION, "Hello World"));

		}catch (Exception e) {
			e.printStackTrace();
			return new EventActionResult(req, new ActionResult(ActionResult.EXCEPTION, e));
		}
	}
}
