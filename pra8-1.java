package prac0914;

import java.io.BufferedWriter;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import com.agile.api.IAgileSession;
import com.agile.api.IDataObject;
import com.agile.api.IItem;
import com.agile.api.INode;
import com.agile.api.ItemConstants;
import com.agile.px.ActionResult;
import com.agile.px.EventActionResult;
import com.agile.px.IEventAction;
import com.agile.px.IEventDirtyCell;
import com.agile.px.IEventInfo;
import com.agile.px.IUpdateEventInfo;

public class Pra8 implements IEventAction{

	@Override
	public EventActionResult doAction(IAgileSession session, INode actionNode, IEventInfo affectedObject) {
		// TODO Auto-generated method stub
		try {
			System.out.println("------------------------------------------------");
			System.out.println("程式開始");

			int i = 0;
			String ans = "";

			String path = "C:\\Users\\Administrator\\Desktop\\data\\test.txt";
			@SuppressWarnings("resource")
			Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(path, true), "UTF-8"));


			IUpdateEventInfo info = (IUpdateEventInfo) affectedObject;
			IDataObject obj = info.getDataObject();
			IItem item = (IItem) session.getObject(ItemConstants.CLASS_PART, obj.getName());

			// getCells()
			IEventDirtyCell[] cells = info.getCells();

			// getAttributeIds()
			Integer[] attrs = info.getAttributeIds();

			for(IEventDirtyCell Cell:cells) {
				Cell.getValue();
				if (!Cell.getAttribute().toString().equals("Title Block.Description")) {
					ans = item.getCell(Integer.parseInt(attrs[i].toString())).getName() + ": "+item.getCell(Integer.parseInt(attrs[i].toString()));;
					System.out.println(ans);
					System.out.println("更改後變成:" + Cell.getValue());
					out.write(ans + "\n" + "更改後變成:" + Cell.getValue() + "\n" );
					out.flush();
					i += 1;
				}
//				else {
//					throw new Exception("變更到 Description");
//				}
			}


			out.close();

			System.out.println("程式結束");

			return new EventActionResult(info, new ActionResult(ActionResult.STRING, "pra8"));
		}catch (Exception e) {
			e.printStackTrace();
			return new EventActionResult(affectedObject, new ActionResult(ActionResult.EXCEPTION, e));
		}
	}

}
