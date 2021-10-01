package prac0914;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Calendar;
import java.util.Iterator;

import com.agile.api.ChangeConstants;
import com.agile.api.IAgileSession;
import com.agile.api.IAttachmentFile;
import com.agile.api.IDataObject;
import com.agile.api.IItem;
import com.agile.api.INode;
import com.agile.api.IRow;
import com.agile.api.ITable;
import com.agile.api.ITwoWayIterator;
import com.agile.api.ItemConstants;
import com.agile.px.ActionResult;
import com.agile.px.ICustomAction;
import com.agile.px.IObjectEventInfo;

public class ActionPra9 implements ICustomAction{

	@Override
	public ActionResult doAction(IAgileSession session, INode arg1, IDataObject affectedObject) {
		try {
			// TODO Auto-generated method stub
			System.out.println("------------------------------------------------");
			System.out.println("程式開始");

			IItem item = (IItem) session.getObject(ItemConstants.CLASS_PART, affectedObject.getName());
			ITable attTable = item.getAttachments();
			ITwoWayIterator it = attTable.getTableIterator();

			while (it.hasNext()) {
				byte[] buffer = new byte[1024];
		        int size = 0;
		        String file = "";

		        //time
				Calendar Cld = Calendar.getInstance();
				String YY = Integer.toString(Cld.get(Calendar.YEAR));
				String MM = Integer.toString(Cld.get(Calendar.MONTH)+1);
				String DD = Integer.toString(Cld.get(Calendar.DATE));
				String HH = Integer.toString(Cld.get(Calendar.HOUR_OF_DAY));
				String mm = Integer.toString(Cld.get(Calendar.MINUTE));
				String SS = Integer.toString(Cld.get(Calendar.SECOND));
				String MI = Integer.toString(Cld.get(Calendar.MILLISECOND));

				IRow row = (IRow)it.next();
				String filename = row.getValue(ChangeConstants.ATT_ATTACHMENTS_FILENAME).toString();
				String[] strs = filename.split("\\.");
				String time = YY + MM + DD + HH + mm + SS + MI;
				file = "C:\\Users\\Administrator\\Desktop\\data\\" + time + "." + strs[1];
				System.out.println(file);

				InputStream in = ((IAttachmentFile)row).getFile();

				while ((size = in.read(buffer)) != -1){
					OutputStream output = new FileOutputStream(file, true);
					System.out.println(output);
					output.write(buffer, 0, size);
					System.out.println("儲存完成");
		        }

				Thread.sleep(1000);
				System.out.println("////////////////");
			}

			System.out.println("程式結束");
			return new ActionResult(ActionResult.STRING, "download");

		}catch (Exception e) {
			e.printStackTrace();
			return new ActionResult(ActionResult.STRING, "download");
		}
	}

}
