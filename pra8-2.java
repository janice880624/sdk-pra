package prac0914;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStreamReader;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.agile.api.IAgileSession;
import com.agile.api.IDataObject;
import com.agile.api.IItem;
import com.agile.api.INode;
import com.agile.api.ItemConstants;
import com.agile.px.ActionResult;
import com.agile.px.ICustomAction;
import com.agile.px.IObjectEventInfo;

public class ActionPra8 implements ICustomAction{

	@SuppressWarnings("resource")
	@Override
	public ActionResult doAction(IAgileSession session, INode arg1, IDataObject affectedObject) {

		// TODO Auto-generated method stub
		String fileName = "C:\\Users\\Administrator\\Desktop\\data\\test.txt";
		BufferedReader reader = null;
		String inputtext = "";

		try {
			System.out.println("程式開始");
			System.out.println("-----------");

			IItem item = (IItem) session.getObject(ItemConstants.CLASS_PART, affectedObject.getName());
			System.out.println("表單名稱: " + item.getName());

			// 讀取 txt
			reader = new BufferedReader(new InputStreamReader(new FileInputStream(fileName), "UTF-8"));
			String str = null;

			while ((str = reader.readLine()) != null) {
				inputtext = inputtext + str + "\n";
			 }

			System.out.println(inputtext);
			session.disableAllWarnings();
			item.setValue(ItemConstants.ATT_TITLE_BLOCK_DESCRIPTION, inputtext);
			session.enableAllWarnings();
			System.out.println("-----------");
			System.out.println("程式結束");

			return new ActionResult(ActionResult.STRING, "finish");
		}catch (Exception e) {
			e.printStackTrace();
			return new ActionResult(ActionResult.EXCEPTION, e);
		}
	}

}
