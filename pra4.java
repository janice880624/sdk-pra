package prac0914;

import java.math.BigDecimal;
import java.util.Iterator;

import com.agile.api.IAgileSession;
import com.agile.api.IDataObject;
import com.agile.api.IItem;
import com.agile.api.INode;
import com.agile.api.IRow;
import com.agile.api.ITable;
import com.agile.api.ItemConstants;
import com.agile.px.ActionResult;
import com.agile.px.ICustomAction;

public class HelloWorld implements ICustomAction{

	@Override
	public ActionResult doAction(IAgileSession session, INode arg1, IDataObject affectedObject) {
		try {
			System.out.println("------------------------------------------------");
			System.out.println("程式開始");

			// set
			int money = 0;
			int money2 = 0;
			int total = 0;
			int total2 = 0;

			IItem item = (IItem) session.getObject(ItemConstants.CLASS_PART, affectedObject.getName());
 			ITable bomTable = item.getTable(ItemConstants.TABLE_BOM);
 			Iterator it = (Iterator) bomTable.iterator();
 			while(it.hasNext()) {

				// 計算成本 numeric
 				IRow row = (IRow) it.next();
 				String qty = (String) row.getValue(ItemConstants.ATT_BOM_QTY);
 				Double cost = (Double) row.getValue(ItemConstants.ATT_BOM_ITEM_P2_NUMERIC01);
 				money = Integer.valueOf(qty) * (int) Math.round (cost);
 				total = total + money;

 				// 計算單價 money
 	 			String salePrice = row.getValue(ItemConstants.ATT_BOM_ITEM_P2_MONEY01).toString();
 				String price = salePrice.replace("USD","");
 				int nprice = new BigDecimal(price).intValue();
 				money2 = Integer.valueOf(qty) * nprice;
 				total2 = total2 + money2;

 			}
 			System.out.println(total);

 			String ans = "利用 numeric 計算總和為" + String.valueOf(total) + "利用 money 計算總和為" + String.valueOf(total2);

			return new ActionResult(ActionResult.STRING, ans);
		} catch (Exception e) {
			e.printStackTrace();
			return new ActionResult(ActionResult.EXCEPTION, e);
		}
	}
}
