import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.ICell;

import com.agile.api.AgileSessionFactory;
import com.agile.api.ChangeConstants;
import com.agile.api.IAdmin;
import com.agile.api.IAgileClass;
import com.agile.api.IAgileSession;
import com.agile.api.IAutoNumber;
import com.agile.api.IChange;
import com.agile.api.IItem;
import com.agile.api.IRedlined;
import com.agile.api.IRedlinedTable;
import com.agile.api.IRow;
import com.agile.api.ITable;
import com.agile.api.ITwoWayIterator;
import com.agile.api.ItemConstants;
public class session {

	public static void main(String[] args) {

		try {
			AgileSessionFactory instance = AgileSessionFactory.getInstance("http://janice-anselm:7001/Agile/");
			HashMap params = new HashMap();
			params.put(AgileSessionFactory.USERNAME, "admin");
			params.put(AgileSessionFactory.PASSWORD, "agile936");
			IAgileSession session = instance.createSession(params);
			System.out.println("連結成功");
			System.out.println("-----------");

			/*****************************************************************************
			* Step.1 讀取 Excel
			* Step.2 抓取舊料號的 where use 父階料號
			* Step.3 開立 eco
			* Step.4 將 where use 中的物件加到 affected items
			* Step.5 讀取 redlines
			******************************************************************************/
			// init
			DataFormatter df = new DataFormatter();
			ArrayList<String> oldlist = new ArrayList<String>(5);
			ArrayList<String> newlist = new ArrayList<String>(5);
			ArrayList<String> number = new ArrayList<String>(5);
			ArrayList<String> change_new = new ArrayList<String>(5);
			HashMap<String, String> map = new HashMap<String, String>();
			Map params1 = new HashMap();

			// 抓 ECR 表單
			IChange change = (IChange)session.getObject(ChangeConstants.CLASS_ECR, "R00004");
			ITable table_ecr = change.getTable(ChangeConstants.TABLE_RELATIONSHIPS);
			Iterator it_ecr = table_ecr.iterator();


			// 讀取 Excel
			String file = "C:\\Users\\janic\\OneDrive\\桌面\\anselm\\excel2.xlsx";

			// Read Excel File into workbook
			FileInputStream inp = new FileInputStream(file);
			XSSFWorkbook wb = new XSSFWorkbook(inp);
			inp.close();
			// get wb sheet(0)
			XSSFSheet sheet = wb.getSheetAt(0);
			XSSFRow row = sheet.getRow(0);

			// 抓出舊料號
			for (int i = 1; i<=sheet.getLastRowNum(); i++) {;
			    XSSFRow row1 = sheet.getRow(i);
			    for (int j = 0; j<row.getLastCellNum(); j++) {
			    	if (!df.formatCellValue(row1.getCell(j)).equals("end")) {
			    		if (j%2==0) {
			    			newlist.add(df.formatCellValue(row1.getCell(j)));
			    		} else {
			    			oldlist.add(df.formatCellValue(row1.getCell(j)));
			    			map.put(df.formatCellValue(row1.getCell(j)), df.formatCellValue(row1.getCell(j-1)));
			    		}
			    	} else {
			    		break;
			    	}
			  	}
			}
			System.out.println(map); // {BBB=AAA, DDD=CCC, FFF=EEE}
			System.out.println(newlist); // [AAA, CCC, EEE]
			System.out.println(oldlist); // [BBB, DDD, FFF]


			// 抓出所有父階料號
			for (int k=0; k<newlist.size(); k++) {
//				System.out.print(oldlist.get(k) + ":");
				IItem item = (IItem) session.getObject(ItemConstants.CLASS_PART, oldlist.get(k));
				ITable table_part = item.getTable(ItemConstants.TABLE_WHEREUSED);
				Iterator it_part = table_part.iterator();
				while(it_part.hasNext()) {
					IRow row1 = (IRow) it_part.next();
//					System.out.print(row1.getValue(ItemConstants.ATT_WHERE_USED_ITEM_NUMBER) + " ");
					if (!number.contains(row1.getValue(ItemConstants.ATT_WHERE_USED_ITEM_NUMBER))) {
						number.add((String) row1.getValue(ItemConstants.ATT_WHERE_USED_ITEM_NUMBER));
					}
				}
//				System.out.println("");
			}
			System.out.println(number); // [00000003, 00000005, 00000006]

			// 將 where use 中的物件加到 affected items
//			for (int l=0; l<number.size(); l++) {
//				IChange change2 = (IChange)session.getObject(ChangeConstants.CLASS_ECO, "C00060");
//				IItem item = (IItem)session.getObject("HDD", number.get(l));
//				ITable affectedItems = change2.getTable(ChangeConstants.TABLE_AFFECTEDITEMS);
//				session.disableAllWarnings();
//				IRow affectedItemRow = affectedItems.createRow(item);
//				session.enableAllWarnings();
//			}

			// 刪除物件
			IChange change1 = (IChange)session.getObject(ChangeConstants.CLASS_ECO, "C00060");
			ITable table_eco = change1.getTable(ChangeConstants.TABLE_AFFECTEDITEMS);
			Iterator it_eco = table_eco.iterator();

			while(it_eco.hasNext()) {
				IRow row1 = (IRow) it_eco.next();
				IItem item = (IItem) row1.getReferent();
				item.setRevision(change1);

				ITable table = (ITable) item.getTable(ItemConstants.TABLE_REDLINEBOM);
				Iterator<?> it = table.iterator();
				((IRedlinedTable)table).undoAllRedline();
				item.refresh();
				change1.refresh();
				Collection<IRow> removeRows = new ArrayList<IRow>();
				while (it.hasNext()) {
					IRow delRow = (IRow)it.next();
					String itemred = (String) delRow.getValue(ItemConstants.ATT_BOM_ITEM_NUMBER);
					System.out.println(itemred);
					boolean contains = map.containsKey(itemred);
				    removeRows.add(delRow);
				}
				table.removeAll(removeRows);
			}

			System.out.println("-----------");

			// 看有沒有在 excel 中
			Iterator add_red = table_eco.iterator();
			while(add_red.hasNext()) {
				IRow row1 = (IRow) add_red.next();
				IItem item = (IItem) row1.getReferent();
				item.setRevision(change1);
				Collection<IRow> redLineCol = new ArrayList<IRow>();
				ITable add_table = (ITable) item.getTable(ItemConstants.TABLE_REDLINEBOM);
				Iterator<?> add_it = add_table.iterator();
				while (add_it.hasNext()) {
					IRow delRow = (IRow)add_it.next();
					String bomItemNumber = delRow.getReferent().getName();
					if (delRow.isFlagSet(ItemConstants.FLAG_IS_REDLINE_REMOVED)) {
						if (!map.containsKey(bomItemNumber)) {
							redLineCol.add(delRow);
						} else {
							System.out.println(map.get(bomItemNumber));
//							IItem item1 = (IItem) session.getObject(ItemConstants.CLASS_PART, map.get(bomItemNumber));
//							IRow redlineRow = add_table.createRow(item1);
//							addRows.add(bomItemNumber);
						}
					}
				}

				if (redLineCol.size() > 0) {
					((IRedlinedTable)add_table).undoRedline(redLineCol);
				}

			}

			System.out.println("-----------");

			// 添加替代料
			Iterator add_red1 = table_eco.iterator();
//			System.out.println(table_eco.size());
//			for(int i=0; i<table_eco.size(); i++) {
//
//			}
			while(add_red1.hasNext()) {
				IRow row1 = (IRow) add_red1.next();
				IItem item = (IItem) row1.getReferent();
				item.setRevision(change1);
				System.out.println(item);
				ITable add2_table = (ITable) item.getTable(ItemConstants.TABLE_REDLINEBOM);
				ITwoWayIterator add_it1 = add2_table.getTableIterator();
				item.refresh();

				Map<Integer, String> addmap = new HashMap();
				Collection<IRow> redLineadd = new ArrayList<IRow>();

				while (add_it1.hasNext()) {
					IRow Row = (IRow)add_it1.next();
					String bomItemNumber = Row.getReferent().getName();
					if (Row.isFlagSet(ItemConstants.FLAG_IS_REDLINE_REMOVED)) {
						if (map.containsKey(bomItemNumber)) {
							System.out.println(map.get(bomItemNumber));
							IItem item1 = (IItem) session.getObject(ItemConstants.CLASS_PART, map.get(bomItemNumber));
//							addmap.put(item1);
							((Collection) addmap).add(item1);

						}
					}
				}
//				Collection<String> valueset = addmap.values();
//				System.out.println(valueset);
				add2_table.addAll((Collection) addmap);

//				IRow row2 = add2_table.createRow(addmap);
			}


			System.out.println("-----------");
			System.out.println("程式結束");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// create new change
	static IChange newChange(IAgileSession session, String classApiName) throws Exception {
		IAdmin admin = session.getAdminInstance();
		IAgileClass iac = admin.getAgileClass(classApiName);
		IAutoNumber[] numSources = iac.getAutoNumberSources();
		String nextAutoNumber = numSources[0].getNextNumber();
		IChange change = (IChange) session.createObject(iac, nextAutoNumber);
		change.setWorkflow(change.getWorkflows()[0]);
		return change;
	}
}
