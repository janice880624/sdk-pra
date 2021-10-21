package prac0914;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.agile.api.APIException;
import com.agile.api.ChangeConstants;
import com.agile.api.IAdmin;
import com.agile.api.IAgileClass;
import com.agile.api.IAgileSession;
import com.agile.api.IAutoNumber;
import com.agile.api.IAttachmentFile;
import com.agile.api.IChange;
import com.agile.api.IDataObject;
import com.agile.api.IItem;
import com.agile.api.INode;
import com.agile.api.IRedlinedTable;
import com.agile.api.IRow;
import com.agile.api.IStatus;
import com.agile.api.ITable;
import com.agile.api.ITwoWayIterator;
import com.agile.api.IUser;
import com.agile.api.IWorkflow;
import com.agile.api.ItemConstants;
import com.agile.px.ActionResult;
import com.agile.px.ICustomAction;
import com.anselm.tools.record.Ini;
import com.anselm.tools.record.Log;

public class updateBOMbyBatch implements ICustomAction {

	@Override
	public ActionResult doAction(IAgileSession session, INode actionNode, IDataObject affectedObject) {

		// use log
		Ini ini = new Ini("C:\\Agile\\Config.ini");
		Log log = new Log();

		try {
			String maxinum = ini.getValue("updateBOMbyBatch", "maxinum");
			System.out.println(maxinum);

			log.logSeparatorBar();
			log.setTopic("AnselmTools_ini_log_Tutorial_");
			log.log("-------------------------程式開始-------------------------");

			/*****************************************************************************
			 * Step.1 讀取 Excel
			 * Step.2 抓取舊料號的 where use 父階料號
			 * Step.3 開立 eco
			 * Step.4 將 where use 中的物件加到 affected items
			 * Step.5 替換料件
			 * Step.6 進站
			 ******************************************************************************/

			// init
			DataFormatter df = new DataFormatter();
			ArrayList<String> oldlist = new ArrayList<String>(5);
			ArrayList<String> newlist = new ArrayList<String>(5);
			ArrayList<String> number = new ArrayList<String>(5);
			HashMap<String, String> map = new HashMap<String, String>();
			Collection<IChange> addeco = new ArrayList<IChange>();
			IChange change_eco = null;

			// 抓 ECR 表單
			IChange change = (IChange) session.getObject(ChangeConstants.CLASS_ECR, affectedObject.getName());
			ITable table_ecr = change.getTable(ChangeConstants.TABLE_RELATIONSHIPS);
			Iterator it_ecr = table_ecr.iterator();

			// downlaod excel
			ITable attTable = change.getAttachments();
			ITwoWayIterator attit = attTable.getTableIterator();
			while (attit.hasNext()) {
				byte[] attbuffer = new byte[1024];
				int attsize = 0;
				String attfile = "";

				IRow attrow = (IRow) attit.next();
				attfile = "C:\\Users\\Administrator\\Desktop\\data\\excel2.xlsx";
				log.log("檔名為" + attfile);
				InputStream in = ((IAttachmentFile) attrow).getFile();

				while((attsize=in.read(attbuffer)) != -1) {
					OutputStream output = new FileOutputStream(attfile, true);
					System.out.println(output);
					output.write(attbuffer, 0, attsize);
					System.out.println("儲存完成");
				}
			}

			// 讀取 Excel
			String file = "C:\\Users\\Administrator\\Desktop\\data\\excel2.xlsx";

			// Read Excel File into workbook
			FileInputStream inp = new FileInputStream(file);
			XSSFWorkbook wb = new XSSFWorkbook(inp);
			inp.close();

			// get wb sheet(0)
			XSSFSheet sheet = wb.getSheetAt(0);
			XSSFRow row = sheet.getRow(0);

			// 抓出舊料號
			for (int i = 1; i <= sheet.getLastRowNum(); i++) {
				XSSFRow row1 = sheet.getRow(i);
				for (int j = 0; j < row.getLastCellNum(); j++) {
					if (!df.formatCellValue(row1.getCell(j)).equals("end")) {
						if (j % 2 == 0) {
							newlist.add(df.formatCellValue(row1.getCell(j)));
						} else {
							oldlist.add(df.formatCellValue(row1.getCell(j)));
							map.put(df.formatCellValue(row1.getCell(j)), df.formatCellValue(row1.getCell(j - 1)));
						}
					} else {
						break;
					}
				}
			}
			log.log(map);

			// 抓出所有父階料號
			for (int k = 0; k < newlist.size(); k++) {

				IItem item = (IItem) session.getObject(ItemConstants.CLASS_PART, oldlist.get(k));
				ITable table_part = item.getTable(ItemConstants.TABLE_WHEREUSED);
				Iterator it_part = table_part.iterator();
				while (it_part.hasNext()) {
					IRow row1 = (IRow) it_part.next();
					if (!number.contains(row1.getValue(ItemConstants.ATT_WHERE_USED_ITEM_NUMBER))) {
						number.add((String) row1.getValue(ItemConstants.ATT_WHERE_USED_ITEM_NUMBER));
					}
				}
			}
			log.log(number);

			for (int l = 0; l < number.size(); l++) {
				if (l % Integer.parseInt(maxinum) == 0) {
					change_eco = newChange(session, "ECO");
					addWorkflow(change_eco);
					addeco.add(change_eco);
					log.log("開單");
				}
				IItem item = (IItem) session.getObject("HDD", number.get(l));
				ITable table_eco = change_eco.getTable(ChangeConstants.TABLE_AFFECTEDITEMS);
				session.disableAllWarnings();
				IRow affectedItemRow = table_eco.createRow(item);
				session.enableAllWarnings();

				Iterator it_eco = table_eco.iterator();

				while (it_eco.hasNext()) {
					IRow row1 = (IRow) it_eco.next();
					IItem item1 = (IItem) row1.getReferent();
					item1.setRevision(change_eco);


					ITable table = (ITable) item1.getTable(ItemConstants.TABLE_REDLINEBOM);
					Iterator<?> it = table.iterator();
					((IRedlinedTable) table).undoAllRedline();
					item1.refresh();
					change_eco.refresh();
					while (it.hasNext()) {
						IRow delRow = (IRow) it.next();
						String itemred = (String) delRow.getValue(ItemConstants.ATT_BOM_ITEM_NUMBER);
						System.out.println(itemred);
						boolean contains = map.containsKey(itemred);
						String bomItemNumber = delRow.getReferent().getName();
						if (contains) {
							delRow.setValue(ItemConstants.ATT_BOM_ITEM_NUMBER, map.get(bomItemNumber));
							log.log("替換完成");
						}
					}

					System.out.println(row1.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_ITEM_NUMBER));
					System.out.println("OLD_REV：" + row1.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV));
					System.out.println("NEW_REV：" + row1.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV));

					// 字串是否是數字
					boolean integerOrNot1 = ((String) row1.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV))
							.matches("[+-]?\\d*(\\.\\d+)?");

					// 是否為整數
					boolean integerOrNot2 = ((String) row1.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV))
							.matches("-?\\d+");

					if (row1.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV).equals("")) {
						if (row1.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV).equals("1") == false) {
							row1.setValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV, "1");
						}
					} else if (integerOrNot1 == true && integerOrNot2 == true) {
						int i = Integer.valueOf((String) row1.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV)) + 1;
						if (row1.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV)
								.equals(Integer.toString(i)) == false) {
							row1.setValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV, Integer.toString(i));
						}
					} else if (integerOrNot1 == true && integerOrNot2 == false) {
						double num2 = Double
								.parseDouble((String) row1.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV));
						double j = num2 + 0.1;
						j = Math.round(j * 100.0) / 100.0;
						if (row1.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV)
								.equals(Double.toString(j)) == false) {
							row1.setValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV, Double.toString(j));
						}
					} else if (integerOrNot1 == false && integerOrNot2 == false) {
						String first = ((String) row1.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV)).substring(0,
								1);
						String x = ((String) row1.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV))
								.replaceAll("[^\\d]", "");
						int k = Integer.parseInt(x);
						if (k < 10) {
							k = Integer.parseInt(x) + 1;
							if (row1.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV)
									.equals(first + "0" + Integer.toString(k)) == false) {
								row1.setValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV,
										first + "0" + Integer.toString(k));
							}
						} else if (k == 99) {
							char result = first.charAt(0);
							int asciiValue = result;
							int asciiValue2 = asciiValue;
							String second = Character.toString((char) (asciiValue2 + 1));
							if (row1.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV)
									.equals(second + "01") == false) {
								row1.setValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV, second + "01");
							}
						} else {
							k = Integer.parseInt(x) + 1;
							if (row1.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV)
									.equals(first + Integer.toString(k)) == false) {
								row1.setValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV, first + Integer.toString(k));
							}
						}
					}
				}
				log.log("-------------------------新舊料替換完成-------------------------");
			}


			// 加到 ECR Relationships
			ITable table_rel = change.getTable(ChangeConstants.TABLE_RELATIONSHIPS);
			table_rel.addAll(addeco);
			log.log("------------------- 加到 ECR Relationships -------------------");

			// eco進站
			IStatus nextStatus_eco = change_eco.getDefaultNextStatus();
			System.out.println("Next default status = " + nextStatus_eco.getName());
			session.disableAllWarnings();
			change_eco.changeStatus(nextStatus_eco, false, "", true, true, null, null, null, false);
			session.enableAllWarnings();
			log.log("----------------- 將全部的 BOM 替換單進站發行 -----------------");

			// ecr進站
			IStatus nextStatus_ecr = change.getDefaultNextStatus();
			System.out.println("Next default status = " + nextStatus_ecr.getName());
			session.disableAllWarnings();
			change.changeStatus(nextStatus_ecr, false, "", true, true, null, null, null, false);
			session.enableAllWarnings();
			log.log("----------------- ecr 進站發行 -----------------");

			removeAdmin(change, ini);

			log.log("-------------------------程式結束-------------------------");

		} catch (Exception e) {
			e.printStackTrace();
			log.logException(e);
		}
		return new ActionResult(ActionResult.STRING, "Hello World");
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

	// 解構 & 釋放
	public static void close(Log log, Ini ini) {
		try {
			log.close();
			log = null;
			ini = null;
		} catch (Exception e) {
		}
	}

	// remove admin (解除卡站)
	@SuppressWarnings("deprecation")
	private static void removeAdmin(IChange change, Ini ini) throws Exception {
		IUser admin = (IUser) change.getSession().getObject(IUser.OBJECT_TYPE, ini.getValue("AgileAP", "username"));
		IUser[] adminuser = new IUser[] { admin };
		IUser[] userlist = change.getApprovers(change.getStatus());
		for (int i = 0; i < userlist.length; i++) {
			if (userlist[i].getName().equals(admin.getName())) {
				change.removeApprovers(change.getStatus(), adminuser, null, "Removing admin from approvers.");
			}
		}
	}

	// 設定 workflow
	private static IWorkflow addWorkflow(IChange change) throws APIException {
		IWorkflow[] wfs = change.getWorkflows();
		IWorkflow workflow = null;
		for (int i = 0; i < wfs.length; i++) {
			if (wfs[i].getName().equals("ECO_BOM"))
				workflow = wfs[i];
		}
		change.setWorkflow(workflow);
		return workflow;
	}

}
