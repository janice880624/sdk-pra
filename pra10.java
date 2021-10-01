package prac0914;

import java.io.FileOutputStream;

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

public class ActionPra10 implements ICustomAction{

	@Override
	public ActionResult doAction(IAgileSession session, INode arg1, IDataObject affectedObject) {
		// TODO Auto-generated method stub
		try {
			System.out.println("------------------------------------------------");
			System.out.println("程式開始");
			IItem item = (IItem) session.getObject(ItemConstants.CLASS_PART, affectedObject.getName());

			String data = item.getCell(1016).toString().split(" ")[0].replace("-", "/");
			String[] exceldata = {item.getCell(1001).toString(), item.getCell(1081).toString(), item.getCell(1271).toString(), item.getCell(1016).toString()};
			String[][] InputData = {{" ", "Part Information Sheet"},
									{"Part Number", item.getCell(1001).toString()},
									{"Part Type", item.getCell(1081).toString()},
									{"create User", item.getCell(1271).toString()},
									{"Rev Releasd Date", data}};

			// 大量資料用 SXSSFWorkbook
			String path = "C:\\Users\\Administrator\\Desktop\\data\\template.xlsx";

			// 設定幾筆之後，就先寫到硬碟的暫存檔
			SXSSFWorkbook wb = new SXSSFWorkbook(100);
			SXSSFSheet sheet = wb.createSheet();
			FileOutputStream fileOut = new FileOutputStream(path);
			Row row   = null;
			Cell cell = null;
			int c = 0;
			for(int i = 0; i < InputData.length; i++) {
				row = sheet.createRow(i);
	            for(c = 0; c < InputData[0].length; c++) {
	                System.out.print(InputData[i][c] + " ");
					cell = row.createCell(c);
			        cell.setCellValue(InputData[i][c]);
	            }
	            System.out.println();
	        }

			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
			wb.dispose();

			System.out.println("程式結束");
			return new ActionResult(ActionResult.STRING, "finish");
		}catch (Exception e) {
			e.printStackTrace();
			return new ActionResult(ActionResult.STRING, "error");
		}
	}

}
