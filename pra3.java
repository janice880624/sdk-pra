package prac0914;

import java.io.*;

import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.agile.api.APIException;
import com.agile.api.AgileSessionFactory;
import com.agile.api.ChangeConstants;
import com.agile.api.CommonConstants;
import com.agile.api.IAgileList;
import com.agile.api.IAgileSession;
import com.agile.api.ICell;
import com.agile.api.IChange;
import com.agile.api.IDataObject;
import com.agile.api.INode;
import com.agile.api.IProgram;
import com.agile.api.IUser;
import com.agile.px.ActionResult;
import com.agile.px.ICustomAction;


public class EventWorkflow implements ICustomAction{
	@Override
	public ActionResult doAction(IAgileSession session, INode actionNode, IDataObject affectedObject) {
		try {

			System.out.println("------------------------------------------------");
			System.out.println("程式開始");

			// To create a Change by class ID
			IChange change = (IChange)session.getObject(ChangeConstants.CLASS_ECO, affectedObject.getName());
			System.out.println("當前名稱" + affectedObject.getName());


			ICell cell = change.getCell(CommonConstants.ATT_PAGE_THREE_LIST01);
			//IAgileList cl = (IAgileList)cell.getValue();
			IProgram program = (IProgram) cell.getReferent();

			String file = "C:\\Users\\janice\\Desktop\\data\\excel1.xlsx";

			// test value
			String projectName = "QX-40";
			DataFormatter df = new DataFormatter();
			Collection<IUser> userCollection = new ArrayList<IUser>();


			// Read Excel File into workbook
      FileInputStream inp = new FileInputStream(file);
      XSSFWorkbook wb = new XSSFWorkbook(inp);
      inp.close();

      // get wb sheet(0)
      XSSFSheet sheet = wb.getSheetAt(0);

      // get total num of row
      int rowLength = sheet.getLastRowNum();
      String sheetName = sheet.getSheetName();
      System.out.println(sheetName);

      // get wb row
      XSSFRow row = sheet.getRow(0);

      // total num of cols(cell)
      int cellLength = row.getLastCellNum();
      System.out.println("cellLength: " + cellLength);
      System.out.println("rowLength: " + rowLength);

      // get wb cols(cell
      XSSFCell cell2 = row.getCell(0);
      System.out.println("cell: " + cell2);

      for (int i = 0; i <= rowLength; i++) {
          XSSFRow row1 = sheet.getRow(i);
          if (df.formatCellValue(row1.getCell(0)).equals(program.getName())) {
          	System.out.println("ya~~~~");
          	for (int j = 1; j <cellLength; j++) {
          		XSSFCell cell1 = row1.getCell(j);
          		System.out.print(df.formatCellValue(row1.getCell(j)));

          		IUser item = (IUser) session.getObject(IUser.OBJECT_TYPE, df.formatCellValue(row1.getCell(j)));
          		userCollection.add(item);
          	}
          }
      }

      System.out.println(userCollection);
      change.addReviewers(change.getStatus(), userCollection, null, null, false, "Add Approvers");
			System.out.println("程式結束");

			return new ActionResult(ActionResult.STRING, "Hello World");
		} catch (Exception e) {
			e.printStackTrace();
			return new ActionResult(ActionResult.EXCEPTION, e);
		}
	}
}
