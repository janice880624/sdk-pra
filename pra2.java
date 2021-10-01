package prac0914;

import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;

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


			ICell cell = change.getCell(CommonConstants.ATT_PAGE_THREE_MULTILIST01);
			IAgileList cl = (IAgileList)cell.getValue();
			System.out.println(cl);

			Collection userCollection = new ArrayList();

			IAgileList [] selected = cl.getSelection();

			for (int i=0; i<selected.length; i++){
				IUser people = (IUser)selected[i].getValue();
				userCollection.add(people);
			}
			System.out.println(userCollection);

			change.addReviewers(change.getStatus(), userCollection, null, null, false, "Add Approvers");
			System.out.println("程式結束");
			return new ActionResult(ActionResult.STRING, "Hello World");
		} catch (APIException e) {
			e.printStackTrace();
			return new ActionResult(ActionResult.EXCEPTION, e);
		}

	}
}
