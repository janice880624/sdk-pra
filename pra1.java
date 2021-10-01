package prac0914;

import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;

import com.agile.api.APIException;
import com.agile.api.AgileSessionFactory;
import com.agile.api.ChangeConstants;
import com.agile.api.IAgileSession;
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
			// 利用 API Name 找到  Cover Page 上的人員
//			IChange change = (IChange)session.getObject(IChange.OBJECT_TYPE, affectedObject.getName());
			IChange change = (IChange)session.getObject(ChangeConstants.CLASS_ECO, affectedObject.getName());
			System.out.println(change.getCell("PageTwo.list11"));

			// 取得站點名稱
			String status = change.getStatus().getName();
			System.out.println(status);

			//  增加 Approvers
			if (status.equals("課級主管")) {
				System.out.println("進到課級主管");
				IUser people1_name = (IUser)change.getCell("PageTwo.list11").getReferent();

				Collection userCollection = new ArrayList();
				userCollection.add(people1_name);

				change.addReviewers(change.getDefaultNextStatus(), userCollection, null, null, false, "Add Approvers");
				System.out.println("-----------");
			}
			else if (status.equals("部級主管")) {
				System.out.println("進到部級主管");
				IUser people2_name = (IUser)change.getCell("PageTwo.list12").getReferent();

				Collection userCollection = new ArrayList();
				userCollection.add(people2_name);
				System.out.println(people2_name);

				change.addReviewers(change.getDefaultNextStatus(), userCollection, null, null, false, "Add Approvers");
				System.out.println("-----------");
			}
			else if (status.equals("處級主管")) {
				System.out.println("進到處級主管");
				IUser people3_name = (IUser)change.getCell("PageTwo.list13").getReferent();

				Collection userCollection = new ArrayList();
				userCollection.add(people3_name);
				System.out.println(people3_name);

//				change.addReviewers(change.getDefaultNextStatus(), userCollection, null, null, false, "Add Approvers");
				System.out.println("-----------");
			}
//			System.out.println("-----------");


		} catch (APIException e) {
			e.printStackTrace();
		}
		return new ActionResult(ActionResult.STRING, "Hello World");
	}
}
