import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export class SharePointService {

  public async addDismissedItem(
    listTitle: string,
    announcementId: number,
    userId: number
  ): Promise<void> {

    await sp.web.lists
      .getByTitle(listTitle)
      .items.add({
        lookupAnnouncementIDId: announcementId,
        UserEmailId: userId,
        DismissDate: new Date()
      });
  }
}
