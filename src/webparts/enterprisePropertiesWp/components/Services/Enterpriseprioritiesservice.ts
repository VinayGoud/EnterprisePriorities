import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IPriorityCard {
  Id: number;
  Title: string;
  Description: string;
  Link: string;
  Icon: string; // URL to an image, uploaded by the editor to a SharePoint document library
  Active: boolean;
  DisplayOrder: number;
  StartDate: string | undefined;
  EndDate: string | undefined;
}

interface IRawPriorityListItem {
  Id: number;
  Title: string;
  Description: string;
  Link: { Url: string } | string | undefined;
  IconUrl: { Url: string } | string | undefined;
  Active: boolean;
  DisplayOrder: number;
  StartDate: string | undefined;
  EndDate: string | undefined;
}

const LIST_NAME = "EnterprisePriorities";
const MAX_VISIBLE_CARDS = 3; // confirmed: "three visible cards unless changed"

export class EnterprisePrioritiesService {
  private sp: SPFI;

  constructor(context: WebPartContext) {
    this.sp = spfi().using(SPFx(context));
    // NOTE: if your project already has a shared spfi() setup (like FeedbackService.ts),
    // prefer reusing that pattern instead of re-initializing here for consistency.
  }

  /**
   * Returns up to MAX_VISIBLE_CARDS active priority cards, sorted by DisplayOrder.
   * Icon is read from the IconUrl column — editors upload the image to a
   * SharePoint document library first, then paste that file's URL into this column.
   *
   * NOTE: "Active eq 1" filter syntax is unverified for this environment —
   * same open question flagged on MyToolsService.ts. Test against real data;
   * may need to be "Active eq true" depending on PnPjs/REST behavior here.
   */
  public async getActivePriorityCards(): Promise<IPriorityCard[]> {
    try {
      const items: IRawPriorityListItem[] = await this.sp.web.lists
        .getByTitle(LIST_NAME)
        .items.select(
          "Id",
          "Title",
          "Description",
          "Link",
          "IconUrl",
          "Active",
          "DisplayOrder",
          "StartDate",
          "EndDate"
        )
        .filter("Active eq 1")
        .orderBy("DisplayOrder", true)
        .top(MAX_VISIBLE_CARDS)();

      return items.map((item) => {
        const linkValue =
          typeof item.Link === "object" && item.Link !== null
            ? item.Link.Url
            : (item.Link as string) ?? "";

        const iconValue =
          typeof item.IconUrl === "object" && item.IconUrl !== null
            ? item.IconUrl.Url
            : (item.IconUrl as string) ?? "";

        return {
          Id: item.Id,
          Title: item.Title,
          Description: item.Description,
          Link: linkValue,
          Icon: iconValue,
          Active: item.Active,
          DisplayOrder: item.DisplayOrder,
          StartDate: item.StartDate,
          EndDate: item.EndDate,
        };
      });
    } catch (error) {
      console.error("Error fetching Enterprise Priority cards:", error);
      // Per BRD error/empty states: "no active cards" should render a defined
      // empty state in the UI, not throw — returning [] lets the component
      // handle that state.
      return [];
    }
  }
}