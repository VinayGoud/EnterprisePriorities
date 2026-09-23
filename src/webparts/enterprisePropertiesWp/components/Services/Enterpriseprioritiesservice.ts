import { spfi, SPFI, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IPriorityCard {
  Id: number;
  Title: string;
  Description: string;
  URL: string;
  DisplayOrder: number;
  Active: boolean;
  BackgroundColor: string;
  Icon: string;
  IconURL: string;
  PublishingStartDate?: string;
  PublishingEndDate?: string;
  ShowNewLabel: boolean;
  NewLabelStartDate?: string;
  NewLabelEndDate?: string;
  IsNew: boolean;
}

interface IRawPriorityListItem {
  Id: number;
  Title: string;
  Description: string;
  URL: { Url: string } | string | undefined;
  DisplayOrder: number;
  Active: boolean;
  BackgroundColor: string;
  Icon: string;
  IconURL: { Url: string } | string | undefined;
  PublishingStartDate?: string;
  PublishingEndDate?: string;
  ShowNewLabel: boolean;
  NewLabelStartDate?: string;
  NewLabelEndDate?: string;
}

const LIST_NAME = 'EnterprisePriorities';
const MAX_VISIBLE_CARDS = 3;

export class EnterprisePrioritiesService {
  private sp: SPFI;

  constructor(context: WebPartContext) {
    this.sp = spfi().using(SPFx(context));
  }

  public async getActivePriorityCards(): Promise<IPriorityCard[]> {
    try {
      const items: IRawPriorityListItem[] = await this.sp.web.lists
        .getByTitle(LIST_NAME)
        .items.select(
          'Id',
          'Title',
          'Description',
          'URL',
          'DisplayOrder',
          'Active',
          'BackgroundColor',
          'Icon',
          'IconURL',
          'PublishingStartDate',
          'PublishingEndDate',
          'ShowNewLabel',
          'NewLabelStartDate',
          'NewLabelEndDate'
        )
        .filter('Active eq 1')
        .orderBy('DisplayOrder', true)();

      const currentDate = new Date();

      return items
        .filter((item) => {
          const startDate = item.PublishingStartDate
            ? new Date(item.PublishingStartDate)
            : undefined;

          const endDate = item.PublishingEndDate
            ? new Date(item.PublishingEndDate)
            : undefined;

          const isStarted = !startDate || currentDate >= startDate;
          const isNotExpired = !endDate || currentDate <= endDate;

          return isStarted && isNotExpired;
        })
        .slice(0, MAX_VISIBLE_CARDS)
        .map((item) => {
          const url =
            typeof item.URL === 'object' && item.URL !== null
              ? item.URL.Url
              : item.URL || '';

          const iconUrl =
            typeof item.IconURL === 'object' && item.IconURL !== null
              ? item.IconURL.Url
              : item.IconURL || '';

          const newLabelStartDate = item.NewLabelStartDate
            ? new Date(item.NewLabelStartDate)
            : undefined;

          const newLabelEndDate = item.NewLabelEndDate
            ? new Date(item.NewLabelEndDate)
            : undefined;

          const isNew =
            item.ShowNewLabel &&
            (!newLabelStartDate || currentDate >= newLabelStartDate) &&
            (!newLabelEndDate || currentDate <= newLabelEndDate);

          return {
            Id: item.Id,
            Title: item.Title,
            Description: item.Description,
            URL: url,
            DisplayOrder: item.DisplayOrder,
            Active: item.Active,
            BackgroundColor: item.BackgroundColor,
            Icon: item.Icon,
            IconURL: iconUrl,
            PublishingStartDate: item.PublishingStartDate,
            PublishingEndDate: item.PublishingEndDate,
            ShowNewLabel: item.ShowNewLabel,
            NewLabelStartDate: item.NewLabelStartDate,
            NewLabelEndDate: item.NewLabelEndDate,
            IsNew: isNew
          };
        });
    } catch (error) {
      console.error('Error fetching Enterprise Priority cards:', error);
      return [];
    }
  }
}