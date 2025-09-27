import { IWorkflowItemsRequestedItem } from '../models/IWorkflowItemsRequestedItem';
import {
  IFilterConfig as IOverlayFilterConfig,
  IFilterValue
} from '../components/FilterOverlay';

export interface IColumnFilter {
  [columnKey: string]: IOverlayFilterConfig;
}

export class FilterUtils {
  public static extractColumnValues(
    items: IWorkflowItemsRequestedItem[],
    columnKey: string
  ): IFilterValue[] {
    const valueMap = new Map<string, number>();

    items.forEach(item => {
      const value = FilterUtils.getColumnValue(item, columnKey);
      if (value !== null && value !== undefined && value !== '') {
        const stringValue = String(value);
        valueMap.set(stringValue, (valueMap.get(stringValue) || 0) + 1);
      }
    });

    const values: IFilterValue[] = [];
    valueMap.forEach((count, key) => {
      values.push({ key, text: key, count });
    });

    return values.sort((a, b) => a.text.localeCompare(b.text));
  }

  private static getColumnValue(
    item: IWorkflowItemsRequestedItem,
    columnKey: string
  ): any {
    switch (columnKey) {
      case 'title':
        return item.title;
      case 'workflowStatus':
        return item.workflowStatus;
      case 'creditManager':
        return item.creditManager;
      case 'dsr':
        return item.dsr;
      case 'customerService':
        return item.customerService;
      case 'currentAssignedRole':
        return item.currentAssignedRole;
      case 'sourceListTitle':
        return item.sourceListTitle;
      case 'priority':
        return item.priority;
      case 'requestor':
        return item.requestor;
      case 'createdDate':
        return item.createdDate;
      case 'modifiedDate':
        return item.modifiedDate;
      case 'dueDate':
        return item.dueDate;
      default:
        return null;
    }
  }

  public static applyFilters(
    items: IWorkflowItemsRequestedItem[],
    filters: IColumnFilter
  ): IWorkflowItemsRequestedItem[] {
    let result = [...items];
    for (const columnKey of Object.keys(filters)) {
      const filter = filters[columnKey];
      if (!filter) continue;
      if (!FilterUtils.isFilterActive(filter)) continue;
      result = result.filter(item =>
        FilterUtils.itemMatchesFilter(item, columnKey, filter)
      );
    }
    return result;
  }

  public static isFilterActive(filter: IOverlayFilterConfig): boolean {
    switch (filter.type) {
      case 'checkbox':
        return (filter.selectedValues || []).length > 0;
      case 'date':
        return !!(
          filter.dateRange &&
          (filter.dateRange.from || filter.dateRange.to)
        );
      case 'text':
        return !!(filter.textValue && filter.textValue.trim().length > 0);
      default:
        return false;
    }
  }

  private static itemMatchesFilter(
    item: IWorkflowItemsRequestedItem,
    columnKey: string,
    filter: IOverlayFilterConfig
  ): boolean {
    const value = FilterUtils.getColumnValue(item, columnKey);
    switch (filter.type) {
      case 'checkbox': {
        const vals = filter.selectedValues || [];
        if (vals.length === 0) return true;
        const strVal = String(value ?? '');
        return vals.indexOf(strVal) !== -1;
      }
      case 'date': {
        const dateVal = value instanceof Date ? value : undefined;
        if (!filter.dateRange) return true;
        const { from, to } = filter.dateRange;
        if (!dateVal) return false;
        if (from && dateVal < from) return false;
        if (to && dateVal > to) return false;
        return true;
      }
      case 'text': {
        const search = (filter.textValue || '').toLowerCase();
        if (!search) return true;
        return (
          String(value ?? '')
            .toLowerCase()
            .indexOf(search) !== -1
        );
      }
      default:
        return true;
    }
  }

  public static getColumnFilterConfig(
    columnKey: string,
    items: IWorkflowItemsRequestedItem[]
  ): IOverlayFilterConfig {
    switch (columnKey) {
      case 'title': // Treat title like a value list so we can show selectable matches
        return {
          type: 'checkbox',
          values: FilterUtils.extractColumnValues(items, columnKey),
          selectedValues: [],
          showSearch: true
        };
      case 'requestor':
        return { type: 'text', textValue: '', showSearch: true };
      case 'workflowStatus':
      case 'creditManager':
      case 'dsr':
      case 'customerService':
      case 'currentAssignedRole':
      case 'sourceListTitle':
        return {
          type: 'checkbox',
          values: FilterUtils.extractColumnValues(items, columnKey),
          selectedValues: [],
          showSearch: true
        };
      case 'createdDate':
      case 'dueDate':
        return { type: 'date', dateRange: {} };
      default:
        return { type: 'text', textValue: '', showSearch: true };
    }
  }

  public static initializeFilters(
    items: IWorkflowItemsRequestedItem[],
    columnKeys: string[]
  ): IColumnFilter {
    const filters: IColumnFilter = {};
    for (const key of columnKeys) {
      filters[key] = FilterUtils.getColumnFilterConfig(key, items);
    }
    return filters;
  }

  public static updateColumnFilter(
    filters: IColumnFilter,
    columnKey: string,
    newConfig: Partial<IOverlayFilterConfig>
  ): IColumnFilter {
    return {
      ...filters,
      [columnKey]: {
        ...(filters[columnKey] || { type: 'text' }),
        ...newConfig
      }
    };
  }

  public static clearAllFilters(filters: IColumnFilter): IColumnFilter {
    const cleared: IColumnFilter = {};
    for (const key of Object.keys(filters)) {
      const f = filters[key];
      switch (f.type) {
        case 'checkbox':
          cleared[key] = { ...f, selectedValues: [] };
          break;
        case 'date':
          cleared[key] = { ...f, dateRange: {} } as any;
          break;
        case 'text':
          cleared[key] = { ...f, textValue: '' };
          break;
        default:
          cleared[key] = f;
      }
    }
    return cleared;
  }

  public static getActiveFilterCount(filters: IColumnFilter): number {
    return Object.keys(filters)
      .map(k => filters[k])
      .filter(FilterUtils.isFilterActive).length;
  }

  public static getFilterSummary(filters: IColumnFilter): string {
    const n = FilterUtils.getActiveFilterCount(filters);
    if (n === 0) return '';
    return n === 1 ? '1 filter active' : `${n} filters active`;
  }
}
