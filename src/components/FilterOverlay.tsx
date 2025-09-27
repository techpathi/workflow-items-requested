import * as React from 'react';
import {
  Callout,
  DirectionalHint,
  Checkbox,
  SearchBox,
  DatePicker,
  Stack,
  Text,
  IconButton,
  Separator,
  DefaultButton,
  PrimaryButton,
  IDatePickerStyles
} from '@fluentui/react';

export interface IFilterValue {
  key: string;
  text: string;
  count?: number;
}

export interface IDateRangeFilter {
  from?: Date;
  to?: Date;
}

export interface IFilterConfig {
  type: 'checkbox' | 'date' | 'text';
  values?: IFilterValue[];
  selectedValues?: string[];
  dateRange?: IDateRangeFilter;
  textValue?: string;
  showSearch?: boolean;
}

export interface IFilterOverlayProps {
  target: React.RefObject<HTMLElement>;
  isVisible: boolean;
  onDismiss: () => void;
  onApply: (filterConfig: IFilterConfig) => void;
  onClear: () => void;
  columnKey: string;
  columnName: string;
  filterConfig: IFilterConfig;
}

export interface IFilterOverlayState {
  searchText: string;
  tempSelectedValues: string[];
  tempDateRange: IDateRangeFilter;
  tempTextValue: string;
}

export class FilterOverlay extends React.Component<
  IFilterOverlayProps,
  IFilterOverlayState
> {
  constructor(props: IFilterOverlayProps) {
    super(props);

    this.state = {
      searchText: '',
      tempSelectedValues: [...(props.filterConfig.selectedValues || [])],
      tempDateRange: props.filterConfig.dateRange
        ? { ...props.filterConfig.dateRange }
        : {},
      tempTextValue: props.filterConfig.textValue || ''
    };
  }

  public componentDidUpdate(prevProps: IFilterOverlayProps): void {
    // Reset temp state when filter config changes
    if (prevProps.filterConfig !== this.props.filterConfig) {
      this.setState({
        tempSelectedValues: [...(this.props.filterConfig.selectedValues || [])],
        tempDateRange: this.props.filterConfig.dateRange
          ? { ...this.props.filterConfig.dateRange }
          : {},
        tempTextValue: this.props.filterConfig.textValue || ''
      });
    }
  }

  private _onSearchChange = (
    event?: React.ChangeEvent<HTMLInputElement>,
    newValue?: string
  ): void => {
    this.setState({ searchText: newValue || '' });
  };

  private _onCheckboxChange = (key: string, checked: boolean): void => {
    const { tempSelectedValues } = this.state;
    if (checked) {
      this.setState({
        tempSelectedValues: [...tempSelectedValues, key]
      });
    } else {
      this.setState({
        tempSelectedValues: tempSelectedValues.filter(v => v !== key)
      });
    }
  };

  private _onSelectAll = (): void => {
    const { filterConfig } = this.props;

    if (filterConfig.values) {
      const filteredValues = this._getFilteredValues();
      const allKeys = filteredValues.map(v => v.key);
      this.setState({ tempSelectedValues: allKeys });
    }
  };

  private _onSelectNone = (): void => {
    this.setState({ tempSelectedValues: [] });
  };

  private _onDateFromChange = (date: Date | null | undefined): void => {
    this.setState({
      tempDateRange: {
        ...this.state.tempDateRange,
        from: date || undefined
      }
    });
  };

  private _onDateToChange = (date: Date | null | undefined): void => {
    this.setState({
      tempDateRange: {
        ...this.state.tempDateRange,
        to: date || undefined
      }
    });
  };

  private _onTextChange = (
    event?: React.ChangeEvent<HTMLInputElement>,
    newValue?: string
  ): void => {
    this.setState({ tempTextValue: newValue || '' });
  };

  private _onApply = (): void => {
    const { filterConfig, onApply } = this.props;
    const { tempSelectedValues, tempDateRange, tempTextValue } = this.state;

    const newConfig: IFilterConfig = {
      ...filterConfig,
      selectedValues: tempSelectedValues,
      dateRange: tempDateRange,
      textValue: tempTextValue
    };

    onApply(newConfig);
  };

  private _onClear = (): void => {
    this.setState({
      tempSelectedValues: [],
      tempDateRange: {},
      tempTextValue: ''
    });
    this.props.onClear();
  };

  private _getFilteredValues = (): IFilterValue[] => {
    const { filterConfig } = this.props;
    const { searchText } = this.state;

    if (!filterConfig.values) return [];

    if (!searchText) return filterConfig.values;

    return filterConfig.values.filter(
      value => value.text.toLowerCase().indexOf(searchText.toLowerCase()) !== -1
    );
  };

  private _renderCheckboxFilter = (): JSX.Element => {
    const { filterConfig } = this.props;
    const { tempSelectedValues, searchText } = this.state;
    const filteredValues = this._getFilteredValues();

    const allSelected =
      filteredValues.length > 0 &&
      filteredValues.every(v => tempSelectedValues.indexOf(v.key) !== -1);

    return (
      <Stack tokens={{ childrenGap: 8 }}>
        {filterConfig.showSearch && (
          <SearchBox
            placeholder={`Search ${this.props.columnName.toLowerCase()}...`}
            value={searchText}
            onChange={this._onSearchChange}
            styles={{ root: { width: '100%' } }}
          />
        )}

        <Stack horizontal horizontalAlign='space-between'>
          <DefaultButton
            text='Select All'
            onClick={this._onSelectAll}
            disabled={allSelected}
            styles={{ root: { minWidth: 80 } }}
          />
          <DefaultButton
            text='None'
            onClick={this._onSelectNone}
            disabled={tempSelectedValues.length === 0}
            styles={{ root: { minWidth: 60 } }}
          />
        </Stack>

        <Separator />

        <div style={{ maxHeight: 200, overflowY: 'auto' }}>
          <Stack tokens={{ childrenGap: 4 }}>
            {filteredValues.map(value => (
              <Checkbox
                key={value.key}
                label={
                  value.text +
                  (value.count !== undefined ? ` (${value.count})` : '')
                }
                checked={tempSelectedValues.indexOf(value.key) !== -1}
                onChange={(ev, checked) =>
                  this._onCheckboxChange(value.key, !!checked)
                }
                styles={{
                  root: { width: '100%' },
                  label: { width: '100%' }
                }}
              />
            ))}
          </Stack>
        </div>
      </Stack>
    );
  };

  private _renderDateFilter = (): JSX.Element => {
    const { tempDateRange } = this.state;

    const datePickerStyles: Partial<IDatePickerStyles> = {
      root: { width: '100%' }
    };

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <Text variant='medium' styles={{ root: { fontWeight: 600 } }}>
          Date Range
        </Text>

        <Stack tokens={{ childrenGap: 8 }}>
          <Text variant='small'>From:</Text>
          <DatePicker
            value={tempDateRange.from}
            onSelectDate={this._onDateFromChange}
            placeholder='Select start date...'
            ariaLabel='From date'
            styles={datePickerStyles}
          />
        </Stack>

        <Stack tokens={{ childrenGap: 8 }}>
          <Text variant='small'>To:</Text>
          <DatePicker
            value={tempDateRange.to}
            onSelectDate={this._onDateToChange}
            placeholder='Select end date...'
            ariaLabel='To date'
            styles={datePickerStyles}
          />
        </Stack>

        <Stack horizontal horizontalAlign='start' tokens={{ childrenGap: 8 }}>
          <DefaultButton
            text='Last 7 days'
            onClick={() =>
              this.setState({
                tempDateRange: {
                  from: new Date(Date.now() - 7 * 24 * 60 * 60 * 1000),
                  to: new Date()
                }
              })
            }
            styles={{ root: { fontSize: '12px', padding: '4px 8px' } }}
          />
          <DefaultButton
            text='This month'
            onClick={() => {
              const now = new Date();
              const firstDay = new Date(now.getFullYear(), now.getMonth(), 1);
              this.setState({
                tempDateRange: {
                  from: firstDay,
                  to: now
                }
              });
            }}
            styles={{ root: { fontSize: '12px', padding: '4px 8px' } }}
          />
        </Stack>
      </Stack>
    );
  };

  private _renderTextFilter = (): JSX.Element => {
    const { tempTextValue } = this.state;

    return (
      <Stack tokens={{ childrenGap: 12 }}>
        <Text variant='medium' styles={{ root: { fontWeight: 600 } }}>
          Search Text
        </Text>

        <SearchBox
          placeholder={`Search in ${this.props.columnName.toLowerCase()}...`}
          value={tempTextValue}
          onChange={this._onTextChange}
          styles={{ root: { width: '100%' } }}
        />
      </Stack>
    );
  };

  public render(): JSX.Element {
    const { target, isVisible, onDismiss, filterConfig, columnName } =
      this.props;
    const { tempSelectedValues, tempDateRange, tempTextValue } = this.state;

    // Determine if there are changes to apply
    const hasChanges = (() => {
      switch (filterConfig.type) {
        case 'checkbox':
          return (
            JSON.stringify(tempSelectedValues.sort()) !==
            JSON.stringify((filterConfig.selectedValues || []).sort())
          );
        case 'date':
          return (
            JSON.stringify(tempDateRange) !==
            JSON.stringify(filterConfig.dateRange || {})
          );
        case 'text':
          return tempTextValue !== (filterConfig.textValue || '');
        default:
          return false;
      }
    })();

    if (!isVisible) return <></>;

    return (
      <Callout
        target={target.current}
        onDismiss={onDismiss}
        directionalHint={DirectionalHint.bottomLeftEdge}
        isBeakVisible={true}
        setInitialFocus={true}
        styles={{
          calloutMain: {
            padding: 16,
            minWidth: 250,
            maxWidth: 300
          }
        }}
      >
        <Stack tokens={{ childrenGap: 16 }}>
          {/* Header */}
          <Stack
            horizontal
            horizontalAlign='space-between'
            verticalAlign='center'
          >
            <Text variant='medium' styles={{ root: { fontWeight: 600 } }}>
              Filter: {columnName}
            </Text>
            <IconButton
              iconProps={{ iconName: 'Clear' }}
              title='Close'
              onClick={onDismiss}
              styles={{ root: { width: 24, height: 24 } }}
            />
          </Stack>

          {/* Filter Content */}
          {filterConfig.type === 'checkbox' && this._renderCheckboxFilter()}
          {filterConfig.type === 'date' && this._renderDateFilter()}
          {filterConfig.type === 'text' && this._renderTextFilter()}

          {/* Action Buttons */}
          <Separator />
          <Stack
            horizontal
            horizontalAlign='space-between'
            tokens={{ childrenGap: 8 }}
          >
            <DefaultButton
              text='Clear Filter'
              onClick={this._onClear}
              iconProps={{ iconName: 'ClearFilter' }}
            />
            <PrimaryButton
              text='Apply'
              onClick={this._onApply}
              disabled={!hasChanges}
              iconProps={{ iconName: 'Filter' }}
            />
          </Stack>
        </Stack>
      </Callout>
    );
  }
}
