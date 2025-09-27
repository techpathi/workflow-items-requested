import * as React from 'react';

export interface IListInfo {
  id: string;
  title: string;
  webUrl: string;
}

export interface IAppConfigContext {
  selectedLists: IListInfo[];
  currentSiteUrl: string;
  refreshData: () => Promise<void>;
  isLoading: boolean;
  error: string | null;
}

export const AppConfigContext = React.createContext<
  IAppConfigContext | undefined
>(undefined);

export interface IAppConfigProviderProps {
  selectedListTitles: string[];
  currentSiteUrl: string;
  onRefreshData?: () => Promise<void>;
  children?: React.ReactNode;
}

export const AppConfigProvider: React.FC<IAppConfigProviderProps> = ({
  selectedListTitles,
  currentSiteUrl,
  onRefreshData,
  children
}) => {
  const [selectedLists, setSelectedLists] = React.useState<IListInfo[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | null>(null);

  // Convert list titles to list info objects
  React.useEffect(() => {
    const convertTitlesToListInfo = () => {
      const listInfos: IListInfo[] = selectedListTitles.map(title => ({
        id: title, // For now, using title as ID - this should be enhanced to get actual list IDs
        title: title,
        webUrl: currentSiteUrl
      }));
      setSelectedLists(listInfos);
    };

    convertTitlesToListInfo();
  }, [selectedListTitles, currentSiteUrl]);

  const refreshData = React.useCallback(async (): Promise<void> => {
    setIsLoading(true);
    setError(null);
    try {
      if (onRefreshData) {
        await onRefreshData();
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Unknown error occurred');
    } finally {
      setIsLoading(false);
    }
  }, [onRefreshData]);

  const contextValue: IAppConfigContext = {
    selectedLists,
    currentSiteUrl,
    refreshData,
    isLoading,
    error
  };

  return (
    <AppConfigContext.Provider value={contextValue}>
      {children}
    </AppConfigContext.Provider>
  );
};

export const useAppConfig = (): IAppConfigContext => {
  const context = React.useContext(AppConfigContext);
  if (!context) {
    throw new Error('useAppConfig must be used within an AppConfigProvider');
  }
  return context;
};
