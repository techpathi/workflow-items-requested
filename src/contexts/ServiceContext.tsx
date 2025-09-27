import * as React from 'react';
import { SPFI } from '@pnp/sp';
import { GraphFI } from '@pnp/graph';

export interface IServiceContext {
  sp: SPFI;
  graph?: GraphFI;
}

export const ServiceContext = React.createContext<IServiceContext | undefined>(
  undefined
);

export const ServiceProvider: React.FC<
  IServiceContext & { children?: React.ReactNode }
> = ({ sp, graph, children }) => {
  // The provider simply passes down the instances we're given.
  return (
    <ServiceContext.Provider value={{ sp, graph }}>
      {children}
    </ServiceContext.Provider>
  );
};

export const useServices = (): IServiceContext => {
  const context = React.useContext(ServiceContext);
  if (!context) {
    throw new Error('useServices must be used within a ServiceProvider');
  }
  return context;
};
