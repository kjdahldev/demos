import * as React from 'react';
import {AppRepository} from './AppRepository';

interface IAppGlobalState {    
    appRepository: AppRepository;
}

export const appGlobalState : IAppGlobalState = {    
    appRepository: null
};

export const appGlobalStateContext = React.createContext(appGlobalState);