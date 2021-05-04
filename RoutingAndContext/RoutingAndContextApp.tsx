import * as React from 'react';
import { HashRouter, Redirect, Route, Switch } from 'react-router-dom';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {appGlobalState, appGlobalStateContext} from './GlobalState';
import {AppRepository} from './AppRepository';
import FormComponent from './FormComponent';
import ListViewComponent from './ListViewComponent';

interface IRoutingAndContextAppProps {
    context: WebPartContext;
}
const RoutingAndContextApp: React.FC<IRoutingAndContextAppProps> = ({context}) => {
    const [loading, setLoading] = React.useState(true);

    React.useEffect(() => {
        appGlobalState.appRepository = new AppRepository(context);
        setLoading(false);
    }, []);

    if (loading) {
        return <div>Loading...</div>;
    }

    return (
        <div>
            <appGlobalStateContext.Provider value={appGlobalState}>
                <HashRouter>
                    <Switch>
                        <Route sensitive exact path="/">
                            <Redirect to={`/items`}></Redirect>
                        </Route>
                        <Route path="/items" sensitive exact component={ListViewComponent} />
                        <Route path="/items/:id" exact sensitive component={FormComponent} /> 
                    </Switch>
                </HashRouter>
            </appGlobalStateContext.Provider>   
        </div>
    );
};
export default RoutingAndContextApp;
