import * as React from 'react';
import {IDemoItem} from './IDemoItem';
import {appGlobalStateContext} from './GlobalState';
import { Link } from 'react-router-dom';

interface IListViewComponentProps {}

const ListViewComponent: React.FC<IListViewComponentProps> = () => {
    const {appRepository} = React.useContext(appGlobalStateContext);
    const [items, setItems] = React.useState<IDemoItem[]>([]);

    React.useEffect(() => {
        const getItems = async () => {            
            const allItems = appRepository.getAllItems();            
            setItems(allItems);
        };

        getItems();
    }, []);

    if (items.length === 0) {
        return <div>Loading...</div>;
    }

    return (
        <div>
            {items.map(item => {
                return (
                    <p>
                        <Link to={`/items/${item.Id}`}>{item.Title}</Link>
                    </p>
                );
            })}
        </div>
    );
};
export default ListViewComponent;