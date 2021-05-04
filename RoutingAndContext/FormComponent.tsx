import * as React from 'react';
import {IDemoItem} from './IDemoItem';
import {appGlobalStateContext} from './GlobalState';
import { useParams } from 'react-router-dom';

interface IFormComponentProps {}
interface IFormComponentRouterProps {
    id: string;
}
const FormComponent: React.FC<IFormComponentProps> = () => {
    const {id} = useParams<IFormComponentRouterProps>();
    const {appRepository} = React.useContext(appGlobalStateContext);
    const [item, setItem] = React.useState<IDemoItem>(null);

    React.useEffect(() => {
        const getItem = async () => {
            const result = appRepository.getItemById(Number(id));
            setItem(result);
        };

        getItem();
    }, []);
    
    if (!item) return <div>Loading...</div>;

    return (
        <div>
            <p><b>Title: </b>{item.Title}</p>
            <p><b>Description: </b>{item.Description}</p>
        </div>
    );
};
export default FormComponent;