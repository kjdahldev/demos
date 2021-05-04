import { WebPartContext } from '@microsoft/sp-webpart-base';
import {IDemoItem} from './IDemoItem';

export class AppRepository {        
    private _context: WebPartContext;
    private _dummyItems: IDemoItem[];

    constructor(context: WebPartContext) {
        // Context to fetch from SharePoint. 
        // In demo we do dummy data.

        this._context = context;
        this._dummyItems = [];
        this.buildDummyData();
    }

    private buildDummyData () : void {
        for (let i = 0; i < 10; i++) {
            this._dummyItems.push({
                Id: i,
                Title: "Test" + i,
                Description: "Description" + i
            });
        }
    }

    public getAllItems () : IDemoItem[] {
        return this._dummyItems;
    }

    public getItemById (id: number) : IDemoItem {
        return this._dummyItems.filter(x => x.Id === id)[0];
    }
}