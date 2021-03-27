import * as React from 'react';
import { graph } from "@pnp/graph/presets/all";
import { MSGraphClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { resultItem } from 'office-ui-fabric-react/lib/components/ExtendedPicker/PeoplePicker/ExtendedPeoplePicker.scss';

export interface IDemoPostToTeamChannelProps {
    context: WebPartContext;    
    unifiedGroupId: string;
}

export const DemoPostToTeamChannel: React.FC<IDemoPostToTeamChannelProps> = ({context, unifiedGroupId}) => {
    const cellStyle = `padding:10px;border:1px solid #ccc;`;
    const msg = `
        <table>
            <tr>
                <th style="${cellStyle}">Heading 1</th>
                <th style="${cellStyle}">Heading 2</th>
            </tr>
            <tr>
                <td style="${cellStyle}">Value 1</td>
                <td style="${cellStyle}">Value 2 with <b>formatting</b></td>
            </tr>
        </table>
    `;

    const post = async () => {
        const demoRepository = new DemoGraphRepository(context, unifiedGroupId);
        const result = await demoRepository.postToChannel(msg);

        console.log(result);
    }

    return <button onClick={post}>Post Message</button>;
};

class DemoGraphRepository {    
    private _unifiedGroupId: string;
    private _repository: DemoGraphService;

    constructor(context: WebPartContext, unifiedGroupId: string) {
        this._unifiedGroupId = unifiedGroupId;                
        this._repository = new DemoGraphService(context);
    }

    public async postToChannel (msg: string, channelName: string = null): Promise<any> {
        const channels = await this._repository.getChannelsForTeam(this._unifiedGroupId);

        if (channels && channels.value && channels.value.length > 0) {
            let channel = channels.value.filter(x => x.displayName === channelName || "General")[0];
            
            if (!channel) {
                channel = channels.value[0];
            }

            const chatMsg = {
                body: {
                    contentType: "html",
                    content: msg
                }
            };

            const result = await this._repository.postToChannel(this._unifiedGroupId, chatMsg, channel.id);
            return result;
        }
    }
}

class DemoGraphService {
    private _graph;
    private _client: MSGraphClient;    

    constructor(context: WebPartContext) {
        graph.setup({
            spfxContext: context
        });

        this._graph = context.msGraphClientFactory;
        this._client = null;
    }

    private getGraphClient = async () => {
        return new Promise<any>((resolve, reject) => {
            if (this._client) {                
                resolve(this._client);
            }
            else {
                this._graph.getClient().then((client: MSGraphClient) => {
                    this._client = client;
                    resolve(this._client);
                }).catch((error: any) => {
                    console.log(error);
                    reject(null);
                });
            }
        });
    }

    public getChannelsForTeam = async (groupId: string) : Promise<any> => {
        const client = await this.getGraphClient();
        const channels = await client.api(`teams/${groupId}/channels/`).get();        
        return channels;
    }

    public postToChannel = async (groupId: string, chatMsg: Object, channelId: string) : Promise<any> => {
        const client = await this.getGraphClient();
        const result = await client.api(`teams/${groupId}/channels/${channelId}/messages`).post(chatMsg)
        return result;
    }
}
