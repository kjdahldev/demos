import * as React from 'react';
import "@pnp/graph/planner";
import "@pnp/graph/calendars";
import { graph, ITaskAddResult } from "@pnp/graph/presets/all";
import { MSGraphClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDemoComponentProps {
    context: WebPartContext;    
    unifiedGroupId: string;
}

export const DemoComponent: React.FC<IDemoComponentProps> = ({context, unifiedGroupId}) => {
    const createTask = async () => {
        const demo = new Demo(context, unifiedGroupId);
        const result = await demo.createTask(
            "admin@tenant.onmicrosoft.com", 
            "You need to fix this!", 
            "Task body description"
        );
        console.log(result);
    }

    return <button onClick={createTask}>Create Task</button>;
}


class Demo {
    private _firstPlannerId: string;
    private _unifiedGroupId: string;
    private _repository: DemoGraphRepository;

    constructor(context: WebPartContext, unifiedGroupId: string) {
        this._unifiedGroupId = unifiedGroupId;        
        this._firstPlannerId = null;
        this._repository = new DemoGraphRepository(context);
    }

    public async createTask (upn: string, taskName: string, taskDescription: string, 
        duedate: string = null, plannerId: string = null): Promise<any> {

        if (!plannerId) {
            plannerId = await this.getFirstPlannerPlanInGroup();
        }

        if (!plannerId) {            
            return;
        }

        const userGuid = await this._repository.getUser(upn, "id");
        const result = await this._repository.createPlannerTask(plannerId, taskName, userGuid, taskDescription, duedate);

        return result;
    }

    private async getFirstPlannerPlanInGroup () : Promise<string> {
        return new Promise((resolve, reject) => {
            if (this._firstPlannerId) {
                resolve(this._firstPlannerId);
            }
            else {
                this._repository.getPlannerPlansForGroup(this._unifiedGroupId)
                .then(result => {
                    if (result && result.value && result.value.length > 0) {
                        this._firstPlannerId = result.value[0].id;
                        resolve(this._firstPlannerId);
                    }
                })
                .catch((error: any) => {
                    console.log(error);
                    reject(error);
                });
            }
        });
    }
}

class DemoGraphRepository {
    private _graph;
    private _client: MSGraphClient;    

    constructor(context: WebPartContext) {
        graph.setup({
            spfxContext: context
        });

        this._graph = context.msGraphClientFactory;
        this._client = null;
    }

    public getGraphClient = async () => {
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

    public getPlannerPlansForGroup = async (groupId: string) => {
        const client = await this.getGraphClient();
        const result = await client.api(`groups/${groupId}/planner/plans`).get();

        return result;
    }

    public getUser = async (upn: string, property?: string) : Promise<any> => {
        const user = await graph.users.getById(upn).get();
        return property ? user[property] : user;
    }

    public createPlannerTask = async (planId: string, title: string, userId: string, description?: string, duedate?: string) : Promise<ITaskAddResult> => {
        let assignments = {};
        assignments["assignments"] = {
            [userId]: {
                "@odata.type":"#microsoft.graph.plannerAssignment",
                "orderHint":" !"
            }
        };

        const newTask = await graph.planner.tasks.add(planId, title, assignments);
        
        if (duedate) {
            await graph.planner.tasks.getById(newTask.data.id).update({                
                dueDateTime: duedate
            }, newTask.data["@odata.etag"]);
        }
        
        if (description) {            
            const details = await this.getTaskDetails(newTask.data.id, 5);
            if (details) {                
                const etag = details["@odata.etag"];                   

                await graph.planner.tasks.getById(newTask.data.id).details.update({
                    description: description
                }, etag);            
            }
        }

        return newTask;
    }

    public getTaskDetails =  async (taskId: string, attempts: number): Promise<any> => {
        return new Promise((resolve, reject) => {
            if (attempts == 0) {
                reject();
            }

            graph.planner.tasks.getById(taskId).details.get().then(result => {                
                resolve(result);
            }).catch(error => {
                console.log(error);
                this.getTaskDetails(taskId, attempts - 1)
                    .then(resolve)
                    .catch(reject);
            });
        });
    }
}
