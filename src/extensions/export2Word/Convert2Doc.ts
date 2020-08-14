import { SPListOperations, SPCore, SPCommonOperations, ISPBaseResponse, SPFieldOperations, IField, IFields } from 'spfxhelper';
import { Log } from '@microsoft/sp-core-library';
import { IListItemsResponse, IListView } from 'spfxhelper/dist/SPFxHelper/Props/ISPListProps';
import { SPHttpClient } from '@microsoft/sp-http';


class Convert2Doc {

    private _logSource: string = undefined;
    private _webURL: string = undefined;
    private _client: SPHttpClient = undefined;
    private listName: string = undefined;
    private response: IListItemsResponse[] = [];
    private currentViewFields: string[] = [];
    private listFieldDetails: any[] = [];

    // Returns the List Operations object
    private _listOperations: SPListOperations = undefined;
    private get listOperations(): SPListOperations {
        if (!this._listOperations) {
            this._listOperations = new SPListOperations(this._client, this._webURL, this._logSource);
        }
        return this._listOperations;
    }

    // Returns the common operations object
    private _commonOperations: SPCommonOperations = undefined;
    private get commonOperations(): SPCommonOperations {
        if (!this._commonOperations) {
            this._commonOperations = new SPCommonOperations(this._client, this._webURL, this._logSource);
        }
        return this._commonOperations;
    }

    // Returns the field Operations object
    private _fieldOperations: SPFieldOperations = undefined;
    private get fieldOperations(): SPFieldOperations {
        if (!this._fieldOperations) {
            this._fieldOperations = new SPFieldOperations(this._client, this._webURL, this._logSource);
        }
        return this._fieldOperations;
    }

    /**
    * Creates the select query for retreiving the records
    */
    private get createSelectQuery(): string {

        let select: string[] = [];
        let expand: string[] = [];
        Log.verbose(this._logSource, `initiating query creation based on the current view fields...`);

        //Iterate over each query and create the query
        this.listFieldDetails.forEach(i => {

            switch (i.TypeAsString) {
                case "User":
                case "Person or Group":
                    select.push(`${i.InternalName}/Title`);
                    expand.push(i.InternalName);
                    break;
                case "Lookup":
                    select.push(`${i.InternalName}/${i.LookupField}`);
                    expand.push(i.InternalName);
                    break;
                default:
                    select.push(i.InternalName);
            }
        });

        Log.verbose(this._logSource, `Query generated: ?$select=${select.join(',')}&$expand=${expand.join(',')}`);
        return `?$select=${select.join(',')}&$expand=${expand.join(',')}`;
    }

    constructor(spHttp: SPHttpClient, webUrl: string, logSource: string, listName: string) {

        this._logSource = logSource;
        this._webURL = webUrl;
        this._client = spHttp;
        this.listName = listName;
    }

    /**
     * Creates the document for the export for the current list with current view fields
     */
    public async createDocument(): Promise<void> {

        Log.verbose(this._logSource, "initiating get of all items in the list...");
        let items: IListItemsResponse = await this.getItems();


        if (await this.validateColumnTypes()) {

            let showQnA: boolean = confirm("QnA format can be printed with the selected view. Do you want to proceed with the QnA format ?\n Press Ok to continue with QnA format, cancel to continue with List Format");
            // export in the format Q&A
            showQnA ? this.generateQnAFormat(items) : this.generateTableFormat(items);
        }
        else {
            // SHow in grid format
            this.generateTableFormat(items);
        }
    }

    /**
     * Returns all the items in a list
     * @param listName list name on which query needs to be performed
     */
    private async getItems(): Promise<IListItemsResponse> {

        let allItems: IListItemsResponse = { ok: true, result: [] };

        try {

            // Await to get the response for all the queries
            await this.getAllItems();
            Log.verbose(this._logSource, `got all items.`);
            Log.verbose(this._logSource, "collecting all items...");
            // Iterate over the responses and accumlate all the reveived items
            this.response.forEach(i => {
                if (i.ok) {
                    allItems.result = [...allItems.result, ...i.result];
                }
                else {
                    Log.error(this.listOperations.LogSource, i.error);
                }
            });
            Log.verbose(this._logSource, `items collected with the count ${allItems.result.length}`);
        }
        catch (e) {
            Log.error(this._logSource, new Error(`Error in the method Convert2Doc.getItems()`));
            Log.error(this._logSource, e);
        }
        finally {
            return Promise.resolve(allItems);
        }
    }

    /**
     * Recursively gets all the items in the list
     * @param nextLink 
     */
    private async getAllItems(nextLink?: string): Promise<void> {

        try {
            if (nextLink) {
                // Get the next batch of items using the next link
                Log.verbose(this._logSource, `getting the next set of 5000 records`);
                this.response.push(await this.listOperations.getListItemsByNextLink(nextLink));
                Log.verbose(this._logSource, `response received`);
            }
            else {
                // Get the current view fields with all field details
                Log.verbose(this._logSource, `retreiving all the fields for the current view...`);
                await this.getCurrentViewFields();
                Log.verbose(this._logSource, `all fields retreived`);

                // Call the first set of items
                Log.verbose(this._logSource, `getting the first set of 5000 records`);
                this.response.push(await this.listOperations.getListItemsByQuery(this.listName, `${this.createSelectQuery}&$top=5000`));
                Log.verbose(this._logSource, `response received`);
            }

            // Check if the recent revieved response has the next link
            if (this.response[this.response.length - 1].nextLink) {
                this.getAllItems(this.response[this.response.length - 1].nextLink);
            }
        }
        catch (e) {
            Log.error(this._logSource, new Error(`Error in the method Convert2Doc.getAllItems()`));
            Log.error(this._logSource, e);
        }
    }

    /**
     * Returns the fields in the current View
     * @param listName 
     */
    public async getCurrentViewFields(): Promise<void> {

        try {
            let viewId: string = SPCore.getParameterValue(location.href, "viewid");
            if (viewId) {
                let view: ISPBaseResponse = await this.commonOperations.queryGETResquest(`${this._webURL}/_api/web/lists/getByTitle('${this.listName}')/Views('${viewId}')/ViewFields`);
                this.currentViewFields = view.result["Items"];
            }
            else {
                let defaultView: IListView = await this.listOperations.getDefaultView(this.listName);
                let viewFields: IFields = await this.fieldOperations.getFieldsByView(this.listName, defaultView.view["Title"]);
                this.currentViewFields = viewFields.details["Items"];
            }

            let fieldDetails: IFields = await this.fieldOperations.getFieldsByList(this.listName);
            this.listFieldDetails = fieldDetails.details.filter(i => this.currentViewFields.indexOf(i.InternalName) > -1);

            // Keep the same orfer that of the fields in the view
            let orderedFields: any[] = [];
            this.currentViewFields.forEach(i => {
                orderedFields.push(this.listFieldDetails.filter(j => j.InternalName == i)[0]);
            });
            this.listFieldDetails = orderedFields;
        }
        catch (e) {
            Log.error(this._logSource, new Error(`Error in the method Convert2Doc.getCurrentViewFields()`));
            Log.error(this._logSource, e);
        }
    }

    /**
    * Validates the column types for QnA format
    * @param listName listname
    */
    private async validateColumnTypes(): Promise<boolean> {

        let isSingleLine: boolean = false;
        let isMultiline: boolean = false;
        let isValidforAnswerMode: boolean = false;
        Log.verbose(this._logSource, `Check if the QnA format cane be created from the current view ?`);

        this.listFieldDetails.forEach(i => {

            switch (i.TypeDisplayName.toLowerCase()) {
                case "single line of text":
                case "computed":
                    isSingleLine = i.Title === "Title" ? true : false;
                    break;
                case "multiple lines of text":
                    isMultiline = i.Title === "Answer" ? true : false;
            }
        });

        if ((this.currentViewFields.length == 2 && isSingleLine && isMultiline)) {
            isValidforAnswerMode = true;
        }
        Log.verbose(this._logSource, `${isValidforAnswerMode}, is the response`);
        return Promise.resolve(isValidforAnswerMode);
    }

    /**
     * Generates the table format for the output
     * @param items 
     */
    public generateTableFormat(items: IListItemsResponse): void {

        let html: string = '<table>';
        let index: number = 0;

        items.result.forEach(i => {

            html += `<tr style="height:30px"></tr>`;

            let isAlternate: boolean = index % 2 == 0;

            this.listFieldDetails.forEach(k => {

                let value: string = '';

                switch (k.TypeAsString) {
                    case "User":
                    case "Person or Group":
                        value = i[k.InternalName]["Title"];
                        break;
                    case "Lookup":
                        value = i[k.InternalName][k.LookupField];
                        break;
                    case "TaxonomyFieldType":
                        value = i[k.InternalName]["Label"];
                        break;
                    case "URL":
                        value = `<a href="${i[k.InternalName]["Url"]}" style="cursor:pointer;">${i[k.InternalName]["Description"]}</a>`;
                        break;
                    case "DateTime":
                        value = new Date(i[k.InternalName]).toLocaleString();
                        break;
                    default:
                        value = i[k.InternalName];
                }

                html += `<tr style="background-color:${isAlternate ? '#f3f3f3' : '#ffffff'}">`;
                html += `<td style="width:30%; border:${isAlternate ? '1px solid #ffffff' : '1px solid #bcb7b7'};">${k.Title}</td>`;
                html += `<td style="width:70%;border:${isAlternate ? '1px solid #ffffff' : '1px solid #bcb7b7'};">${value}</td>`;
                html += `</tr>`;
            });
            index += 1;
        });

        html = `${html}</table>`;
        this.generateDocument(html);
    }

    /**
     * Generates the QnA format for the output
     * @param items 
     */
    public generateQnAFormat(items: IListItemsResponse): void {

        let QnA: string = '';

        items.result.forEach(i => {

            this.currentViewFields.forEach(k => {

                if (k.toLowerCase().indexOf('title') > -1) {
                    QnA += `<h3>${i[k]}</h3>`;
                }
                else {
                    QnA += `<p>${i[k]}</p>`;
                }
            });
        });


        this.generateDocument(QnA);
    }

    /**
    * Generates the document and download it
    * @param sourceHTML 
    */
    public generateDocument(sourceHTML: string) {

        let headerHTML: string = `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
        <head><meta charset='utf-8'><title>${this.listName}</title></head><body>`;

        let titleHTML: string = `<h1><center>${this.listName}</center></h1><hr></hr>`;

        let footerHTML: string = "</body></html>";

        var sourceHTML = headerHTML + titleHTML + `<div id="source-html">${sourceHTML}</div>` + footerHTML;

        var source = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(sourceHTML);
        var fileDownload = document.createElement("a");
        document.body.appendChild(fileDownload);
        fileDownload.href = source;
        fileDownload.download = `${this.listName}.doc`;
        fileDownload.click();
        document.body.removeChild(fileDownload);
    }
}

export { Convert2Doc };