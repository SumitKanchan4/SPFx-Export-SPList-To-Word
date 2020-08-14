import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters,
  RowAccessor
} from '@microsoft/sp-listview-extensibility';
import { Convert2Doc } from './Convert2Doc';

/**
 * 
 */
export interface IExport2WordCommandSetProperties {

}

const LOG_SOURCE: string = 'Export2WordCommandSet';

export default class Export2WordCommandSet extends BaseListViewCommandSet<IExport2WordCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized Export2WordCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const export2WordCommand: Command = this.tryGetCommand('Export2Word');
    if (export2WordCommand) {
      // This command should be hidden if selected any rows.
      // export2WordCommand.visible = !(event.selectedRows.length > 0);
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'Export2Word':
        let cnvrt2docx: Convert2Doc = new Convert2Doc(this.context.spHttpClient as any, this.context.pageContext.web.absoluteUrl, LOG_SOURCE, this.context.pageContext.list.title);
        event.selectedRows.length == 0 ? cnvrt2docx.createDocument() : this.createDocumentSelectedItems(event, cnvrt2docx);
        break;
      default:
        throw new Error('Unknown command');
    }
  }



/**
 * Creates the documents for the selected items only
 * @param event 
 * @param cnvrt2docx 
 */
  private createDocumentSelectedItems(event: IListViewCommandSetExecuteEventParameters, cnvrt2docx: Convert2Doc) {

    let html: string = '<table>';
    let index: number = 0;


    event.selectedRows.forEach(i => {

      html += `<tr style="height:30px"></tr>`;
      let isAlternate: boolean = index % 2 == 0;

      i.fields.forEach(k => {

        let value: string = '';
        let fieldValue: any = i.getValue(k);

        switch (k.fieldType) {
          case "User":
          case "Person or Group":
            value = fieldValue && fieldValue.length > 0 ? fieldValue[0].title : '';
            break;
          case "Lookup":
            value = fieldValue && fieldValue.length > 0 ? fieldValue[0].lookupValue : '';
            break;
          case "TaxonomyFieldType":
            value = i.getValue(k).Label;
            break;
          case "URL":
            value = `<a href="${i.getValue(k)}" style="cursor:pointer;">${i.getValue(k)}</a>`;
            break;
          case "DateTime":
            value = new Date(i.getValue(k)).toLocaleString();
            break;
          default:
            value = i.getValue(k);
        }

        html += `<tr style="background-color:${isAlternate ? '#f3f3f3' : '#ffffff'}">`;
        html += `<td style="width:30%; border:${isAlternate ? '1px solid #ffffff' : '1px solid #bcb7b7'};">${k.displayName}</td>`;
        html += `<td style="width:70%;border:${isAlternate ? '1px solid #ffffff' : '1px solid #bcb7b7'};">${value}</td>`;
        html += `</tr>`;
      });
      index += 1;
    });

    html = `${html}</table>`;
    cnvrt2docx.generateDocument(html);

  }
}
