import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';
import * as $ from 'jquery';

import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import { SPComponentLoader } from '@microsoft/sp-loader';
export interface SPFxLists {
  value: SPFxList[];
}

export interface SPFxList {
  Title: string;
  Id: string;
}

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  protected onInit(): Promise<void> {
    SPComponentLoader.loadCss("https://spfx2.sharepoint.com/sites/spfx/Style%20Library/style.css");
    return super.onInit();
  }

  private _getListData(): Promise<SPFxLists> {
    return this.context.spHttpClient.get(`https://spfx2.sharepoint.com/sites/spfx/_api/web/lists/getbytitle('List')/Items`, SPHttpClient.configurations.v1)
      .then(
        (response: SPHttpClientResponse) => {
          return response.json();
        }
      );
  }

  private _renderList(items: SPFxList[]): void {
    let html: string = '';
    items.forEach(
      (item: SPFxList) => {
        html += `<ul><li><span>${item.Title}</span></li></ul>`;
      }
    );

    const listContainer: Element = this.domElement.querySelector("#spListContainer");
    listContainer.innerHTML = html;
  }

  private _renderListAsync(): void {
    this._getListData().then((response) => {
      this._renderList(response.value);
    });
  }

  public render(): void {
    this.domElement.innerHTML = `
        <div class="${ styles.helloWorld}">
      <div class="${ styles.container}">
        <div class="${ styles.row}">
          <div class="${ styles.column}">
            <span class="${ styles.title}">Welcome to SharePoint!</span>
    <p class="${ styles.subTitle}">Customize SharePoint experiences using Web Parts.</p>
      <p class="${ styles.description}">${escape(this.properties.description)}</p>
        <a href="https://aka.ms/spfx" class="${ styles.button}">
          <span class="${ styles.label}">Learn more</span>
            </a>
            </div>
            <div class="${ styles.column}">Add by linyu.</div>
            </div>
            </div>
            </div>
            <div id="btn">Click</div>
            <div><input id="myInput" type="text"/></div>
            <div id="spListContainer"></div>
            `;
    //this._renderListAsync();

    $(function () {
      $("#btn").click(function () {
        var currentValue = $("#myInput").val();
        $.ajax({
          url: "https://spfx2.sharepoint.com/sites/spfx/_api/contextinfo",
          method: "POST",
          headers: { "Accept": "application/json; odata=verbose" },
          success: function (data) {
            $.ajax({
              url: "https://spfx2.sharepoint.com/sites/spfx/_api/web/lists/getbyTitle('List')/items",
              type: "POST",
              data: JSON.stringify({ '__metadata': { 'type': 'SP.Data.ListListItem' }, 'Title': currentValue }),
              headers: {
                "accept": "application/json;odata=verbose",
                "content-type": "application/json;odata=verbose",
                "X-RequestDigest": data.d.GetContextWebInformation.FormDigestValue,
                "X-HTTP-Method": "POST"
              },
              success: function (data) {
                alert("success");
                //console.log("success");
              }
            });
          },
          error: function (data) {
          }
        });

      });
      // $("#btn").click(function () {
      //   //Load SharePoint List Data
      //   var endPointUrl = "https://spfx2.sharepoint.com/sites/spfx/_api/web/lists/getbyTitle('List')/items";
      //   var headers = {
      //     "accept": "application/json;odata=verbose"
      //   };
      //   $.ajax({
      //     url: endPointUrl,
      //     type: "GET",
      //     headers: headers,
      //     success: function (data) {
      //       var result = data.d.results;
      //       var myData = [];
      //       var html = "";
      //       result.forEach(function (currentValue, index, arr) {
      //         html+=`<div>`+currentValue.Title+`</div>`;
      //       })
      //       $("#spListContainer").html(html);
      //     }
      //   });
      // });
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
