import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http'
;
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './NewsWebPartWebPart.module.scss';
import * as strings from 'NewsWebPartWebPartStrings';

export interface INewsItem {
  title : string;
  description: string;
}

export interface INewsWebPartWebPartProps {
  description: string;
}

export default class NewsWebPartWebPart extends BaseClientSideWebPart<INewsWebPartWebPartProps> {
  private newsItemUrl : string = "";

  protected addNewsItem(title : string) : Promise<INewsItem> {
      // Impl
      let reqBody : string = JSON.stringify({
          'Title' : title
      });

      let resprom : any = null;
      console.log(`addNewsItem(): Making AJAX Call to : ${this.newsItemUrl} \nWith: ${JSON.stringify(reqBody)}`);


      this.context.spHttpClient.post(this.newsItemUrl,SPHttpClient.configurations.v1,{
          headers : {
              'Accept' : 'application/json;odata=nometadata',
              'Content-Type' : 'application/json;odata=nometadata',
              'odata-version' : ''
          },
          body : reqBody
      }).then((res:SPHttpClientResponse) : Promise<INewsItem> => {
          console.log("Added List Item and returned the following - " + JSON.stringify(res.json()));
          resprom = res.json() as Promise<INewsItem>;
          return resprom;
      }).catch((err) => {
          console.log("Error inserting new News Item into List");
      });
      
      return resprom;
  }

  public render(): void {
    this.newsItemUrl = `${this.context.pageContext.web.absoluteUrl}/_api/Lists/GetByTitle('News')/Items`;

    this.domElement.innerHTML = `
      <div class="${ styles.newsWebPart }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            News Text : <input id ="newsText" type="text"></input>&nbsp;&nbsp;
                        <button type="button" id="btnAdd">Add News</button>
          </div>
        </div>
      </div>
      `;
      this.setAddNewsClicked();
  }

  private setAddNewsClicked(): void {  
    const webPart: NewsWebPartWebPart = this;  

    this.domElement.querySelector('#btnAdd').addEventListener('click', () => {  

        var newsText =  document.getElementById("newsText")["value"];

        alert("This will add this News Item to the News List : " + newsText);

        this.addNewsItem(newsText)
        .then((data) => {
          alert("News Item Added Successfully!!");
        })
        .catch((err)=>{
          alert("Error: News Item could not be Added!!!");
        });
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
