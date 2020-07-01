import Axios, { AxiosRequestConfig } from "axios";
import qs = require("qs");
import { IDocument } from "../model/IDocument";

export default class GraphController {
  private async authenticate() {
    const loginUrl: string = `https://login.microsoftonline.com/${process.env.TENANT}/oauth2/v2.0/token`;
    const headers = {
      'Content-Type': 'application/x-www-form-urlencoded'
    }
    const body = {
      grant_type: 'client_credentials',
      client_id: process.env.CLIENT_ID,
      client_secret: process.env.CLIENT_SECRET,
      scope: 'https://graph.microsoft.com/.default'
    }

    return Axios.post(loginUrl, qs.stringify(body), { headers: headers })
      .then(response=> {
        return response.data.access_token;
      })
      .catch(err => {
        console.log(err);
      });
  }

  public async getFiles(token: string, siteID: string, listID: string): Promise<IDocument[]> {
    if (token === null || token === '') {
      const token = await this.authenticate();
    }
    const requestUrl: string = `https://graph.microsoft.com/v1.0/sites/${siteID}/lists/${listID}/items?$filter=fields/NextReview lt '2020-06-26'&expand=fields`;
    
    return Axios.get(requestUrl, {
      headers: {          
          Authorization: `Bearer ${token}`
      }})
      .then(response => {
        let docs: IDocument[] = [];
        console.log(response.data.value);
        response.data.value.forEach(element => {
          docs.push({
            name: element.fields.FileLeafRef,
            description: element.fields.Description0,
            author: element.createdBy.user.displayName,
            url: element.webUrl,
            id: element.id,
            modified: new Date(element.lastModifiedDateTime)
          });
        });
        return docs;
      })
      .catch(err => {
        console.log(err);
        return [];
      });
  }

  public async updateItem(siteID: string, listID: string, itemID: string, nextReview: string) {
    const token = await this.authenticate();
    const requestUrl: string = `https://graph.microsoft.com/v1.0/sites/${siteID}/lists/${listID}/items/${itemID}/fields`;
    const config: AxiosRequestConfig = {  headers: {      
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    }};
    const fieldValueSet = {
      LastReviewed: new Date().toISOString(),
      NextReview: new Date(nextReview).toISOString()
    };  
    Axios.patch(requestUrl, 
                fieldValueSet,
                config
    )
    .then((response) => {
      console.log(response);
    })
    .catch((error) => {
      console.log(error);
    });
  }
}