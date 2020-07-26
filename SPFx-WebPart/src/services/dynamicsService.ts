import axios from "axios";
import { Observable } from "../hooks/observable";
import { AadTokenProviderFactory,AadHttpClient,IHttpClientOptions } from "@microsoft/sp-http";
export default class DynamicsService {
  public readonly errorMessage = new Observable<string>(null);
  private accessToken: string;
  public aadTokenProviderFactory: AadTokenProviderFactory;
  public aadHttpClient: AadHttpClient;
  public azureFunctionUri: string;
  // update ResourceURI based on your dynamic crm uri
  public resourceUri:string;

  public async getAccessToken(){
    const token = sessionStorage.getItem("dynamic365Token");
    if(token)
      this.accessToken = token;
    else{
      await 
      this.aadTokenProviderFactory
      .getTokenProvider()
      .then((tokenProvider) => {
        tokenProvider
          .getToken(this.resourceUri)
          .then((t) => {
            this.accessToken = t;
            sessionStorage.setItem("dynamic365Token",t);
          })
          .catch((err) => console.log("Error: " + err));
      });
    }
  }

  public async getAccounts() {
    const headers: Headers = new Headers();
        headers.append("Accept", "application/json");
        const requestOptions:IHttpClientOptions  = {
            headers          
        };
        const response = await this.aadHttpClient.get(
            `${this.azureFunctionUri}?resource=${this.resourceUri}`,
            AadHttpClient.configurations.v1,
            requestOptions
        );
        const result = await response.text();
        const accounts = JSON.parse(result);
        return accounts.value;
  }

  public async getContacts(id:string){
    const url = `${this.resourceUri}/api/data/v9.0/contacts?$top=5&$select=fullname,emailaddress1,telephone1,address1_city&$filter=_accountid_value eq '${id}'`;
    const response = await axios({
      url,
      method: "GET",
      headers:{"Authorization":`Bearer ${this.accessToken}`}
    });
    return response.data.value;
  }

  public async searchAccounts (name:string){
    const url = `${this.resourceUri}/api/data/v9.0/accounts?$top=20&$select=name,emailaddress1&$filter=contains(name,'${name}')`;
    const response = await axios({
      url,
      method: "GET",
      headers:{"Authorization":`Bearer ${this.accessToken}`}
    });
    return response.data.value;
  }
}
