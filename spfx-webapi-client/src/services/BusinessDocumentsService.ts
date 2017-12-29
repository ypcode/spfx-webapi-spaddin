import { IApiConfigService, ApiConfigServiceKey } from './ApiConfigService';
import HttpClient from '@microsoft/sp-http/lib/httpClient/HttpClient';
import { IBusinessDocument } from '../entities/IBusinessDocument';
import { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';

export interface IBusinessDocumentsService {
	getAllBusinessDocuments(): Promise<IBusinessDocument[]>;
	getMyBusinessDocuments(): Promise<IBusinessDocument[]>;
	getBusinessDocument(id: number): Promise<IBusinessDocument>;
	createBusinessDocument(businessDocument: IBusinessDocument): Promise<any>;
	updateBusinessDocument(id: number, update: IBusinessDocument): Promise<any>;
	removeBusinessDocument(id: number): Promise<any>;
}

export class BusinessDocumentsService implements IBusinessDocumentsService {
  private httpClient: HttpClient;
  private apiConfig: IApiConfigService;


	constructor(private serviceScope: ServiceScope) {
		serviceScope.whenFinished(() => {
      this.httpClient = serviceScope.consume(HttpClient.serviceKey);
      this.apiConfig = serviceScope.consume(ApiConfigServiceKey);
		});
	}

	public getAllBusinessDocuments(): Promise<IBusinessDocument[]> {
		return this.httpClient.get(this.apiConfig.apiUrl, HttpClient.configurations.v1, {
      mode: 'cors',
      credentials: 'include'
    }).then((resp) => resp.json());
	}

	public getMyBusinessDocuments(): Promise<IBusinessDocument[]> {
		return this.httpClient.get(this.apiConfig.apiMyDocumentsUrl, HttpClient.configurations.v1, {
      mode: 'cors',
      credentials: 'include'
    }).then((resp) => resp.json());
	}

	public getBusinessDocument(id: number): Promise<IBusinessDocument> {
		return this.httpClient.get(`${this.apiConfig.apiUrl}/${id}`, HttpClient.configurations.v1,{
      mode: 'cors',
      credentials: 'include'
    }).then((resp) => resp.json());
	}

	public createBusinessDocument(businessDocument: IBusinessDocument): Promise<any> {
		return this.httpClient
			.post(`${this.apiConfig.apiUrl}`, HttpClient.configurations.v1, {
        body: JSON.stringify(businessDocument),
        headers: [
          ['Content-Type','application/json']
        ],
				mode: 'cors',
				credentials: 'include'
			})
			.then((resp) => resp.json());
	}

	public updateBusinessDocument(id: number, update: IBusinessDocument): Promise<any> {
		return this.httpClient
			.fetch(`${this.apiConfig.apiUrl}/${id}`, HttpClient.configurations.v1, {
        body: JSON.stringify(update),
        headers: [
          ['Content-Type','application/json']
        ],
				mode: 'cors',
				credentials: 'include',
				method: 'PUT'
			});
	}

	public removeBusinessDocument(id: number): Promise<any> {
		return this.httpClient
			.fetch(`${this.apiConfig.apiUrl}/${id}`, HttpClient.configurations.v1, {
				mode: 'cors',
        credentials: 'include',
        method:'DELETE'
			});
	}
}

export const BusinessDocumentsServiceKey = ServiceKey.create<IBusinessDocumentsService>(
	'ypcode:bizdocs-service',
	BusinessDocumentsService
);
