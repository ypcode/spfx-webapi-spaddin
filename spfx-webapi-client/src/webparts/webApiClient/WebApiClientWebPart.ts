import { ApiConfigServiceKey } from './../../services/ApiConfigService';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-webpart-base';

import * as strings from 'WebApiClientWebPartStrings';
import WebApiClient from './components/WebApiClient';
import { IWebApiClientProps } from './components/IWebApiClientProps';

export interface IWebApiClientWebPartProps {
	remoteApiHost: string;
	appInstanceId: string;
}

export default class WebApiClientWebPart extends BaseClientSideWebPart<IWebApiClientWebPartProps> {
	public onInit(): Promise<any> {
		return (
			super
				.onInit()
				// When configuration is done, we get the properly configured instances of the services we want to use
				.then(() => {
					this.context.serviceScope.whenFinished(() => {
						let apiConfig = this.context.serviceScope.consume(ApiConfigServiceKey);
						apiConfig.configure(
							this.context.pageContext.web.absoluteUrl,
							this.properties.remoteApiHost,
							this.properties.appInstanceId
						);
					});
				})
		);
	}

	public render(): void {
		this.domElement.innerHTML = 'Loading...';
		this.context.serviceScope.whenFinished(() => {
			const element: React.ReactElement<IWebApiClientProps> = React.createElement(WebApiClient, {
				serviceScope: this.context.serviceScope
			});

			ReactDom.render(element, this.domElement);
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
								PropertyPaneTextField('remoteApiHost', {
									label: 'Remote API Host URL'
                }),
                PropertyPaneTextField('appInstanceId', {
									label: 'App Instance ID'
								})
							]
						}
					]
				}
			]
		};
	}
}
