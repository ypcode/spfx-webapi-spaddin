import * as React from 'react';
import styles from './WebApiClient.module.scss';
import { IWebApiClientProps } from './IWebApiClientProps';
import { escape } from '@microsoft/sp-lodash-subset';

import {
	CommandBar,
	DetailsList,
	ISelection,
	Selection,
	SelectionMode,
	Panel,
	TextField,
	PrimaryButton,
	DefaultButton
} from 'office-ui-fabric-react';
import { IBusinessDocument } from '../../../entities/IBusinessDocument';
import { BusinessDocumentsServiceKey, IBusinessDocumentsService } from '../../../services/BusinessDocumentsService';
import { ApiConfigServiceKey, IApiConfigService } from '../../../services/ApiConfigService';

export interface IWebApiClientState {
	businessDocuments?: IBusinessDocument[];
	selectedDocument?: IBusinessDocument;
	selection?: ISelection;
	isAdding?: boolean;
	isEditing?: boolean;
	selectedView?: 'All' | 'My';
}

export default class WebApiClient extends React.Component<IWebApiClientProps, IWebApiClientState> {
	private businessDocsService: IBusinessDocumentsService;
	private apiConfig: IApiConfigService;
	private authenticated: boolean;

	constructor(props: IWebApiClientProps) {
		super(props);
		this.state = {
			businessDocuments: [],
			selectedDocument: null,
			isAdding: false,
			isEditing: false,
			selectedView: 'All',
			selection: new Selection({
				onSelectionChanged: this._onSelectionChanged.bind(this)
			})
		};
	}

	public componentWillMount() {
		this.props.serviceScope.whenFinished(() => {
			this.businessDocsService = this.props.serviceScope.consume(BusinessDocumentsServiceKey);
			this.apiConfig = this.props.serviceScope.consume(ApiConfigServiceKey);
			console.log('business service instance', this.businessDocsService);
			this._loadDocuments();
		});
	}

	private _loadDocuments(stateRefresh?: IWebApiClientState, forceView?: 'All' | 'My') {
		let { selectedView } = this.state;

		let effectiveView = forceView || selectedView;
		// After being authenticated
		this._executeOrDelayUntilAuthenticated(() => {
			switch (effectiveView) {
				case 'All':
					// Load all business documents when component is being mounted
					this.businessDocsService.getAllBusinessDocuments().then((docs) => {
						let state = stateRefresh || {};
						state.businessDocuments = docs;
						this.setState(state);
					});
					break;
				case 'My':
					// Load My business documents when component is being mounted
					this.businessDocsService.getMyBusinessDocuments().then((docs) => {
						let state = stateRefresh || {};
						state.businessDocuments = docs;
						this.setState(state);
					});
					break;
			}
		});
	}

	private _executeOrDelayUntilAuthenticated(action: Function): void {
		if (this.authenticated) {
			console.log('Is authenticated');
			action();
		} else {
			console.log('Still not authenticated');
			setTimeout(() => {
				this._executeOrDelayUntilAuthenticated(action);
			}, 1000);
		}
	}

	private _onSelectionChanged() {
		let { selection } = this.state;
		let selectedDocuments = selection.getSelection() as IBusinessDocument[];

		let selectedDocument = selectedDocuments && selectedDocuments.length == 1 && selectedDocuments[0];

		console.log('SELECTED DOCUMENT: ', selectedDocument);
		this.setState({
			selectedDocument: selectedDocument || null
		});
	}

	private _buildCommands() {
		let { selectedDocument } = this.state;

		const add = {
			key: 'add',
			name: 'Create',
			icon: 'Add',
			onClick: () => this.addNewBusinessDocument()
		};

		const edit = {
			key: 'edit',
			name: 'Edit',
			icon: 'Edit',
			onClick: () => this.editCurrentBusinessDocument()
		};

		const remove = {
			key: 'remove',
			name: 'Remove',
			icon: 'Remove',
			onClick: () => this.removeCurrentBusinessDocument()
		};

		let commands = [ add ];

		if (selectedDocument) {
			commands.push(edit, remove);
		}

		return commands;
	}

	private _buildFarCommands() {
		let { selectedDocument, selectedView } = this.state;

		const views = {
			key: 'views',
			name: selectedView == 'All' ? 'All' : "I'm in charge of",
			icon: 'View',
			subMenuProps: {
				items: [
					{
						key: 'viewAll',
						name: 'All',
						icon: 'ViewAll',
						onClick: () => this.selectView('All')
					},
					{
						key: 'inChargeOf',
						name: "I'm in charge of",
						icon: 'AccountManagement',
						onClick: () => this.selectView('My')
					}
				]
			}
		};

		let commands = [ views ];

		return commands;
	}

	public selectView(view: 'All' | 'My') {
		this.setState({
			selectedView: view
		});

		this._loadDocuments(null, view);
	}

	public addNewBusinessDocument() {
		console.log('ADD NEW DOCUMENT');
		this.setState({
			isAdding: true,
			selectedDocument: {
				Id: 0,
				Name: 'New document.docx',
				Purpose: '',
				InCharge: ''
			}
		});
	}

	public editCurrentBusinessDocument() {
		console.log('EDIT DOCUMENT');
		let { selectedDocument } = this.state;
		if (!selectedDocument) {
			return;
		}

		console.log('SELECTED DOCUMENT: ', selectedDocument);

		this.setState({
			isEditing: true
		});
	}

	public removeCurrentBusinessDocument() {
		let { selectedDocument } = this.state;
		if (!selectedDocument) {
			return;
		}

		if (confirm('Are you sure ?')) {
			this._executeOrDelayUntilAuthenticated(() => {
				this.businessDocsService
					.removeBusinessDocument(selectedDocument.Id)
					.then(() => {
						alert('Document is removed !');
						this._loadDocuments();
					})
					.catch((error) => {
						console.log(error);
						alert('Document CANNOT be removed !');
					});
			});
		}
	}

	private onValueChange(property: string, value: string) {
		let { selectedDocument } = this.state;
		if (!selectedDocument) {
			return;
		}

		selectedDocument[property] = value;
	}

	private onApply() {
		let { selectedDocument, isAdding, isEditing } = this.state;

		if (isAdding) {
			this._executeOrDelayUntilAuthenticated(() => {
				this.businessDocsService
					.createBusinessDocument(selectedDocument)
					.then(() => {
						alert('Document is created !');
						this._loadDocuments({
							selectedDocument: null,
							isAdding: false,
							isEditing: false
						});
					})
					.catch((error) => {
						console.log(error);
						alert('Document CANNOT be created !');
					});
			});
		} else if (isEditing) {
			this._executeOrDelayUntilAuthenticated(() => {
				this.businessDocsService
					.updateBusinessDocument(selectedDocument.Id, selectedDocument)
					.then(() => {
						alert('Document is updated !');
						this._loadDocuments({
							selectedDocument: null,
							isAdding: false,
							isEditing: false
						});
					})
					.catch((error) => {
						console.log(error);
						alert('Document CANNOT be updated !');
					});
			});
		}
	}

	private onCancel() {
		this.setState({
			selectedDocument: null,
			isAdding: false,
			isEditing: false
		});
	}

	public render(): React.ReactElement<IWebApiClientProps> {
		let { businessDocuments, selection, selectedDocument, isAdding, isEditing } = this.state;
		return (
			<div className={styles.webApiClient}>
				<div className={styles.container}>
					<iframe
						src={this.apiConfig.appRedirectUri}
						style={{ display: 'none' }}
						onLoad={() => (this.authenticated = true)}
					/>
					<CommandBar items={this._buildCommands()} farItems={this._buildFarCommands()} />
					<DetailsList
						items={businessDocuments}
						columns={[
							{
								key: 'id',
								name: 'Id',
								fieldName: 'Id',
								minWidth: 15,
								maxWidth: 30
							},
							{
								key: 'docName',
								name: 'Name',
								fieldName: 'Name',
								minWidth: 100,
								maxWidth: 200
							},
							{
								key: 'docPurpose',
								name: 'Purpose',
								fieldName: 'Purpose',
								minWidth: 100,
								maxWidth: 200
							},
							{
								key: 'inChargeOf',
								name: "Who's in charge",
								fieldName: 'InCharge',
								minWidth: 100,
								maxWidth: 200
							}
						]}
						selectionMode={SelectionMode.single}
						selection={selection}
					/>
					{selectedDocument &&
					(isAdding || isEditing) && (
						<Panel isOpen={true}>
							<TextField
								label="Name"
								value={selectedDocument.Name}
								onChanged={(v) => this.onValueChange('Name', v)}
							/>
							<TextField
								label="Purpose"
								value={selectedDocument.Purpose}
								onChanged={(v) => this.onValueChange('Purpose', v)}
							/>
							<TextField
								label="InCharge"
								value={selectedDocument.InCharge}
								onChanged={(v) => this.onValueChange('InCharge', v)}
							/>
							<PrimaryButton text="Apply" onClick={() => this.onApply()} />
							<DefaultButton text="Cancel" onClick={() => this.onCancel()} />
						</Panel>
					)}
				</div>
			</div>
		);
	}
}
