import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { getIconClassName } from '@uifabric/styling';

import styles from './EmployeesWebPart.module.scss';
import * as strings from 'EmployeesWebPartStrings';

export interface IEmployeesWebPartProps {
	description: string;
}

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { sum } from 'lodash';

export interface ISPLists {
	value: ISPList[];
}

export interface ISPList {
	ProjectName: string;
	count_today: number;
	count_last_month: number;
	sum_today: number;
	sum_last_month: number;
}

export default class EmployeesWebPart extends BaseClientSideWebPart<IEmployeesWebPartProps> {
	private _getListData(): Promise<ISPLists> {
		return this.context.spHttpClient
			.get(
				this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('InfoPanel')/Items`,
				SPHttpClient.configurations.v1
			)
			.then((response: SPHttpClientResponse) => {
				debugger;
				return response.json();
			});
	}

	private _renderListAsync(): void {
		this._getListData().then((response) => {
			this._renderList(response.value);
		});
	}

	private _renderList(items: ISPList[]): void {
		let html: string = `<ul class="${styles.project__list}">`;
		const total = items[0].sum_today;
		const total_last_month = items[0].sum_last_month;

		items.forEach((item: ISPList) => {
			html += `	
        <li class="${styles.project__item}">
          <div class="${styles.project__title}">${item.ProjectName}</div>
          <div class="${styles.project__number}">
						${item.count_today}
						<div class="${styles.info}">
							<i class="${getIconClassName('InfoSolid')}"></i>
							<div class="${styles.tooltiptext}">
								<span class="${styles.text}">${item.count_last_month} за предыдущий месяц</span>
							</div>
						</div>  
					</div>
				</li>
      `;
		});

		html += `
			<li class="${styles.project__item}">
				<div class="${styles.project__title} ${styles.general}">
					Общее количество работников:
				</div>
				<div class="${styles.project__number} ${styles.general}">
					${total}
					<div class="${styles.info}">
						<i class="${getIconClassName('InfoSolid')}"></i>
						<div class="${styles.tooltiptext}">
							<span class="${styles.text}">${total_last_month}  за предыдущий месяц</span>
						</div>
					</div>  
				</div>
			</li>							
		</ul>`;

		const listContainer: Element = this.domElement.querySelector('#spListContainer');

		listContainer.innerHTML = html;
	}

	public render(): void {
		this.domElement.innerHTML = `
      <div class="${styles.employees}">
        <div class="${styles.container}">
          <div class="${styles.data}">
				    <div class="${styles.title}">Мы - одна команда!</div>
            <div id="spListContainer"> </div> 
          </div>
        </div>
      </div>`;
		this._renderListAsync();
	}

	protected get dataVersion(): Version {
		return Version.parse('1.0');
	}

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		return {
			pages: [
				{
					header: {
						description: strings.PropertyPaneDescription,
					},
					groups: [
						{
							groupName: strings.BasicGroupName,
							groupFields: [
								PropertyPaneTextField('description', {
									label: strings.DescriptionFieldLabel,
								}),
							],
						},
					],
				},
			],
		};
	}
}
