import * as React from 'react';
import { IClaringtonStaffDirectoryProps } from './IClaringtonStaffDirectoryProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IColumn } from 'office-ui-fabric-react/lib/components/DetailsList/DetailsList.types';
import { IPersonaSharedProps, Persona, PersonaSize, PersonaPresence } from '@fluentui/react/lib/Persona';
import { DetailsList, SelectionMode } from 'office-ui-fabric-react/lib/components/DetailsList';

export default class ClaringtonStaffDirectory extends React.Component<IClaringtonStaffDirectoryProps, any> {

  constructor(props) {
    super(props);
    this.state = {
      users: this.props.users,
      persona: [],
      columns: [
        {
          key: 'column1',
          name: 'Name',
          fieldName: 'displayName',
          minWidth: 50,
          isSorted: false,
          isResizable: true,
          isSortedDescending: false,
          sortAscendingAriaLabel: 'Sorted A to Z',
          sortDescendingAriaLabel: 'Sorted Z to A',
          onColumnClick: this._onColumnClick,
          onRender: (item: any) => (
            <Persona
              {...item}
              size={PersonaSize.size40}
            />
          ),
        },
        {
          key: 'column2',
          name: 'Department',
          fieldName: 'department',
          minWidth: 200,
          isResizable: true,
          isSorted: false,
          isSortedDescending: false,
          sortAscendingAriaLabel: 'Sorted A to Z',
          sortDescendingAriaLabel: 'Sorted Z to A',
          onColumnClick: this._onColumnClick,
        },
        {
          key: 'column4',
          name: 'Email',
          fieldName: 'mail',
          minWidth: 200,
          isResizable: true,
          isSorted: false,
          isSortedDescending: false,
          sortAscendingAriaLabel: 'Sorted A to Z',
          sortDescendingAriaLabel: 'Sorted Z to A',
          onColumnClick: this._onColumnClick,
        },
        {
          key: 'column5',
          name: 'Phone',
          fieldName: 'businessPhones',
          minWidth: 50,
          isSorted: false,
          isSortedDescending: false,
          sortAscendingAriaLabel: 'Sorted A to Z',
          sortDescendingAriaLabel: 'Sorted Z to A',
          onColumnClick: this._onColumnClick,
          onRender: (item: any) => (
            <div>{item.businessPhones.map(f => { return <div title={f}>{f}</div>; })}</div>
          ),
        },
      ]
    };

    this._queryAllUsers();
  }

  /**
     * Filter out guest and group users.  Return only active users.
     * @param response Response from Graph API.
     */
  private _filterUsers(response): any {
    let claringtonUsers = response.value.filter(value => { return value.mail != null && value.jobTitle != null; });
    claringtonUsers = claringtonUsers.filter(value => { return value.mail.includes('clarington.net'); });
    return claringtonUsers;
  }

  private async _queryUsers(): Promise<any> {
    let client = await this.props.context.msGraphClientFactory.getClient();
    return await client.api('users').top(200).select(['displayName', 'surname', 'givenName', 'mail', 'jobTitle', 'businessPhones', 'department', 'mobilePhone', 'userPrincipalName']).get();
  }

  private async _queryNextLink(nextLink): Promise<any> {
    let client = await this.props.context.msGraphClientFactory.getClient();
    return await client.api(nextLink).get();
  }

  private async _queryAllUsers(nextLink?, users?): Promise<any> {
    let usersOutput = users ? users : [];

    if (nextLink) {
      let queryNextLinkResult = await this._queryNextLink(nextLink);
      usersOutput.push(...this._filterUsers(queryNextLinkResult));

      if (queryNextLinkResult["@odata.nextLink"]) {
        this._queryAllUsers(queryNextLinkResult["@odata.nextLink"], usersOutput);
      }
      else {
        this._setUserState(usersOutput);
      }
    }
    else {
      // Make initial query. 
      let queryUserResult = await this._queryUsers();
      usersOutput.push(...this._filterUsers(queryUserResult));
      if (queryUserResult["@odata.nextLink"]) {
        this._queryAllUsers(queryUserResult["@odata.nextLink"], usersOutput);
      }
      else {
        this._setUserState(usersOutput);
      }
    }
  }

  private _setUserState(usersOutput): void {
    this.setState({
      users: usersOutput,
      persona: [...usersOutput.map(user => {
        return {
          imageUrl: "https://www.google.ca",
          imageInitials: "ZZ",
          text: user.displayName,
          secondaryText: user.jobTitle,
          ...user
        };
      })]
    });
  }

  //#region Grid Methods
  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { columns, users } = this.state;
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
        this.setState({
          announcedMessage: `${currColumn.name} is sorted ${currColumn.isSortedDescending ? 'descending' : 'ascending'}`,
        });
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newUsers = this._copyAndSort(users, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      columns: newColumns,
      users: newUsers,
    });
  }

  private _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
  }
  //#endregion

  public render(): React.ReactElement<IClaringtonStaffDirectoryProps> {
    return (
      <div>
        <DetailsList
          items={this.state.persona}
          columns={this.state.columns}
          selectionMode={SelectionMode.none}
          // selection={this._selection}
          onShouldVirtualize={() => false}
        />
      </div>
    );
  }
}
