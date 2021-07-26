import * as React from 'react';
import { IColumn } from 'office-ui-fabric-react/lib/components/DetailsList/DetailsList.types';
import { Persona, PersonaSize, PersonaPresence } from '@fluentui/react/lib/Persona';
import { DetailsList, SelectionMode } from 'office-ui-fabric-react/lib/components/DetailsList';
import { Shimmer } from 'office-ui-fabric-react';
import { IconButton } from '@fluentui/react/lib/Button';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { MessageBar } from '@fluentui/react/lib/components/MessageBar/MessageBar';
import { Link } from '@fluentui/react/lib/components/Link/Link';
import { IClaringtonStaffDirectoryProps, IClaringtonStaffDirectoryState } from './IClaringtonStaffDirectory';


class StaffGrid extends React.Component<any> {
  public render(): React.ReactElement<any> {
    return <div>
      <SearchBox placeholder="Search by Name, Job Title, or Department" onChange={this.props.onSearchChange} />
      <DetailsList
        items={this.props.items}
        columns={this.props.columns}
        selectionMode={SelectionMode.none}
        onShouldVirtualize={() => false}
      />
    </div>;
  }
}

class MyShimmer extends React.Component {
  public render() {
    return (<div>
      <div style={{ marginBottom: '15px' }}>
        <Shimmer style={{ marginBottom: '5px' }} />
        <Shimmer width="75%" style={{ marginBottom: '5px' }} />
        <Shimmer width="50%" style={{ marginBottom: '5px' }} />
      </div>
      <div style={{ marginBottom: '15px' }}>
        <Shimmer style={{ marginBottom: '5px' }} />
        <Shimmer width="75%" style={{ marginBottom: '5px' }} />
        <Shimmer width="50%" style={{ marginBottom: '5px' }} />
      </div>
      <div style={{ marginBottom: '15px' }}>
        <Shimmer style={{ marginBottom: '5px' }} />
        <Shimmer width="75%" style={{ marginBottom: '5px' }} />
        <Shimmer width="50%" style={{ marginBottom: '5px' }} />
      </div>
    </div>);
  }
}

//TODO: Replace allPersona with users variable. 
export default class ClaringtonStaffDirectory extends React.Component<IClaringtonStaffDirectoryProps, IClaringtonStaffDirectoryState> {

  constructor(props) {
    super(props);

    this.state = {
      users: this.props.users,
      persona: null,
      allPersonas: [],
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
          onRender: (item: any) => (
            <div>
              {/* <Icon aria-label="Mail" iconName="MailIcon" /> */}
              <IconButton href={`mailto:${item.mail}`} iconProps={{ iconName: 'Mail' }} title={item.mail} ariaLabel="Mail" />
              <span>{item.mail}</span>
            </div>
          )
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
    let claringtonUsers = response.value.filter(value => {
      return value.mail != null
        && value.jobTitle != null
        && value.surname != null
        && value.givenName != null
        && value.department != null
        && value.accountEnabled === true;
    });
    claringtonUsers = claringtonUsers.filter(value => { return value.mail.includes('clarington.net'); });
    return claringtonUsers;
  }

  private async _queryUsers(): Promise<any> {
    let client = await this.props.context.msGraphClientFactory.getClient();
    return await client.api('users').top(200).select(['displayName', 'surname', 'givenName', 'mail', 'jobTitle', 'businessPhones', 'department', 'mobilePhone', 'userPrincipalName', 'accountEnabled']).get();
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

  // TODO: See if I can get away with only calling this method once.
  private _setUserState(usersOutput, callback?: Function): void {
    let persona = [...usersOutput.map(user => {
      return {
        imageUrl: "https://www.google.ca",
        imageInitials: `${user.givenName.charAt(0)}${user.surname.charAt(0)}`,
        text: user.displayName,
        secondaryText: user.jobTitle,
        ...user
      };
    })];

    this.setState({
      users: usersOutput,
      persona: persona,
      allPersonas: persona
    }, callback && callback());
  }

  //#region Grid Methods
  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { columns, persona } = this.state;
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];

    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });

    const newUsers = this._copyAndSort(persona, currColumn.fieldName!, currColumn.isSortedDescending);

    this.setState({
      persona: newUsers,
      columns: newColumns
    });
  }

  private _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
  }
  //#endregion

  //#region Search Box Events
  /**
   * Take the users input from the search box and filter users.
   * This method will update the state object to display the correct users.
   * @param newValue User input from search box
   */
  private _applySearchFilter = (newValue: string) => {
    debugger;
    let visibleUsers = this.state.persona;
    if (newValue) {
      newValue = newValue.toLowerCase();
      // All users =  this.state.allPersonas;
      // Visible users = this.state.persona;
      visibleUsers = this.state.allPersonas.filter(user => {
        // start with display name but I should also use jobTitle and department
        return user.displayName.toLowerCase().includes(newValue)
          || user.jobTitle.toLowerCase().includes(newValue)
          || (user.department && user.department.toLowerCase().includes(newValue));
      });
    }
    else {
      visibleUsers = this.state.allPersonas;
    }

    // Apply any sorting. 
    let sortedColumn = this.state.columns.find(col => { return col.isSorted; });

    if (sortedColumn) {
      visibleUsers = this._copyAndSort(visibleUsers, sortedColumn.fieldName!, sortedColumn.isSortedDescending);
    }

    debugger;
    this.setState({ persona: visibleUsers });
  }

  private _onSearchChange = (event: any, newValue: string) => {
    this._applySearchFilter(newValue);
  }
  //#endregion

  public render(): React.ReactElement<IClaringtonStaffDirectoryProps> {
    return (
      <div style={{ maxWidth: '1300px', margin: 'auto' }}>
        {
          this.state.persona === null ?
            <MyShimmer /> :
            <StaffGrid onSearchChange={this._onSearchChange} items={this.state.persona} columns={this.state.columns} />
        }
      </div>
    );
  }
}
