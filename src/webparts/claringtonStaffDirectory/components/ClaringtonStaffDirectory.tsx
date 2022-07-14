import * as React from 'react';
import { IColumn } from 'office-ui-fabric-react/lib/components/DetailsList/DetailsList.types';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { DetailsList, SelectionMode } from 'office-ui-fabric-react/lib/components/DetailsList';
import { Shimmer } from 'office-ui-fabric-react';
import { IconButton, SearchBox } from '@fluentui/react';

import { IClaringtonStaffDirectoryProps, IStaffGridState } from './IClaringtonStaffDirectory';

import { PnPClientStorage } from "@pnp/core";



class MyShimmer extends React.Component {
  public render() {
    return (<div>
      <div style={{ marginBottom: '15px', marginTop: '20px' }}>
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

class StaffGrid extends React.Component<any, IStaffGridState> {
  constructor(props) {
    super(props);
    this.state = {
      loadingUsers: true, // Set this to true by default.  It will be set to false if/when the AD query is complete.
      columns: [
        {
          key: 'column1',
          name: 'Name',
          fieldName: 'displayName',
          minWidth: 200,
          isSorted: true,
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
      ],
      groups: [],
      persona: null,
      allPersonas: null
    };

    this._queryAllUsers();
  }

  private storage = new PnPClientStorage();
  private STORAGE_KEY = 'myTestKey';

  public componentDidUpdate(prevProps, prevState, snapshot) {
    if (prevProps.searchString !== this.props.searchString) {
      this._applySearchFilter(this.props.searchString);
    }
  }

  //#region Get Users
  /**
   * When this method is called for the first time BOTH parameters should be null.
   * Step 1: Run _queryUsers() to get a list of users. 
   * Step 2: If there are more users to be queried via '@odata.nextLink' run this method again. 
   * Step 3: Repeat Step 2 until '@odata.nextLink' is not set. 
   * Step 4: Run _setUserState() to render users on the page. 
   * Step 5: ???
   * @param nextLink A string that will tell AD to query the next batch of users.
   * @param users An array of users that have already been queried.
   */
  private async _queryAllUsers(nextLink?, users?): Promise<any> {
    let usersOutput = users ? users : [];

    if (nextLink) {
      // Run the next query for more users.
      let queryNextLinkResult = await this._queryNextLink(nextLink);

      // Take the results of the next query and append them to the running total.
      usersOutput.push(...this._filterUsers(queryNextLinkResult));

      // Render the running total of users that we have queried so far.
      // * This will continue to render users until we have queried all of them.
      this._setUserState(usersOutput);

      // After running the next query, check if there is another next query.  
      if (queryNextLinkResult["@odata.nextLink"]) {
        this._queryAllUsers(queryNextLinkResult["@odata.nextLink"], usersOutput);
      }
      else if (usersOutput.length > 0) {
        // ! This is what hides the loading icons and displays the list of users.
        /** 
         * nextLink was provided as a parameter AND queryNextLinkResult["@odata.nextLink"] is not found AND there are users found in usersOutput.
         * This means that we are done querying our users and it's time to hide the loading icons and show our users.
        */
        this.setState({ loadingUsers: false });

        alert('ALL DONE!  NO MORE NEXTLINK!');
        // Whenever the users have been filtered save the filtered result in local storage. 
        this._saveUsersInLocalStorage(usersOutput);
      }
    }
    // Make initial query. 
    else {
      // Run the initial query for users. 
      debugger;
      let queryUserResult = await this._queryUsers();
      debugger;

      if (queryUserResult.hasOwnProperty('value')) {
        // Append the results of the initial query to a running list of users.
        usersOutput.push(...this._filterUsers(queryUserResult));
      }
      else if (queryUserResult.length > 0) {
        usersOutput.push(...this._filterUsers({ value: queryUserResult }));
      }
      else {
        // This shouldn't happen.... I hope.
        alert('Something went wrong.  Please contact helpdesk@clarington.net');
      }

      // Render the list of users that we have queried. 
      this._setUserState(usersOutput);

      // If there are more users to be discovered, we will query them here. 
      if (queryUserResult["@odata.nextLink"]) {
        this._queryAllUsers(queryUserResult["@odata.nextLink"], usersOutput);
      }
    }
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

  private async _queryUsers(nextLink?: string): Promise<any> {
    let client = await this.props.context.msGraphClientFactory.getClient();

    // Check to see if 'nextLink' has been passed into this function.  If it has been passed we can assume that we're query from AD.
    if (nextLink) {
      return await client.api(nextLink).get();
    }
    else {
      // Since 'nextLink' has not been provided this means we are running our first search.
      // Before querying AD, check to see if there are any users in local storage. 
      let usersFromLocalStorage = this._getUsersFromLocalStorage();

      // If there are any users in local storage return those users BEFORE we query AD.
      // TODO: Uncomment the if statement below when ready.
      if (usersFromLocalStorage) {
        // ! This is what hides the loading icons and displays the list of users.
        this.setState({ loadingUsers: false });
        return usersFromLocalStorage;
      }
      else {
        return await client.api('users').top(200).select(['displayName', 'surname', 'givenName', 'mail', 'jobTitle', 'businessPhones', 'department', 'mobilePhone', 'userPrincipalName', 'accountEnabled']).get();
      }
    }
  }

  private async _queryNextLink(nextLink): Promise<any> {
    return await this._queryUsers(nextLink);
  }

  private _setUserState(usersOutput, callback?: Function): void {
    let persona = [...usersOutput.map(user => {
      return {
        imageUrl: `/_layouts/15/userphoto.aspx?size=L&username=${user.mail}`,
        imageInitials: `${user.givenName.charAt(0)}${user.surname.charAt(0)}`,
        text: user.displayName,
        secondaryText: user.jobTitle,
        ...user
      };
    })];

    this.setState({
      //users: usersOutput,
      allPersonas: persona
    }, () => {
      if (callback) {
        callback();
      }
      this._applySearchFilter(this.props.searchString);
    });
  }

  /**
   * Get a list of users from local storage. 
   */
  private _getUsersFromLocalStorage = () => {
    // TODO: This method should check and return any users found in local storage.
    // How it should work. 
    return this.storage.local.get(this.STORAGE_KEY);

    // But what happens when it doesn't work. 
    //return this.storage.local.get('badkeythatdoesntexistqwerty');
  }

  /**
   * Set users that have been queried from AD.
   * This method should override any existing values that are being stored in local storage.
   */
  private _saveUsersInLocalStorage = (input: any) => {
    this.storage.local.put(this.STORAGE_KEY, input, new Date(Date.now() + (6.048e+8)));
  }

  /**
   * Delete any and all users saved in local storage, query AD for a list of users, save the new result in local storage.
   */
  private _clearLocalStorageAndQueryAD = () => {
    alert('_clearLocalStorageAndQueryAD');
  }
  //#endregion

  //#region Help Methods
  /**
   * Generate an array that has the users grouped by a given field. 
   * Source: https://stackoverflow.com/a/65834042
   * 
   * @param itemsList Visible Users.
   * @param fieldName Grouped Field.
   */
  public groupsGenerator(itemsList, fieldName) {
    // Array of group objects
    const groupObjArr = [];

    // Get the group names from the items list
    const groupNames = new Set(itemsList.map(item => item[fieldName]));

    // Iterate through each group name to build proper group object
    groupNames.forEach(gn => {
      // Count group items
      const groupLength = itemsList.filter(item => item[fieldName] === gn).length;

      // Find the first group index
      const groupIndex = itemsList.map(item => item[fieldName]).indexOf(gn);

      // Generate a group object
      groupObjArr.push({
        key: gn, name: gn, level: 0, count: groupLength, startIndex: groupIndex
      });
    });

    this.setState({ groups: groupObjArr });

    // The final groups array returned
    return groupObjArr;
  }

  /**
   * Sort the visible users by a given column.
   * This method will always apply two sorts to the array of users.  The first will always be the department columns.  The second is whatever the user wants.  
   * @param items Visible Users.
   * @param columnKey Field Name.
   * @param isSortedDescending Is Sorted Descending (bool)
   */
  private _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;

    if (columnKey === 'department') {
      // Sory by just Department.
      return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
    }
    else {
      // Sort by Department AND columnKey.  
      let output = [];

      // Group everything by departments. 
      let group = items.reduce((r, a) => {
        r[a['department']] = [...r[a['department']] || [], a];
        return r;
      }, {});

      // Iterate over each group/department. 
      for (var departmentKey in group) {
        if (group.hasOwnProperty(departmentKey)) {
          // This should sort the department users but maintain their department grouping.s
          output.push(...group[departmentKey].slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1)));
        }
      }
      return output;
    }
  }

  /**
   * Click event for sorting columns. 
   * @param ev Event
   * @param column Column Clicked
   */
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
    }, () => this.groupsGenerator(newUsers, 'department'));
  }
  //#endregion

  //#region Search Methods
  /**
   * Take the users input from the search box and filter users.
   * This method will update the state object to display the correct users.
   * @param newValue User input from search box
   */
  private _applySearchFilter = (newValue: string) => {
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

    // ALWAYS sort by department first.  This will ensure that the list of users is first sorted by department, then sorted by other columns. 

    visibleUsers = visibleUsers.slice(0).sort((a, b) => ((a['department'] > b['department'] ? 1 : -1)));

    // Apply any sorting. 
    let sortedColumn = this.state.columns.find(col => { return col.isSorted; });

    if (sortedColumn) {
      visibleUsers = this._copyAndSort(visibleUsers, sortedColumn.fieldName!, sortedColumn.isSortedDescending);
    }

    // * This is where we set what users will be displayed. 
    this.setState({ persona: visibleUsers }, () => {
      this.groupsGenerator(this.state.persona, "department");
    });
  }
  //#endregion

  public render(): React.ReactElement<any> {
    return <div>
      {
        (this.state.persona === null || this.state.loadingUsers === true) ?
          <MyShimmer /> :
          <DetailsList
            items={this.state.persona}
            columns={this.state.columns}
            groups={this.state.groups}
            selectionMode={SelectionMode.none}
            onShouldVirtualize={() => false}
          />
      }
    </div>;
  }
}

export default class ClaringtonStaffDirectory extends React.Component<IClaringtonStaffDirectoryProps, any> {

  constructor(props) {
    super(props);
    this.state = {
      searchString: undefined
    };
  }

  public render(): React.ReactElement<IClaringtonStaffDirectoryProps> {
    return (
      <div style={{ maxWidth: '1300px', margin: 'auto' }}>
        <SearchBox
          placeholder={"Search by Name, Job Title, or Department"}
          onChange={(event: any, newValue: string) => this.setState({ searchString: newValue })}
        />
        <StaffGrid {...this.props} searchString={this.state.searchString} />
      </div>
    );
  }
}
