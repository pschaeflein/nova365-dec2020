import * as React from 'react';
import styles from './GraphConsumer.module.scss';
import * as strings from "GraphConsumerWebPartStrings";
import {
  BaseButton,
  Button,
  CheckboxVisibility,
  DetailsList,
  DetailsListLayoutMode,
  PrimaryButton,
  SelectionMode,
  TextField,
} from "office-ui-fabric-react";
import { MSGraphClient } from "@microsoft/sp-http";
import { IGraphConsumerProps } from './IGraphConsumerProps';
import { IGraphConsumerState } from './IGraphConsumerState';
import { User } from '@microsoft/microsoft-graph-types';
import { escape } from '@microsoft/sp-lodash-subset';

// Configure the columns for the DetailsList component
let _usersListColumns = [
  {
    key: "displayName",
    name: "Display name",
    fieldName: "displayName",
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
  {
    key: "mail",
    name: "Mail",
    fieldName: "mail",
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
  },
  {
    key: "userPrincipalName",
    name: "User Principal Name",
    fieldName: "userPrincipalName",
    minWidth: 100,
    maxWidth: 200,
    isResizable: true,
  },
];

export default class GraphConsumer extends React.Component<IGraphConsumerProps, IGraphConsumerState> {

  constructor(props: IGraphConsumerProps, state: IGraphConsumerState) {
    super(props);

    // Initialize the state of the component
    this.state = {
      users: [],
      searchFor: ""
    };
  }

  public render(): React.ReactElement<IGraphConsumerProps> {
    return (
      <div className={styles.graphConsumer}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Search for a user!</span>
              <p className={styles.form}>
                <TextField
                  label={strings.SearchFor}
                  required={true}
                  onChange={this._onSearchForChanged}
                  onGetErrorMessage={this._getSearchForErrorMessage}
                  value={this.state.searchFor}
                />
              </p>
              <p className={styles.form}>
                <PrimaryButton
                  text='Search'
                  title='Search'
                  onClick={this._search}
                />
              </p>
              {
                (this.state.users != null && this.state.users.length > 0) ?
                  <p className={styles.form}>
                    <DetailsList
                      items={this.state.users}
                      columns={_usersListColumns}
                      setKey='set'
                      checkboxVisibility={CheckboxVisibility.hidden}
                      selectionMode={SelectionMode.none}
                      layoutMode={DetailsListLayoutMode.fixedColumns}
                      compact={true}
                    />
                  </p>
                  : null
              }
            </div>
          </div>
        </div>
      </div>
    );
  }

  private _onSearchForChanged = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {

    // Update the component state accordingly to the current user's input
    this.setState({
      searchFor: newValue,
    });
  }

  private _getSearchForErrorMessage = (value: string): string => {
    // The search for text cannot contain spaces
    return (value == null || value.length == 0 || value.indexOf(" ") < 0)
      ? ''
      : `${strings.SearchForValidationErrorMessage}`;
  }

  private _search = (event: React.MouseEvent<HTMLAnchorElement | HTMLButtonElement | HTMLDivElement | BaseButton | Button, MouseEvent>): void => {
    this._searchWithGraph();
  }

  private _searchWithGraph = (): void => {
    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        // From https://github.com/microsoftgraph/msgraph-sdk-javascript sample
        client
          .api("users")
          .version("v1.0")
          .select("displayName,mail,userPrincipalName")
          .filter(`startswith(givenName,'${escape(this.state.searchFor)}') or startswith(surname,'${escape(this.state.searchFor)}') or startswith(displayName,'${escape(this.state.searchFor)}')`)
          .get((err, res) => {

            if (err) {
              console.error(err);
              return;
            }

            // Prepare the output array
            var users: Array<User> = new Array<User>();

            // Map the JSON response to the output array
            res.value.map((item: any) => {
              users.push({
                displayName: item.displayName,
                mail: item.mail,
                userPrincipalName: item.userPrincipalName,
              });
            });

            // Update the component state accordingly to the result
            this.setState(
              {
                users: users,
              }
            );
          });
      });
  }
}
