import * as React from "react";
import { BaseComponent } from "office-ui-fabric-react/lib/Utilities";
import { IPersonaProps, PersonaCoin } from "office-ui-fabric-react/lib/Persona";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import { IInputs } from "./generated/ManifestTypes";
import { Guid } from "guid-typescript";
import {
  parseGuid,
  retrieve,
  retrieveMultiple,
  WebApiConfig,
} from "xrm-webapi";

import {
  IBasePickerSuggestionsProps,
  IBasePicker,
  NormalPeoplePicker,
  ValidationState,
} from "office-ui-fabric-react/lib/Pickers";

export interface IPeoplePersona {
  text?: string;
  secondaryText?: string;
  id?: string;
  entity?: string;
}

export interface IPeopleProps {
  people?: any;
  preselectedpeople?: any;
  context?: ComponentFramework.Context<IInputs>;
  peopleList?: (newValue: any) => void;
}

export interface IPeoplePickerState {
  currentPicker?: number | string;
  delayResults?: boolean;
  peopleList: IPersonaProps[];
  mostRecentlyUsed: IPersonaProps[];
  currentSelectedItems?: IPersonaProps[];
  isPickerDisabled?: boolean;
}

//Loading comments text
const suggestionProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: "Suggested Records",
  mostRecentlyUsedHeaderText: "Suggested Records",
  noResultsFoundText: "No results found",
  loadingText: "Loading",
  showRemoveButtons: true,
  suggestionsAvailableAlertText: "Suggestions available",
  suggestionsContainerAriaLabel: "Suggested Records",
};

let query: any;
let queryPrimaryField: string;
let querySecondaryField: string;
let queryResultIdField: string;
let searchType: string;
let entityName: string;

export class PeoplePickerTypes extends BaseComponent<any, IPeoplePickerState> {
  // All pickers extend from BasePicker specifying the item type.
  private _picker = React.createRef<IBasePicker<IPersonaProps>>();

  constructor(props: IPeopleProps) {
    super(props);

    this.state = {
      currentPicker: 2,
      delayResults: false,
      peopleList: this.props.people,
      mostRecentlyUsed: [],
      currentSelectedItems: [],
      isPickerDisabled: false,
    };
    initializeIcons();
    this.handleChange = this.handleChange.bind(this);
  }

  public render() {
    return (
      <NormalPeoplePicker
        onResolveSuggestions={this._onFilterChanged}
        onEmptyInputFocus={this._returnMostRecentlyUsed}
        getTextFromItem={this._getTextFromItem}
        pickerSuggestionsProps={suggestionProps}
        className={"ms-PeoplePicker"}
        key={"normal"}
        onRemoveSuggestion={this._onRemoveSuggestion}
        onValidateInput={this._validateInput}
        removeButtonAriaLabel={"Remove"}
        defaultSelectedItems={this.props.preselectedpeople}
        onItemSelected={this._onItemSelected}
        inputProps={{
          onBlur: (ev: React.FocusEvent<HTMLInputElement>) => { },
          onFocus: (ev: React.FocusEvent<HTMLInputElement>) => { },
          "aria-label": "People Picker",
        }}
        componentRef={this._picker}
        onInputChange={this._onInputChange}
        resolveDelay={300}
        disabled={this.state.isPickerDisabled}
        onChange={() => this.handleChange()}
      />
    );
  }

  private handleChange() {
    if (this._picker !== undefined) {
      this.getPeopleRelatedVal();
    }
  }

  private _onItemSelected = (item: any): Promise<IPersonaProps> => {
    const processedItem = { ...item };
    processedItem.text = `${item.text}`;
    return new Promise<IPersonaProps>((resolve, reject) =>
      resolve(processedItem)
    );
  };

  private async getPeopleRelatedVal() {
    try {
      let tempPeople: IPeoplePersona[] = [];
      await Promise.all(
        this._picker.current!.items!.map((item) => {
          tempPeople.push({
            text: item.text,
            secondaryText: item.secondaryText,
            id: item.id,
            entity: item.entity,
          });
        })
      );
      if (this.props.peopleList) {
        this.props.peopleList(tempPeople);
      }
    } catch (err) {
      console.log(err);
    }
  }

  private _getTextFromItem(persona: IPersonaProps): string {
    return persona.text as string;
  }

  private _onRemoveSuggestion = (item: IPersonaProps): void => {
    const { peopleList, mostRecentlyUsed: mruState } = this.state;
    const indexPeopleList: number = peopleList.indexOf(item);
    const indexMostRecentlyUsed: number = mruState.indexOf(item);

    if (indexPeopleList >= 0) {
      const newPeople: IPersonaProps[] = peopleList
        .slice(0, indexPeopleList)
        .concat(peopleList.slice(indexPeopleList + 1));
      this.setState({ peopleList: newPeople });
    }

    if (indexMostRecentlyUsed >= 0) {
      const newSuggestedPeople: IPersonaProps[] = mruState
        .slice(0, indexMostRecentlyUsed)
        .concat(mruState.slice(indexMostRecentlyUsed + 1));
      this.setState({ mostRecentlyUsed: newSuggestedPeople });
    }
  };

  private _onFilterChanged = (
    filterText: string,
    currentPersonas: IPersonaProps[] | undefined,
    limitResults?: number
  ): any => {
    if (filterText) {
      if (filterText.length >= this.props.context.parameters.textFilterLength.raw!) {
        entityName = this.props.context.parameters.entityName.raw!;
        query = this.props.context.parameters.filterQuery.raw!;
        queryPrimaryField = this.props.context.parameters.queryPrimaryField.raw!;
        querySecondaryField = this.props.context.parameters.querySecondaryField.raw!;
        queryResultIdField = this.props.context.parameters.queryResultIdField.raw!;
        searchType = this.props.context.parameters.searchType.raw!;
        if (this.props.context.parameters.entityName.raw! == "account" && this.props.context.parameters.queryPrimaryField.raw! == null) {
          return this._searchAccounts(filterText);
        } else if (this.props.context.parameters.entityName.raw! == "contact") {
          return this._searchContacts(filterText);
        } else if (this.props.context.parameters.entityName.raw! == "systemuser") {
          return this._searchUsers(filterText);
        } else if (this.props.context.parameters.entityName.raw! == "customer") {
          return this._searchCustomers(filterText);
        } else {
          return this._searchEntity(filterText);
        }
      }
    } else {
      return [];
    }
  };

  private _searchUsers(
    filterText: string
  ): IPersonaProps[] | Promise<IPersonaProps[]> {
    return new Promise(async (resolve: any, reject: any) => {
      let People: IPersonaProps[] = [];

      var disabled = false;
      if (this.props.context.parameters.includeDisabled.raw! == 1) {
        disabled = true;
      }
      const config = new WebApiConfig("9.1");

      const usersQuery: string =
        "?$select=fullname,internalemailaddress,systemuserid&$filter=startswith(fullname,'" +
        filterText +
        "') and isdisabled eq " +
        disabled +
        " and " +
        this.props.context.parameters.filterQuery.raw! +
        "";
      retrieveMultiple(config, "systemusers", usersQuery).then(
        (results) => {
          const People: IPersonaProps[] = [];

          for (const recordUser of results.value) {
            People.push({
              text: recordUser.fullname,
              secondaryText: recordUser.internalemailaddress,
              id: recordUser.systemuserid,
              entity: "systemuser"
            });
            resolve(People);
          }

          const teamQuery: string =
            "?$select=name, description, teamid&$filter=startswith(name,'" +
            filterText +
            "')";
          retrieveMultiple(config, "teams", teamQuery).then(
            (resultsTeam) => {
              const People: IPersonaProps[] = [];

              for (const record of resultsTeam.value) {
                People.push({
                  text: record.name,
                  secondaryText: record.description,
                  id: record.teamid,
                  entity: "team"
                });
                resolve(People);
              }
            },
            (error) => {
              reject(People);
              console.log(error);
            }
          );
        },
        (error) => {
          reject(People);
          console.log(error);
        }
      );
    });
  }

  private _searchCustomers(
    filterText: string
  ): IPersonaProps[] | Promise<IPersonaProps[]> {
    return new Promise(async (resolve: any, reject: any) => {
      let Records: IPersonaProps[] = [];
      var disabled = 0;
      if (this.props.context.parameters.includeDisabled.raw! == 1) {
        disabled = 1;
      }
      const config = new WebApiConfig("9.1");

      const accountQuery: string = "?$select=name,df_tickerivp&$filter=startswith(name,'" + filterText + "') and statecode eq " + disabled + "";
      retrieveMultiple(config, "accounts", accountQuery).then(
        (results) => {
          const Records: IPersonaProps[] = [];

          for (const record of results.value) {
            Records.push({
              text: record.name,
              secondaryText: record.df_tickerivp,
              id: record.accountid,
              entity: "account"
            });
            resolve(Records);
          }

          const contactQuery: string = "?$select=fullname,emailaddress1&$filter=contains(fullname,'" + filterText + "') and statecode eq " + disabled + "";
          retrieveMultiple(config, "contacts", contactQuery).then(
            (resultsTeam) => {
              const Records: IPersonaProps[] = [];

              for (const record of resultsTeam.value) {
                Records.push({
                  text: record.fullname,
                  secondaryText: record.emailaddress1,
                  id: record.contactid,
                  entity: "contact"
                });
                resolve(Records);
              }
            },
            (error) => {
              reject(Records);
              console.log(error);
            }
          );
        },
        (error) => {
          reject(Records);
          console.log(error);
        }
      );
    });
  }

  private _searchAccounts(
    filterText: string
  ): IPersonaProps[] | Promise<IPersonaProps[]> {
    return new Promise(async (resolve: any, reject: any) => {
      let Account: IPersonaProps[] = [];
      const config = new WebApiConfig("9.1");
      var disabled = 0;
      if (this.props.context.parameters.includeDisabled.raw! == 1) {
        disabled = 1;
      }
      let accountQuery;
      //Check to see which type is supplied.
      if (this.props.context.parameters.filterQuery.raw! != null && this.props.context.parameters.searchType.raw! == "ticker") {
        accountQuery = "?$select=name,tickersymbol&$filter=startswith(tickersymbol, '" + filterText + "') and statecode eq " + disabled + " and " + this.props.context.parameters.filterQuery.raw! + "";
      } else if (this.props.context.parameters.filterQuery.raw! != null && this.props.context.parameters.searchType.raw! == null) {
        accountQuery = "?$select=name,tickersymbol&$filter=startswith(name,'" + filterText + "') and statecode eq " + disabled + " and " + this.props.context.parameters.filterQuery.raw! + "";
      } else if (this.props.context.parameters.searchType.raw! == "ticker" && this.props.context.parameters.filterQuery.raw! == null) {
        accountQuery = "?$select=name,tickersymbol&$filter=startswith(tickersymbol, '" + filterText + "') and statecode eq " + disabled + "";
      } else {
        accountQuery = "?$select=name,tickersymbol&$filter=startswith(name,'" + filterText + "') and statecode eq " + disabled + "";
      }
      retrieveMultiple(config, "accounts", accountQuery).then(
        (results) => {
          const Account: IPersonaProps[] = [];

          for (const account of results.value) {
            Account.push({
              text: account.name,
              secondaryText: account.tickersymbol,
              id: account.accountid,
              entity: "account"
            });
            resolve(Account);
          }
        },
        (error) => {
          reject(Account);
          console.log(error);
        });
    });

  }

  private _searchContacts(
    filterText: string
  ): IPersonaProps[] | Promise<IPersonaProps[]> {
    return new Promise(async (resolve: any, reject: any) => {
      let Contact: IPersonaProps[] = [];
      const config = new WebApiConfig("9.1");
      var disabled = 0;
      if (this.props.context.parameters.includeDisabled.raw! == 1) {
        disabled = 1;
      }
      let contactQuery;
      if (this.props.context.parameters.filterQuery.raw! != null) {
        contactQuery = "?$select=fullname,emailaddress1&$filter=startswith(fullname,'" + filterText + "') and statecode eq " + disabled + " and " + this.props.context.parameters.filterQuery.raw! + "";
      } else {
        contactQuery = "?$select=fullname,emailaddress1&$filter=startswith(fullname,'" + filterText + "') and statecode eq " + disabled + "";
      }
      retrieveMultiple(config, "contacts", contactQuery).then(
        (results) => {
          const Contact: IPersonaProps[] = [];

          for (const entity of results.value) {
            Contact.push({
              text: entity.fullname,
              secondaryText: entity.emailaddress1,
              id: entity.contactid,
              entity: "contact"
            });
            resolve(Contact);
          }
        },
        (error) => {
          reject(Contact);
          console.log(error);
        });
    });

  }

  private _searchEntity(filterText: string): IPersonaProps[] | Promise<IPersonaProps[]> {
    return new Promise(async (resolve: any, reject: any) => {
      let Entity: IPersonaProps[] = [];
      const queryEntity: string = query.replace('searchText', filterText);
      //if (!entityName.endsWith('s') && !entityName.includes('_')){
      //          entityName = entityName + 's';
      //      } 
      Xrm.WebApi.online.retrieveMultipleRecords(entityName, queryEntity).then(
        function success(results) {
          for (var i = 0; i < results.entities.length; i++) {
            let secondary: string;
            if (results.entities[i][querySecondaryField] == null) {
              secondary = "";
            }
            else {
              secondary = results.entities[i][querySecondaryField];
            }
            Entity.push({
              text: results.entities[i][queryPrimaryField],
              secondaryText: secondary,
              id: results.entities[i][queryResultIdField],
              entity: entityName
            });
            resolve(Entity);
          }
        },
        function (error) {
          reject(Entity);
          console.log(error);
        });
    });
  }

  //TODO: update logic/input - add columns to search/return along with entity name -

  private _returnMostRecentlyUsed = (
    currentPersonas: any
  ): IPersonaProps[] | Promise<IPersonaProps[]> => {
    let { mostRecentlyUsed } = this.state;
    mostRecentlyUsed = this._removeDuplicates(
      mostRecentlyUsed,
      currentPersonas
    );
    return this._filterPromise(mostRecentlyUsed);
  };

  private _filterPromise(
    personasToReturn: IPersonaProps[]
  ): IPersonaProps[] | Promise<IPersonaProps[]> {
    if (this.state.delayResults) {
      return this._convertResultsToPromise(personasToReturn);
    } else {
      return personasToReturn;
    }
  }

  private _listContainsPersona(
    persona: IPersonaProps,
    personas: IPersonaProps[]
  ) {
    if (!personas || !personas.length || personas.length === 0) {
      return false;
    }
    return personas.filter((item) => item.text === persona.text).length > 0;
  }

  private _filterPersonasByText(filterText: string): IPersonaProps[] {
    return this.state.peopleList.filter((item) =>
      this._doesTextStartWith(item.text as string, filterText)
    );
  }

  private _doesTextStartWith(text: string, filterText: string): boolean {
    return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
  }

  private _convertResultsToPromise(
    results: IPersonaProps[]
  ): Promise<IPersonaProps[]> {
    return new Promise<IPersonaProps[]>((resolve, reject) =>
      setTimeout(() => resolve(results), 2000)
    );
  }

  private _removeDuplicates(
    personas: IPersonaProps[],
    possibleDupes: IPersonaProps[]
  ) {
    return personas.filter(
      (persona) => !this._listContainsPersona(persona, possibleDupes)
    );
  }

  private _validateInput = (input: string): ValidationState => {
    if (input.indexOf("@") !== -1) {
      return ValidationState.valid;
    } else if (input.length > 1) {
      return ValidationState.warning;
    } else {
      return ValidationState.invalid;
    }
  };

  /**
   * Takes in the picker input and modifies it in whichever way
   * the caller wants, i.e. parsing entries copied from Outlook (sample
   * input: "Aaron Reid <aaron>").
   *
   * @param input The text entered into the picker.
   */
  private _onInputChange(input: string): string {
    const outlookRegEx = /<.*>/g;
    const emailAddress = outlookRegEx.exec(input);

    if (emailAddress && emailAddress[0]) {
      return emailAddress[0].substring(1, emailAddress[0].length - 1);
    }
    return input;
  }
}

class SystemUser {
  fullname: string;
  internalemailaddress: string;
  systemuserid: Guid;

  constructor(entity: any) {
    this.fullname = entity.fullname;
    this.internalemailaddress = entity.interalemailaddress;
    this.systemuserid = entity.systemuserid;
  }
}
