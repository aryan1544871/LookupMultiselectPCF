import * as React from "react";
import {
  Dropdown,
  IDropdownOption,
  IDropdownStyles,
  DropdownMenuItemType,
} from "@fluentui/react/lib/Dropdown";
//import { Stack, IStackTokens } from "@fluentui/react/lib/Stack";
import { IInputs } from "./generated/ManifestTypes";
import { IconButton, IButtonStyles } from "@fluentui/react/lib/Button";
import { Icon } from "@fluentui/react/lib/Icon";
import { SearchBox } from "@fluentui/react/lib/SearchBox";
import { FontSizes, ISearchBoxStyles, ITheme } from "@fluentui/react";
//import { Sticky } from "@fluentui/react";
import { Associate, DisAssociate } from "./WebApiOperations";

function GetAssociatedItem (primaryEntityId : string, relationshipName : string){

console.log(primaryEntityId);
console.log(relationshipName);
}

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: {
    width: "100%",
    selectors: {
      "&:focus": {
        borderColor: "#0078d4",
        boxShadow: "0 0 5px rgba(0, 120, 212, 0.5)",
      },
    },
  },
  root: {
    width: "100%",
  },
  callout: {
    maxHeight: "50vh",
    overflowY: "auto",
  },
  dropdownItemsWrapper: {
    maxHeight: "inherit",
  },
  title: {
    borderColor: "#666666",
    selectors: {
      "&:hover": {
        borderColor: "#333333",
      },
    },
  },
};

const searchBoxStyles: Partial<ISearchBoxStyles> = {
  clearButton: { display: "none" },
};
//const stackTokens: IStackTokens = { childrenGap: 10 };
const buttonStyles: IButtonStyles = { icon: { fontSize: "11px" } };

export interface ILookupMultiSel {
  onChange: (selectedValues: string[]) => void;
  initialValues: string[];
  context: ComponentFramework.Context<IInputs>;
  relatedEntityType: string;
  relatedPrimaryColumns: string[];
  relatedPrimaryColumnsName: string[];
  primaryEntityType: string;
  relationshipName: string;
  primaryEntityId: string;
  disabled: boolean;
  primaryEntityName: string;
  primaryFilterColumn :ComponentFramework.LookupValue[];
  mappedEntityAndColumnForFilter : string[];
}

export const LookupMultiSel = React.memo((props: ILookupMultiSel) => {
  const {
    onChange,
    initialValues,
    context,
    relatedEntityType,
    relatedPrimaryColumns,
     relatedPrimaryColumnsName,
    primaryEntityType,
    relationshipName,
    primaryEntityId,
    disabled,
    primaryEntityName,
    primaryFilterColumn,
    mappedEntityAndColumnForFilter,
  } = props;
  const [selectedValues, setSelectedValues] = React.useState<string[]>([]);
  const [userOptions, setUserOptions] = React.useState<IDropdownOption[]>([]);
  const [associatedRecords, setAssociatedRecords] = React.useState<string[]>([]);
  const onChangeTriggered = React.useRef(false);
  const [searchText, setSearchText] = React.useState<string>("");
  const onChangePrimaryFilterColumn = React.useRef(false);

  /**
   * Gets selected values from props and maintain using state
   */
  React.useEffect(() => {
    setSelectedValues(initialValues);
  }, []);

  /**
   * Retrieves entity records using webapi and maintain using state
   */
  React.useEffect(() => {
    let userOptionsList: IDropdownOption[] = [];
    let associatedRecordLists : string[] = [];
    let associatedString = primaryEntityId? `?$expand=${relationshipName}($select=${ relatedPrimaryColumnsName[1]})&$filter=${primaryEntityName} eq ${primaryEntityId}`:null;
    let primaryFilterColumnValue  = (primaryFilterColumn?primaryFilterColumn[0]:null as any)?.Id?._rawGuid;
    let formattedPrimaryFilterColumnValue = primaryFilterColumnValue?`${primaryFilterColumnValue.slice(0, 8)}-${primaryFilterColumnValue.slice(8, 12)}-${primaryFilterColumnValue.slice(12, 16)}-${primaryFilterColumnValue.slice(16, 20)}-${primaryFilterColumnValue.slice(20)}`:null;
    if (mappedEntityAndColumnForFilter.length === 3){
      let filteredOptionsets : string[] = [];
      let filterGuid = formattedPrimaryFilterColumnValue;
      if (filterGuid!= null){
         context.webAPI
          .retrieveMultipleRecords(mappedEntityAndColumnForFilter[0])
          .then((response) => {
            response.entities.forEach((element) => {
            if(element[`_${mappedEntityAndColumnForFilter[1]}_value`] === filterGuid){
               filteredOptionsets.push(
                  element[`_${mappedEntityAndColumnForFilter[2]}_value`]
              );
            }
            });
            context.webAPI
            .retrieveMultipleRecords(relatedEntityType)
            .then((response) => {
              response.entities.forEach((element) => {
              if(filteredOptionsets.includes(element[relatedPrimaryColumns[0]])){
                  userOptionsList.push({
                  key: element[relatedPrimaryColumns[0]],
                  text: element[relatedPrimaryColumns[1]],
                  data: { value: element[relatedPrimaryColumns[0]] },
                });
              }
              });
              userOptionsList.sort((a,b)=> a.text.localeCompare(b.text));
              setUserOptions(userOptionsList);
            })
        })
      }
      else{
        setUserOptions(userOptionsList);
      }
    }
    else{
      context.webAPI
      .retrieveMultipleRecords(relatedEntityType)
      .then((response) => {
        response.entities.forEach((element) => {
          userOptionsList.push({
            key: element[relatedPrimaryColumns[0]],
            text: element[relatedPrimaryColumns[1]],
            data: { value: element[relatedPrimaryColumns[0]] },
          });
        });
        setUserOptions(userOptionsList);
      })
    }
    if(associatedString){
      context.webAPI
      .retrieveMultipleRecords(primaryEntityType,associatedString)
      .then((response)=>{
        response.entities.forEach((element) => {
            associatedRecordLists.push(
              element[relationshipName]
            );
          });
          setAssociatedRecords(associatedRecordLists);
      })
    }
  }, [primaryFilterColumn]);

  /**
   * Trigger onchange to update the property
   */
  React.useEffect(() => {
    if (onChangeTriggered.current) onChange(selectedValues);
  }, [selectedValues]);


  React.useEffect(() => {
    if (onChangePrimaryFilterColumn.current && primaryFilterColumn.length === 0 ) {
      setSelectedValues([]); 
    } else {
      onChangePrimaryFilterColumn.current = true; // mark as mounted
    }
  }, [primaryFilterColumn]);
  
  /**
   * Triggers on change of dropdown
   * @param ev Event of the dropdown
   * @param option Selected option from dropdown
   * @param eventId Event to identify is it for dropdown or cancel icon
   */
  const onChangeDropDownOrOnIconClick = (
    ev: unknown,
    option?: IDropdownOption,
    eventId?: number
  ) => {
    if (eventId === 1) {
      let iconEvent = ev as React.MouseEvent<HTMLButtonElement>;
      iconEvent.stopPropagation() ;
    }

    if (option) {
      onChangeTriggered.current = true;
      setSelectedValues(
        option.selected
          ? [...selectedValues, option.key as string]
          : selectedValues.filter((key) => key != option.key)
      );
    }
  /*
    if (option?.selected ){
     Associate(
            context,
            option.key,
            primaryEntityType,
            relatedEntityType,
            relationshipName,
            primaryEntityId
          );
    }
    else if (!option?.selected){
      DisAssociate(
                context,
                option?.key!,
                primaryEntityType,
                relationshipName,
                primaryEntityId
              );
    } 
              */
  };
  /**
   *Render icon of the dropdown to search
   * @returns Icon
   */
  const onRenderCaretDown = () => {
    let associatedItems: any[] = [];
    let associatedItemsArray :any[] = [];
    let recordsToBeDissociated: any[] = [];
    if (selectedValues.length === 0){

       associatedRecords.forEach(n=>{
        for (var i=0 ; i < n.length; i++ ){
          associatedItems.push(n[i])
            }
           })
        associatedItems.forEach((element) => {
          associatedItemsArray.push(element[relatedPrimaryColumnsName[0]])
        });
        associatedItemsArray.forEach((element)=>{
                recordsToBeDissociated.push(element)
            })
        recordsToBeDissociated.forEach((key)=>{
            DisAssociate(
                    context,
                    key,
                    primaryEntityType,
                    relationshipName,
                    primaryEntityId
                  );
          })
    }
    return <Icon iconName="Search"></Icon>;
  };

  /**
   * Render drop down item event
   * @param option Drop down item
   * @returns
   */
  const onRenderOption = (option?: IDropdownOption) => {
    return option?.itemType === DropdownMenuItemType.Header &&
      option.key === "FilterHeader" ? (
      <SearchBox
        onChange={(ev, newValue?: string) => setSearchText(newValue!)}
        underlined={true}
        placeholder="Search options"
        autoFocus={true}
        styles={searchBoxStyles}
      ></SearchBox>
    ) : (
      <>{option?.text}</>
    );
  };

  /**
   * Render custom title
   * @param options Selected option from dropdown
   * @returns
   */
  const onRenderTitle = (options: any) => {
    let option: any[] = [];
    let selectedList: IDropdownOption[] = options;
    let selectedListArray: any[] =[];
    let associatedItems: any[] = [];
    let associatedItemsArray :any[] = [];
    let recordsToBeAssociated: any[] = [];
    let recordsToBeDissociated: any[] = [];
    selectedList.forEach((element)=>{
      selectedListArray.push(element.key)
    });

    associatedRecords.forEach(n=>{
      for (var i=0 ; i < n.length; i++ ){
        associatedItems.push(n[i])
      }
    })
    associatedItems.forEach((element) => {
      associatedItemsArray.push(element[relatedPrimaryColumnsName[0]])
    });
    selectedListArray.forEach((element)=>{
      if(!associatedItemsArray.includes(element)){
        recordsToBeAssociated.push(element)
      }
    })
     associatedItemsArray.forEach((element)=>{
      if(!selectedListArray.includes(element)){
        recordsToBeDissociated.push(element)
      }
    })
    recordsToBeAssociated.forEach((key) =>{
      Associate(
              context,
              key,
              primaryEntityType,
              relatedEntityType,
              relationshipName,
              primaryEntityId
            );
    } )
    recordsToBeDissociated.forEach((key)=>{
      DisAssociate(
              context,
              key,
              primaryEntityType,
              relationshipName,
              primaryEntityId
            );
    })
    //let url: string = `main.aspx?pagetype=entityrecord&etn=${entityType}&id=`;
    selectedList.forEach((element) => {
      option.push(
        <span>
          {element.text}
          <IconButton
            iconProps={{ iconName: "Cancel" }}
            title={element.text}
            onClick={(ev) => onChangeDropDownOrOnIconClick(ev, element, 1)}
            className="IconButtonClass"
            styles={buttonStyles}
          ></IconButton>
        </span>
      );
    });
    return <div>{option}</div>;
  };

  return (
    <>
      {/* <Stack horizontal tokens={stackTokens}> */}
      <Dropdown
        {...userOptions}
        options={[
          {
            key: "FilterHeader",
            text: "-",
            itemType: DropdownMenuItemType.Header,
          },
          {
            key: "divider_filterHeader",
            text: "-",
            itemType: DropdownMenuItemType.Divider,
          },
          ...userOptions.filter(
            (opt) =>
              opt.text
                .toLocaleLowerCase()
                .indexOf(searchText.toLocaleLowerCase()) > -1
          ),
        ]}
        styles={dropdownStyles}
        multiSelect={true}
        onChange={onChangeDropDownOrOnIconClick}
        selectedKeys={selectedValues}
        calloutProps={{ directionalHintFixed: true }}
        onRenderTitle={onRenderTitle}
        //dropdownWidth="auto"
        id="MainDropDown"
        placeholder="Look for records"
        onRenderCaretDown={onRenderCaretDown}
        onRenderOption={onRenderOption}
        onDismiss={() => setSearchText("")}
        disabled={disabled}
      />
      {/* </Stack> */}
    </>
  );
});

LookupMultiSel.displayName = "LookupMultiSel";
