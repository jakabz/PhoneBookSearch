import * as React from 'react';
import styles from './PhoneBookSearch.module.scss';
import { IPhoneBookSearchProps } from './IPhoneBookSearchProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import {
  IPersonaSharedProps,
  Persona,
  PersonaSize
} from 'office-ui-fabric-react/lib/Persona';
import {
  TooltipHost,
  TooltipDelay,
  DirectionalHint
} from 'office-ui-fabric-react/lib/Tooltip';
import { Icon } from 'office-ui-fabric-react/lib/Icon';

export interface INotificationItemsState {
  searchResults: any[];
}

export default class PhoneBookSearch extends React.Component<IPhoneBookSearchProps, INotificationItemsState> {  
  constructor(props: IPhoneBookSearchProps) {
    super(props);
    this.state = {
      searchResults: []
    };
  }

  public _search = (text: string) => {
    let result: any = [];
    if(text.length > 2) {
      const fields: any = ['Title', 'Email', 'Phone', 'Position', 'Organization'];
      this.props.listData.map((item) => { // Összes listaelemen végigmegyünk
        let addedItem: boolean = false;
        fields.map(async (field) => { // Végigmegyünk azokon a fildeken amikben keresünk
          if (item[field] && !addedItem) { //Csak akkor vizsgálunk ha van a adott fieldben adat és még nem lett felvéve a találatok közé
            if (item[field].toUpperCase().indexOf(text.toUpperCase()) !== -1) { //Az adott filben szerepel a keresőszó
              if(item.Attachments) {
                let url = await this._getImageUrl(item.Id).then((res) => res.value[0].ServerRelativeUrl);
                item.AttachmentsImage = url; 
              } else {
                item.AttachmentsImage = null;
              }
              if(item[field] && !addedItem){
                result.push(item);
              }
              addedItem = true;
              this.setState({searchResults: result});
            }
          }
        });
      });
      
    } else {
      result = [];
      this.setState({searchResults: result});
    }
  }  

  public _getImageUrl = (id: number):Promise<any> => {
    return this.props.context.spHttpClient.get(this.props.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('PhoneBook')/Items(${id})/AttachmentFiles`, 
      SPHttpClient.configurations.v1).then((response) => {
        return response.json();
      });
  }

  public render(): React.ReactElement<IPhoneBookSearchProps> {
    console.info(this.state.searchResults);

    return (
      <div className={styles.phoneBookSearch}>
        { 
          this.props.wpTitle && <div className={styles.wpTitle}>{this.props.wpTitle}</div>
        }
        <div></div>
        <SearchBox
          placeholder='Személyek keresése'
          onChange={ (text) => this._search(text) }
        />
        <div className={styles.searchResult}>
          {
            this.state.searchResults.map((result,i) => {
              const examplePersona: IPersonaSharedProps = {
                imageUrl: result.AttachmentsImage,
                imageInitials: result.Title.split(' ')[0].slice(0,1)+result.Title.split(' ')[1].slice(0,1),
                text: result.Title,
                secondaryText: result.Position,
                tertiaryText: result.Email,
                optionalText: result.Phone
              };
              const mailLink: string = 'mailto:' + result.Email;
              const phoneLink: string = 'tel:' + result.Phone;
              let position: DirectionalHint;
              switch(this.props.position) {
                case 'topCenter':
                  position = DirectionalHint.topCenter;
                  break;
                case 'bottomCenter':
                  position = DirectionalHint.bottomCenter;
                  break;
                case 'leftCenter':
                  position = DirectionalHint.leftCenter;
                  break;
                default:
                  position = DirectionalHint.rightCenter;
              }
              return ( <div className={styles.persona}>
                        <TooltipHost 
                          tooltipProps={{onRenderContent: () => { 
                              return (
                                <div className={styles.tooltipContent}>
                                  <div><Icon iconName="Contact" /> <strong>{result.Title}</strong></div>
                                  {result.Position && <div><Icon iconName="Work" /> <i>{result.Position}</i></div>}
                                  {result.Organization && <div><Icon iconName="Org" /> {result.Organization}</div>}
                                  {result.Email && <div><Icon iconName="Mail" /> <a href={mailLink}>{result.Email}</a></div>}
                                  {result.Phone && <div><Icon iconName="Phone" /> <a href={phoneLink}>{result.Phone}</a></div>}
                                </div>
                              ) 
                            } 
                          }}
                          id={'tooltip-'+i} 
                          calloutProps={ { gapSpace: 0 } } 
                          closeDelay={ 500 }
                          directionalHint={position}
                          className={ styles.tooltip }
                        >
                          <Persona
                            {...examplePersona}
                          />
                        </TooltipHost>
                        </div>
                      );
            })
          }
        </div>
      </div>
    );
  }
}
