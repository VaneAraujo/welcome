import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'WelcomeWebPartStrings';
import styles from './WelcomeWebPart.module.scss';

export interface IWelcomeWebPartProps {
  title: string;
  messagestyle: string;
  textalignment: string;
  showtimebasedmessage: boolean;
  morningmessage: string;
  afternoonmessage: string;
  afternoonbegintime: number;
  eveningmessage: string;
  eveningbegintime: number;
  message: string;
  showname: string;
  showfirstname: boolean;
}

export default class WelcomeWebPart extends BaseClientSideWebPart<IWelcomeWebPartProps> {

 

 public render(): void {
    
    let message = this.properties.message;

    if (this.properties.showtimebasedmessage) {
      const today: Date = new Date();
      if (today.getHours() >= this.properties.eveningbegintime) {
        message = this.properties.eveningmessage;
      }
      if (
        today.getHours() >= this.properties.afternoonbegintime &&
        today.getHours() <= this.properties.eveningbegintime
      ) {
        message = this.properties.afternoonmessage;
      }
      if (today.getHours() < this.properties.afternoonbegintime) {
        message = this.properties.morningmessage;
      }
    }

    const nameparts = this.context.pageContext.user.displayName.split(" ");

    let name = "";

    switch (this.properties.showname) {
      case "full": {
        name = this.context.pageContext.user.displayName;
        break;
      }
      case "first": {
        name = nameparts[0];
        break;
      }
    }


    const textalign =
      this.properties.textalignment === "left"
        ? styles.left
        : this.properties.textalignment === "right"
        ? styles.right
        : styles.center;
//
        const messagecontent = `<${this.properties.messagestyle}  class=${textalign}> ${message} ${name}</${this.properties.messagestyle}>`;
   
        this.domElement.innerHTML = `
          <div> 
          ${messagecontent}
          </div>`;
  }


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    
    const messagefields = [];

    if (this.properties.showtimebasedmessage) {
      messagefields.push(
        PropertyPaneTextField("morningmessage", {
          label: strings.MorningMessageLabel,
        })
      );
      messagefields.push(
        PropertyPaneTextField("afternoonmessage", {
          label: strings.AfternoonMessageLabel,
        })
      );

      messagefields.push(
        PropertyPaneTextField("eveningmessage", {
          label: strings.EveningMessageLabel,
        })
      );
      messagefields.push(
        PropertyPaneSlider("afternoonbegintime", {
          label: strings.AfternoonBeginTimeLabel,
          min: 11,
          max: 14,
        })
      );
      messagefields.push(
        PropertyPaneSlider("eveningbegintime", {
          label: strings.EveningBeginTimeLabel,
          min: 16,
          max: 19,
        })
      );
    } else {
      messagefields.push(
        PropertyPaneTextField("message", {
          label: strings.MessageLabel,
        })
      );
    }

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.NamePropertiesGroupName,
              groupFields: [
                PropertyPaneTextField("title", {
                  label: strings.TitleLabel,
                }),
                PropertyPaneChoiceGroup("showname", {
                  label: strings.ShowNameLabel,
                  options: [
                    {
                      key: "full",
                      text: "Full Name",
                    },
                    {
                      key: "first",
                      text: "First name only",
                    },
                    {
                      key: "none",
                      text: "No name",
                    },
                  ],
                }),
                PropertyPaneToggle("showtimebasedmessage", {
                  label: strings.ShowTimeBasedMessageLabel,
                }),
                ...messagefields,
              ],
            },
            {
              groupName: strings.StylePropertiesGroupName,
              groupFields: [
                //Font size
                PropertyPaneChoiceGroup("messagestyle", {
                  label: strings.MessageStyleLabel,
                  options: [
                    {
                      key: "h1",
                      text: "Extra-Large",
                    },
                    {
                      key: "h2",
                      text: "Large",
                    },
                    {
                      key: "h3",
                      text: "Medium",
                    },
                    {
                      key: "h4",
                      text: "Small",
                    },
                    {
                      key: "p",
                      text: "Extra-Small",
                    },
                  ],
                }),
                //Color Picker
                
               //Text Aling
                PropertyPaneChoiceGroup("textalignment", {
                  label: strings.TextAlignmentLabel,
                  options: [
                    {
                      key: "left",
                      text: "Left",
                      iconProps: {
                        officeFabricIconFontName: "AlignLeft",
                      },
                    },
                    {
                      key: "centre",
                      text: "Center",
                      iconProps: {
                        officeFabricIconFontName: "AlignCenter",
                      },
                    },
                    {
                      key: "right",
                      text: "Right",
                      iconProps: {
                        officeFabricIconFontName: "AlignRight",
                      },
                    },
                  ],
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}

