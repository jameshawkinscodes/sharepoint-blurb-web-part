import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'BlurbWebPartStrings';
import { Blurb } from './components/Blurb';
import { IBlurbProps } from './components/IBlurbProps';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { PropertyFieldIconPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldIconPicker';
import { initializeIcons } from '@fluentui/react/lib/Icons';

export interface IBlurbWebPartProps {
  description: string;
  containerCount: number;
  key:string;
  containers: Array<{
    text: string;
    icon: string;
    backgroundColor: string;
    borderColor: string;
    title: string; 
  }>;
}

export default class BlurbWebPart extends BaseClientSideWebPart<IBlurbWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private selectedContainerIndex: number = -1;  // Store the selected container index

  public render(): void {
    console.log('Containers:', this.properties.containers);
    const element: React.ReactElement<IBlurbProps> = React.createElement(
      Blurb,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        containers: this.properties.containers || [],
        containerCount: this.properties.containerCount || 1,
        onContainerClick: (index: number) => {
          console.log('Container clicked:', index); // Verify this log
          this.selectedContainerIndex = index;
          this._isEditMode = true;
          this.context.propertyPane.refresh();
          this.context.propertyPane.open();
        }        
      }
    );

    ReactDom.render(element, this.domElement);
  }

  // Ensure that we unmount the component when disposing of the web part
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected async onInit(): Promise<void> {
    initializeIcons();
  
    console.log('Initializing containers'); // Debug log
  
    if (!this.properties.containerCount) {
      this.properties.containerCount = 1; // Default to 1 if not set
    }
  
    if (!this.properties.containers || this.properties.containers.length === 0) {
      this.properties.containers = [];
      for (let i = 0; i < this.properties.containerCount; i++) {
        this.properties.containers.push({
          icon: '',
          backgroundColor: '#ffffff',
          borderColor: '#000000',
          title: `Blurb ${i + 1}`,
          text: 'Add text'
        });
      }
    }
  
    console.log('Containers after initialization:', this.properties.containers); // Debug log
  
    const message = await this._getEnvironmentMessage();
    this._environmentMessage = message;
  
    return await super.onInit();
  }  

  protected onPropertyPaneConfigurationComplete(): void {
    this._isEditMode = false;
    this.selectedContainerIndex = -1;
  }

  private async _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      const context = await this.context.sdks.microsoftTeams.teamsJs.app.getContext();
      let environmentMessage: string = '';
      switch (context.app.host.name) {
        case 'Office':
          environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
          break;
        case 'Outlook':
          environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
          break;
        case 'Teams':
          environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
          break;
        default:
          environmentMessage = strings.UnknownEnvironment;
      }
      return environmentMessage;
    }
  
    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }
  
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'containerCount' && newValue !== oldValue) {
      this.properties.containers = [];
      for (let i = 0; i < newValue; i++) {
        this.properties.containers.push({
          icon: 'Contact',
          backgroundColor: '#ffffff',
          borderColor: '#000000',
          title: `Blurb ${i + 1}`,
          text: 'Add text'
        });
      }
    }
  
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.render();
  }

  private _isEditMode: boolean = false;
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    if (!this._isEditMode || this.selectedContainerIndex === -1) {
      return {
        pages: [
          {
            header: { description: "Select a Blurb to configure its properties" },
            groups: [
              {
                groupName: "Settings",
                groupFields: [
                  PropertyPaneTextField('description', {
                    label: "Description",
                    value: this.properties.description,
                  }),
                  PropertyPaneSlider('containerCount', {
                    label: "Number of Blurbs",
                    min: 1,
                    max: 4,
                    step: 1,
                    value: this.properties.containerCount,
                    showValue: true,
                  }),
                ]
              }
            ]
          }
        ]
      };
    }
  
    // Ensure containers array is initialized
    if (!this.properties.containers || this.properties.containers.length < this.properties.containerCount) {
      this.properties.containers = [];
      for (let i = 0; i < this.properties.containerCount; i++) {
        this.properties.containers.push({
          icon: '',
          backgroundColor: '#ffffff',
          borderColor: '#000000',
          title: `Container ${i + 1}`,
          text: ''
        });
      }
    }
  
    // Get the selected container based on the selected index
    const selectedContainer = this.properties.containers[this.selectedContainerIndex] || {};
    return {
      pages: [
        {
          header: {
            description: `Configure Blurb ${this.selectedContainerIndex + 1}`
          },
          groups: [
            {
              groupName: "Blurb Settings",
              groupFields: [
                PropertyFieldIconPicker(`containers[${this.selectedContainerIndex}].icon`, {
                  label: "Icon",
                  currentIcon: this.properties.containers[this.selectedContainerIndex].icon || '',
                  onSave: (iconName: string) => {
                    this.onPropertyPaneFieldChanged(
                      `containers[${this.selectedContainerIndex}].icon`,
                      this.properties.containers[this.selectedContainerIndex].icon,
                      iconName
                    );
                    this.properties.containers[this.selectedContainerIndex].icon = iconName;
                    this.render(); // Re-render after selecting the icon
                  },
                  buttonLabel: "Icon",
                  renderOption: "panel", // Use "panel" or "dialog"
                  key: `iconPicker-${this.selectedContainerIndex}`,
                  properties: this.properties,
                  onPropertyChange: function (propertyPath: string, oldValue: string, newValue: string): void {
                    throw new Error('Function not implemented.');
                  }
                }),
                
                PropertyPaneTextField(`containers[${this.selectedContainerIndex}].title`, {
                  label: `Blurb Heading ${this.selectedContainerIndex + 1}`,
                  value: selectedContainer.title || ''
                }),
                PropertyFieldColorPicker(`containers[${this.selectedContainerIndex}].backgroundColor`, {
                  label: "Background Color",
                  selectedColor: selectedContainer.backgroundColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  showPreview: true,
                  key: `backgroundColor-${this.selectedContainerIndex}`
                }),
                PropertyFieldColorPicker(`containers[${this.selectedContainerIndex}].borderColor`, {
                  label: "Border Color",
                  selectedColor: selectedContainer.borderColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  showPreview: true,
                  key: `borderColor-${this.selectedContainerIndex}`
                }),
                PropertyPaneTextField(`containers[${this.selectedContainerIndex}].text`, {
                  label: `Blurb Text ${this.selectedContainerIndex + 1}`,
                  value: selectedContainer.text || ''
                })
              ]
            }
          ]
        }
      ]
    };
  }
}  