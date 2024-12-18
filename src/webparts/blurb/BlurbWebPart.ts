import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneSlider, PropertyPaneDropdown } from '@microsoft/sp-property-pane';
import * as strings from 'BlurbWebPartStrings';
import { Blurb } from './components/Blurb';
import { IBlurbProps } from './components/IBlurbProps';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { PropertyFieldIconPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldIconPicker';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import { DisplayMode } from '@microsoft/sp-core-library';

export interface IBlurbWebPartProps {
  description: string;
  containerCount: number;
  containers: Array<{
    fontColor: string;
    icon: string;
    backgroundColor: string;
    borderColor: string;
    borderRadius: string;
    title: string;
    text: string;
    linkUrl?: string;
    linkTarget?: string;
  }>;
}
export default class BlurbWebPart extends BaseClientSideWebPart<IBlurbWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private selectedContainerIndex: number = -1;
  private _isEditMode: boolean = false;

  public render(): void {
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
        isEditMode: this.displayMode === DisplayMode.Edit,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        isFullWidth: (this.context.manifest as any).supportedHosts?.includes('SharePointFullPage'),
        displayMode: this.displayMode, // Pass display mode
        onContainerClick: async (index: number) => {
          if (this.displayMode === DisplayMode.Edit) {
            // Avoid closing and reopening the property pane if the same container is clicked
            if (this.selectedContainerIndex === index && this.context.propertyPane.isRenderedByWebPart()) {
              return;
            }
        
            if (this.context.propertyPane.isRenderedByWebPart()) {
              this.context.propertyPane.close();
            }
        
            // Add a small delay to ensure smooth UI transitions
            await new Promise(resolve => setTimeout(resolve, 10));
            
            // Set the selected container index and update edit mode
            this.selectedContainerIndex = index;
            this._isEditMode = true;
        
            // Refresh and open the property pane for the selected container
            this.context.propertyPane.refresh();
            this.context.propertyPane.open();
          }
        },
        onEditClick: async (index: number) => {
          if (this.displayMode === DisplayMode.Edit) {
            // Avoid redundant updates if the same container is already being edited
            if (this.selectedContainerIndex === index && this.context.propertyPane.isRenderedByWebPart()) {
              return;
            }
        
            // Update the selected container index and edit mode
            this.selectedContainerIndex = index;
            this._isEditMode = true;
        
            // Close the property pane if it is already rendered
            if (this.context.propertyPane.isRenderedByWebPart()) {
              this.context.propertyPane.close();
            }
        
            // Add a slight delay for smoother transitions
            await new Promise(resolve => setTimeout(resolve, 10));
        
            // Refresh and open the property pane for the selected container
            this.context.propertyPane.refresh();
            this.context.propertyPane.open();
          }
        
        },
        onMoveClick: (index: number, direction: 'up' | 'down') => {
          if (this.displayMode === DisplayMode.Edit) {
            if (direction === 'up' && index > 0) {
              const temp = this.properties.containers[index];
              this.properties.containers[index] = this.properties.containers[index - 1];
              this.properties.containers[index - 1] = temp;
            } else if (direction === 'down' && index < this.properties.containers.length - 1) {
              const temp = this.properties.containers[index];
              this.properties.containers[index] = this.properties.containers[index + 1];
              this.properties.containers[index + 1] = temp;
            }
            this.render();
          }
        },
        onRemoveClick: (index: number, updatedCount: number) => {
          if (this.displayMode === DisplayMode.Edit) {
            this.properties.containers.splice(index, 1);
            this.properties.containerCount = updatedCount;
            this.render();
          }
        },
      }
    );
  
    ReactDom.render(element, this.domElement);
  }
  

  protected async onInit(): Promise<void> {
    initializeIcons();
    if (!this.properties.containerCount) {
      this.properties.containerCount = 1;
    }

    if (!this.properties.containers) {
      this.properties.containers = [];
    }

    const currentContainerCount = this.properties.containers.length;
    if (this.properties.containerCount > currentContainerCount) {
      for (let i = currentContainerCount; i < this.properties.containerCount; i++) {
        this.properties.containers.push({
          icon: 'CannedChat',
          backgroundColor: '#FAF9F8',
          borderColor: '#EDEBE9',
          borderRadius: '0',
          fontColor: '#323130',
          title: ``,
          text: '',
          linkUrl: '',
          linkTarget: '_self',
        });
      }
    } else if (this.properties.containerCount < currentContainerCount) {
      this.properties.containers.splice(this.properties.containerCount);
    }

    const message = await this._getEnvironmentMessage();
    this._environmentMessage = message;

    return super.onInit();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: string | number, newValue: string | number): void {
    if (propertyPath === 'containerCount' && newValue !== oldValue) {
      const newContainerCount = newValue as number;
      const currentContainerCount = this.properties.containers.length;

      if (newContainerCount > currentContainerCount) {
        for (let i = currentContainerCount; i < newContainerCount; i++) {
          this.properties.containers.push({
            icon: 'CannedChat',
            backgroundColor: '#FAF9F8',
            borderColor: '#EDEBE9',
            borderRadius: '0',
            fontColor: '#323130',
            title: ``,
            text: '',
            linkUrl: '',
            linkTarget: '_self',
          });
        }
      } else if (newContainerCount < currentContainerCount) {
        this.properties.containers.splice(newContainerCount);
      }
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.render();
  }
  // The blurb properties pane
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    if (this._isEditMode && this.selectedContainerIndex !== -1) {
      const selectedContainer = this.properties.containers[this.selectedContainerIndex] || {};
      return {
        pages: [
          {
            header: { description: `Configure Blurb ${this.selectedContainerIndex + 1}` },
            groups: [
              {
                groupName: "Blurb Settings",
                groupFields: [
                  PropertyFieldIconPicker(`containers[${this.selectedContainerIndex}].icon`, {
                    label: "Select Icon",
                    currentIcon: selectedContainer.icon,
                    onSave: (iconName: string) => {
                      this.properties.containers[this.selectedContainerIndex].icon = iconName;
                      this.render();
                    },
                    properties: this.properties,
                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                    buttonLabel: "Select Icon",
                    renderOption: "panel",
                    key: `iconPicker-${this.selectedContainerIndex}`
                  }),
                  PropertyPaneTextField(`containers[${this.selectedContainerIndex}].title`, {
                    label: `Blurb Title ${this.selectedContainerIndex + 1}`,
                    value: selectedContainer.title || ''
                  }),
                  PropertyPaneTextField(`containers[${this.selectedContainerIndex}].text`, {
                    label: `Blurb Text ${this.selectedContainerIndex + 1}`,
                    value: selectedContainer.text || '',
                    multiline: true,
                    resizable: true
                  }),
                  PropertyPaneTextField(`containers[${this.selectedContainerIndex}].linkUrl`, {
                    label: `Blurb Link URL ${this.selectedContainerIndex + 1}`,
                    value: selectedContainer.linkUrl || '',
                    placeholder: "Enter a clickable link URL",
                  }),
                  PropertyPaneDropdown(`containers[${this.selectedContainerIndex}].linkTarget`, {
                    label: `Link Target ${this.selectedContainerIndex + 1}`,
                    options: [
                      { key: '_self', text: 'Open in same tab' },
                      { key: '_blank', text: 'Open in new tab' }
                    ],
                    selectedKey: selectedContainer.linkTarget || '_self',
                  }),                  
                  PropertyFieldColorPicker(`containers[${this.selectedContainerIndex}].fontColor`, {
                    label: "Font Color",
                    selectedColor: selectedContainer.fontColor,
                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                    properties: this.properties,
                    style: PropertyFieldColorPickerStyle.Inline,
                    showPreview: true,
                    key: `fontColor-${this.selectedContainerIndex}`
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
                  PropertyPaneSlider(`containers[${this.selectedContainerIndex}].borderRadius`, {
                    label: "Border Radius",
                    min: 0,
                    max: 50,
                    step: 1,
                    value: selectedContainer.borderRadius ? parseInt(selectedContainer.borderRadius, 10) : 0,
                    showValue: true,
                  }),
                ]
              }
            ]
          }
        ]
      };
    }
    // The main web part properties pane
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
                  max: 10,
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
}
