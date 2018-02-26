import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-webpart-base';

import * as strings from 'CswpWebPartStrings';
import Cswp from './components/Cswp';
import { ICswpProps } from './components/ICswpProps';
import pnp from "sp-pnp-js";
import * as DOMPurify from "dompurify";

export interface ICswpWebPartProps {
  query: string;
  itemTemplate: string;
  controlTemplate: string;
  cssStyles: string;
  maxNumberResults: number;
  noResultsTemplate: string;
}

export default class CswpWebPart extends BaseClientSideWebPart<ICswpWebPartProps> {

  protected onInit(): Promise<void> {

    return super.onInit().then(_ => {
      pnp.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    const element: React.ReactElement<ICswpProps > = React.createElement(
      Cswp,
      {
        query: this.properties.query,
        itemTemplate: this.properties.itemTemplate,
        controlTemplate: this.properties.controlTemplate,
        cssStyles: this.properties.cssStyles,
        maxNumberResults: this.properties.maxNumberResults,
        noResultsTemplate: this.properties.noResultsTemplate
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _validateHTML(value: string): string {
    var clean = DOMPurify.sanitize(value);
    if(clean != value){
      return "Please make sure your HTML is valid, use double quotes and remove any JavaScript.";
    } else return "";
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean { 
    return true; 
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
         /*   header: {
            description: strings.PropertyPaneDescription
          },  */
          groups: [
            {
              //groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('query', {
                  label: strings.QueryFieldLabel,
                  multiline: true
                }),
                PropertyPaneSlider('maxNumberResults', {
                  showValue: true,
                  min: 1,
                  max: 20,
                  label: strings.MaximumNumberResultsFieldLabel
                }),
                PropertyPaneTextField('itemTemplate', {
                  label: strings.ItemTemplateFieldLabel,
                  multiline: true,
                  validateOnFocusOut:  true,
                  onGetErrorMessage: this._validateHTML.bind(this) 
                }),
                PropertyPaneTextField('controlTemplate', {
                  label: strings.ControlTemplateFieldLabel,
                  multiline: true,
                  validateOnFocusOut:  true,
                  onGetErrorMessage: this._validateHTML.bind(this)
                }),
                PropertyPaneTextField('cssStyles', {
                  label: strings.CssStylesFieldLabel,
                  multiline: true,
                  validateOnFocusOut:  true
                }),
                PropertyPaneTextField('noResultsTemplate', {
                  label: strings.NoResultsTemplateFieldLabel,
                  multiline: true,
                  validateOnFocusOut:  true,
                  onGetErrorMessage: this._validateHTML.bind(this)
                })
              ]
            }
          ]
        }
      ]
    };
  }
} 
 