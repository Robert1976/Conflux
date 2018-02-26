import * as React from 'react';
import styles from './Cswp.module.scss';
import { ICswpProps } from './ICswpProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {ICswpState} from './CswpState';
import pnp, { SearchQuery, SearchResults, SearchQueryBuilder, SearchResult } from "sp-pnp-js";
import { HtmlParser } from "./tools/HtmlParser";
import { loadStyles } from '@microsoft/load-themed-styles';

export default class Cswp extends React.Component<ICswpProps, ICswpState> {

  constructor(props:ICswpProps) {
    super(props);
    this.state = {
      htmlResult: null,
      error: false,
      errorMessage: null,
      cssStyles: null
    };

    this._searchSharePoint.bind(this);
    this._generateHtml.bind(this);
  }

  public componentDidUpdate(prevProps, prevState): void {
    if (JSON.stringify(prevProps) !== JSON.stringify(this.props)) {
      this._generateHtml();
    }
  }

 public componentDidMount(): void {
    this._generateHtml();
  }

  private _generateHtml() {
    let newParser = new HtmlParser(this.props.itemTemplate, this.props.controlTemplate, this.props.noResultsTemplate);
    newParser.getPlaceholders()
      .then((placeholders)=>{
        return this._searchSharePoint(placeholders);})
      .then((searchResults)=>{
        return newParser.replacePlaceholdersWithSearchResultValues(searchResults);})
      .then((renderstuff)=>{
          this.setState({...this.state, htmlResult: renderstuff});
          loadStyles(this.props.cssStyles);
      }).catch(error=>{
        this.setState({...this.state, error: true, errorMessage: error});
      });
  }
 

  private _searchSharePoint(placeHolders:string[]): Promise<SearchResult[]> {
    return new Promise<SearchResult[]>((resolve, reject) => {
      try {
        let q = SearchQueryBuilder.create().text(this.props.query).selectProperties(...placeHolders).rowLimit(this.props.maxNumberResults);
        pnp.sp.search(q).then(r => { 
          if(r.RowCount > 0){
            let primarySearchResults = r.PrimarySearchResults;
            resolve(primarySearchResults);
          } else resolve(null);
        });
      } catch(error){
        reject(error);
      }
    });
  }

  public render(): React.ReactElement<ICswpProps> {
    const errorState = this.state.error;
    return (
      <div className='conflux-cswp'>
         {errorState ? (
            <div>Error: { this.state.errorMessage }</div>
          ) : (
            <div dangerouslySetInnerHTML={{__html: this.state.htmlResult}}></div>
          )}          
      </div>   
    );
  }
}
