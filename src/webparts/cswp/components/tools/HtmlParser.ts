import pnp, { SearchResult } from "sp-pnp-js";

export class HtmlParser {
    private _placeHolders: string[];
    private _htmlItemsString: string;
    private _htmlControlString: string;
    private _noResultsTemplate: string;

    constructor(htmlItemsString:string, htmlControlString: string, noResultsTemplate:string){
        this._htmlItemsString = htmlItemsString;
        this._htmlControlString = htmlControlString;
        this._noResultsTemplate = noResultsTemplate;
        this._placeHolders = this._getPlaceholders(this._htmlItemsString);
    }

    private _getPlaceholders(htmlString:string): string[] {
        var match, matches = [];
        var regex = /_#(.*?)#_/g;

        while (match = regex.exec(htmlString))
        {
            matches.push(match[1]);    
        }

        return matches;
    }

    public getPlaceholders(): Promise<string[]> {
        return new Promise<string[]>((resolve, reject) => {
            try{
                resolve(this._placeHolders.map(str => str.replace(/\s+/g, '')));   
            }
            catch(error){
                reject(error);
            }
        });
    }
    
    public replacePlaceholdersWithSearchResultValues(results:SearchResult[]): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            try{
                if(results==null) { 
                    resolve(this._noResultsTemplate); 
                }
                let resultHtml: string[] = [];
                results.map((result:SearchResult) => {
                    let template_copy = (' ' + this._htmlItemsString).slice(1);
                    this._placeHolders.map((placeholder:string) => {
                        template_copy = template_copy.replace('_#' + placeholder + '#_', (result[placeholder.trim()] || ""));
                    });
                    resultHtml.push(template_copy);
                });
                resolve(this._htmlControlString.replace(/{items}/i, resultHtml.join('')));
            }
            catch(error){
                reject(error);
            }
        });
    }
}