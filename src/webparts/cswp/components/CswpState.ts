import pnp, { SearchResult } from "sp-pnp-js";

export interface ICswpState {
    htmlResult: string;
    error: boolean;
    errorMessage: string;
    cssStyles:string;
}