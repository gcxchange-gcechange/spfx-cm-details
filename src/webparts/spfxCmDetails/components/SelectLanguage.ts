/* eslint-disable @typescript-eslint/no-var-requires */

import * as strings from 'SpfxCmDetailsWebPartStrings';
const english = require("../loc/en-us.js")
const french = require("../loc/fr-fr.js")

export function SelectLanguage(lang: string): ISpfxCmDetailsWebPartStrings {
    switch (lang) {
        case "en-us": {
            return english;
        }
        case "fr-fr": {
            return french;
        }
        default: {
            return strings;
        }
    }
}