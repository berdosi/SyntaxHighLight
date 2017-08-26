
/**
 * Returns the text with highlighted syntax
 * 
 * @param {string} rawText The raw source code
 * @param {string} language The language to highlight the syntax for.
 */
function syntaxHighlight(rawText, language = "VB") {
    /** @typedef Token 
     * @property {RegExp} match regexp to mathc for at the beginning of the sequence
     * @property {string} type type of the token "keyword"|"whitespace"|"delimiter"|"escapedChar"|"string"|"number"
     * @property {?string} allowedParents like "string", e.g. to  only match escaped sequences inside of strings
     * @property {?string} disallowedParents like "string", e.g. to avoid matching keywords inside of strings.
     * 
    */

    /** @type {Token[]} */
    let tokenList = [
        { match: /^""/, type: "escapedChar", allowedParents: ["string"]}, // only in strings.
        { match: /^"/, type: "string", disallowedParents: ["string"] }, // sets or unsets the "string" parent as well.
        { match: /^Sub\s+|^Function\s+|^If\s+|^Else\s+|^While\s+|^Wend\s+|^For\s+|^Loop\s+|^In\s+|^To\s+|^Each\s+|^End\s+|^Exit /,
            type: "keyword", disallowedParents: ["string"]},
        { match: /^String|^Long|^Byte|^Short|^Array|^Integer|^Object|^Variant/, type: "type", disallowedParents: ["string"] },
        { match: /^[:\r\n]/,
            type: "delimiter", disallowedParents: ["string"]}, 
        { match: /^[<>=+-\/*]/,
        type: "operator", disallowedParents: ["string"]},
        { match: /^\s+/,
        type: "whitespace"},
        { match: /^[^"]+/, type: "stringText", allowedParents: ["string"]}
    ];

    let currentParent = "";
    let output = [];
    while (rawText.length > 0) {
        for (let token of tokenList) {
            if (token.match.test(rawText) 
                && (!token.allowedParents 
                    || (token.allowedParents.indexOf(currentParent) != -1))
                && (!token.disallowedParents 
                    || (token.disallowedParents.indexOf(currentParent) == -1))) {
                currentParent = token.type;
                
            }
        }
    }

    
}

syntaxHighlight()