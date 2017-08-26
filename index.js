
/**
 * Returns the text with highlighted syntax
 * 
 * @param {string} rawText The raw source code
 */
function syntaxHighlight(rawText) {
    /** @typedef Token 
     * @property {RegExp} match regexp to mathc for at the beginning of the sequence
     * @property {string} type type of the token "keyword"|"whitespace"|"delimiter"|"escapedChar"|"string"|"number"
     * @property {?string[]} allowedParents like "string", e.g. to  only match escaped sequences inside of strings
     * @property {?string[]} disallowedParents like "string", e.g. to avoid matching keywords inside of strings.
     * @property {?string} setParent like "string", e.g. to avoid matching keywords inside of strings.
     * @property {?string} unsetParent like "string", e.g. to avoid matching keywords inside of strings.
     * 
    */

    /** @type {Token[]} */
    let tokenList = [
        {
            match: /^""/,
            type: "escapedChar",
            allowedParents: ["stringText", "string"]
        }, // only in strings.
        {
            match: /^"/,
            type: "stringStart",
            disallowedParents: ["string", "stringText"],
            setParent: "string"
        }, // sets or the "string" parent as well.
        {
            match: /^"/, // escapedChar will be matched first, if any.
            type: "stringEnd",
            allowedParents: ["string", "stringText"],
            unsetParent: "string"
        }, // unsets or the "string" parent, if we were in a string, or in a stringtext. If there was an escapedChar, it was matched before     
        {
            match: /^Sub\s+|^Function\s+|^If\s+|^Else\s+|^While\s+|^Wend\s+|^For\s+|^Loop\s+|^In\s+|^To\s+|^Each\s+|^End\s+|^Exit\s+|^Call\s+/i,
            type: "keyword", disallowedParents: ["string"]
        },
        {
            match: /^String|^Long|^Byte|^Short|^Array|^Integer|^Object|^Variant/,
            type: "type", 
            disallowedParents: ["string"]
        },
        {
            match: /^[:\r\n]/,
            type: "delimiter", 
            disallowedParents: ["string"]
        },
        {
            match: /^\s+/,
            type: "whitespace"
        },
        {
            match: /^[^"]+/,
            type: "stringText", allowedParents: ["string"]
        },
        {
            match: /[+\-]?[0-9]+\.?[0-9]*/,
            type: "generic",
            disallowedParents: ["string"]
        },
        {
            match: /^[<>=+\-\/*]/,
            type: "operator", 
            disallowedParents: ["string"]
        },
        {
            match: /[a-z0-9]+/i,
            type: "generic",
            disallowedParents: ["string"]
        },
        {
            match: /[\[\]\(\){}]/,
            type: "parentheses",
            disallowedParents: ["string"]
        },
        {
            match: /'.*$/,
            type: "comment",
            disallowedParents: ["string"]
        },
        {
            match: /\./,
            type: "dot",
            disallowedParents: ["string"]
        }
    ];

    /** all the parents(, grandparents, etc.) of the element, in reverse order 
     * (so that the closest parent can be referenced at index 0) 
     * @type {string[]}
     * */
    let ancestry = ["root"];
    let output = [];
    while (rawText.length > 0) {

        const currentParent = ancestry[0];
        const nextToken = tokenList.find(token => (
            token.match.test(rawText)
            && (!token.allowedParents
                || (token.allowedParents.indexOf(currentParent) != -1))
            && (!token.disallowedParents
                || (token.disallowedParents.indexOf(currentParent) == -1))));
                (match, p1) =>(match, p1) => p1 p1
        if (!nextToken) throw new Error("parse error");
        // if the token changes the hierarchy (e.g. starts a string inside of which other things match)
        // record this.
        if (nextToken.setParent) ancestry.unshift(nextToken.setParent);
        if (nextToken.unsetParent) ancestry.shift(nextToken.unsetParent);

        // get the matched token
        // - create a regexp, which replaces the rawText with the token only, get the tokenText
        const captureRegexp = new RegExp(`(${nextToken.match.source}).*`, nextToken.match.flags + "m");
        // - replace the token in rawText (first match only)
        const tokenContent = rawText.split(/[\r\n]/)[0].replace(captureRegexp, (match, p1) => p1); 
        rawText = rawText.replace(nextToken.match, "");
        // remove the matched token from the beginning of rawText
        output.push({ content: tokenContent, token: nextToken })
    }


    return output;
}

let output = syntaxHighlight(`Sub Macro1()
'
' Macro1 Macro
'
'
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Hello World"
    Range("B2").Select
End Sub`)