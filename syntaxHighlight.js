/** @typedef Token 
 * @property {RegExp} match regexp to mathc for at the beginning of the sequence
 * @property {string} type type of the token "keyword"|"whitespace"|"delimiter"|"escapedChar"|"string"|"number"
 * @property {?string[]} allowedParents like "string", e.g. to  only match escaped sequences inside of strings
 * @property {?string[]} disallowedParents like "string", e.g. to avoid matching keywords inside of strings.
 * @property {?string} setParent like "string", e.g. to avoid matching keywords inside of strings.
 * @property {?string} unsetParent like "string", e.g. to avoid matching keywords inside of strings.
 * @property {?string} content as a return value the matched content from the input
 */

/** Expose methods for tokenizing a source code and generating HTML from it.
 * @constructor 
 * 
 * @param {Token[]} tokenList - Define the tokens to look in the text for.
 */
function SyntaxHighlight(tokenList) {
    /** Return tokenized text
     * @method tokenize
     * @param {string} rawText Raw source code
     * @return {Token[]} An array of detected tokens.
     */
    this.tokenize = function(rawText) {
        /** all the parents(, grandparents, etc.) of the element, in reverse order 
         * (so that the closest parent can be referenced at index 0) 
         * @type {string[]}
         * */
        let ancestry = ["root"];
        let output = [];
        while (rawText.length > 0) {
            const line = rawText.split(/[\r\n]/)[0] + "\n"; // always only on the first line
            const currentParent = ancestry[0];
            const nextToken = tokenList.find(token => (
                token.match.test(line)
                && (!token.allowedParents
                    || (token.allowedParents.indexOf(currentParent) != -1))
                && (!token.disallowedParents
                    || (token.disallowedParents.indexOf(currentParent) == -1))));
            if (!nextToken) throw new Error("parse error");
            // if the token changes the hierarchy (e.g. starts a string inside of which other things match)
            // record this.
            if (nextToken.setParent) ancestry.unshift(nextToken.setParent);
            if (nextToken.unsetParent) ancestry.shift(nextToken.unsetParent);

            // get the matched token
            // - create a regexp, which replaces the rawText with the token only, get the tokenText
            const captureRegexp = new RegExp(`(${nextToken.match.source}).*\n`, nextToken.match.flags);
            // - replace the token in rawText (first match only)
            const tokenContent = line.replace(captureRegexp, (match, p1) => p1); 
            rawText = rawText.replace(nextToken.match, "");
            // remove the matched token from the beginning of rawText
            output.push(Object.assign({ content: tokenContent }, nextToken ));
        }

        return output;
    }

    /** Generate a node with code annotated with syntax-highlightable classes inside 
     * @method generateHtml
     * @param {Token[]} tokens List of tokens to generate source code from.
     * @return {HTMLDivElement} DIV element containing the rendered source code.
     */
    this.generateHtml = function(tokens) {
        console.log(tokens);
        const containerElement = document.createElement("div");
        containerElement.className = "code";
        /** @type {Token[][]} */
        const codeLines = tokens.reduce((prev, current) => {
            if (current.type === "lineDelimiter") {
                prev.push([]);
                return prev;
            } else {
                prev[prev.length-1].push(current);
                return prev;
            }
        }, [[]]);

        codeLines.forEach(codeLine => {
            const line = document.createElement("div");
            codeLine.forEach(token => {
                const tokenElement = document.createElement("span");
                tokenElement.className = token.type;
                tokenElement.innerHTML = token.content;
                line.appendChild(tokenElement);
            });

            // make sure consecutive linebreaks are preserved when copying into Notepad
            line.appendChild(document.createElement("br")); 
            
            containerElement.appendChild(line);
            // keep the original line breaks from the input ()
            containerElement.appendChild(document.createTextNode("\n")); 
        });
        return containerElement;
    }
}
