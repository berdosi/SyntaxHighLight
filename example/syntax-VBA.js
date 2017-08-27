/** @type {Token[]} */
let tokenListVBA = [
    { // escapedChar (only quote in this case)
        match: /^""/,
        type: "escapedChar",
        allowedParents: ["stringText", "string"]
    }, // only in strings.
    { // stringStart
        match: /^"/,
        type: "stringStart",
        disallowedParents: ["string", "stringText"],
        setParent: "string"
    }, // sets or the "string" parent as well.
    { // stringEnd
        match: /^"/, // escapedChar will be matched first, if any.
        type: "stringEnd",
        allowedParents: ["string", "stringText"], // only match inside a string (otherwise the stringStart above will be matched)
        unsetParent: "string" // remove the "string" parentNode
    },
    { // comment
        match: /^'.*/,
        type: "comment",
        disallowedParents: ["string"]
    }, 
    { // keywords vol 1
        match: /^Sub\b|^Function\b|^If\b|^Else\b|^While\b|^Wend\b/i,
        type: "keyword", disallowedParents: ["string"]
    },
    { // keywords continue
        match: /^For\b|^Loop\b|^In\b|^To\b|^Each\b|^End\b|^Exit\b|^Call\b|^Class\b|^As\b|^Dim\b|^Const\b|^Step\b/i,
        type: "keyword", disallowedParents: ["string"]
    },
    { // keywords continue
        match: /^Private\b|^Public\b|^Friend\b|^Option\b|^Explicit\b|^End\b|^Exit\b|^Call\b|^Class\b/i,
        type: "keyword", disallowedParents: ["string"]
    },
    { // types
        match: /^String\b|^Long\b|^Byte\b|^Short\b|^Array\b|^Integer\b|^Object\b|^Variant\b/,
        type: "type", 
        disallowedParents: ["string"]
    },
    { // line delimiters (NB they get special treatment when rendering the code)
        match: /^[\r\n]/,
        type: "lineDelimiter", 
        disallowedParents: ["string"]
    },
    { // inline command delimiter
        match: /^:/,
        type: "inlineDelimiter", 
        disallowedParents: ["string"]
    },
    { // whitespace (to be preserved)
        match: /^[\t ]+/,
        type: "whitespace"
    },
    { // text inside a string (due to allowedParents property only matches inside of a string )
        match: /^[^"]+/,
        type: "stringText", allowedParents: ["string"]
    },
    { // number
        match: /^[+\-]?[0-9]+\.?[0-9]*/,
        type: "number",
        disallowedParents: ["string"]
    },
    { // operator
        match: /^[<>=+\-\/*&_]/,
        type: "operator", 
        disallowedParents: ["string"]
    },
    { // generic, unmatched-before word-like things.
        match: /^[a-z0-9]+/i,
        type: "generic",
        disallowedParents: ["string"]
    },
    { // parentheses 
        match: /^[\[\]\(\){}]/,
        type: "parentheses",
        disallowedParents: ["string"]
    },
    { // dot notation
        match: /^\./,
        type: "dot",
        disallowedParents: ["string"]
    }
];