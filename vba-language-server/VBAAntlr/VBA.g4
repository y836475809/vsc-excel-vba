grammar VBA;

options {
    caseInsensitive = true;
}

// startRule
//     : (typeStmt
//     | propertyGetStmt | propertyLetStmt | propertySetStmt 
//     | .
//     )*
//     EOF
//     ;
// startRule
//     : typeStmt
//     | propertyGetStmt | propertyLetStmt | propertySetStmt 
//     | endOfLine*
//     EOF
//     ;

startRule
    : (WS? endOfLine* module endOfLine* WS?)* EOF
    ;
module
    : comment
    | typeStmt
    | propertyGetStmt | propertyLetStmt | propertySetStmt 
    | variableStmt
    | subStmt | constStmt
    | functionStmt
    | declareStmt
    | macroConstStmt | macroIfBlockStmt | macroElseIfBlockStmt | macroElseBlockStmt 
    | macroEndBlockStmt
    ;
startIORule
    : (WS? endOfLine* ioModule endOfLine* WS?)* EOF
    ;
ioModule
    : comment
    | openStmt | printStmt | writeStmt | closeStmt
    ;
openStmt
    : OPEN WS ambiguousIdentifier WS FOR WS OpenMode 
       (WS ACCESS WS AccessMode)?
       (WS LockMode)?
       WS AS WS fileNumber
       (WS 'LEN' WS? EQ WS? ambiguousIdentifier)?
    ;
OPEN
    : 'WRIOPENTE'
    ;
FOR
    : 'FOR'
    ;
ACCESS
    : 'ACCESS'
    ;
printStmt
    : 'PRINT' WS fileNumber WS? (',' | ';') (WS? ambiguousIdentifier)?
    ;
writeStmt
    : 'WRITE' WS fileNumber (WS? (',' | ';') WS? ambiguousIdentifier?)+
    ;
closeStmt
    : 'CLOSE' (WS fileNumber (WS? ',' WS? fileNumber)*)?
    ;
OpenMode
    : 'APPEND' | 'BINARY' 
    | 'INPUT' | 'OUTPUT' 
    | 'RANDOM'
    ;
AccessMode
    : 'READ' | 'WRITE' 
    | ('READ' WS 'WRITE')
    ;
LockMode
    : 'SHARED' 
    | ('LOCK' WS 'READ') 
    | ('LOCK' WS 'WRITE') 
    | ('LOCK' WS 'READ' WS 'WRITE')
    ;
fileNumber
    : '#'? ambiguousIdentifier
    ;
variableStmt
    : (DIM | STATIC | visibility) WS (WITHEVENTS WS)? variableListStmt endOfStatement?
    ;
// variableStmt
//     : (DIM | STATIC | visibility) WS (WITHEVENTS WS)? .+?
//     ;
constStmt
    : (visibility WS)? CONST WS constSubStmt (WS? ',' WS? constSubStmt)*
    ;
constSubStmt
    : ambiguousIdentifier (WS asTypeClause)? WS? EQ WS? ambiguousIdentifier
    ;
variableListStmt
    : variableSubStmt (WS? ',' WS? variableSubStmt)*
    ;
variableSubStmt
    : ambiguousIdentifier 
    (WS? LPAREN WS? variableAryStmt? RPAREN WS?)?
    (WS asTypeClause)?
    ;
variableAryStmt
    : AryStmt (WS? ',' WS? AryStmt)*
    | AryToStmt (WS? ',' WS? AryToStmt)*
    ;
AryToStmt
    : DIGIT+ WS 'TO' WS DIGIT+
    ;
AryStmt
    : DIGIT+
    ;
fragment DIGIT
    : [0-9]
    ;
typeStmt
    : (visibility WS)? TYPE WS ambiguousIdentifier endOfStatement typeStmt_Element+ typeEndStmt
    ;
// typeStmt_Element
//     : ambiguousIdentifier (WS asTypeClause)? endOfStatement
//     ;
typeStmt_Element
    : variableSubStmt endOfStatement
    ;
typeEndStmt
    : END_TYPE
    ;
propertyGetStmt
    : (visibility WS)? (STATIC WS)? PROPERTY_GET WS ambiguousIdentifier LPAREN WS? RPAREN (
        WS asTypeClause
    )? endOfStatement (blockLetSetStmt*?|.*?) endPropertyStmt
    ;
blockLetSetStmt
    : letStmt | setStmt
    ;
// blockStmt
//     : letStmt | setStmt
//     ;
letStmt
    : (LET WS)? ambiguousIdentifier WS? EQ WS? ambiguousIdentifier endOfStatement
    ;
setStmt
    : SET WS ambiguousIdentifier WS? EQ WS? ambiguousIdentifier endOfStatement
    ;
propertyLetStmt
    : (visibility WS)? (STATIC WS)? PROPERTY_LET WS ambiguousIdentifier (WS? argList)? endOfStatement .*? endPropertyStmt
    ;
propertySetStmt
    : (visibility WS)? (STATIC WS)? PROPERTY_SET WS ambiguousIdentifier (WS? argList)? endOfStatement .*? endPropertyStmt
    ;
endPropertyStmt
    : END_PROPERTY
    ;
subStmt
    : (visibility WS)? (STATIC WS)? SUB WS? ambiguousIdentifier (WS? argList)? endOfStatement .*? END_SUB
    ;
functionStmt
    : (visibility WS)? (STATIC WS)? FUNCTION WS? ambiguousIdentifier (WS? argList)? (
        WS? asTypeClause
    )? endOfStatement .*? END_FUNCTION
    ;
enumerationStmt
    : (visibility WS)? ENUM WS ambiguousIdentifier endOfStatement enumerationStmt_Constant* END_ENUM
    ;
enumerationStmt_Constant
    : ambiguousIdentifier (WS? EQ WS? ambiguousIdentifier)? endOfStatement
    ;
// declareStmt
//     : (visibility WS)? DECLARE WS (PTRSAFE WS)? (FUNCTION | SUB) WS ambiguousIdentifier WS LIB WS STRINGLITERAL (
//         WS ALIAS WS STRINGLITERAL
//     )? (WS? argList)? (WS asTypeClause)?
//     ;
declareStmt
    : (visibility WS)? DECLARE WS (PTRSAFE WS)? (FUNCTION | SUB) WS ambiguousIdentifier WS LIB WS ambiguousIdentifier (
        WS ALIAS WS ambiguousIdentifier
    )? (WS? argList)? (WS asTypeClause)?
    ;
asTypeClause
    : AS WS (NEW WS)? ambiguousIdentifier
    ;
// asTypeClause
//     : AS WS UNDERSCORE WS NEWLINE ambiguousIdentifier
//     ;
// argList
//     : LPAREN (ambiguousIdentifier | WS | ',')* RPAREN
//     ;
argList
    : LPAREN (WS? arg (WS? ',' WS? arg)*)? WS? RPAREN
    ;
arg
    : (ambiguousIdentifier WS)? ambiguousIdentifier (
        WS? LPAREN WS? RPAREN
    )? (WS? asTypeClause)? (WS? argDefaultValue)?
    ;
argDefaultValue
    : EQ WS? ambiguousIdentifier
    ;
DECLARE
    : 'DECLARE'
    ;
LIB
    : 'LIB'
    ;
ALIAS
    : 'ALIAS'
    ;
PTRSAFE
    : 'PTRSAFE'
    ;
ENUM
    : 'ENUM'
    ;
END_ENUM
    : 'END' WS 'ENUM'
    ;
CONST
    : 'CONST'
    ;
SUB
    : 'SUB'
    ;
END_SUB
    : 'END' WS 'SUB'
    ;
FUNCTION
    : 'FUNCTION'
    ;
END_FUNCTION
    : 'END' WS 'FUNCTION'
    ;
DIM
    : 'DIM'
    ;
WITHEVENTS
    : 'WITHEVENTS'
    ;
PROPERTY_GET
    : 'PROPERTY' WS 'GET'
    ;
PROPERTY_LET
    : 'PROPERTY' WS 'LET'
    ;
PROPERTY_SET
    : 'PROPERTY' WS 'SET'
    ; 
END_PROPERTY
    : 'END' WS 'PROPERTY'
    ;
// PP
//     : [(PROPERTY) WS+ (SET)]
//     ;
EQ
    : '='
    ;
LET
    : 'LET'
    ;
SET
    : 'SET'
    ;
LPAREN
    : '('
    ;
RPAREN
    : ')'
    ;
STATIC
    : 'STATIC'
    ;
NEW
    : 'NEW'
    ;
AS
    : 'AS'
    ;
PRIVATE
    : 'PRIVATE'
    ;
PUBLIC
    : 'PUBLIC'
    ;
UNDERSCORE
    : '_'
    ;
TYPE
    : 'TYPE'
    ;
COLON
    : ':'
    ;
REM
    : 'REM'
    ;
visibility
    : PRIVATE
    | PUBLIC
    ;

ambiguousIdentifier
    : (IDENTIFIER)+
    ;
L_SQUARE_BRACKET
    : '['
    ;
R_SQUARE_BRACKET
    : ']'
    ;
NEWLINE
    : [\r\n\u2028\u2029]+
    ;
macroConstStmt
    : MACRO_CONST WS? ambiguousIdentifier WS? EQ WS? ambiguousIdentifier
    ;
macroIfBlockStmt
    : MACRO_IF WS ambiguousIdentifier WS THEN endOfStatement
    ;
macroElseIfBlockStmt
    : MACRO_ELSEIF WS ambiguousIdentifier WS THEN endOfStatement
    ;
macroElseBlockStmt
    : MACRO_ELSE endOfStatement
    ;
macroEndBlockStmt
    : MACRO_END_IF endOfStatement
    ;
MACRO_CONST
    : '#CONST'
    ;
MACRO_IF
    : '#IF'
    ;
MACRO_ELSEIF
    : '#ELSEIF'
    ;
MACRO_ELSE
    : '#ELSE'
    ;
MACRO_END_IF
    : '#END' WS? 'IF'
    ;
THEN
    : 'THEN'
    ;
// IDENTIFIER
//     : ~[\]()\r\n\t.,'"|!@#$%^&*\-+:=; ]+
//     | L_SQUARE_BRACKET (~[!\]\r\n])+ R_SQUARE_BRACKET
//     ;
// IDENTIFIER
//     : ~[\]()\r\n\t,'"|!@#$%^:; ]+
//     | L_SQUARE_BRACKET (~[!\]\r\n])+ R_SQUARE_BRACKET
//     | STRINGLITERAL
//     ;
IDENTIFIER
    : ~[\]()\r\n\t,'"|!@#%^:; ]+
    | L_SQUARE_BRACKET (~[!\]\r\n])+ R_SQUARE_BRACKET
    | STRINGLITERAL
    ;
END_TYPE
    : 'END' WS 'TYPE'
    ;
COMMENT
    : SINGLEQUOTE (LINE_CONTINUATION | ~[\r\n\u2028\u2029])*
    ;
SINGLEQUOTE
    : '\''
    ;
remComment
    : REMCOMMENT
    ;
REMCOMMENT
    : COLON? REM WS (LINE_CONTINUATION | ~[\r\n\u2028\u2029])*
    ;
comment
    : COMMENT
    ;
endOfLine
    : WS? (NEWLINE | comment | remComment) WS?
    ;
endOfStatement
    : (endOfLine | WS? COLON WS?)*
    ;

// whitespace, line breaks, comments, ...
LINE_CONTINUATION
    : [ \t]+ UNDERSCORE '\r'? '\n' WS* -> skip
    ;

// WS
//     : ([ \t] | LINE_CONTINUATION)+
//     ;
WS
    : ([ \t] | UNDERSCORE [ \t]? NEWLINE | LINE_CONTINUATION)+
    ;
STRINGLITERAL
    : '"' (~["\r\n] | '""')* '"'
    ;
// TOKEN
//     : ~[+\-\u0000-\u001f <>:"/\\|?*#@] ~[\u0000-\u001f <>:"/\\|?*#@]+
//     ;
// PP
//     :~'Property'
//     ;
// PPP :'Property'
//     ;
// UNKNOWN: . -> skip;
// ignored : . ;
// IGNORED : . ;