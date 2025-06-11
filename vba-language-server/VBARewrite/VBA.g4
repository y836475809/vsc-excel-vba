grammar VBA;

options {
    caseInsensitive = true;
}

startRule
    : (typeStmt 
    // | propertyGetStmt | propertyLetStmt | propertySetStmt
    | propertyGetStmt | propertySetStmt
    | endPropertyStmt
    | moduleAttributes | moduleOption 
    | openStmt | outputStmt | closeStmt 
    | inputStmt | lineInputStmt
    // | setStmt | letStmt
    | dimStmt | redimStmt
    | subStmt | endSubStmt
    | functionStmt | endFunctionStmt
    | .)*?
    EOF
    ;
// startFileIoRule
//     : (openStmt |  outputStmt | closeStmt 
//     | inputStmt | lineInputStmt
//     | .)*?
//     EOF
//     ;
moduleAttributes
    : (attributeStmt endOfLine+)+
    ;
attributeStmt
    : ATTRIBUTE WS identifier WS? EQ WS? identifier
    ;
ATTRIBUTE
    : 'ATTRIBUTE'
    ;
moduleOption
    : 'OPTION' WS 'EXPLICIT'
    ;
dimStmt
    // : DIM WS identifier (WS? argList)? (WS asTypeClause)?
    : DIM WS identifier WS? LPAREN WS? RPAREN (WS asTypeClause)?
    ;
redimStmt
    // : (DIM | REDIM) WS identifier (WS? argList)? (WS asTypeClause)?
    : REDIM WS (PRESERVE WS)? identifier (WS? redimToArgList | redimArgList)? (WS asTypeClause)?
    ;
redimToArgList
    : LPAREN (WS? redimToArg (WS? ',' WS? redimToArg)*)? WS? RPAREN
    ;
redimToArg
    : identifier WS REDIMTO WS identifier
    ;
redimArgList
    : LPAREN (WS? identifier (WS? ',' WS? identifier)*)? WS? RPAREN
    ;
DIM
    : 'DIM'
    ;
REDIM
    : 'REDIM'
    ;
PRESERVE
    : 'PRESERVE'
    ;
REDIMTO
    : 'TO'
    ;
subStmt
    // : (visibility WS)? SUB WS identifier WS? LPAREN WS? RPAREN WS? endOfStatement (redimStmt | .)*? END_SUB
    // : (visibility WS)? SUB WS identifier WS? .*? END_SUB
    // : (visibility WS)? 'SUB' WS identifier WS? LPAREN WS? RPAREN WS? endOfStatement
    : (visibility WS)? SUB WS identifier WS? LPAREN
    ;
functionStmt
    // : (visibility WS)? FUNCTION WS identifier argList? (WS? asTypeClause)? endOfStatement (.*?) END_FUNCTION
    // : (visibility WS)? 'FUNCTION' WS identifier argList? (WS? asTypeClause)? endOfStatement
    : (visibility WS)? FUNCTION WS identifier WS? LPAREN
    ;
endSubStmt
    : 'END' WS 'SUB'
    ;
endFunctionStmt
    : 'END' WS 'FUNCTION'
    ;
SUB
    : 'SUB'
    ;
// END_SUB
//     : 'END' WS 'SUB'
//     ;
FUNCTION
    : 'FUNCTION'
    ;
// END_FUNCTION
//     : 'END' WS 'FUNCTION'
//     ;
// Option Explicit
// fileIoModule
//     : comment
//     | openStmt | printStmt | writeStmt | closeStmt
//     ;
fileIoModule
    : openStmt
    | outputStmt
    ;
openStmt
    : OPEN WS openPath WS FOR WS ('APPEND' | 'BINARY' | 'INPUT' | 'OUTPUT' | 'RANDOM')
    (WS ACCESS WS ('READ' | 'WRITE' | 'READ WRITE'))?
    (WS ('SHARED' | 'LOCK READ' | 'LOCK WRITE' | 'LOCK READ WRITE'))?
    WS AS WS fileNumber
    //    openLenStmt?
    ;
// openLenStmt
//     : 'LEN' WS? EQ WS? identifier
//     ;put
OPEN
    : 'OPEN'
    ;
FOR
    : 'FOR'
    ;
ACCESS
    : 'ACCESS'
    ;
outputStmt
    : OUTPUT (WS fileNumber)? WS? ',' WS? outputArg? outputList*
    ;
outputList
    : (WS? (';' | ',' | outputArg))+
    ;
outputArg
    : (identifier WS)? identifier (
        WS? LPAREN (WS? outputArg WS? ','?)* WS? RPAREN
    )?
    ;
OUTPUT
    : PRINT | WRITE
    ;
PRINT
    : 'PRINT'
    ;
WRITE
    : 'WRITE'
    ;
inputStmt
    : INPUT WS fileNumber (WS? ',' WS? identifier)+
    ;
lineInputStmt
    : LINE_INPUT WS fileNumber WS? ',' WS? identifier
    ;
INPUT
    : 'INPUT'
    ;
LINE_INPUT
    : 'LINE' WS 'INPUT'
    ;
closeStmt
    : CLOSE (WS fileNumber (WS? ',' WS? fileNumber)*)?
    ;
CLOSE
    : 'CLOSE'
    ;
openPath
    : identifier
    ;
openMode
    : identifier
    ;
fileNumber
    : '#'? identifier
    ;
FILENUMBERSYMBOL
    : '#'
    ;
typeStmt
    : (visibility WS)? TYPE WS identifier endOfStatement (blockTypeStmt | .)*? typeEndStmt
    ;
blockTypeStmt
    : (visibility WS)? identifier .*? WS asTypeClause endOfStatement
    ;
typeEndStmt
    : END_TYPE
    ;
propertyGetStmt
    // : (visibility WS)? (STATIC WS)? PROPERTY_GET WS identifier LPAREN WS? RPAREN (
    //     WS asTypeClause
    // )? endOfStatement (blockLetSetStmt | .)*? endPropertyStmt
    : (visibility WS)? (STATIC WS)? PROPERTY WS GET WS identifier argList (
        WS asTypeClause (WS? arrayStmt)?
    )?
    ;
// blockLetSetStmt
//     : letStmt | setStmt
//     ;
letStmt
    : (LET WS)? identifier WS? EQ WS? identifier argList? endOfStatement
    ;
setStmt
    : SET WS identifier WS? EQ WS? identifier argList? endOfStatement
    ;
// propertyLetStmt
//     // : (visibility WS)? (STATIC WS)? PROPERTY_LET WS identifier (WS? argList)? endOfStatement .*? endPropertyStmt
//     : (visibility WS)? (STATIC WS)? PROPERTY_LET WS identifier (WS? argList)? endOfStatement
//     ;
propertySetStmt
    // : (visibility WS)? (STATIC WS)? PROPERTY_SET WS identifier (WS? argList)? endOfStatement .*? endPropertyStmt
    // : (visibility WS)? (STATIC WS)? PROPERTY_SET WS identifier (WS? argList)? endOfStatement
    : (visibility WS)? (STATIC WS)? PROPERTY WS (SET | LET) WS identifier WS? argList
    ;
endPropertyStmt
    : END_PROPERTY
    ;
asTypeClause
    : AS WS? (NEW WS)? WS? identifier
    ;
argList
    : LPAREN (WS? arg (WS? ',' WS? arg)*)? WS? RPAREN
    ;
arg
    : ((BYVAL | BYREF) WS)? identifier (
        // WS? LPAREN WS? RPAREN
        WS? arrayStmt
    )? (WS? asTypeClause)? (WS? argDefaultValue)?
    ;
argDefaultValue
    : EQ WS? identifier
    ;
arrayStmt
    : redimArgList | redimToArgList
    ;
BYVAL
    :  'BYVAL'
    ;
BYREF
    :  'BYREF'
    ;
PROPERTY
    : 'PROPERTY'
    ;
GET
    : 'GET'
    ;
// PROPERTY_SET
//     : 'SET'
//     ;
// PROPERTY_LET
//     : 'LET'
//     ;
// PROPERTY_SET
//     : 'PROPERTY' WS 'SET'
//     ; 
// PROPERTY_GET
//     : 'PROPERTY' WS 'GET'
//     ;
// PROPERTY_LET
//     : 'PROPERTY' WS 'LET'
//     ;
// PROPERTY_SET
//     : 'PROPERTY' WS 'SET'
//     ; 
END_PROPERTY
    : 'END' WS 'PROPERTY'
    ;
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

identifier
    : IDENTIFIER+
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
IDENTIFIER
    : FILENUMBERSYMBOL? ~[\]()\r\n\t.,'"|!@%^:;=#/ ]+
    | L_SQUARE_BRACKET (~[!\]\r\n])+ R_SQUARE_BRACKET
    | STRINGLITERAL
    ;
DOT
    : '.'
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
LINE_CONTINUATION
    : [ \t]+ UNDERSCORE '\r'? '\n' WS* -> skip
    ;
WS
    : ([ \t] | UNDERSCORE [ \t]? NEWLINE | LINE_CONTINUATION)+
    ;
STRINGLITERAL
    : '"' (~["\r\n] | '""')* '"'
    ;
UNKNOWN: . -> skip;