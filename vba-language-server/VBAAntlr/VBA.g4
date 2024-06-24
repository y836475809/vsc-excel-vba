grammar VBA;

options {
    caseInsensitive = true;
}

startRule
    : (typeStmt 
    | propertyGetStmt | propertyLetStmt | propertySetStmt 
    | .)*?
    EOF
    ;
startFileIoRule
    : (fileIoModule 
    | .)*?
    EOF
    ;
// fileIoModule
//     : comment
//     | openStmt | printStmt | writeStmt | closeStmt
//     ;
fileIoModule
    : openStmt
    ;
openStmt
    : OPEN WS identifier WS FOR WS OpenMode 
       (WS ACCESS WS AccessMode)?
       (WS LockMode)?
       WS AS WS fileNumber
       (WS 'LEN' WS? EQ WS? identifier)?
    ;
OPEN
    : 'OPEN'
    ;
FOR
    : 'FOR'
    ;
ACCESS
    : 'ACCESS'
    ;
printStmt
    : 'PRINT' WS fileNumber WS? (',' | ';') (WS? identifier)?
    ;
writeStmt
    : 'WRITE' WS fileNumber (WS? (',' | ';') WS? identifier?)+
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
    : (visibility WS)? (STATIC WS)? PROPERTY_GET WS identifier LPAREN WS? RPAREN (
        WS asTypeClause
    )? endOfStatement (blockLetSetStmt | .)*? endPropertyStmt
    ;
blockLetSetStmt
    : letStmt | setStmt
    ;
letStmt
    : (LET WS)? identifier WS? EQ WS? identifier argList? endOfStatement
    ;
setStmt
    : SET WS identifier WS? EQ WS? identifier argList? endOfStatement
    ;
propertyLetStmt
    : (visibility WS)? (STATIC WS)? PROPERTY_LET WS identifier (WS? argList)? endOfStatement .*? endPropertyStmt
    ;
propertySetStmt
    : (visibility WS)? (STATIC WS)? PROPERTY_SET WS identifier (WS? argList)? endOfStatement .*? endPropertyStmt
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
    : (identifier WS)? identifier (
        WS? LPAREN WS? RPAREN
    )? (WS? asTypeClause)? (WS? argDefaultValue)?
    ;
argDefaultValue
    : EQ WS? identifier
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
    : FILENUMBERSYMBOL? ~[\]()\r\n\t,'"|!@%^:;=#/ ]+
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