grammar VBADocumentSymbol;

options {
    caseInsensitive = true;
}

startRule
    : (constStmt | dimStmt | filedVariant
    | typeStmt 
    | enumStmt
    | propertyGetStmt | propertySetStmt | endPropertyStmt
    | subStmt | endSubStmt
    | functionStmt | endFunctionStmt
    | .)*?
    EOF
    ;
constStmt
    : CONST WS identifier WS? EQ WS? (identifier | NUMBER)
    ;
CONST
    : 'CONST'
    ;
dimStmt
    : DIM WS identifier (WS? argList)? (WS asTypeClause)?
    ;
DIM
    : 'DIM'
    ;
filedVariant
    : (visibility WS)? identifier argList? WS asTypeClause
    ;
enumStmt
    : (visibility WS)? ENUM WS identifier endOfStatement blockEnumStmt*? endEnumStmt
    ;
endEnumStmt
    : 'END' WS 'ENUM'
    ;
ENUM
    : 'ENUM'
    ;
blockEnumStmt
    : identifier (WS? EQ WS? NUMBER)? endOfStatement
    ;
subStmt
    : (visibility WS)? SUB WS identifier WS? argList WS? endOfStatement
    // : (visibility WS)? SUB WS identifier WS? .*? END_SUB
    // : (visibility WS)? 'SUB' WS identifier WS? LPAREN WS? RPAREN WS? endOfStatement
    // : (visibility WS)? SUB WS identifier WS? LPAREN
    ;
functionStmt
    : (visibility WS)? FUNCTION WS identifier WS? argList WS? (WS? asTypeClause)? endOfStatement
    // : (visibility WS)? 'FUNCTION' WS identifier argList? (WS? asTypeClause)? endOfStatement
    // : (visibility WS)? FUNCTION WS identifier WS? LPAREN
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
FUNCTION
    : 'FUNCTION'
    ;
FILENUMBERSYMBOL
    : '#'
    ;
typeStmt
    : (visibility WS)? TYPE WS identifier endOfStatement blockTypeStmt*? typeEndStmt
    ;
blockTypeStmt
    // : (visibility WS)? identifier .*? WS asTypeClause endOfStatement
    : (visibility WS)? identifier (WS? LPAREN WS? .*? WS? RPAREN)? (WS asTypeClause)? endOfStatement
    ;
typeEndStmt
    : END_TYPE
    ;
propertyGetStmt
    : (visibility WS)? (STATIC WS)? PROPERTY WS GET WS identifier argList? (
        WS asTypeClause
    )? endOfStatement
    ;
propertySetStmt
    : (visibility WS)? (STATIC WS)? PROPERTY WS (SET | LET) WS identifier (WS? argList)? endOfStatement
    // : (visibility WS)? (STATIC WS)? PROPERTY_SET WS identifier (WS? argList)? endOfStatement
    // : (visibility WS)? (STATIC WS)? PROPERTY WS (SET | LET) WS identifier WS? LPAREN WS? arg WS? RPAREN endOfStatement
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
        WS? LPAREN WS? RPAREN
    )? (WS? asTypeClause)? (WS? argDefaultValue)?
    ;
argDefaultValue
    : EQ WS? identifier
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
NUMBER
    : [0-9]+
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