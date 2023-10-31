grammar MyExcelMAUIApp;

/*
 * Parser Rules
 */

compileUnit : expression EOF;

expression :
    LPAREN expression RPAREN #ParenthesizedExpr
    | expression operatorToken=(MULTIPLY | DIVIDE) expression #MultiplicativeExpr
    | expression operatorToken=(ADD | SUBTRACT) expression #AdditiveExpr
    | IDENTIFIER #IdentifierExpr
    | expression EQ expression #EqualityExpr
    | expression LT expression #LessThanExpr
    | expression GT expression #GreaterThanExpr
    | NOT expression #NotExpr
    | expression AND expression #AndExpr
    | expression OR expression #OrExpr
    | INC LPAREN expression RPAREN #IncrementExpr
    | DEC LPAREN expression RPAREN #DecrementExpr
    | NUMBER #NumberExpr
    | TRUE #TrueExpr
    | FALSE #FalseExpr
    ;

/*
 * Lexer Rules
 */

NUMBER : INT ('.' INT)?;
IDENTIFIER : [a-zA-Z]+[1-9][0-9]*;

INT : ('0'..'9')+;

EQ : '=';
LT : '<';
GT : '>';
NOT : 'not';
AND : 'and';
OR : 'or';
INC : 'inc';
DEC : 'dec';
TRUE : 'True';
FALSE : 'False';

EXPONENT : '^';
MULTIPLY : '*';
DIVIDE : '/';
SUBTRACT : '-';
ADD : '+';
LPAREN : '(';
RPAREN : ')';

WS : [ \t\r\n] -> channel(HIDDEN);
