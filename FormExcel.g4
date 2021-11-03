grammar FormExcel;
/*
 * Parser Rules
 */

compileUnit : expression EOF;

expression :
	LPAREN expression RPAREN #ParenthesizedExpr
    | expression operatorToken=(MULTIPLY | DIVIDE) expression #MultiplicativeExpr
	| operatorToken=(MOD | DIV) LPAREN expression DESP expression RPAREN #MoDivExpr
	| expression UNARYPLUS expression #UnaryPExpr
	| expression UNARYMINUS expression #UnaryMExpr
	| expression operatorToken=(ADD | SUBTRACT) expression #AdditiveExpr
	| operatorToken=(MAX | MIN) LPAREN expression DESP expression RPAREN #MaxMinExpr
	| NUMBER #NumberExpr
	| IDENTIFIER #IdentifierExpr
	; 

/*
 * Lexer Rules
 */

NUMBER : INT ('.' INT)?; 
IDENTIFIER : [a-zA-Z]+[1-9][0-9]+;

INT : ('0'..'9')+;

MULTIPLY : '*';
DIVIDE : '/';
MOD: 'mod';
DIV: 'div';
SUBTRACT : '-';
ADD : '+';
LPAREN : '(';
RPAREN : ')';
MAX : 'MAX';
MIN : 'MIN';
UNARYPLUS : '++';
UNARYMINUS : '--';
DESP: ';'|',';

WS : [ \t\r\n] -> channel(HIDDEN);