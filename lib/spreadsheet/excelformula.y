class ExcelFormulaParser

prechigh
  nonassoc UMINUS
  right    '^'
  left     '&'
  left     '*' '/'
  left     '+' '-'
  left     '<' '>' '<=' '>=' '<>'
preclow

rule

  formula      : expr_list
  
  expr_list    :
               | expr_list expr EOL
               | expr_list EOL

  expr         : expr '+' expr
               | expr '-' expr
               | expr '*' expr
               | expr '/' expr
               | expr '^' expr
               | expr '&' expr
               | expr LT  expr
               | expr GT  expr
               | expr LE  expr
               | expr GE  expr
               | expr NE  expr
               | primary

  primary      : '(' expr ')'
               | '-' expr  = UMINUS
               | FUNC
               | NUMBER
               | STRING
               | REF2D
               | REF3D
               | RANGE2D
               | RANGE3D
               | funcall

  funcall      : FUNC '(' args ')'
               | FUNC '(' ')'

  args         : expr
               | args ',' expr
end
