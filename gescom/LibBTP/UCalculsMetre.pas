unit UCalculsMetre;

interface
uses Classes,sysutils,UTOB,Hent1;

type
  TNodeType = (ntOperator, ntValue);
  TOpType = (otPlus, otMinus, otMul, otDiv, otPower, otFunction);
  TFuncType = (ftExp, ftLn, ftSin, ftCos, ftTan, ftNeg, ftFact, ftDeg, ftSqr,
    ftSqrt);

  PNode = ^TNode;
  TNode = packed record
    NodeType: TNodeType;
    case TNodeType of
      ntOperator:
      (
        Operator: TOpType;
        case TOpType of
          otPlus, otMinus, otMul, otDiv, otPower:
          (
            Left: PNode;
            Right: PNode;
          );
          otFunction:
          (
            Func: TFuncType;
            Param: PNode;
          );
      );
      ntValue:
      (
        Value: Real;
      );
  end;


type
  TByteCodeInstr = (biPush, biCall);
  TByteCodeCall = (boExp, boLn, boSin, boCos, boTan, boNeg, boFact, boDeg, boSqr,
    boSqrt, boPlus, boMinus, boMul, boDiv, boPower);

const
  BCCallStr: array [TByteCodeCall] of string =
    ('exp', 'ln', 'sin', 'cos', 'tan', 'neg', 'fact', 'deg', 'sqr',
     'sqrt', 'add', 'sub', 'mul', 'div', 'pow');
  BCInstrStr: array [TByteCodeInstr] of string =
    ('push', 'call');


type
{ Packing this record may affect the virtual machine performance }
  TByteCode = {$IFDEF PACKCODE} packed {$ENDIF} record
    Instr: TByteCodeInstr;
    case TByteCodeInstr of
      biPush: (Value: Real);
      biCall: (Oper: TByteCodeCall);
  end;

  TCompiledExpr = array of TByteCode;

function CalculeFormuleLigne (chaine : string) : double;

implementation

const
  N = 100000;


function Fact(N: double): double;
begin
  if N > 0 then
    Result := N * Fact(N - 1)
  else
    Result := 1;
end;

function Compile(var Expr: PChar): PNode;

procedure Error;
begin
  Expr := #0;
end;

procedure SkipBlanks;
begin
  while (Expr^ <> #0) and (Expr^ in [#32, #9]) do
    Inc(Expr);
end;

function Constant: PNode;
var
  Str: string;
  Err: Integer;
begin
  SkipBlanks;
  New(Result);
  Result^.NodeType := ntValue;
  Str := '';
  if Expr^ in ['0'..'9'] then
  begin
    while (Expr^ <> #0) and (Expr^ in ['0'..'9']) do
    begin
      Str := Str + Expr^;
      Inc(Expr);
    end;
    if Expr^ = '.' then
    begin
      Str := Str + Expr^;
      Inc(Expr);
      while (Expr^ <> #0) and (Expr^ in ['0'..'9']) do
      begin
        Str := Str + Expr^;
        Inc(Expr);
      end;
    end;
    Val(Str, Result^.Value, Err);
  end
  else
  if Expr^ in ['a'..'z', 'A'..'Z'] then
  begin
    while (Expr^ <> #0) and (Expr^ in ['a'..'z', 'A'..'Z']) do
    begin
      Str := Str + UpCase(Expr^);
      Inc(Expr);
    end;
    if Str = 'PI' then
      Result^.Value := Pi
    else
    if Str = 'E' then
      Result^.Value := Exp(1);
  end
  else
  begin
    Result^.Value := 0;
  end;
end;

function Atom: PNode;
begin
  SkipBlanks;
  if Expr^ = '(' then
  begin
    Inc(Expr);
    Result := Compile(Expr);
    SkipBlanks;
    if Expr^ = ')' then
      Inc(Expr);
  end else
  begin
    Result := Constant;
  end;
end;

function RightUnary: PNode;
var
  N: PNode;
begin
  Result := Atom;
  SkipBlanks;               
  while (Expr^ <> #0) and (Expr^ in ['!', '°', '²']) do
  begin
    New(N);
    N^.NodeType := ntOperator;
    N^.Operator := otFunction;
    case Expr^ of
      '!' : N^.Func := ftFact;
      '°' : N^.Func := ftDeg;
      '²' : N^.Func := ftSqr;
    end;
    N^.Param := Result;
    Result := N;
    Inc(Expr);
    SkipBlanks;
  end;
end;

function LeftUnary: PNode;
var
  P: PChar;
  Str: string;
  F: TFuncType;
begin
  SkipBlanks;
  if Expr^ in ['+', '-'] then
  begin
    if Expr^ = '-' then
    begin
      Inc(Expr);
      New(Result);
      Result^.NodeType := ntOperator;
      Result^.Operator := otFunction;
      Result^.Func := ftNeg;
      Result^.Param := LeftUnary;
    end else
    begin
      Inc(Expr);
      Result := LeftUnary;
    end;
  end
  else
  if Expr^ in ['a'..'z', 'A'..'Z'] then
  begin
    P := Expr;
    Str := '';
    while (Expr^ <> #0) and (Expr^ in ['a'..'z', 'A'..'Z']) do
    begin
      Str := Str + UpCase(Expr^);
      Inc(Expr);
    end;
    if Str = 'EXP' then
      F := ftExp
    else
    if Str = 'LN' then
      F := ftLn
    else
    if Str = 'SIN' then
      F := ftSin
    else
    if Str = 'COS' then
      F := ftCos
    else
    if Str = 'TAN' then
      F := ftTan
    else
    if Str = 'SQRT' then
      F := ftSqrt
    else
    begin
      Expr := P; 
      Result := RightUnary;
      Exit;
    end;
    New(Result);
    Result^.NodeType := ntOperator;
    Result^.Operator := otFunction;
    Result^.Func := F;
    Result^.Param := LeftUnary;
  end
  else
  begin
    Result := RightUnary;
  end;
end;

function Power: PNode;
var
  N: PNode;
begin
  Result := LeftUnary;
  SkipBlanks;
  while (Expr^ <> #0) and (Expr^ = '^') do
  begin
    Inc(Expr);
    New(N);
    N^.NodeType := ntOperator;
    N^.Operator := otPower;     
    N^.Left := Result;
    N^.Right := LeftUnary;
    Result := N;
    SkipBlanks;
  end;
end;

function Factor: PNode;
var
  N: PNode;
begin
  Result := Power;
  SkipBlanks;
  while (Expr^ <> #0) and (Expr^ in ['*', '/']) do
  begin     
    New(N);
    case Expr^ of
      '*': N^.Operator := otMul;
      '/': N^.Operator := otDiv;
    end;
    Inc(Expr);
    N^.NodeType := ntOperator;
    N^.Left := Result;
    N^.Right := Power;
    Result := N;
    SkipBlanks;
  end;
end;

var
  N: PNode;

begin
  Result := Factor;
  SkipBlanks;
  while (Expr^ <> #0) and (Expr^ in ['+', '-']) do
  begin
    New(N);
    case Expr^ of
      '-': N^.Operator := otMinus;
      '+': N^.Operator := otPlus;
    end;
    Inc(Expr);
    N^.NodeType := ntOperator;
    N^.Left := Result;
    N^.Right := Factor;
    Result := N;
    SkipBlanks;
  end;
end;

{ Console listing }

procedure Display(var Code: TCompiledExpr);
var
  i: Integer;
begin
  for i := 0 to Length(Code) - 1 do
  begin
    Write(BCInstrStr[Code[i].Instr], '      ');
    case Code[i].Instr of
      biCall: Writeln(BCCallStr[Code[i].Oper]);
      biPush: Writeln(Code[i].Value:1:5);
    end;
  end;
end;

{ Code generation }

procedure Generate(Root: PNode; var Code: TCompiledExpr);

  procedure AddPush(What: double);
  begin
    SetLength(Code, Length(Code) + 1);
    with Code[Length(Code) - 1] do
    begin
      Instr := biPush;
      Value := What;
    end;
  end;

  procedure AddCall(Call: TByteCodeCall);
  begin
    SetLength(Code, Length(Code) + 1);
    with Code[Length(Code) - 1] do
    begin
      Instr := biCall;
      Oper := Call;
    end;
  end;

begin
  if Root^.NodeType = ntValue then
  begin
    AddPush(Root^.Value);
  end
  else
  begin
    if Root^.Operator = otFunction then
    begin
      Generate(Root^.Param, Code);
      case Root^.Func of
        ftExp  : AddCall(boExp);
        ftLn   : AddCall(boLn);
        ftSin  : AddCall(boSin);
        ftCos  : AddCall(boCos);
        ftTan  : AddCall(boTan);
        ftNeg  : AddCall(boNeg);
        ftFact : AddCall(boFact);
        ftDeg  : AddCall(boDeg);
        ftSqr  : AddCall(boSqr);
        ftSqrt : AddCall(boSqrt);
      end;
    end
    else
    begin
      Generate(Root^.Right, Code);
      Generate(Root^.Left, Code);       
      case Root^.Operator of
        otPlus     : AddCall(boPlus);
        otMinus    : AddCall(boMinus);
        otMul      : AddCall(boMul);
        otDiv      : AddCall(boDiv);
        otPower    : AddCall(boPower);
      end;
    end;
  end;
end;

function Tan(X: double): double;
begin
  Result := Sin(X) / Cos(X);
end;


function Eval(var Code: TCompiledExpr): double;

const
  STACK_SIZE = 200;

var
  i, Peek: Integer;
  Stack: array [0..STACK_SIZE - 1] of double;

  procedure Push(Value: double);
  begin
    Inc(Peek);
    Stack[Peek] := Value;
  end;

  function Pop: double;
  begin
    Result := Stack[Peek];
    Dec(Peek);
  end;

begin
  Peek := -1;
  for i := 0 to Length(Code) - 1 do
  case Code[i].Instr of
    biPush: Push(Code[i].Value);
    biCall:
    case Code[i].Oper of
      boPlus  : Push(Pop + Pop);
      boMinus : Push(Pop - Pop);
      boMul   : Push(Pop * Pop);
      boDiv   : Push(Pop / Pop);
      boPower : Push(Exp(Pop * Ln(Pop)));
      boExp   : Push(Exp(Pop));
      boLn    : Push(Ln(Pop));
      boSin   : Push(Sin(Pop));
      boCos   : Push(Cos(Pop));
      boTan   : Push(Tan(Pop));
      boNeg   : Push(-Pop);
      boFact  : Push(Fact(Trunc(Pop)));
      boDeg   : Push(Pop * Pi / 180);
      boSqr   : Push(Sqr(Pop));
      boSqrt  : Push(Sqrt(Pop));
    end;
  end;
  Result := Pop;
end;

procedure FreeTree(Root: PNode);
begin
  if Root = nil then
    Exit;
  if Root^.NodeType = ntOperator then
  begin
    if Root^.Operator = otFunction then
    begin
      FreeTree(Root^.Param);
    end
    else
    begin
      FreeTree(Root^.Left);
      FreeTree(Root^.Right);
    end;
  end;
  Dispose(Root);
end;

function TreeSize(Root: PNode): Integer;
begin
  if Root = nil then
  begin
    Result := 0;
    Exit;
  end;
  Result := SizeOf(TNode);
  if Root^.NodeType = ntOperator then
  begin
    if Root^.Operator = otFunction then
    begin
      Result := Result + TreeSize(Root^.Param);
    end else
    begin
      Result := Result + TreeSize(Root^.Left);
      Result := Result + TreeSize(Root^.Right);
    end;
  end;
end;

function CalculeFormuleLigne (chaine : string) : double;
var strLoc : string;
		IPosdeuxpoint : integer;
    P: PChar;
    Root: PNode;
    Code: TCompiledExpr;
begin
  IPosDeuxPoint := Pos (':',Chaine);
  if IposDeuxPoint > 0 then StrLoc := copy(Chaine,IPosdeuxpoint+1,255)
  										 else StrLoc := Chaine;

  P := PChar(StrLoc);
  Root := Compile(P);
  SetLength(Code, 0);
  Generate(Root, Code);
  FreeTree(Root);
  Result := Eval(Code);
end;

end.
