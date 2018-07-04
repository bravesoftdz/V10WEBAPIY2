unit UtilsTOB;

interface
uses Classes,uTob,Variants,HMsgBox

;

type
  TresultCmp = (TrcEqual,TrcLess,TrcMore);

function TOBFindInto(TheTOB : TOB; Clef : Array of string; TheKeys: array of variant): Integer;
function TOBDeleteInto(TheTOB : TOB; Clef : array of string; TheKeys: array of Variant): Boolean;
function TOBInsertInto (TOBFather,TOBInsert: TOB; Clef: array of string) : Integer;
function TOBGetInto(TheTOB : TOB; Clef : Array of string;TheKeys: array of variant): TOB;

implementation

uses SysUtils;


procedure ErrClefNoDef (debug : boolean = False);
begin
  if debug then PgiError('ATTENTION : Vous n''avez pas définie de clef');
end;

procedure ErrNbElements (debug : boolean = False);
begin
	if debug then PgiError('ATTENTION : le nombre d''argument ne correspondent pas à ceux de la clef');
end;

procedure ErrTOBNotOk (debug : boolean = False);
begin
	if debug then PgiError('ATTENTION : La TOB n''est pas du type défini');
end;

procedure ErrValueExist (debug : boolean = False);
begin
	if debug then PgiError('ATTENTION : Vous tentez d''inserer un valeur existante..');
end;


function OkKeys(Clef : Array of string; TheKeys: array of Variant): Boolean;
begin
	result := (High(TheKeys) =High(Clef));
end;

function ConstitueCleLoc(TheInsert: TOB; Clef : array of string ; var ClefLoc: array of Variant): boolean;
var i : Integer;
		NbClef : Integer;
begin
  Result := True;
  NbClef := High(Clef) ;
	for I := 0 to NbClef do
  begin
		if not TheInsert.FieldExists(clef[I]) then
    begin
      Result := false;
      Break;
    end;
    ClefLoc[i] := TheInsert.GetValue(clef[I]);
  end;
end;

function CmpValues(TOBCmp: TOB; Clef: array of string; Values: array of Variant): TresultCmp;
var NbClef : Integer;
		I : Integer;
begin
  Result := TrcEqual;
  NbClef := High(Clef) ;
  for i := 0 to NbClef do
  begin
    if TOBCmp.GetValue(Clef[i]) > Values[I] then BEGIN result := TrcMore; Break; end;
    if TOBCmp.GetValue(Clef[i]) < Values[I] then BEGIN result := TrcLess; Break; end;
  end;
end;

function TOBFindIntoGt(TheTOB : TOB; Clef : array of string; TheKeys: array of Variant): Integer;
var Milieu   : integer;
    Ok_Find  : Boolean;
    BorneInf : Integer;
    BorneSup : Integer;
    ResultCmp : TresultCmp;
begin
  Result := -1;
  if TheTOB= nil then Exit;
	if High(clef)=-1 then BEGIN ErrClefNoDef; Exit; end; // Si aucune clef définie on sort
  if not OkKeys (Clef, TheKeys) then BEGIN ErrNbElements; Exit; end; // Si le nombre d'elements ne correspondent pas a celle de la clef

  Ok_Find   := false;
  //
  if TheTOB.detail.count = 0 then Exit;
  //
  BorneInf  := 0;
  BorneSup  := TheTOB.detail.count-1;
  // recherche sur le 1er au cas ou
  ResultCmp := CmpValues (TheTOB.detail[0],Clef,TheKeys);
  if ResultCmp = TrcMore then
  begin
  	Result :=BorneInf;
  end else if ResultCmp = trcEqual then
  begin
  	Result := -2;
    exit;
  end;
  // recherche sur le dernier au cas ou
  ResultCmp := CmpValues (TheTOB.detail[BorneSup],Clef,TheKeys);
  if ResultCmp = TrcMore then
  begin
  	Result :=BorneSup;
  end else if ResultCmp = trcEqual then
  begin
  	Result := -2;
    exit;
  end;
  if TheTOB.Detail.Count < 3 then Exit;
  //
  //Boucle de recherche
  Repeat
    Milieu := (Borneinf + Bornesup) div 2;
  	if (BorneInf=Milieu) Or (BorneSup=Milieu) then Break;
    ResultCmp := CmpValues (TheTOB.detail[Milieu],Clef,TheKeys);
    if ResultCmp = trcEqual Then
    begin
  		Result := -2;
      break;
    end else
    begin
      if ResultCmp = TrcMore then
      begin
        Result := Milieu;   // on stocke le dernier supérieur
        BorneSup := Milieu
      end else
      begin
        BorneInf := Milieu;
      end;
    end;
  until (Ok_Find);
end;

function TOBFindInto(TheTOB : TOB; Clef : Array of string; TheKeys: array of variant): Integer;
var Milieu   : integer;
    Ok_Find  : Boolean;
    BorneInf : Integer;
    BorneSup : Integer;
    ResultCmp : TresultCmp;
begin
  Result := -1;
  if TheTOB = nil then Exit;
	if High(clef)=-1 then BEGIN ErrClefNoDef; Exit; end; // Si aucune clef définie on sort
  if not OkKeys (Clef, TheKeys) then BEGIN ErrNbElements; Exit; end; // Si le nombre d'elements ne correspondent pas a celle de la clef

  Ok_Find   := false;
  //
  if TheTOB.detail.count = 0 then Exit;
  //
  BorneInf  := 0;
  BorneSup  := TheTOB.detail.count-1;
  // recherche sur le 1er au cas ou
  if CmpValues (TheTOB.detail[0],Clef,TheKeys) = trcEqual then
  begin
  	result := 0;
    exit;
  end;
  // recherche sur le dernier au cas ou
  if CmpValues (TheTOB.detail[BorneSup],Clef, TheKeys) = trcEqual then
  begin
  	result := TheTOB.Detail[Bornesup].GetIndex;
    exit;
  end;

  //Boucle de recherche
  Repeat
    Milieu := (Borneinf + Bornesup) div 2;
    ResultCmp := CmpValues (TheTOB.detail[Milieu],Clef,TheKeys);
    if ResultCmp = trcEqual Then
    begin
      result := TheTOB.Detail[milieu].GetIndex;
      break;
    end else
    begin
    	if (BorneInf=Milieu) Or (BorneSup=Milieu) then Break;
      if ResultCmp = TrcMore then
        BorneSup := Milieu
      else
        BorneInf := Milieu;
    end;
  until (Ok_Find);
end;

function TOBDeleteInto(TheTOB : TOB; Clef : array of string; TheKeys: array of Variant): Boolean;
var Indice : Integer;
begin
  Result := false;
  if TheTOB = nil then Exit;
	if High(clef)=-1 then BEGIN ErrClefNoDef; Exit; end; // Si aucune clef définie on sort
  if not OkKeys (Clef, TheKeys) then BEGIN ErrNbElements; Exit; end; // Si le nombre d'elements ne correspondent pas a celle de la clef
  Indice := TOBFindInto(TheTOB,Clef,TheKeys);
  if Indice >= 0 then TheTOB.detail[Indice].Free;
end;

function TOBInsertInto (TOBFather,TOBInsert: TOB; Clef: array of string) : Integer;
var ClefLoc : array of Variant;
begin
  Result := -1;
  if TOBFather = nil then Exit;
	if High(clef)=-1 then BEGIN ErrClefNoDef; Exit; end; // Si aucune clef définie on sort
  TRY
		SetLength(ClefLoc,High(Clef)+1); // Définition du nombre de clef
  	if not ConstitueCleLoc (TOBInsert,Clef,ClefLoc) then BEGIN ErrTOBNotOk; Exit; END; // erreur
		Result := TOBFindIntoGt(TOBFather,Clef,ClefLoc); if Result = -2 then BEGIN ErrValueExist; Exit; end;
    TOBInsert.ChangeParent(TOBFather,Result);
  finally
    SetLength (ClefLoc,0);
  end;
end;


function TOBGetInto(TheTOB : TOB; Clef : Array of string;TheKeys: array of variant): TOB;
var Indice : Integer;
begin
  Result := nil;
  if TheTOB = nil then Exit;
	if High(clef)=-1 then BEGIN ErrClefNoDef; Exit; end; // Si aucune clef définie on sort
  if not OkKeys (Clef,TheKeys) then BEGIN ErrNbElements; Exit; end; // Si le nombre d'elements ne correspondent pas a celle de la clef
	Indice := TOBFindInto(TheTOB,Clef,TheKeys);
  if Indice >=0 then result := TheTOB.detail[Indice];
end;

end.
