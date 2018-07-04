{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 01/04/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : TIERSFRAIS (TIERSFRAIS)
Mots clefs ... : TOM;TIERSFRAIS
*****************************************************************}
Unit TIERSFRAIS_TOM ;

Interface

Uses
  StdCtrls,
  Controls,
  Classes,

  {$IFNDEF EAGLCLIENT}
  db,
  dbTables,
  Fe_Main,
  Fiche,
  FichList,
  Mul,
  {$ELSE}
  MainEagl,
  eFiche,
  eFichList,
  eMul,
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOM,
  UTob
  ;

Type
  TOM_TIERSFRAIS = Class (TOM)
    procedure OnArgument ( S: String )   ; override ;
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
  private
    NatureAuxi, Tiers            : String;
    procedure SetLastError (Num : integer; ou : string );

  end ;

Const
	TexteMessage: array[1..1] of string = (
	  {1}  	''
		);

  ErrCodePort       = 1;

Implementation

Uses
  LookUp,
  wCommuns
  ;

procedure TOM_TIERSFRAIS.OnArgument ( S: String ) ;
var
  Critere: string;
  ChampMul, ValMul: string;
  x: integer;
begin
  Inherited ;
  { Récup des arguments }
  repeat
    Critere:=Trim(ReadTokenSt(S));
    if Critere<>'' then
    begin
      x:=pos('=',Critere);
      if x<>0 then
      begin
        ChampMul := UpperCase(copy(Critere,1,x-1));
        ValMul := Copy(Critere,x+1,length(Critere));
        if (ChampMul = 'NATUREAUXI') then NatureAuxi := ValMul;
        if (ChampMul = 'TIERS') then Tiers := ValMul;
      end;
    end;
  until Critere = '';
end ;

procedure TOM_TIERSFRAIS.OnNewRecord ;
begin
  Inherited ;
  { init clé}
  SetField('GTF_TIERS',Tiers);
  SetField('GTF_NATUREAUXI',NatureAuxi);
end ;

procedure TOM_TIERSFRAIS.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_TIERSFRAIS.OnUpdateRecord ;
begin
  Inherited;
  if Assigned(Ecran) then
  begin
    { Contrôles de saisie }
    if not LookupValueExist(GetControl('GTF_CODEPORT')) then
    begin
      SetLastError(ErrCodePort, 'GTF_CODEPORT');
      EXIT;
    end;
  end;
end ;

procedure TOM_TIERSFRAIS.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_TIERSFRAIS.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_TIERSFRAIS.OnChangeField ( F: TField ) ;
begin
  Inherited ;
  If (F.FieldName='GTF_CODEPORT') then
  begin
    if (F.FieldName<> '') and (not LookupValueExist(GetControl('GTF_CODEPORT')) ) then
      SetLastError(ErrCodePort, 'GTF_CODEPORT')
    else
    begin
      SetField('GTF_LIBELLE',RECHDOM('GCPORT',GetField('GTF_CODEPORT'),false) );
    end;
  end;
end ;


procedure TOM_TIERSFRAIS.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_TIERSFRAIS.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_TIERSFRAIS.SetLastError(Num: integer; ou: string);
begin
  if ou<>'' then SetFocusControl(ou);
  LastError := Num;
  LastErrorMsg := TexteMessage[LastError];
end;

Initialization
  registerclasses ( [ TOM_TIERSFRAIS ] ) ; 
end.
