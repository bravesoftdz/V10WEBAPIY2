{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 03/03/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : AFINDICE (AFINDICE)
Mots clefs ... : TOM;AFINDICE
*****************************************************************}
Unit uTomAfIndice ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,
{$Else}
     MainEagl,
{$ENDIF}
     forms, 
     sysutils,
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOM,
     UTob ;

Type
  TOM_AFINDICE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    end ;
const
	TexteMsg: array[1..2] of string 	= (
          {1}        'Suppression impossible certaines valeurs sont saisies sur ce code',
          {2}        'Suppression impossible ce code est utilisé dans une formule'

          );

procedure AglLanceFicheAFINDICE(cle,Action : string ) ;

Implementation

procedure TOM_AFINDICE.OnNewRecord ;
begin
  Inherited ;
  SetField('AIN_INDDATECREA',date());
  SetField('AIN_INDDATEFIN', iDate2099);
  SetControlText('AIN_INDMAJ_LIB', 'Saisie');
  SetField('AIN_INDMAJ', 'SAI');
  SetField('AIN_DATEMAJ',date());
end ;
 
procedure TOM_AFINDICE.OnDeleteRecord ;
var st : string ;
    i  : integer ;
begin
  Inherited ;
  st:='SELECT * FROM afvalindice WHERE afv_indcode = "'+getfield('AIN_INDCODE')+'"' ;
  If ExisteSQL(st) then
     begin
     LastError    := 1;
     LastErrorMsg := TexteMsg[LastError];
     end;
   st:='select * from afformule Where ';
   for i:=1 to 10 do
    begin
     St:=st+' AFE_INDCODE'+inttostr(i)+'="'+getfield('AIN_INDCODE')+'"' ;
     if i<10 then St:=St+' Or ' ;
     end ;
  If ExisteSQL(st) then
     begin
     LastError    := 2;
     LastErrorMsg := TexteMsg[LastError];
     end;
end ;

procedure TOM_AFINDICE.OnUpdateRecord ;
begin
  Inherited;
end;         

procedure TOM_AFINDICE.OnAfterUpdateRecord ;
begin
  Inherited;
end;

procedure TOM_AFINDICE.OnLoadRecord ;
begin
  Inherited;
  SetFocusControl('AIN_INDCODE');
  SetControlText('AIN_INDMAJ_LIB', RechDom('AFTMAJINDICE', GetField('AIN_INDMAJ'), false));
end;

procedure TOM_AFINDICE.OnChangeField ( F: TField ) ;
begin
  Inherited ;

end ;

procedure TOM_AFINDICE.OnArgument ( S: String ) ;

begin
  Inherited ;
end ;

procedure TOM_AFINDICE.OnClose ;
begin
  Inherited ;
  AvertirTable('AFINDICE');
end ;

procedure TOM_AFINDICE.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure AglLanceFicheAFINDICE(cle,Action : string ) ;
begin
  AglLanceFiche('AFF','AFINDICE','',cle,Action);
end ;


Initialization
  registerclasses ( [ TOM_AFINDICE ] ) ;
end.
