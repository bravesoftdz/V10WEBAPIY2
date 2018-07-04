{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 23/11/2017
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSAISMOIAVANC ()
Mots clefs ... : TOF;BTSAISMOIAVANC
*****************************************************************}
Unit BTSAISMOIAVANC_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     uDbxDataSet, 
     mul, 
{$else}
     eMul, 
{$ENDIF}
     uTob, 
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF ;

Type
  TOF_BTSAISMOIAVANC = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    TOBP : TOB;
  end ;

Implementation

procedure TOF_BTSAISMOIAVANC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISMOIAVANC.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISMOIAVANC.OnUpdate ;
var SQL,TheMois : string;
begin
  Inherited ;
  if (VALEUR(GetControlText('TMOIS')) = 0) or (VALEUR(GetControlText('ANNEE')) = 0) then
  begin
    PGIError('Vous devez renseigner une année et un mois de saisie');
    Ecran.ModalResult := 0;
    Exit;
  end;

  TheMois := GetControlText('ANNEE')+GetControlText('TMOIS');
  SQl := 'SELECT 1 FROM BAVANCTECH WHERE '+
         'BM9_NATUREPIECEG="'+TOBP.GetString('NATUREPIECEG')+'" AND '+
         'BM9_SOUCHE="'+TOBP.GetString('SOUCHE')+'" AND '+
         'BM9_NUMERO='+TOBP.GetString('NUMERO')+' AND '+
         'BM9_INDICEG='+TOBP.GetString('INDICEG')+' AND '+
         'BM9_MOIS >='+TheMois;
  if ExisteSql (Sql) then
  begin
    PGIError('Ce mois est déjà présent ou une période ultérieure à été renseignée');
    Ecran.ModalResult := 0;
    Exit;
  end;
  TOBP.SetString('MOIS',TheMois);
end ;

procedure TOF_BTSAISMOIAVANC.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISMOIAVANC.OnArgument (S : String ) ;
var YY,MM,JJ : Word;
  Mois : string;
begin
  Inherited ;
  TOBP := LaTOB;
  DecodeDate(Now,YY,MM,JJ);
  SetControlText('ANNEE',inttostr(YY));
  Mois := Format('%.02d',[MM]);
  SetControlText('TMOIS',Mois);
end ;

procedure TOF_BTSAISMOIAVANC.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISMOIAVANC.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISMOIAVANC.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTSAISMOIAVANC ] ) ; 
end.
