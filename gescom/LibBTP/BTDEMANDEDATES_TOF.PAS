{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 03/09/2010
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTDEMANDEDATES ()
Mots clefs ... : TOF;BTDEMANDEDATES
*****************************************************************}
Unit BTDEMANDEDATES_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
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
     AGLInit,
     UTOF ;

Type
  TOF_BTDEMANDEDATES = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  	private
    	TOBData : TOB;
  end ;

function DemandeDateLivraison (TOBpiece : TOB) : Boolean;


Implementation
uses saisUtil;

function DemandeDateLivraison (TOBpiece : TOB) : Boolean;
var TobDates : TOB;
begin
  TOBDates := TOB.Create ('LES DATES', nil,-1);
  TOBDates.AddChampSupValeur('RETOUROK','-');
  TOBDates.AddChampSupValeur('DATFAC',TOBpiece.GetValue('GP_DATELIVRAISON'));
  TOBDates.AddChampSupValeur('DATESITUATION','-');
  TOBDates.AddChampSupValeur('TYPEDATE','Date de livraison');
  TRY
    TheTOB := TOBDates;
    AGLLanceFiche('BTP','BTDEMANDEDATES','','','');
    TheTOB := nil;
    if TOBDates.getValue('RETOUROK')='X' then
    begin
    	TOBpiece.PutValue('GP_DATELIVRAISON', TOBDates.GetDateTime('DATEFAC'));
    end;
  FINALLY
  	result := (TOBDates.getValue('RETOUROK')='X');
  	freeAndNil(TOBDates);
  END;
end;

procedure TOF_BTDEMANDEDATES.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTDEMANDEDATES.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTDEMANDEDATES.OnUpdate ;
begin
  Inherited ;
  if Not IsValidDate(GetControlText('DATEFAC')) Then
  begin
    PgiInfo ('La date est incorrecte. Merci de la modifier');
    Ecran.ModalResult := 0;
    Exit;
  end;
  if ( ControleDate (GetControlText('DATEFAC'))<>0) Then
  begin
    PgiInfo ('La date n''est pas dans un exercice ouvert. Merci de la modifier');
    Ecran.ModalResult := 0;
    Exit;
  end;
  TOBData.SetDateTime ('DATFAC',StrToDate(GetControlText('DATEFAC')));
  TOBData.putValue('RETOUROK','X');
end ;

procedure TOF_BTDEMANDEDATES.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTDEMANDEDATES.OnArgument (S : String ) ;
begin
  Inherited ;
  TOBData := LaTOB;
  if TOBData.getValue('DATFAC') <> iDate1900 then SetControlText('DATEFAC',TOBData.getValue('DATFAC'))
                                              else SetControlText('DATEFAC','  /  /    ');
  if TOBData.getValue('DATESITUATION')='X' then
  begin
  	ecran.Caption := TraduireMemoire ('Date de situation');
    SetControlCaption('TDATEFAC','Date situation');
  end;
  if TOBData.getValue('TYPEDATE')<>'' then
  begin
  	ecran.Caption := TraduireMemoire (TOBData.getValue('TYPEDATE'));
    SetControlCaption('TDATEFAC',TOBData.getValue('TYPEDATE'));
  end;
end ;

procedure TOF_BTDEMANDEDATES.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTDEMANDEDATES.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTDEMANDEDATES.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTDEMANDEDATES ] ) ; 
end.
