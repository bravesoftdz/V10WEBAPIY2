{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 03/06/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AMREGROUPEMENT ()
Mots clefs ... : TOF;AMREGROUPEMENT
Suite ........ : BTY 07/05 FQ 16284 Gérer montants selon nb décimales de la monnaie du dossier
*****************************************************************}
Unit AMREGROUPEMENT_TOF ;

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
     MaineAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     uTOB ;

Type
  TOF_AMREGROUPEMENT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

procedure AMLanceFiche_FicheRegroupement(Lequel: string);

Implementation

procedure AMLanceFiche_FicheRegroupement(Lequel: string);
begin
  AGLLanceFiche('AM', 'AMREGROUPEMENT', '', '', Lequel);
end;

procedure TOF_AMREGROUPEMENT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AMREGROUPEMENT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AMREGROUPEMENT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AMREGROUPEMENT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AMREGROUPEMENT.OnArgument (S : String ) ;
var T : TOB;
    StLibelle : string;
begin
  Inherited ;
  T := TOB.Create ('IMMO',nil,-1);
  try
    if (T.SelectDB('"'+S+'"',nil,True)) then
    begin
      StLibelle := Copy ( S, 11, 3 );   // à revoir pour faire ça plus proprement
      Ecran.Caption := stLibelle + ' - '+ RechDom('AMREGROUPEMENT',StLibelle,False);
      SetControlCaption('LI_COMPTEIMMO',GetControlText('LI_COMPTEIMMO')+T.GetValue('I_COMPTEIMMO'));
      SetControlCaption('LI_LIEUGEO',GetControlText('LI_LIEUGEO')+T.GetValue('I_LIEUGEO'));
      SetControlCaption('LI_CODEPOSTAL',GetControlText('LI_CODEPOSTAL')+T.GetValue('I_CODEPOSTAL'));
      SetControlCaption('LI_VILLE',GetControlText('LI_VILLE')+T.GetValue('I_VILLE'));
      SetControlCaption('LI_PAYS',GetControlText('LI_PAYS')+T.GetValue('I_PAYS'));
      SetControlCaption('LI_ETABLISSEMENT',GetControlText('LI_ETABLISSEMENT')+T.GetValue('I_ETABLISSEMENT'));
      SetControlCaption('LI_METHODEECO',GetControlText('LI_METHODEECO')+T.GetValue('I_METHODEECO'));
      SetControlCaption('LI_DUREEECO',GetControlText('LI_DUREEECO')+IntToStr(T.GetValue('I_DUREEECO'))+' '+TraduireMemoire ('mois'));
      SetControlCaption('LI_DATEDEBUTAMORT',GetControlText('LI_DATEDEBUTAMORT')+DateToStr(T.GetValue('I_DATEAMORT')));
      SetControlCaption('LI_DATEFINAMORT',GetControlText('LI_DATEFINAMORT')+DateToStr(T.GetValue('I_DATEDERMVTECO')));
      // BTY 07/05 Fiche 16284 Nb décimales
      //SetControlCaption('LI_QUOTEPART',GetControlText('LI_QUOTEPART')+StrFMontant(T.GetValue('I_QUOTEPART'),13,2,'',True));
      //SetControlCaption('LI_BASEDEDUCTIBLE',GetControlText('LI_BASEDEDUCTIBLE')+StrFMontant(T.GetValue('I_REINTEGRATION'),13,2,'',True));
      SetControlCaption('LI_QUOTEPART',GetControlText('LI_QUOTEPART')+StrFMontant(T.GetValue('I_QUOTEPART'),13,V_PGI.OkDecV,'',True));
      SetControlCaption('LI_BASEDEDUCTIBLE',GetControlText('LI_BASEDEDUCTIBLE')+StrFMontant(T.GetValue('I_REINTEGRATION'),13,V_PGI.OkDecV,'',True));
      SetControlCaption('LI_TABLE0',GetControlText('LI_TABLE0')+T.GetValue('I_TABLE0'));
      SetControlCaption('LI_TABLE1',GetControlText('LI_TABLE1')+T.GetValue('I_TABLE1'));
    end;
  finally
    T.Free;
  end;
end ;

procedure TOF_AMREGROUPEMENT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AMREGROUPEMENT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AMREGROUPEMENT.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_AMREGROUPEMENT ] ) ;
end.
