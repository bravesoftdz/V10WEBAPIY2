{***********UNITE*************************************************
Auteur  ......  :
Créé le ...... : 18/07/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : ELTNATDOSLISTE ()
Mots clefs ... : TOF;TOF_PGSAISLISTEELTNATDOS
*****************************************************************
}
Unit UTofPGSaisielisteEltNatDos;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ELSE}
      MaineAgl,
{$ENDIF}
     forms,ParamDat,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOB,UTOF,uTableFiltre,SaisieList,
     hTB97,PGOutilsFormation,EntPaie,LookUp;


Type
  TOF_PGSAISLISTEELTNATDOS = Class (TOF)

    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    Private
    TF :  TTableFiltre;

  end ;
var PGTobSalListe,PGTobTableDynListe : Tob;


Implementation


procedure TOF_PGSAISLISTEELTNATDOS.OnLoad ;
var i : Integer;
begin
  Inherited ;
  If (TF.Reccount = 0) and (TF.TypeAction <> TaConsult) then
  begin
    TF.Insert;
    SetFocuscontrol('PHD_NEWVALEUR');
  end;
end ;

procedure TOF_PGSAISLISTEELTNATDOS.OnArgument (S : String ) ;
var Libelle,Element,Salarie,Action,Niveau,DateVal : String;

begin
  Inherited ;
  TF  :=  TFSaisieList(Ecran).LeFiltre;
  Action := Trim(ReadTokenPipe(S, ';'));
  Niveau := Trim(ReadTokenPipe(S, ';'));
  Salarie := Trim(ReadTokenPipe(S, ';'));
  Element := Trim(ReadTokenPipe(S, ';'));
  DateVal := Trim(ReadTokenPipe(S, ';'));
  Libelle := Trim(ReadTokenPipe(S, ';'));
  TFSaisieList(Ecran).Caption := 'Saisie de l''élément : '+Libelle;
  UpdateCaption(TFSaisieList(Ecran));
  TF.WhereTable := 'WHERE PED_TYPENIVEAU="SAL" AND PED_VALEURNIVEAU="'+Salarie+'" AND PED_CODEELT="'+Element+'"';
  TFSaisieList(Ecran).Retour := '';
end ;

Initialization
  registerclasses ( [ TOF_PGSAISLISTEELTNATDOS ] ) ;
end.

