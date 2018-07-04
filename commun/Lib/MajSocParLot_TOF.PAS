{***********UNITE*************************************************
Auteur  ...... : JS
Créé le ...... : 29/07/2003
Modifié le ... :  29/07/03
Description .. : Source TOF de la FICHE : MAJSOCPARLOT ()
Mots clefs ... : TOF;MAJSOCPARLOT
*****************************************************************}
Unit MajSocParLot_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,FE_Main,
  {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
     MajTable,
{$ELSE}
     UtileAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1, galSystem,
     PGIAppli,
{$IFDEF VER150}
     variants,
{$ENDIF}
     HMsgBox,HTB97,HPanel, HsysMenu,
     UTOF;

Type
  TOF_MAJSOCPARLOT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

    private
      Stop : Boolean;
      Recap : TListBox;
      GSoc : THGrid;
      HMTrad: THSystemMenu;
      procedure bStopClick(Sender: TObject);
      function  Traitement(libdoss: String; var compteur: Integer ; toMinutes : integer): Boolean;
      procedure TraiteDossiers;
      procedure Affiche(msg: String);
  end ;

Function YYLanceFiche_MajSocParLot(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Implementation

uses UtilPgi; //js1 120107 pour utilisation de VersionBase

const NomExec = 'PGIMajVer.exe';

Function YYLanceFiche_MajSocParLot(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_MAJSOCPARLOT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_MAJSOCPARLOT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_MAJSOCPARLOT.OnUpdate ;
begin
  Inherited;
  SaveSynRegKey('MajSocParLot_HrDebut',GetControlText('HRDEBUT'),True);
  SaveSynRegKey('MajSocParLot_HrFin',GetControlText('HRFIN'),True);
  SaveSynRegKey('MajSocParLot_NbErreurs',GetControlText('NBERREURS'),True);
  SaveSynRegKey('MajSocParLot_TimeOutMn',GetControlText('TIMEOUTMN'),True);
  SaveSynRegKey('MajSocParLot_ChkDebut',GetControlText('CHKDEBUT'),True);
  Recap.Clear;
  if PGIAsk('Confirmez vous la mise à jour ?') = mrYes then TraiteDossiers;
end ;

procedure TOF_MAJSOCPARLOT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_MAJSOCPARLOT.OnArgument (S : String ) ;
var stKey : string;
    ListSoc : TListBox;
    iInd : integer;
begin
  Inherited ;
  Stop := False;
  ListSoc := TListBox(GetControl('LBSOCIETE'));
  GSoc := THGrid(GetControl('GSOC'));
  Recap := TListBox(GetControl('LBRECAP'));
  GSoc.Cells[0,0] := TraduireMemoire('Société');
  GSoc.Cells[1,0] := TraduireMemoire('Version');
  ChargeDossier(ListSoc.Items,True);       
  for iInd := 0 to ListSoc.Items.Count -1 do
  begin
    GSoc.RowCount := GSoc.RowCount + 1;
    GSoc.Cells[0,iInd+1] := ListSoc.Items[iInd];
    GSoc.Cells[1,iInd+1] := VersionBase(ListSoc.Items[iInd]);
    if (VersionBase(ListSoc.Items[iInd]) = -1) then Affiche('Dossier '+ListSoc.Items[iInd]+' inaccessible.');
  end;
  GSoc.RowCount := GSoc.RowCount -1;
  HMTrad.ResizeGridColumns(GSoc);
  TToolbarButton97(GetControl('BSTOP')).OnClick := bStopClick;


  // Valeurs par défaut
  SetControlText('DATDEBUT',DateToStr(Date));
  SetControlText('DATFIN',DateToStr(Date+1));
  stKey := GetSynRegKey('MajSocParLot_HrDebut','',True);
 
 

  if not IsValidHeureMinute(stKey) then stKey := '20:00';
  SetControlText('HRDEBUT',stKey);
  stKey := GetSynRegKey('MajSocParLot_HrFin','',True);
  if not IsValidHeureMinute(stKey) then stKey := '08:00';
  SetControlText('HRFIN',stKey);
  stKey := GetSynRegKey('MajSocParLot_NbErreurs','',True);
  if stKey = '' then stKey := '10';
  SetControlText('NBERREURS',stKey);
  stKey := GetSynRegKey('MajSocParLot_TimeOutMn','',True);
  if stKey = '' then stKey := '15';
  SetControlText('TIMEOUTMN',stKey);
  stKey := GetSynRegKey('MajSocParLot_ChkDebut','',True);
  if stKey = '' then stKey := 'X';
  SetControlText('CHKDEBUT',stKey);
end ;

procedure TOF_MAJSOCPARLOT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_MAJSOCPARLOT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_MAJSOCPARLOT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_MAJSOCPARLOT.bStopClick(Sender: TObject);
begin
  Stop := True;
end;

procedure TOF_MAJSOCPARLOT.TraiteDossiers;
var libdoss: String;
    iInd, nberr, cpteur, nberrtot, tominutes : Integer;
    DatDebut, DatFin: TDateTime;
begin
  Application.ProcessMessages; // pour appui sur stop
  if Stop then begin Affiche('Traitement interrompu par l''utilisateur'); exit; end;
  nberr := 0;
  nberrtot := ValeurI(GetControlText('NBERREURS'));
  tominutes := ValeurI(GetControlText('TIMEOUTMN'));
  DatDebut := StrToDateTime(GetControlText('DATDEBUT') + ' ' + GetControlText('HRDEBUT'));
  DatFin := StrToDateTime(GetControlText('DATFIN') + ' ' + GetControlText('HRFIN'));

  // attente
  if GetCheckBoxState('CHKDEBUT') = cbUnchecked then
  begin
    if Now<DatDebut then Affiche('Attente heure de déclenchement '+ DateTimeToStr(DatDebut));
    While Now<DatDebut do
    begin
      Application.ProcessMessages; // pour appui sur stop
      if Stop then begin Affiche('Traitement interrompu par l''utilisateur'); exit; end;
    end;
  end;

  for iInd := 1 to GSoc.RowCount -1 do
  begin
    if not GSoc.IsSelected(iInd) then continue;
    libdoss := GSoc.Cells[0,iInd];
    //nodoss := Qry.FindField('DOS_NODOSSIER').AsString;
    if not Traitement(libdoss, cpteur, tominutes) then nberr := nberr+1;
    if nberr>nberrtot then
    begin
      Affiche('Le nombre d''échecs de traitement est ');
      Affiche('supérieur à la tolérance ('+IntToStr(nberrtot)+')');
      break;
    end;
    Application.ProcessMessages;
    if Stop Then break;
  end;
end;

function  TOF_MAJSOCPARLOT.Traitement(libdoss: String; var compteur: Integer ; toMinutes : integer): Boolean;
var VDoss: Variant;
    lgcde : String;
begin
  Result := False;
  Affiche(IntToStr(compteur)+'. Mise à jour du dossier '+libdoss+' à '+TimeToStr(Now));
  compteur := compteur + 1;
  // se positionne
  //V_PGI_Env.NoDossier := nodoss;
{  if (V_PGI_Env.NoDossier<>nodoss) or (V_PGI_Env.DBSocName='') then
    begin
    Affiche('*** Impossible de sélectionner le dossier '+nodoss);
    exit;
    end;}
  VDoss := VersionBase(libdoss);
  if (VDoss = -1) then Affiche('Dossier '+libdoss+' inaccessible.');  
  if VDoss = -1 then exit;
  if V_PGI.NumVersionSocRef  = VDoss then
  begin
    Affiche('--> Dossier '+libdoss+' à jour.');
    Result := True;
    exit;
  end;


  // ************ Traitement **********
  //lgcde := V_PGI_Env.RepLocal+'\APP\'+NomExec; // Pas du tout(PCS)  =>   <-  attention App est déjà dans le chemin de recherche
  // c'est PGIMAJVER qui se relance lui meme
  lgcde := ExtractFilePath(Application.ExeName) + NomExec;
  // lgcde := V_PGI_Env.RepLocal + NomExec;
  lgcde := lgcde + ' /USER='+V_PGI.UserLogin+' /PASSWORD='+V_PGI.Password+' /DATE='+DateToStr(V_PGI.DateEntree)+' /DOSSIER='+libdoss+' /MAJSTRUCTURE=TRUE /MAJSTRUCTAUTO=TRUE';
  // paramètres particuliers à l'appli
    //lgcde := lgcde + ' ' + Param;
  Affiche(lgcde);

  FileExecAndWaitUntil(lgcde, tominutes, Stop);


  if Stop then
  begin
    Affiche('Traitement interrompu par l''utilisateur');
    exit;
  end;
  // Vérifie version de la base obtenue
  VDoss := VersionBase(libdoss);
  if (VDoss = -1) then Affiche('Dossier '+libdoss+' inaccessible.');  
  if V_PGI.NumVersionSocRef = VDoss then Affiche('--> Dossier '+libdoss+' à jour.')
  else
  begin
    Affiche('La version du dossier '+libdoss+' est '+IntToStr(VDoss)+' au lieu de '+IntToStr(V_PGI.NumVersionSocRef));
    Affiche('Echec, ou mise à jour pas encore terminée pour ce dossier...');
    exit;
  end;
  Result := True;
end;

procedure TOF_MAJSOCPARLOT.Affiche(msg: String);
begin
  Recap.Items.Add('***  ' + msg + '  ***');
  Recap.ItemIndex := Recap.Items.Count-1;
  Recap.Refresh;
  Application.ProcessMessages;
end;

Initialization
  registerclasses ( [ TOF_MAJSOCPARLOT ] ) ;
end.
