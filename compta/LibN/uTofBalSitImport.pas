
{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 29/10/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : BALSIT_IMPORT ()
Mots clefs ... : TOF;BALSIT_IMPORT
*****************************************************************}
Unit uTofBalSitImport ;

Interface

Uses
  StdCtrls,
  Controls,
  Classes,
  db,
  forms,
  sysutils,
  ComCtrls,
  Graphics,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF ,
  ImpExpBDS,
  LookUp,
  uTOB,
  Ent1,
{$IFDEF EAGLCLIENT}
  MainEagl, // AGLLanceFiche
{$ELSE}
  FE_Main,  // AGLLanceFiche
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  HTB97;

Type
  TOF_BALSIT_IMPORT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
      FBalList : TOB;
      procedure OnChangeModeImport(Sender : TObject );
      procedure OnInfoImportBalance (Sender : TObject; Msg : string; bErr : boolean );
  end ;

procedure CPLanceFiche_BALSIT_IMPORT;

Implementation

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/07/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure CPLanceFiche_BALSIT_IMPORT;
begin
  AGLLanceFiche('CP', 'BALSIT_IMPORT', '', '', '');
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_BALSIT_IMPORT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BALSIT_IMPORT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BALSIT_IMPORT.OnUpdate ;
var FileName : string;
    Import : TImpExpBDS;
    i : integer;
    ModeImport : TModeImpExpBds;
begin
  Inherited ;
  { Initialisation de l'importation }
  if not VH^.TenueEuro then
  begin
    MessageAlerte('Le dossier doit être en Euro pour importer une balance de situation.');
    TToolBarButton97(GetControl('BFERME')).Click;
  end;
  if GetControlText('FMODEIMPORT')='HIS' then
  begin
    ModeImport := miHistobal;
    i := THValComboBox(GetControl('FHISTOBAL')).ItemIndex;
    if i=0 then  FileName := ''
    else FileName := FBalList.Detail[i-1].GetValue('HB_EXERCICE') + ';'+
                    DateToStr(FBalList.Detail[i-1].GetValue('HB_DATE1'))+';'+
                    DateToStr(FBalList.Detail[i-1].GetValue('HB_DATE2'));
  end else
  begin
    if GetControlText('FMODEIMPORT')='SI2' then ModeImport:=miSiscoII
    else ModeImport := miPGI;
    FileName  := GetControlText ('FFICHIERIMPORT');
    if not FileExists(FileName) then
    begin
      PGIBox ('Fichier à importer incorrect.',ECRAN.Caption);
      SetFocusControl ('FFICHIERIMPORT');
      exit;
    end;
  end;
  { Importation }
  Import := TImpExpBDS.Create ( ModeImport ) ;
  try
    Import.OnInformation := OnInfoImportBalance;
    Import.Importation ( FileName , GetControlText ('FCODEBAL'), GetControlText ('FLIBELLEBAL'),GetControlText ('FABREGEBAL'),GetCheckBoxState('FSOLDEDEBCRECOLL')=cbChecked);
    if Import.GetLastError = 0 then  // Tout est OK
    begin
      PGIInfo ('Importation terminée avec succès.',ECRAN.Caption);
    end else PGIBox ( 'Erreur lors de l''importation.'+#10#13+Import.LastErrorMsg, ECRAN.Caption);
  finally
    Import.Free;
  end;
end ;

procedure TOF_BALSIT_IMPORT.OnLoad ;
var Q : TQuery;
    St : string;
    i : integer;
    T : TOB;
    Histo : THValComboBox;
begin
  Inherited ;
  if not VH^.TenueEuro then
  begin
    MessageAlerte('Le dossier doit être en Euro pour importer une balance de situation.');
  end;
  { Chargement de la liste des balances de HISTOBAL }
  FBalList := TOB.Create('', nil, - 1);
  St := 'SELECT DISTINCT EX1.EX_LIBELLE C0, HB_EXERCICE, HB_DATE1, HB_DATE2, HB_TYPEBAL FROM HISTOBAL LEFT OUTER JOIN EXERCICE EX1 ON HB_EXERCICE=EX1.EX_EXERCICE AND EX1.EX_NATEXO="" WHERE HB_TYPEBAL="BDS"';
  Q:= OpenSQL (St, True);
  try
    FBalList.LoadDetailDB('','','',Q,False);
  finally
    Ferme (Q);
  end;
  { Mise à jour de la Combo de choix HISTOBAL }
  Histo := THValComboBox (GetControl('FHISTOBAL'));
  if FBalList.Detail.Count > 0 then Histo.Items.Add ('<<Tous>>');
  for i:=0 to FBalList.Detail.Count - 1 do
  begin
    T := FBalList.Detail[i];
    Histo.Items.Add ( 'du '+DateToStr(T.GetValue('HB_DATE1'))+' au '+DateToStr(T.GetValue('HB_DATE2')));
  end;
end ;

procedure TOF_BALSIT_IMPORT.OnArgument (S : String ) ;
begin
  Inherited ;
  THValComboBox (GetControl ('FMODEIMPORT')).OnChange := OnChangeModeImport;
  THValComboBox (GetControl ('FMODEIMPORT')).ItemIndex := 0; // Sisco II par défaut
end ;

procedure TOF_BALSIT_IMPORT.OnClose ;
begin
  FBalList.Free;
  Inherited ;
end ;

procedure TOF_BALSIT_IMPORT.OnInfoImportBalance(Sender: TObject;
  Msg: string; bErr : boolean);
var lCommentaire : THLabel;
begin
  lCommentaire := THLabel(GetControl ('FCOMMENTAIRE'));
  if bErr then lCommentaire.Font.Color := clRed
  else lCommentaire.Font.Color := clGreen;
  lCommentaire.Caption := Msg;
end;

procedure TOF_BALSIT_IMPORT.OnChangeModeImport(Sender: TObject);
var St : string;
begin
  if GetControlText('FMODEIMPORT')='HIS' then
  begin
    St := '&Balance';
    SetControlVisible('FFICHIERIMPORT',False);
    SetControlVisible('FHISTOBAL',True);
  end
  else
  begin
    St := '&Fichier';
    SetControlVisible('FFICHIERIMPORT',True);
    SetControlVisible('FHISTOBAL',False);
  end;
  SetControlProperty('TFFICHIERIMPORT','Caption',TraduireMemoire(St));
end;

Initialization
  registerclasses ( [ TOF_BALSIT_IMPORT ] ) ;
end.
