{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFREVPREPAIMPORT ()
Mots clefs ... : TOF;AFREVPREPAIMPORT
*****************************************************************}
unit utofAfRevPrepaImport;

interface

uses StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fe_Main,
  {$ELSE}
  MainEagl,
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF, utob, paramsoc, FileCtrl, HTB97, buttons,
  utofafRevImportIndice;

type
  TOF_AFREVPREPAIMPORT = class(TOF)
    BTNValider: TToolbarButton97;
    BTNchoix: Tbitbtn;
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
    procedure BTNValiderClick(sender: Tobject);
    function CherchePlusRecent: string;

    private
      fStFichierImp : String;

  end;

procedure AFLanceFiche_AFREVPREPAIMPORT;

implementation

function TOF_AFREVPREPAIMPORT.CherchePlusRecent: string;
var
  i: integer;
  sr: TSearchRec;
  chemin_Du_Fichier: string;
  listeFic: TstringList;
  ok1: boolean;
begin

  chemin_Du_Fichier := GetparamSoc('SO_AFINDIMPEXP');
  if chemin_Du_Fichier = '' then
    chemin_Du_Fichier := ExtractFileDrive(Application.ExeName) + '\';

  if chemin_Du_Fichier[length(chemin_Du_Fichier)] <> '\' then
    chemin_Du_Fichier := chemin_Du_Fichier + '\';
    
  listeFic := TstringList.create;
  listeFic.Sort;
  ListeFic.Count;
  ok1 := false;
  if FindFirst(chemin_Du_Fichier + 'IND*.*', faAnyFile, sr) = 0 then
  begin
    ok1 := true;
    listeFic.Add(sr.Name);
    while FindNext(sr) = 0 do listeFic.Add(sr.Name);
  end;
  listeFic.Sort;
  if ok1 then result := chemin_Du_Fichier + listeFic[listeFic.count - 1] else result := '';
  for i := listeFic.Count - 1 downto 0 do
    listeFic.Delete(i);
  listeFic.clear;
  FindClose(sr);
end;

procedure TOF_AFREVPREPAIMPORT.BTNValiderClick(sender: Tobject);
begin
  AFLanceFiche_AFREVIMPORTINDICE(fStFichierImp);
end;

procedure TOF_AFREVPREPAIMPORT.OnNew;
begin
  inherited;
end;

procedure TOF_AFREVPREPAIMPORT.OnDelete;
begin
  inherited;
end;

procedure TOF_AFREVPREPAIMPORT.OnUpdate;
begin
  inherited;
end;

procedure TOF_AFREVPREPAIMPORT.OnLoad;
begin
  inherited;
end;

procedure TOF_AFREVPREPAIMPORT.OnArgument(S: string);
begin
  inherited;
  BTNValider := TToolbarButton97(Getcontrol('BVALIDER'));
  BTNValider.OnClick := BTNValiderClick;
  fStFichierImp := CherchePlusRecent;
  setControltext('FICHIERIMP', fStFichierImp);
  SetFocusControl('FICHIERIMP');
end;

procedure TOF_AFREVPREPAIMPORT.OnClose;
begin
  inherited;
end;

procedure TOF_AFREVPREPAIMPORT.OnDisplay();
begin
  inherited;
end;

procedure TOF_AFREVPREPAIMPORT.OnCancel();
begin
  inherited;
end;
                                        
procedure AFLanceFiche_AFREVPREPAIMPORT;
begin
  AglLanceFiche('AFF', 'AFREVPREPAIMPORT', '', '', '');
end;

initialization
  registerclasses([TOF_AFREVPREPAIMPORT]);
end.
