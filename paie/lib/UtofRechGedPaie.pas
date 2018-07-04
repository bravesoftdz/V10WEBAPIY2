{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 06/12/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : RECHDOCGED ()
Mots clefs ... : TOF;RECHDOCGED
*****************************************************************
PT1   26/12/2007 FC V_810 Concept accessibilité fiche salarié
}
unit UtofRechGedPaie;

interface

uses
{$IFNDEF EAGLCLIENT}
  Db, {$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
{$ENDIF}
  Controls,
  Classes,
{$IFNDEF EAGLCLIENT}
  mul,
  Fe_Main, //AGLLanceFiche
{$ELSE}
  eMul,
  MaineAGL, //AGLLanceFiche
{$ENDIF}
  forms,
  variants, //VarIsNull
  sysutils,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  UTOB,
  HTB97, //TToolbarButton97
  PNewDocument, //ShowNewDocument
  UGedFileViewer, //ShowGedFileViewer
  UtilGEDPaie //SupprimeDocumentGedPaie

  ;

type
  TOF_RechGedPaie = class(TOF)
  private
    Reference1: string;
    Objet: string;
    Action:String;//PT1
    procedure SupprimeDocument;

  public
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
    procedure FListeOnDblClick(Sender: TObject);
    procedure BtOnClick(Sender: TObject);
    procedure DispatchGedGRC(NumTag: Integer);
  end;

const
  // Tags
  cTagSupprimer = 5; // Supprimer le document
  cTagProprietes = 4; // Propriétés du document
  cTagDocExterne = 1; // Insertion de document externe
  cTagApercu = 2; // Visualiser le document

function LanceFiche_RechGedPaie(Nat, Cod: string; Range, Lequel, Argument: string): string;

implementation

procedure TOF_RechGedPaie.OnNew;
begin
  inherited;
end;

procedure TOF_RechGedPaie.OnDelete;
begin
  inherited;
end;

procedure TOF_RechGedPaie.OnUpdate;
begin
  inherited;
end;

procedure TOF_RechGedPaie.OnLoad;
begin
  inherited;

end;

procedure TOF_RechGedPaie.OnArgument(S: string);
var
  Critere: string;
  ChampMul: string;
  ValMul: string;
  x, i: integer;
  Ind, Indice, st: string;
  Q: TQuery;
begin
  inherited;
  Action := '';//PT1
  repeat
    Critere := uppercase(Trim(ReadTokenSt(S)));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));
        if ChampMul = 'SALARIE' then
        begin
          Reference1 := ValMul;
        end;
        if ChampMul = 'ACTION' then //PT1
          Action := ValMul;         //PT1
      end;
    end;
  until Critere = '';
  for i := 1 to 3 do
  begin
    Indice := '00' + IntTostr(i);
    Ind := IntTostr(i);
    st := RechDom('PGLIBGED', Indice, FALSE);
    if st <> '' then
    begin
      SetControlText('TRTD_TABLELIBREGED' + Ind, st);
    end
    else
    begin
      SetControlVisible('TRTD_TABLELIBREGED' + Ind, FALSE);
      SetControlVisible('RTD_TABLELIBREGED' + Ind, FALSE);
    end;
  end;
  Objet := 'SAL';
  if Reference1 <> '' then
  begin
    Q := OpenSql('SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM FROM SALARIES WHERE PSA_AUXILIAIRE = "' + reference1 + '"', TRUE);
    if not Q.EOF then
    begin
      SetControlEnabled('PSA_SALARIE', FALSE);
      SetcontrolText('PSA_SALARIE', Q.Findfield('PSA_SALARIE').AsString);
      SetControlEnabled('PSA_LIBELLE', FALSE);
      Ecran.Caption := Ecran.Caption + ' de '+Q.Findfield('PSA_LIBELLE').AsString + ' ' + Q.Findfield('PSA_PRENOM').AsString;
      SetcontrolVisible('PSA_LIBELLE', FALSE);
      SetcontrolVisible('TPSA_LIBELLE', FALSE);
    end;
    Ferme(Q);
    SetControlText('XX_WHERE', 'RTD_TIERS = "' + Reference1 + '"');
    SetControlText('RTD_TIERS', Reference1);
  end
  else SetControlText('XX_WHERE', 'RTD_TIERS <> ""');

  THGrid(GetControl('FLISTE')).OnDblClick := FListeOnDblClick;
  TToolbarButton97(GetControl('BPROPRIETE')).OnClick := BtOnClick;
  TToolbarButton97(GetControl('BDELETE')).OnClick := BtOnClick;
  TToolbarButton97(GetControl('BINSERT')).OnClick := BtOnClick;

  //DEB PT1
  {$IFDEF EMANAGER}
  SetControlVisible('BPROPRIETE', False);
  SetControlVisible('BDELETE',    False);
  SetControlVisible('BINSERT',    False);
  {$ENDIF}
  if Action = 'CONSULTATION' then
  begin
    SetControlEnabled('BInsert',false);
    SetControlEnabled('BDelete',false);
    SetControlEnabled('BPROPRIETE',false);
  end;
  //FIN PT1
end;

procedure TOF_RechGedPaie.OnClose;
begin
  inherited;
end;

procedure TOF_RechGedPaie.OnDisplay();
begin
  inherited;
end;

procedure TOF_RechGedPaie.OnCancel();
begin
  inherited;
end;

procedure TOF_RechGedPaie.FListeOnDblClick(Sender: TObject);
var
  sTempFileName: string;
  sTitre: string;

begin
  inherited;
  if not VarIsNull(GetField('RTD_DOCGUID')) then
  begin
    sTempFileName := ExtraitDocumentPaie(GetField('RTD_DOCGUID'), sTitre);
    if sTempFileName = '' then
      exit;
    ShowGedFileViewer(sTempFileName, False, sTitre, True, False, True, True);
  end;
end;

procedure TOF_RechGedPaie.BtOnClick(Sender: TObject);
var
  Num: Integer;

begin
  Num := 0;
  if Sender is TToolBarButton97 then
    Num := TToolBarButton97(Sender).Tag;
  DispatchGedGRC(Num);
end;

procedure TOF_RechGedPaie.DispatchGedGRC(NumTag: Integer);
var
  St: string;
  Infos, LeTiers: string;
  ParGed: TParamGedDoc;
  Q : TQuery;
begin
  case NumTag of

  // Supprimer
    cTagSupprimer:
      begin
        if not VarIsNull(GetField('RTD_DOCGUID')) then
        begin
          SupprimeDocument;
          TFMul(Ecran).BChercheClick(nil);
        end;
      end;

    // Propriétés
    cTagProprietes:
      begin
        if not VarIsNull(GetField('RTD_DOCGUID')) then
        begin
          ParGed.SDocGUID := GetField('RTD_DOCGUID');
          ParGed.SFileGUID := GetField('YDF_FileGUId');
          ParGed.NoDossier := '';
          ParGed.CodeGed := '';
          ParGed.TypeGed := '';
          ParGed.Objet := Objet;
          ParGed.ModifLien := True;
          if Objet = 'SAL' then
            Infos := 'SALARIE=' + GetField('RTD_TIERS') + ';';
          Parged.Infos := Infos;
          St := ShowNewDocument(ParGed);
          TFMul(Ecran).BChercheClick(nil);
        end;
      end;

    // Insertion de document externe
    cTagDocExterne:
      begin
        ParGed.SDocGUID := '';
        ParGed.SFileGUID := '';
        ParGed.NoDossier := '';
        ParGed.CodeGed := '';
        ParGed.Objet := Objet;
        if Objet = '' then
          ParGed.ModifLien := True
        else
          ParGed.ModifLien := False;
        if Objet = 'SAL' then
        begin
          st := 'SELECT PSA_AUXILIAIRE FROM SALARIES WHERE PSA_SALARIE="'+ GetControlText('PSA_SALARIE')+'"';
          Q := OPENSQL (st, TRUE);
          If NOT Q.EOF then LeTiers := Q.FindField ('PSA_AUXILIAIRE').AsString
          else LeTiers := '';
          Ferme (Q);
          Infos := 'SALARIE=' + LeTiers + ';';
        end;
        Parged.Infos := Infos;
        St := ShowNewDocument(ParGed);
        TFMul(Ecran).BChercheClick(nil);
      end;
  end;
end;

procedure TOF_RechGedPaie.SupprimeDocument;
var
  Msg: string;

begin
  Msg := 'Suppression du document ' + GetField('YDO_LIBELLEDOC') + '.' + #13#10
    + ' Confirmez-vous la suppression ?';

  if PGIAsk(Msg, TitreHalley) = mrNo then
    exit;

  // Suppression en cascade dans les tables, avec tests de dépendance
  SupprimeDocumentGedPaie(GetField('RTD_DOCGUID'));
end;

function LanceFiche_RechGedPaie(Nat, Cod: string; Range, Lequel, Argument: string): string;
begin
  AGLLanceFiche(Nat, Cod, Range, Lequel, Argument);
end;

initialization
  registerclasses([TOF_RechGedPaie]);
end.

