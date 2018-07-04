{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 11/09/2001
Modifié le ... :   /  /
Description .. : Unit de gestion du multi critère qui filtre les différents types
Suite ........ : d'enregistrement possibles (et les dates ) en provenance
Suite ........ : d'un fichier d'import
Mots clefs ... : PAIE;PGDEPORTEE
*****************************************************************}
{
PT1   : 30/04/2004 V_50 PH FQ 11276 Suppression par lot des lignes d'import
PT2   : 09/08/2005 V_60 PH FQ 12079 gestion des bornes de périodes et du matricule salarié
PT3   : 19/09/2005 V_60 PH FQ 12079 Période en cours au lieu du mois

}
unit UTof_mulhistoimport;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls,
  mul,
{$ELSE}
  emul,
{$ENDIF}
  HQry, HStatus, Paramdat,
  Grids, HCtrls, HEnt1, HMsgBox, UTOF, UTOB, UTOM, Vierge, AGLInit;

type
  TOF_MULHISTOIMPORT = class(TOF)
  private
    WW: tHedit;
    procedure SuppressionMvt(Sender: TObject);
    procedure SuppLigne(Q: THQuery);
    procedure ActiveWhere(Sender: TObject);
//    procedure DateElipsisclick(Sender: TObject);
    procedure ExitEdit(Sender: TObject); // PT2
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
  end;

implementation

uses Entpaie, PgOutils2;

procedure TOF_MULHISTOIMPORT.OnArgument(Arguments: string);
var
  Salarie: THEdit;
  BtnDel: TToolbarButton97;
  MoisE, AnneeE, ComboExer: string; // PT3
  Date1,Date2 : THEdit; // PT3
  DebExer, FinExer, ZDate: TDateTime; // PT3
begin
  inherited;

  WW := THEdit(GetControl('XX_WHERE'));
  // DEB PT2
//  Date := THEdit(GetControl('PDD'));
//  if Date <> nil then Date.OnElipsisClick := DateElipsisclick;
//  Date := THEdit(GetControl('PDD'));
//  Date := THEdit(GetControl('PDF'));
//  if Date <> nil then Date.OnElipsisClick := DateElipsisclick;
  Salarie := THEdit(GetControl('PSD_SALARIE'));
  if Salarie <> nil then Salarie.OnExit := ExitEdit;
  // FIN PT2
  // DEB PT3
  Date1 := THEdit(GetControl('PSD_DATEDEBUT'));
  Date2 := THEdit(GetControl('PSD_DATEFIN'));
  if (Date1 <> nil) and (Date2 <> nil) then
  begin
    RendExerSocialEnCours(MoisE, AnneeE, ComboExer, DebExer, FinExer);
    if VH_Paie.PgDatePrepAuto = 'MOI' then
    begin
      if StrToDate(Date2.Text) > FinExer then
      begin
        Date2.Text := DateToStr(FinExer);
        Date1.Text := DateToSTr(DEBUTDEMOIS(FinExer));
      end;
    end
    else
    begin //Recuperation de la periode en cours de la saisie par rubrique
      if not VH_Paie.PGDecalage then ZDate := EncodeDate(StrToInt(AnneeE), StrToInt(MoisE), 01)
      else
      begin
        if MoisE = '12' then ZDate := EncodeDate(StrToInt(AnneeE) - 1, StrToInt(MoisE), 01)
        else ZDate := EncodeDate(StrToInt(AnneeE), StrToInt(MoisE), 01);
      end;
      Date1.Text := DateToSTr(Zdate);
      Date2.Text := DateToSTr(FINDEMOIS(Zdate));
    end;
  end;
  // FIN PT3
// PT1   : 30/04/2004 V_50 PH FQ 11276 Suppression par lot des lignes d'import
  if Arguments = 'SUPP' then
  begin
    Ecran.Caption := 'Suppression de lignes d''import';
    UpdateCaption(Ecran);
    SetControlProperty('FLISTE', 'MultiSelection', TRUE);
    SetControlVisible('DELETEMVT', TRUE);
    SetControlVisible('BSelectAll', TRUE);
    BtnDel := TToolbarButton97(GetControl('DELETEMVT'));
    if BtnDel <> nil then BtnDel.OnClick := SuppressionMvt;
  end;
end;

// PT1   : 30/04/2004 V_50 PH FQ 11276 Suppression par lot des lignes d'import

procedure TOF_MULHISTOIMPORT.SuppressionMvt(Sender: TObject);
var
  OkOk, i: Integer;
begin
  OkOk := PGIAsk('Attention, voulez vraiment supprimer les lignes sélectionnées ?', Ecran.caption);
  if OkOk <> MrYes then exit;
  with TFMul(Ecran) do
  begin
    if FListe.AllSelected then
    begin
      InitMove(Q.RecordCount, '');
      Q.First;
      while not Q.EOF do
      begin
        MoveCur(False);
        SuppLigne(Q);
        Q.NEXT;
      end;
      FListe.AllSelected := False;
    end
    else
    begin
      InitMove(FListe.NbSelected, '');
      for i := 0 to FListe.NbSelected - 1 do
      begin
        FListe.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
        Q.TQ.Seek(FListe.Row - 1);
{$ENDIF}
        MoveCur(False);
        SuppLigne(Q);
      end;
      FListe.ClearSelected;
    end;
    FiniMove;
  end;
  TFMul(Ecran).BChercheClick(nil);
end;

procedure TOF_MULHISTOIMPORT.ActiveWhere(Sender: TObject);
begin
  if WW = nil then exit;
  WW.Text := ' (PSD_ORIGINEMVT = "MHE" OR PSD_ORIGINEMVT = "MLB" OR PSD_ORIGINEMVT = "MFP")';
  // DEB PT2 inutile car les champs de la fiche sont les champs de la table
{  DDebut := ThEdit(getcontrol('PDD'));
  if DDebut = nil then exit;
  DFin := Thedit(getcontrol('PDF'));
  if DFin = nil then exit;
  if isValiddate(DDebut.Text) then
    PDD := StrToDate(DDebut.Text)
  else PDD := 0;
  if isValiddate(DFin.Text) then
    PDF := StrToDate(DFin.Text)
  else PDF := 0;
  if (PDD <> 0) or (PDF <> 0) then
  begin
    WW.Text := WW.text + 'AND PSD_DATEDEBUT >= "' +
      usdatetime(PDD) + '" AND PSD_DATEFIN <= "' + usdatetime(PDF) + '" ';
  end;}
  // FIN PT2
end;
{
procedure TOF_MULHISTOIMPORT.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;
}

procedure TOF_MULHISTOIMPORT.OnLoad;
begin
  inherited;
  ActiveWhere(nil);
end;

// PT1   : 30/04/2004 V_50 PH FQ 11276 Suppression par lot des lignes d'import

procedure TOF_MULHISTOIMPORT.SuppLigne(Q: THQuery);
var
  st, sal, rub, Orig: string;
  D1, D2: TDateTime;
  Ordre: Integer;
begin
  Sal := Q.findfield('PSD_SALARIE').Asstring;
  D1 := Q.findfield('PSD_DATEDEBUT').AsDateTime;
  D2 := Q.findfield('PSD_DATEFIN').AsDateTime;
  rub := Q.findfield('PSD_RUBRIQUE').Asstring;
  Orig := Q.findfield('PSD_ORIGINEMVT').Asstring;
  Ordre := Q.findfield('PSD_ORDRE').AsInteger;
  st := 'DELETE FROM HISTOSAISRUB WHERE PSD_SALARIE="' + Sal + '" AND PSD_RUBRIQUE="' + Rub +
    '" AND PSD_ORIGINEMVT="' + Orig + '" AND PSD_DATEDEBUT="' + UsDateTime(D1) +
    '" AND PSD_DATEFIN="' + UsDateTime(D2) + '" AND PSD_ORDRE=' + IntToStr(Ordre);
  ExecuteSQL(st);
end;
// DEB PT2

procedure TOF_MULHISTOIMPORT.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;
// FIN PT2
initialization
  registerclasses([TOF_MULHISTOIMPORT]);
end.

