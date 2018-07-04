{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 11/10/2001
Modifié le ... :   /  /
Description .. : Réaffectation automatique des ventilations analytiques des
Suite ........ : bulletins de paei
Mots clefs ... : PAIE;ANALYTIQUE
*****************************************************************}
{ Fonction de reaffectation automaitique des ventilations analytiques de la paie
Elle crée ou recrée automatiquement les ventilations des bulletins de paie : VENTANA
et les historiques analytiques de la paie : HISTOANALPAIE
On choisit la liste des paies triées/sélectionnées par etablissement, codes organisation
sur une période de date
Attention, comme on travaille sur un historique alors les infos sont celles du salarié au moment
de la paie et non pas les infos actuelles du salarié.
En fait, cette fonction ne doit être utilisée que pour créer un historique des ventilations analytiques
mais pas pour recalculer des ventilations fausses sur l'ensemble des salaries.
Cela suppose que tous les salariés soit correctement paramètrés au niveau de leur ventilations analytiques.
}
{
 PT1 11/10/01 PH Raffraichissement de la liste avant lancement du traitement
 PT2 16/08/04 PH FQ 11504 Compléte le code matricule salarié avec des 0
 PT3 13/04/06 PH V_65 FQ 12962 fetch en CWAS pour prendre en compte la ligne complète
}
unit UTofPG_MulReaffAnal;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97, Hqry,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, Mul, Fe_Main,
{$ELSE}
  MaineAgl, eMul,
{$ENDIF}
  Grids, HCtrls, HEnt1, EntPaie, HMsgBox, UTOF, UTOB, UTOM, Vierge, P5Util, P5Def,
  AGLInit, PgOutils, PGoutils2;
type
  TOF_PGREAFFANAL = class(TOF)
  private
    Date1, Date2: THEdit;
    procedure ActiveWhere(Sender: TObject);
    procedure LanceFicheReaff(Sender: TObject);
    procedure DateDebutExit(Sender: TObject);
    procedure DateFinExit(Sender: TObject);
    procedure ExitEdit(Sender: TObject);
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
  end;

implementation

procedure TOF_PGREAFFANAL.ActiveWhere(Sender: TObject);
var
  DD, DF, WW, NomD, NomF: THEdit;
  St, Dat1, Dat2: string;
begin
  DD := THEdit(GetControl('DATEDEBUT'));
  DF := THEdit(GetControl('DATEFIN'));
  WW := THEdit(GetControl('XX_WHERE'));
  NomD := THEdit(GetControl('NOM1'));
  NomF := THEdit(GetControl('NOM2'));
  Dat1 := UsDateTime(StrToDate(DD.Text));
  Dat2 := UsDateTime(StrToDate(DF.Text));

  if (DD <> nil) and (WW <> nil) then
  begin
    WW.Text := 'PPU_DATEDEBUT >= "' + Dat1 + '" AND PPU_DATEFIN <= "' + Dat2 + '"';
    if (NomD <> nil) and (NomF <> nil) then
    begin
      if (NomD.Text <> '') then
      begin
        if (NomF.Text <> '') then St := ' AND PPU_SALARIE >="' + NomD.Text + '" AND PPU_SALARIE <="' + NomF.Text + '"'
        else St := ' AND PPU_SALARIE ="' + NomD.Text + '"';
        WW.Text := WW.Text + St;
      end;
    end;
  end;

end;

procedure TOF_PGREAFFANAL.DateDebutExit(Sender: TObject);
var
  DD, DF: THEdit;
  DateDeb, DateFin: TDateTime;
begin
  DD := THEdit(GetControl('DATEDEBUT'));
  DF := THEdit(GetControl('DATEFIN'));
  if (DD <> nil) and (DF <> nil) then
  begin
    DateDeb := StrToDate(DD.Text);
    DateFin := StrToDate(DF.Text);
    if DateFin < DateDeb then
    begin
      PGIBox('Attention, la date de début est supérieure à la date de fin', Ecran.Caption);
      SetFocusControl('DATEDEBUT');
    end;
  end;
end;

procedure TOF_PGREAFFANAL.DateFinExit(Sender: TObject);
var
  DD, DF: THEdit;
  DateDeb, DateFin: TDateTime;
begin
  DD := THEdit(GetControl('DATEDEBUT'));
  DF := THEdit(GetControl('DATEFIN'));
  if (DD <> nil) and (DF <> nil) then
  begin
    DateDeb := StrToDate(DD.Text);
    DateFin := StrToDate(DF.Text);
    if DateFin < DateDeb then
    begin
      PGIBox('Attention, la date de fin est inférieure à la date de début', Ecran.Caption);
      SetFocusControl('DATEDEBUT');
    end;
  end;
end;
// DEB PT2

procedure TOF_PGREAFFANAL.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then //AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;
// FIN PT2

procedure TOF_PGREAFFANAL.LanceFicheReaff(Sender: TObject);
var
  st: string;
  DD, DF: THEdit;
  DateDeb, DateFin: TDateTime;
  BtnCherche: TToolbarButton97;
begin
  // PT1 11/10/01 PH Raffraichissement de la liste avant lancement du traitement
  BtnCherche := TToolbarButton97(GetControl('BCherche'));
  if BtnCherche <> nil then BtnCherche.Click;

  DD := THEdit(GetControl('DATEDEBUT'));
  DF := THEdit(GetControl('DATEFIN'));
  if (DD <> nil) and (DF <> nil) then
  begin
    DateDeb := StrToDate(DD.Text);
    DateFin := StrToDate(DF.Text);
    if DateFin < DateDeb then
    begin
      PGIBox('Attention, la date de début est supérieure à la date de fin', Ecran.Caption);
      SetFocusControl('DATEDEBUT');
      exit;
    end;
    if VH_Paie.PGAnalytique then
    begin
      st := DD.Text + ';' + DF.Text;
{$IFDEF EAGLCLIENT}
// DEB PT3
      if TFMul(Ecran).Fetchlestous then
        TheMulQ := TOB(Ecran.FindComponent('Q'))
      else
      begin
        PgiBox('Vous n''avez pas de ligne total dans votre liste, #13#10 Traitement impossible ', Ecran.Caption);
        exit;
      end;
// FIN PT3      
{$ELSE}
      TheMulQ := THQuery(Ecran.FindComponent('Q'));
{$ENDIF}

      AglLanceFiche('PAY', 'REAFFANAL_AUTO', '', '', st);
      TheMulQ := nil;
    end
    else PGIBox('Vous ne gérez pas l''analytique !', 'Réaffectation impossible');
  end;
end;

procedure TOF_PGREAFFANAL.OnArgument(Arguments: string);
var
  Num: Integer;
  BtnValidMul: TToolbarButton97;
  Salar1, Salar2: THedit;
begin
  inherited;
  for Num := 1 to VH_Paie.PGNbreStatOrg do
  begin
    VisibiliteChampSalarie(IntToStr(Num), GetControl('PPU_TRAVAILN' + IntToStr(Num)), GetControl('TPPU_TRAVAILN' + IntToStr(Num)));
  end;
  BtnValidMul := TToolbarButton97(GetControl('BOuvrir'));
  if BtnValidMul <> nil then BtnValidMul.OnClick := LanceFicheReaff;
  Date1 := THEdit(GetControl('DATEDEBUT'));
  if Date1 <> nil then Date1.OnClick := DateDebutExit;
  Date2 := THEdit(GetControl('DATEFIN'));
  if Date2 <> nil then Date2.OnClick := DateFinExit;
  // DEB PT2
  Salar1 := ThEdit(getcontrol('NOM1'));
  if Salar1 <> nil then Salar1.OnExit := ExitEdit;
  Salar2 := ThEdit(getcontrol('NOM2'));
  if Salar2 <> nil then Salar2.OnExit := ExitEdit;
  // FIN PT2
end;


procedure TOF_PGREAFFANAL.OnLoad;
begin
  inherited;
  ActiveWhere(nil);
end;

initialization
  registerclasses([TOF_PGREAFFANAL]);
end.

