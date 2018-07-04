{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 06/08/2001
Modifié le ... :   /  /
Description .. : Multi critères de séléction des acomptes
Mots clefs ... : PAIE;PGACOMPTE
*****************************************************************}
{PT1 30/12/2002 SB V591 FQ 10401 Controle exercice actif incorrect
PT2  30/01/2004 SB V_50 FQ 11036 Suppression en masse des acomptes
PT3  09/11/2004 PH V_60 FQ 11736 test de la présence du champ rubrique lors de la
                        Suppression en masse des acomptes
PT4  05/05/2006 SB V_65 FQ 13111 Refonte multisuppr. en CWAS
}
unit UTofPG_MULACOMPTE;

interface
uses  Controls, Classes,  sysutils, HTB97,
{$IFNDEF EAGLCLIENT}
  db, HDB, mul,
{$ELSE}
  emul, MaineAgl,
{$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF,Ed_tools, HStatus;

type
  TOF_PGMULACOMPTE = class(TOF)
  private
    WW: THEdit;
    DateDebut, DateFin: THEdit;
    procedure ActiveWhere(Sender: TObject);
    procedure DateDebutExit(Sender: TObject);
    procedure DateFinExit(Sender: TObject);
    procedure ExitEdit(Sender: TObject);
    procedure DeleteRecord(Sender: TObject); //PT2
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
  end;

implementation

uses  EntPaie, PgOutils, PGoutils2, P5Def;

procedure TOF_PGMULACOMPTE.ActiveWhere(Sender: TObject);
var
  st: string;
  LaDate, LaDateFin: TDateTime;
begin
  WW.Text := '';
  if (DateDebut = nil) or (DateDebut = nil) then exit;
  LaDate := StrToDate(DateDebut.Text);
  LaDateFin := StrToDate(DateFin.Text);

  St := ' PSD_DATEDEBUT >="' + UsDateTime(LaDate) + '" AND PSD_DATEDEBUT <="' + UsDateTime(LaDateFin) + '"';
  WW.Text := St;
end;

procedure TOF_PGMULACOMPTE.OnArgument(Arguments: string);
var
  num: integer;
  Defaut: ThEdit;
  Btn: TToolBarButton97;
begin
  inherited;
  WW := THEdit(GetControl('XX_WHERE'));
  DateDebut := THEdit(GetControl('DATEDEBUT'));
  DateFin := THEdit(GetControl('DATEFIN'));
  if DateDebut <> nil then DateDebut.OnClick := DateDebutExit;
  if DateFin <> nil then DateFin.OnClick := DateFinExit;
  //RendExerSocialEnCours ( MoisE, AnneeE, ComboExer, DebExer, FinExer);

  for Num := 1 to 4 do
  begin
    VisibiliteChampSalarie(IntToStr(Num), GetControl('PSA_TRAVAILN' + IntToStr(Num)), GetControl('TPSA_TRAVAILN' + IntToStr(Num)));
  end;
  VisibiliteStat(GetControl('PSA_CODESTAT'), GetControl('TPSA_CODESTAT'));
  Defaut := ThEdit(getcontrol('PSD_SALARIE'));
  if Defaut <> nil then Defaut.OnExit := ExitEdit;

  Btn := TToolBarButton97(GetControl('BSUPPRIMER')); //PT2
  if Btn <> nil then Btn.OnClick := DeleteRecord;

end;

procedure TOF_PGMULACOMPTE.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then //AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;



procedure TOF_PGMULACOMPTE.OnLoad;
var
  Mois, Annee, Jour: WORD;
  LeMois, LAnnee: string;
  LaDate: TDateTime;
begin
  inherited;
  if (DateDebut <> nil) then
  begin
    LaDate := StrToDate(DateDebut.Text);
    DecodeDate(LaDate, Annee, Mois, Jour);
    LeMois := ColleZeroDevant(Mois, 2);
    LAnnee := ColleZeroDevant(Annee, 4);
    SetControlEnabled('BOuvrir', FALSE);
    SetControlEnabled('FListe', FALSE);
    SetControlEnabled('BInsert', FALSE);

    {PT1 Appel d'une nouvelle fonction pour test exercice actif
    if  NOT ControlMoisAnneeExer(LeMois, LAnnee, AnneeOk)}
    if not isPeriodeActif(LaDate, LaDate) then
      SetFocusControl('DATEDEBUT')
    else
    begin
      SetControlEnabled('BOuvrir', TRUE);
      SetControlEnabled('FListe', TRUE);
      SetControlEnabled('BInsert', TRUE);
    end;
  end;
  ActiveWhere(nil);
end;


procedure TOF_PGMULACOMPTE.DateDebutExit(Sender: TObject);
var
  DD, DF: THEdit;
  DateDeb, DateFin: TDateTime;
begin
  DD := THEdit(GetControl('DATEDEBUT'));
  DF := THEdit(GetControl('DATEFIN'));
  if (DD <> nil) and (DF <> nil) then
  begin
    if not (isvaliddate(DD.text)) then
    begin
      PGIBox('Date incorrecte', Ecran.Caption);
      SetFocusControl('DATEDEBUT');
      exit;
    end;
    DateDeb := StrToDate(DD.Text);
    DateFin := StrToDate(DF.Text);
    if DateFin < DateDeb then
    begin
      PGIBox('Attention, la date de début est supérieure à la date de fin', Ecran.Caption);
      SetFocusControl('DATEDEBUT');
    end;
    { if StrToDate (DD.Text) < DebExer then
      begin
      PGIBox ('Attention, la date de début est inférieure à la date de début d''exercice',Ecran.Titre);
      SetFocusControl ('DATEFIN') ;
      end;}
  end;
end;

procedure TOF_PGMULACOMPTE.DateFinExit(Sender: TObject);
var
  DD, DF: THEdit;
  DateDeb, DateFin: TDateTime;
begin
  DD := THEdit(GetControl('DATEDEBUT'));
  DF := THEdit(GetControl('DATEFIN'));
  if (DD <> nil) and (DF <> nil) then
  begin
    if not (isvaliddate(DF.text)) then
    begin
      PGIBox('Date incorrecte', Ecran.Caption);
      SetFocusControl('DATEFIN');
      exit;
    end;
    DateDeb := StrToDate(DD.Text);
    DateFin := StrToDate(DF.Text);
    if DateFin < DateDeb then
    begin
      PGIBox('Attention, la date de fin est inférieure à la date de début', Ecran.Caption);
      SetFocusControl('DATEDEBUT');
    end;
    { if StrToDate (DF.Text) > FinExer then
      begin
      PGIBox ('Attention, la date de fin est supérieure à la date de fin d''exerice','Préparation automatique');
      SetFocusControl ('DATEFIN') ;
      end;}
  end;
end;
{ DEB PT2 }

procedure TOF_PGMULACOMPTE.DeleteRecord(Sender: TObject);
var
  NbEnr, Ordre, i: Integer;
  Salarie, Rub, st: string;
  DateDeb: TDateTime;
  AllSelect : Boolean;
begin
  // DEB PT3
  St := TFMul(Ecran).Q.Champs;
  if POS('PSD_RUBRIQUE', St) <= 0 then
  begin
    PgiInfo('Vous devez rajouter le champ rubrique dans votre liste', Ecran.Caption);
    Exit
  end;
  // FIN PT3
  if (TFMul(Ecran).FListe = nil) then exit;
  if (TFMul(Ecran).FListe.nbSelected < 1) and (TFMul(Ecran).FListe.AllSelected = False) then
  begin
    PgiInfo('Aucun(s) élément(s) selectionné(s).Traitement annulé.', Ecran.caption);
    Exit;
  end;
  //  TFmul(Ecran).Q.Fieldb
  if PgiAsk('Attention vous risquez de supprimer des acomptes réglés. Voulez-vous continuer?', Ecran.caption) = mrNo then
    exit;
  AllSelect := TFMul(Ecran).FListe.AllSelected; { PT4 }
  if ((TFMul(Ecran).FListe.nbSelected) > 0) or (AllSelect) then  { PT4 }
  begin
    NbEnr := TFMul(Ecran).FListe.nbSelected;
    if AllSelect then NbEnr := TFmul(Ecran).Q.RecordCount;
    InitMoveProgressForm(nil, 'Chargement des données en cours...', 'Veuillez patienter SVP...', NbEnr, FALSE, TRUE);
    InitMove(NbEnr, '');
    if AllSelect then TFmul(Ecran).Q.First; { PT4 }
    for i := 0 to NbEnr - 1 do
    begin
    {$IFDEF EAGLCLIENT}
    if not AllSelect then TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);  { PT4 }
    {$ENDIF}
      if not AllSelect then TFMul(Ecran).FListe.GotoLeBOOKMARK(i); { PT4 }
      Salarie := TFmul(Ecran).Q.FindField('PSD_SALARIE').AsString;
      DateDeb := TFmul(Ecran).Q.FindField('PSD_DATEDEBUT').AsDateTime;
      Rub := TFmul(Ecran).Q.FindField('PSD_RUBRIQUE').AsString;
      Ordre := TFmul(Ecran).Q.FindField('PSD_ORDRE').AsInteger;
      MoveCurProgressForm('Suppression acompte salarié :' + Salarie);
      ExecuteSql('DELETE FROM HISTOSAISRUB WHERE PSD_ORIGINEMVT="ACP" ' +
        'AND PSD_SALARIE="' + Salarie + '" ' +
        'AND PSD_DATEDEBUT="' + UsDateTime(DateDeb) + '" ' +
        'AND PSD_RUBRIQUE="' + Rub + '" ' +
        'AND PSD_ORDRE =' + IntToStr(Ordre));
      MoveCur(False);
      if AllSelect then TFmul(Ecran).Q.Next;  { PT4 }
    end;
    TFMul(Ecran).FListe.ClearSelected;
    FiniMove;
    FiniMoveProgressForm;
    if AllSelect then  { PT4 }
    begin
      TFMul(Ecran).FListe.AllSelected := False;
      TFMul(Ecran).bSelectAll.Down := False;
    end;
    TFMul(Ecran).BCherche.Click;
  end;
end;
{ FIN PT2 }

initialization
  registerclasses([TOF_PGMULACOMPTE]);
end.

