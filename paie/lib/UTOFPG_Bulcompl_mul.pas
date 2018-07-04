{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 11/09/2001
Modifié le ... : 23/10/2001
Description .. : Unit de gestion du multi critère saisie bulletin
Suite ........ : complémentaire
Mots clefs ... : PAIE
*****************************************************************}
{
PT1 04/04/02 V571 PH Controle etab par rapport au salarie et non la liste
PT2 11/12/02 V591 PH Recup date debut exercice social sur le 1er mois paie décalée
PT3 15/01/03 V591 PH Test présence du champ PSA_DATESORTIE obligatoire dans la liste FQ 10452
PT4 13/05/03 V_42 PH Optimisation chargement des tob de la paie non fait systématiquement
PT5 06/05/03 V_421 PH FQ 10510 plus de controle si déjà un bulletin complémentaire au 01/01
PT6 20/08/04 V_50  PH FQ 11472 Suppression des contrôles des CP dans le cas du bulletin complémentaire
PT7 05/11/2004 PH V_60 FQ 11766 Bulletin complémentaire sur une date de début <> debut du mois
PT8 15/09/2005 PH V_60 FQ 12494 Bulletin complémentaire pour un salarié entré en cours de mois
PT9 20/09/2005 PH V_60 FQ 12527 Sélection profil exclusif du bulletin que si bulletin interessement
PT10 21/04/2006 GGS V_650 FQ 12750 vide profil exclusif si decoche bulletin interessement
}
unit UTOFPG_BulCompl_mul;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
{$IFNDEF EAGLCLIENT}
  HDB, DBGrids, db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main, mul,
{$ELSE}
  MaineAGL, eMul,
{$ENDIF}
  Grids, HCtrls, HEnt1, vierge, EntPaie, HMsgBox, Hqry, UTOF, UTOB, UTOM, P5Util, P5Def,
  AGLInit, PgOutils, PGoutils2, SaisBul;

type
  TOF_PG_BulCompl_mul = class(TOF)
    procedure OnArgument(stArgument: string); override;
    procedure OnClose; override;
    procedure OnLoad; override;
  private
    LaCheck: TCheckBox;
    Q_Mul: THQuery;
    DateDebut, DateFin: TDateTime;
    procedure ExitEdit(Sender: TObject);
    procedure ExitCheck(Sender: TObject);
    procedure GrilleDblClick(Sender: TObject);
    procedure ActiveWhere;
  end;

implementation

procedure TOF_PG_BulCompl_mul.OnArgument(stArgument: string);
var
  Defaut: THEdit;
  Num: Integer;
  MoisE, AnneeE, ComboExer: string;
  DebExer, FinExer: TDateTime;

{$IFDEF EAGLCLIENT}
  Grille: THGrid;
{$ELSE}
  Grille: THDBGrid;
{$ENDIF}
begin
  inherited;
  Q_Mul := THQuery(Ecran.FindComponent('Q'));
{$IFDEF EAGLCLIENT}
  Grille := THGrid(GetControl('Fliste'));
{$ELSE}
  Grille := THDBGrid(GetControl('Fliste'));
{$ENDIF}
  if Grille <> nil then Grille.OnDblClick := GrilleDblClick;
  // pour un bulletin complémentaire, on fixe par défaut la période de paie à  1 jour
  if RendExerSocialEnCours(MoisE, AnneeE, ComboExer, DebExer, FinExer) = TRUE then
  begin
    DateDebut := EncodeDate(StrToInt(AnneeE), StrToInt(MoisE), 1);
    // PT2 11/12/02 V591 PH Recup date debut exercice social sur le 1er mois paie décalée
    if (VH_Paie.PGDecalage) and (MoisE = '12') then
      DateDebut := EncodeDate(StrToInt(AnneeE) - 1, StrToInt(MoisE), 1);
  end
  else
  begin
    PGIBox('Attention, l''identification de la période de paie impossible', Ecran.caption);
    DateDebut := DebutDeMois(Date);
  end;
  DateFin := DateDebut;

  for Num := 1 to VH_Paie.PGNbreStatOrg do
  begin
    if Num > 4 then Break;
    VisibiliteChampSalarie(IntToStr(Num), GetControl('PSA_TRAVAILN' + IntToStr(Num)), GetControl('TPSA_TRAVAILN' + IntToStr(Num)));
  end;
  VisibiliteStat(GetControl('PSA_CODESTAT'), GetControl('TPSA_CODESTAT'));
  Defaut := ThEdit(getcontrol('PSA_SALARIE'));
  if Defaut <> nil then Defaut.OnExit := ExitEdit;
  LaCheck := TCheckBox(GetControl('CHBXBULCOMPL'));
  // DEB PT9
  if LaCheck <> nil then
  begin
    LaCheck.OnClick := ExitCheck;
    if LaCheck.Checked then SetControlEnabled('LEPROFILPART', TRUE)
    else SetControlEnabled('LEPROFILPART', FALSE);
  end;
  // FIN PT9
  InitLesTOBPaie;
  ChargeLesTOBPaie;
end;

procedure TOF_PG_BulCompl_mul.OnClose;
begin
  VideLesTOBPaie(FALSE);
  //if TOB_Salarie <> NIL Then TOB_Salarie.Free;
  // TOB_Salarie := NIL;
  if TOB_ExerSocial <> nil then TOB_ExerSocial.Free;
  TOB_ExerSocial := nil;
end;

procedure TOF_PG_BulCompl_mul.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then //AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

procedure TOF_PG_BulCompl_mul.ExitCheck(Sender: TObject);
begin
  if LaCheck <> nil then
  begin
    if LaCheck.Checked then SetControlEnabled('LEPROFILPART', TRUE)
    else
    begin
      SetControlText('LEPROFILPART','');                //PT10
      SetControlEnabled('LEPROFILPART', FALSE);
    end;
  end;
end;

procedure TOF_PG_BulCompl_mul.GrilleDblClick(Sender: TObject);
var
  CodeSalarie, stq, Etabl: string;
  ActionBul: TActionBulletin;
  Annee, Mois, Jour: WORD;
  Q: tquery;
  rep: integer;
  GestionCPEtab, AutorisationBulletin, AutorisationSaisie, LePlus: boolean;
begin
{$IFDEF EAGLCLIENT}
  TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
{$ENDIF}
  AutorisationBulletin := True;
  ActionBul := taCreation;
  DecodeDate(DateDebut, Annee, Mois, Jour);

  AutorisationSaisie := TRUE;
  if LaCheck.Checked then
  begin
    if GetControlText('LEPROFILPART') = '' then
    begin
      PGiBox('Vous devez saisir un profil', Ecran.Caption);
      exit;
    end;
  end;
  CodeSalarie := Q_Mul.FindField('PSA_SALARIE').AsString;
  try
    //PT1 04/04/02 V571 PH Controle etab par rapport au salarie et non la liste
    Q := Opensql('SELECT PSA_ETABLISSEMENT FROM SALARIES WHERE PSA_SALARIE="' + CodeSalarie + '"', true);
    if not Q.eof then Etabl := Q.FindField('PSA_ETABLISSEMENT').Asstring
    else Etabl := '';
    Ferme(Q);
    // FIN PT1
    if Etabl = '' then
    begin
      PGiBox('Le salarié doit appartenir à un établissement', Ecran.Caption);
      AutorisationSaisie := FALSE;
    end;
  except
    PGiBox('Le salarié doit appartenir à un établissement', Ecran.Caption);
    AutorisationSaisie := FALSE;
  end;
  Stq := 'SELECT PPU_SALARIE FROM PAIEENCOURS WHERE PPU_ETABLISSEMENT="' + Etabl + '" AND ' +
    'PPU_SALARIE="' + CodeSalarie + '" AND AND PPU_DATEDEBUT>="' + UsDateTime(DateDebut) +
    '" AND PPU_DATEFIN<="' + UsDateTime(FINDEMOIS(DateFin)) + '"';
  Q := Opensql(stq, true);
  if not Q.eof then
  begin
    Stq := 'Attention, il existe déjà un bulletin sur la période du ' + DateToStr(DateDebut) + ' au ' + DateToStr(FINDEMOIS(DateDebut)) +
      '#13#10 Vous allez créer un bulletin complémentaire au ' + DateToStr(DateDebut) +
      '#13#10 Voulez vous continuer ? ';
    Rep := PGIAsk(Stq, Ecran.Caption);
    if rep <> mrYes then
    begin
      Ferme(Q);
      exit;
    end;
  end;
  Ferme(Q);
  // DEB PT7 Controle de la présence d'un bulletin à la même date que le bulletin complémentaire
  LePlus := FALSE;
  Stq := 'SELECT PPU_SALARIE FROM PAIEENCOURS WHERE PPU_ETABLISSEMENT="' + Etabl + '" AND ' +
    'PPU_SALARIE="' + CodeSalarie + '" AND AND PPU_DATEDEBUT="' + UsDateTime(DateDebut) +
    '" AND PPU_DATEFIN="' + UsDateTime(DateDebut) + '"';
  Q := Opensql(stq, true);
  if not Q.eof then
  begin
    Stq := 'Attention, il existe déjà un bulletin sur la période du ' + DateToStr(DateDebut) + ' au ' + DateToStr(DateDebut) +
      '#13#10 Vous êtes obligé de créer un bulletin complémentaire au ' + DateToStr(DateDebut + 1) +
      '#13#10 Voulez vous continuer ? ';
    Rep := PGIAsk(Stq, Ecran.Caption);
    if rep <> mrYes then
    begin
      Ferme(Q);
      exit;
    end;
    LePlus := TRUE;
  end;
  Ferme(Q);
  // FIN PT7
  GestionCPEtab := false;
  { // DEB PT6
  if  VH_Paie.PGCongesPayes then
  begin
  stq :='SELECT ETB_DATECLOTURECPN,ETB_CONGESPAYES FROM ETABCOMPL '+
        'WHERE ETB_ETABLISSEMENT="'+Etabl+'"';
  Q := Opensql (stq,true);
  if not Q.eof then
     begin
     try
     CP:= Q.findfield('ETB_CONGESPAYES').AsVariant ;
     except
     AutorisationSaisie := FALSE;
     PGiBox ('Vous devez indiquer si vous gérez les congés payés ?', Ecran.Caption);
     end;
     GestionCPEtab := (CP='X');
     if (GestionCPEtab) AND (AutorisationSaisie) then
       begin
       try
       DTClot := Q.findfield('ETB_DATECLOTURECPN').AsDateTime ;
       if (DTClot = 0) then
          begin
          AutorisationSaisie := FALSE;
          PGiBox ('Vous devez saisir la date de clôture des congés payés ?', Ecran.Caption);
          end;
       except
       AutorisationSaisie := FALSE;
       PGiBox ('Vous devez saisir la date de clôture des congés payés ?', Ecran.Caption);
       end;
       decodedate(DTCLOT,aa,mm,jj);
       if ((aa<Annee) or ((aa=Annee)and(mm<Mois))) then
         AutorisationBulletin := false;
       end;
     end;
  ferme(Q);
  end; FIN PT6
  }
  if (AutorisationBulletin) then
  begin
    if AutorisationSaisie then
    begin
      if ActionBul = taCreation then
      begin
        // PT1 : 10/09/2001 V547 PH Controle de la date de la paie en création de bulletin
        if not ControlPaieCloture(DateDebut, DateFin) then
        begin
          PgiBox('Vous ne pouvez pas saisir de paie sur une période close.', Ecran.Caption);
          exit;
        end;
        // FIN PT1
      end;
      // PT3 22/10/01 V562 PH  Rajout un paramètre à l'appel de la saisie du bulletin
      if ActionBul = taCreation then
      begin
        // PT3 15/01/03 V591 PH Test présence du champ PSA_DATESORTIE obligatoire dans la liste FQ 10452
{$IFNDEF EAGLCLIENT}
        stq := Q_Mul.Champs; 
{$ELSE}
        stq := Q_MUl.Champs;   
{$ENDIF}
        if (Pos('PSA_DATESORTIE', Stq)) <= 0 then
        begin
          PgiBox('Vous devez rajouter la date de sortie dans la liste #13#10 pour ' +
            'faire un bulletin complémentaire.', 'Bulletin complémentaire');
        end
        else
          // FIN PT3
        begin
          // DEB PT7 +1 sur la date le cas échéant
          if LePLus then SaisieBulletin(CodeSalarie, Etabl, 'X', GetControlText('LEPROFILPART'), DateDebut + 1, DateFin + 1, ActionBul, Q_Mul, FALSE, GestionCPEtab)
          else SaisieBulletin(CodeSalarie, Etabl, 'X', GetControlText('LEPROFILPART'), DateDebut, DateFin, ActionBul, Q_Mul, FALSE, GestionCPEtab);
          // FIN PT7
        end;
      end;
    end;
  end
  else
    HShowMessage('1;;La clôture des congés payés de l''établissement n''a pas été effectuée#13#10 Bulletin impossible;N;C;C;', '', '');
end;

procedure TOF_PG_BulCompl_mul.ActiveWhere;
var
  WW: THEdit;
begin
  WW := THEdit(GetControl('XX_WHERE'));
  // PT5 06/05/03 V_421 PH FQ 10510 plus de controle si déjà un bulletin complémentaire au 01/01
  { car on peut faire autant de bulletins complémentaires que l'on veut
   AND (NOT EXISTS (SELECT PPU_SALARIE FROM PAIEENCOURS WHERE PPU_SALARIE = PSA_SALARIE AND PPU_DATEDEBUT="'+UsDateTime (DateDebut)+'"'+
             ' AND PPU_DATEFIN="'+UsDateTime (DateDebut)+'"))';
  }
  WW.Text := '(PSA_DATEENTREE <="' + UsDateTime(FindeMois(DateFin)) + '")'; // PT8

end;

procedure TOF_PG_BulCompl_mul.OnLoad;
var
  TT: TFMul;
begin
  inherited;
  TT := TFMul(Ecran);
  TFMul(Ecran).Caption := 'Création d''un bulletin complémentaire au ' + DateToStr(DateDebut);
  if TT <> nil then UpdateCaption(TT);
  ActiveWhere;
end;

initialization
  registerclasses([TOF_PG_BulCompl_mul]);
end.

