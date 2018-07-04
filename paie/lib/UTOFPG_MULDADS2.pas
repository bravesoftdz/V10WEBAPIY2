{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/04/2003
Modifié le ... :
Description .. : Source TOF de la FICHE : PGMULDADS2 ()
Suite ........ : Unité de gestion du multicritère TD Bilatéral
Mots clefs ... : TOF;PGMULDADS2;PAIE;PGDADSB
*****************************************************************}
{
PT1   : 31/03/2004 VG V_50 On ne pose plus la question pour la réinitialisation
                           du fichier de log. Le fichier est réinitialisé si le
                           précédent calcul datait de plus de 6 mois.
                           Le fichier est systématiquement ouvert en fin de
                           traitement
PT2   : 14/05/2004 VG V_50 Ajout du traitement lié aux raccourcis clavier
PT3   : 17/10/2006 VG V_70 Suppression du fichier de contrôle - mise en table
                           des erreurs
PT4   : 09/01/2007 VG V_72 Suppression de l'ouverture du fichier de log
                           FQ N°13818
PT5   : 28/03/2007 VG V_70 Portage CWAS
PT6   : 12/11/2007 NA V_80 Mise à jour du code section à 01 dans table ETABLCOMPL si section = 00
}
unit UTOFPG_MULDADS2;

interface

uses StdCtrls,
  Controls,
  Classes,
{$IFNDEF EAGLCLIENT}
  db,
  HDB,
  mul,
  FE_Main,
  ParamSoc,
  ShellAPI,
{$ELSE}
  emul,
  MaineAgl,
  Utob,
{$ENDIF}
  sysutils,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  Hqry,
  HTB97,
  PgDADSCommun,
  PgDADSBilaterale,
  PgOutils2,
  AGLInit,
  ed_tools,
  hstatus,
  EntPaie,
  windows;

type
  TOF_PGMULDADS2 = class(TOF)
  public
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;

  private
    Validite: THValComboBox;
    SALAR: THEdit;
    Q_Mul: THQuery; // Query pour changer la liste associee
    Calculer: TToolbarButton97;

{$IFNDEF EAGLCLIENT}
    Liste: THDBGrid;
{$ELSE}
    Liste: THGrid;
{$ENDIF}
    param: string;

    procedure ActiveWhere(Sender: TObject);
    procedure GrilleDblClick(Sender: TObject);
    procedure CalculerClick(Sender: TObject);
    procedure InitCalcul(NomFic: string);
    procedure Calcul_un;
    procedure SalarieExit(Sender: TObject);
    procedure DateChange(Sender: Tobject);
    procedure Parametrage;
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  end;

implementation

procedure TOF_PGMULDADS2.OnNew;
begin
  inherited;
end;


procedure TOF_PGMULDADS2.OnDelete;
begin
  inherited;
end;


procedure TOF_PGMULDADS2.OnUpdate;
begin
  inherited;
  if (param = 'C') then
  begin
    if Calculer <> nil then
      Calculer.Enabled := True;
  end
  else
    TFMul(Ecran).BOuvrir.Enabled := True;

  if (PGAnnee = '') then
  begin
    TFMul(Ecran).BOuvrir.Enabled := False;
    if Calculer <> nil then
      Calculer.Enabled := False;
  end;
end;

procedure TOF_PGMULDADS2.OnLoad;
begin
  inherited;
  ActiveWhere(nil);
end;

procedure TOF_PGMULDADS2.OnArgument(S: string);
var
  AnneePrec, arg: string;
  TT: TFMul;
  JourJ: TDateTime;
  AnneeA, Jour, MoisM: Word;
begin
  inherited;
  arg := S;
  param := Trim(ReadTokenPipe(arg, ';'));
  TT := TFMul(Ecran);

  if param = 'C' then
    TFMul(Ecran).Caption := 'Calcul TD Bilatéral';
  if param = 'S' then
  begin
    TFMul(Ecran).Caption := 'Saisie périodes d''activité TD Bilatéral';
    SetControlText('CBCALC', 'X');
  end;

  // pt6
   ExecuteSQL('UPDATE ETABCOMPL SET ETB_DADSSECTION="01" WHERE ETB_DADSSECTION="00"');



  if TT <> nil then
    UpdateCaption(TT);
  if (param = 'C') then
  begin
    Calculer := TToolbarButton97(GetControl('BCALCULER'));
    if Calculer <> nil then
    begin
      Calculer.Visible := True;
      Calculer.Enabled := False;
      Calculer.OnClick := CalculerClick;
    end;
    TFMul(Ecran).BOuvrir.Visible := False;
  end;

  if (param = 'S') then
  begin
    TFMul(Ecran).BOuvrir.Enabled := False;
    TFMul(Ecran).BOuvrir.OnClick := GrilleDblClick;

    TFMul(Ecran).bSelectAll.Visible := False;

{$IFNDEF EAGLCLIENT}
    Liste := THDBGrid(GetControl('FListe'));
{$ELSE}
    Liste := THGrid(GetControl('FListe'));
{$ENDIF}
    if Liste <> nil then
    begin
{$IFNDEF EAGLCLIENT}
      Liste.MultiSelection := False;
{$ENDIF}
      Liste.OnDblClick := GrilleDblClick;
    end;
  end;

  TFMul(Ecran).OnKeyDown := FormKeyDown; //PT2

  Q_Mul := THQuery(Ecran.FindComponent('Q'));

  SALAR := THEdit(GetControl('PSA_SALARIE'));

  Validite := THValComboBox(GetControl('VALIDITE'));

  if SALAR <> nil then
    SALAR.OnExit := SalarieExit;

  JourJ := Date;
  DecodeDate(JourJ, AnneeA, MoisM, Jour);
  if MoisM > 9 then
    AnneePrec := IntToStr(AnneeA)
  else
    AnneePrec := IntToStr(AnneeA - 1);

  if Validite <> nil then
  begin
    Validite.value := copy(AnneePrec, 1, 1) + copy(AnneePrec, 3, 2);
    PGAnnee := Validite.value;
    PGExercice := AnneePrec;
    Validite.OnChange := DateChange;
  end;

  SetControlText('L_DDU', '01/01/' + AnneePrec);
  SetControlText('L_DAU', '31/12/' + AnneePrec);

  DebExer := StrToDate(GetControlText('L_DDU'));
  FinExer := StrToDate(GetControlText('L_DAU'));

  ActiveWhere(nil);
end;

procedure TOF_PGMULDADS2.OnClose;
begin
  inherited;
end;

procedure TOF_PGMULDADS2.OnDisplay;
begin
  inherited;
end;

procedure TOF_PGMULDADS2.OnCancel;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/04/2003
Modifié le ... :   /  /
Description .. : XX_WHERE
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}

procedure TOF_PGMULDADS2.ActiveWhere(Sender: TObject);
begin
  if (Validite <> nil) then
  begin
    if Q_Mul <> nil then
      TFMul(Ecran).SetDBListe('PGMULSALDADS2');

    if ((DebExer = 0) or (FinExer = 0)) then
      SetControlText('XX_WHERE', '')
    else
    begin
      if (GetCheckBoxState('CBCALC') = cbChecked) then
        SetControlText('XX_WHERE', ' PD2_VALIDITE="' + RechDom('PGANNEE', PGAnnee, FALSE) + '" AND' +
          ' ((PSA_DATEENTREE <= "' + UsDateTime(FinExer) + '" AND' +
          ' (PSA_DATESORTIE >= "' + UsDateTime(DebExer) + '" OR' +
          ' PSA_DATESORTIE = "' + UsDateTime(IDate1900) + '" OR' +
          ' PSA_DATESORTIE IS NULL)) OR' +
          ' PSA_SALARIE IN (SELECT PPU_SALARIE FROM PAIEENCOURS WHERE' +
          ' PPU_DATEDEBUT>= "' + UsDateTime(DebExer) + '" AND' +
          ' PPU_DATEFIN <= "' + UsDateTime(FinExer) + '"))')
      else
        SetControlText('XX_WHERE', ' (PD2_VALIDITE IS NULL OR' +
          ' PD2_VALIDITE<>"' + RechDom('PGANNEE', PGAnnee, FALSE) + '") AND' +
          ' ((PSA_DATEENTREE <= "' + UsDateTime(FinExer) + '" AND' +
          ' (PSA_DATESORTIE >= "' + UsDateTime(DebExer) + '" OR' +
          ' PSA_DATESORTIE = "' + UsDateTime(IDate1900) + '" OR' +
          ' PSA_DATESORTIE IS NULL)) OR' +
          ' PSA_SALARIE IN (SELECT PPU_SALARIE FROM PAIEENCOURS WHERE' +
          ' PPU_DATEDEBUT>= "' + UsDateTime(DebExer) + '" AND' +
          ' PPU_DATEFIN <= "' + UsDateTime(FinExer) + '"))')
    end;
    if (param = 'C') then
    begin
      if Calculer <> nil then
        Calculer.Enabled := False;
    end
    else
      TFMul(Ecran).BOuvrir.Enabled := False;
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/04/2003
Modifié le ... :   /  /
Description .. : Double-click sur la grille
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOF_PGMULDADS2.GrilleDblClick(Sender: TObject);
var
Annee, Salarie, StSal: string;
begin
Annee:= '';
if ((Q_Mul <> nil) and (Q_Mul.RecordCount = 0)) then
   exit;

{$IFDEF EAGLCLIENT}
TFMul (Ecran).Q.TQ.Seek (TFMul (Ecran).FListe.Row-1);    //PT5
{$ENDIF}

Salarie:= Q_Mul.FindField ('PSA_SALARIE').AsString;

if (PGAnnee <> '') then
   Annee:= RechDom ('PGANNEE', PGAnnee, FALSE);

StSal:= 'SELECT PD2_SALARIE, PD2_VALIDITE, PD2_SIRET, PD2_SECTIONETAB,' +
        ' PD2_TYPEDADS'+
        ' FROM DADS2SALARIES WHERE'+
        ' PD2_SALARIE="'+Salarie+'" AND'+
        ' PD2_VALIDITE="'+Annee+'"';
{$IFNDEF EAGLCLIENT}
TheMulQ:= THQuery (Ecran.FindComponent('Q'));
{$ELSE}
TheMulQ:= TOB (Ecran.FindComponent('Q'));
{$ENDIF}
if (PGAnnee <> '') then
   begin
   if (ExisteSQL (StSal) = FALSE) then
      begin
      ChargeZones (Salarie);
      TFMul (Ecran).BCherche.Click;
      end
   else
      AGLLanceFiche ('PAY', 'DADS2_SALARIE', '', Salarie+';'+Annee,
                     'ACTION=MODIFICATION;'+Salarie+';'+Annee);
   end
else
   PGIBox('L''année n''est pas valide', 'TD Bilatéral');
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/04/2003
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "Calculer"
Suite ........ : ou "Valider" (pour l'instant)
Mots clefs ... : PAIE,PGDADSB
*****************************************************************}

procedure TOF_PGMULDADS2.CalculerClick(Sender: TObject);
var
  BufDest, NomFic: string;
  i: integer;
  Maintenant: TDateTime;
begin
{$IFNDEF EAGLCLIENT}
  Liste := THDBGrid(GetControl('FListe'));
{$ELSE}
  Liste := THGrid(GetControl('FListe'));
{$ENDIF}
{$IFNDEF EAGLCLIENT}
  if Liste <> nil then
  begin
    if (Liste.NbSelected = 0) and (not Liste.AllSelected) then
    begin
      MessageAlerte('Aucun élément sélectionné');
      exit;
    end;

    ForceNumerique(GetParamSoc('SO_SIRET'), BufDest);
    if ControlSiret(BufDest) = False then
      PGIBox('Le SIRET de la société n''est pas valide.#13#10' +
        'Vous devez le vérifier en y accédant par le module#13#10' +
        'Paramètres/menu Comptabilité/commande Paramètres comptables/Coordonnées.#13#10' +
        'Si vous travaillez en environnement multi-dossiers,#13#10' +
        'vous pouvez y accéder par le Bureau PGI/Annuaire',
        'Calcul TD Bilatéral');
{$IFDEF EAGLCLIENT}
    NomFic := VH_Paie.PgCheminEagl + '\' + BufDest + '_TDB_PGI.log';
{$ELSE}
    NomFic := V_PGI.DatPath + '\' + BufDest + '_TDB_PGI.log';
{$ENDIF}
    if (Liste.AllSelected = TRUE) then
    begin
      InitMoveProgressForm(nil, 'Calcul en cours', 'Veuillez patienter SVP ...', TFmul(Ecran).Q.RecordCount, FALSE, TRUE);
      InitMove(TFmul(Ecran).Q.RecordCount, '');
      InitCalcul(NomFic);
      TFmul(Ecran).Q.First;
      while not TFmul(Ecran).Q.EOF do
      begin
        Calcul_un;
        TFmul(Ecran).Q.Next;
      end;
      Liste.AllSelected := False;
      TFMul(Ecran).bSelectAll.Down := Liste.AllSelected;
    end
    else
    begin
      InitMoveProgressForm(nil, 'Calcul en cours', 'Veuillez patienter SVP ...', Liste.NbSelected, FALSE, TRUE);
      InitMove(Liste.NbSelected, '');
      InitCalcul(NomFic);
      for i := 0 to Liste.NbSelected - 1 do
      begin
        Liste.GotoLeBOOKMARK(i);
        Calcul_un;
      end;
      Liste.ClearSelected;
    end;

    PGIBox('Traitement terminé', 'Calcul TD Bilatéral');
    Maintenant := Now;
{PT3
    Writeln(FRapport, 'Calcul TD Bilatéral terminé : ' + DateTimeToStr(Maintenant));
    CloseFile(FRapport);
}    

    FiniMove;
    FiniMoveProgressForm;
{PT4
    ShellExecute(0, PCHAR('open'), PChar('WordPad'), PChar(NomFic), nil, SW_RESTORE);
}    
  end;
{$ENDIF}

  if Calculer <> nil then
    Calculer.Enabled := False;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/05/2004
Modifié le ... :   /  /
Description .. : Initialisation du calcul de la TD Bilatéral
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}

procedure TOF_PGMULDADS2.InitCalcul(NomFic: string);
var
  DateCalcul, Maintenant: TDateTime;
  FileAttrs: integer;
  sr: TSearchRec;
begin
  if FileExists(NomFic) then
  begin
    Maintenant := Now;
    DateCalcul := Now;
    FileAttrs := 0;
    FileAttrs := FileAttrs + faAnyFile;
    if FindFirst(NomFic, FileAttrs, sr) = 0 then
    begin
      if (sr.Attr and FileAttrs) = sr.Attr then
        DateCalcul := FileDateToDateTime(sr.Time);
      sysutils.FindClose(sr);
    end;

    if (PlusMois(Maintenant, -6) > DateCalcul) then
      DeleteFile(PChar(NomFic));
  end;

{PT3
  AssignFile(FRapport, NomFic);
  if FileExists(NomFic) then
  begin
    Append(FRapport);
    Writeln(FRapport, '');
  end
  else
  begin
    ReWrite(FRapport);
    Writeln(FRapport, 'Attention, Le dernier calcul se trouve en fin du fichier');
  end;

  Writeln(FRapport, '_____________________________________');
  Writeln(FRapport, 'Début de calcul : ' + DateTimeToStr(Now));
}  
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/05/2004
Modifié le ... :   /  /
Description .. : Calcul d'un élément
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}

procedure TOF_PGMULDADS2.Calcul_un;
var
  St: string;
  Maintenant: TDateTime;
begin
  St := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
  if St <> '' then
  begin
    if (isnumeric(St) and (VH_PAIE.PgTypeNumSal = 'NUM')) then
      St := ColleZeroDevant(StrToInt(St), 10);
    try
      begintrans;
      ExecuteSQL('DELETE FROM DADS2SALARIES WHERE' +
        ' PD2_SALARIE="' + St + '" AND' +
        ' PD2_VALIDITE="' + RechDom('PGANNEE', PGAnnee, FALSE) + '"');
      ChargeTOBSalB (St);
      DeleteErreur (St, 'PDB');	//PT3
      CalculDADSB (St);
      LibereTOBB;
      CommitTrans;
    except
      Rollback;
      Maintenant := Now;
{PT3
      Writeln(FRapport, 'Salarié ' + St + ' : Calcul TD Bilatéral annulé : ' + DateTimeToStr(Maintenant));
}
    end;
  end;
  MoveCur(False);
  MoveCurProgressForm(St);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/04/2003
Modifié le ... :   /  /
Description .. : Sortie de la zone de saisie du code salarié
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}

procedure TOF_PGMULDADS2.SalarieExit(Sender: TObject);
begin

  if (isnumeric(SALAR.Text) and (VH_PAIE.PgTypeNumSal = 'NUM')) then
    SALAR.Text := ColleZeroDevant(StrToInt(SALAR.Text), 10);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/04/2003
Modifié le ... :   /  /
Description .. : Modification de la date
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}

procedure TOF_PGMULDADS2.DateChange(Sender: TObject);
begin
  Parametrage;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/04/2003
Modifié le ... :   /  /
Description .. : Gestion de modification de la nature
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}

procedure TOF_PGMULDADS2.Parametrage;
begin
  PGExercice := RechDom('PGANNEE', Validite.Value, False);
  PGAnnee := Validite.value;

  SetControlText('L_DDU', '01/01/' + PGExercice);
  SetControlText('L_DAU', '31/12/' + PGExercice);

  DebExer := StrToDate(GetControlText('L_DDU'));
  FinExer := StrToDate(GetControlText('L_DAU'));
end;

//PT2
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 12/05/2004
Modifié le ... :   /  /
Description .. : Complément des raccourcis claviers
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}

procedure TOF_PGMULDADS2.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  TFMul(Ecran).FormKeyDown(Sender, Key, Shift);
  case Key of
    VK_F6: if ((GetControlVisible('BCALCULER')) and
        (GetControlEnabled('BCALCULER'))) then
        Calculer.Click; //Calcul des éléments
  end;
end;
//FIN PT2

initialization
  registerclasses([TOF_PGMULDADS2]);
end.

