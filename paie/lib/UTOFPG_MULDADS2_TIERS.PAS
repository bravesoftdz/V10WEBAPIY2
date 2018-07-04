{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 24/08/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : MUL_DADS2TIERS ()
Mots clefs ... : TOF;MUL_DADS2TIERS
*****************************************************************}
{
PT1   : 17/09/2004 VG V_50 On ne prenait en compte que la première ligne
                           d'écriture - FQ N°11580
PT2   : 08/10/2004 VG V_50 Adaptation cahier des charges 2004
PT3   : 17/11/2004 VG V_60 Suppression du filtre par défaut qui provoquait une
                           erreur - FQ N°11770
PT4   : 02/05/2006 VG V_65 Alimentation optionnelle à partir des journaux de
                           type "OD" - FQ N°12366
PT5   : 19/05/2006 VG V_65 Lors de l'import d'honoraires, édition des honoraires
                           concernés - FQ N°12368
PT6   : 04/09/2006 VG V_70 Edition pour tous les comptes auxiliaires
                           sélectionnés - FQ N°12368
PT7-1 : 17/10/2006 VG V_70 Suppression du fichier de contrôle - mise en table
                           des erreurs
PT7-2 : 17/10/2006 VG V_70 Utilisation d'un type pour la cle DADS-U
PT8   : 20/03/2007 TJ V_80 Passage vers la compta
}
unit UTOFPG_MULDADS2_TIERS;

interface

uses
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     HDB,
     FE_Main,
     EdtREtat,
{$ELSE}
     eMul,
     entpaie,
     MaineAGL,
     UtileAGL,
{$ENDIF}
     uTob,
     sysutils,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
{$IFDEF COMPTA}
     UtilPGI,         //VerifSiret
     UtilTrans,       //EpureChar
{$ELSE}
     PgOutils2,
{$ENDIF}
     PgDADSCommun,

     Hqry,
     HTB97,
     paramsoc,
     ed_tools,
     hstatus,
     ShellAPI,
     windows,
     hdebug,
//PT-8
{$IFDEF COMPTA}
   {$IFDEF MODENT1}
   CPTypeCons,
   {$ENDIF MODENT1}
     vierge,        //TFvierge
     ent1,          // VH^, exotodate, ...
     Forms,
     uLibExercice,   // CQuelExercice
{$ENDIF}
     ComCtrls;

type
  TOF_MUL_DADS2TIERS = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;

  private
    Validite : THValComboBox;
    Q_Mul : THQuery; // Query pour changer la liste associee
    Calculer : TToolbarButton97;
    THonoCompta, TListeHono : TOB;
    Pages : TPageControl;
    LCptDu  : THEdit;
    LCptAu  : THEdit;
    Exo : String;
{$IFDEF COMPTA}
    Insert  : boolean;
    fOnSaveKeyDownEcran   : procedure (Sender : TObject; var Key: Word; Shift: TShiftState) of Object;
{$ENDIF}
{$IFNDEF EAGLCLIENT}
    Liste: THDBGrid;
{$ELSE}
    Liste: THGrid;
{$ENDIF}

    procedure ActiveWhere(Sender: TObject);
    procedure DateChange(Sender: Tobject);
    procedure Parametrage;
    procedure CalculerClick(Sender: TObject);
    procedure InitCalcul({NomFic: string});
    procedure Calcul_un (bRemplace : boolean);
    procedure LibereTOBHonB;
    procedure ChargeTOBHonB(Honoraire: string);
    procedure CalculHonDADSB(Honoraire: string);
{$IFDEF COMPTA}          //PT-8
    procedure LaFormOnKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
{$ENDIF}

  end;

//===============================================
//PT-8
{$IFDEF COMPTA}
  Type
    TOF_QU_DADS2TIERS = Class(TOF)
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure BValideOnClick (Sender : Tobject);
    procedure CaseChecked (Sender : TObject); // BVE 19.06.07
    private

      QCalcul1    : THCheckbox;
      QCalcul2    : THCheckbox;
      BValide     : TToolbarButton97;
  end;
{$ENDIF}
//===============================================

  procedure LanceMul_DADS2Tiers;

implementation

uses TntStdCtrls, DateUtils;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 23/02/2007
Modifié le ... :   /  /
Description .. : procédure pour appel depuis unités externes
Mots clefs ... : 
*****************************************************************}
procedure LanceMul_DADS2Tiers; //PT-8
begin
  AGLLanceFiche('PAY', 'MUL_DADS2TIERS', '', '', '');
end;

procedure TOF_MUL_DADS2TIERS.OnNew;
begin
inherited;
end;

procedure TOF_MUL_DADS2TIERS.OnDelete;
begin
inherited;
end;

procedure TOF_MUL_DADS2TIERS.OnUpdate;
begin
inherited;
if ((PGAnnee = '') or (GetCheckBoxState('CBCALC') = cbChecked)) then
   SetControlEnabled('BCALCULER', False)
else
   SetControlEnabled('BCALCULER', True);
end;

procedure TOF_MUL_DADS2TIERS.OnLoad;
begin
inherited;
ActiveWhere(nil);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 24/08/2004
Modifié le ... :   /  /
Description .. : OnArgument
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOF_MUL_DADS2TIERS.OnArgument(S: string);
var
  JourJ : TDateTime;
  AnneeA, Jour, MoisM : Word;
  AnneePrec : string;
begin
inherited;
TFMul(Ecran).FiltreDisabled:= True; //PT3
SetControlText('T_NATUREAUXI', 'FOU');

  // éléments spécifiques à la compta
{$IFDEF COMPTA}    //PT-8
  TFMul(Ecran).Caption  := 'Récupération des données comptables';
  UpdateCaption(TFMul(Ecran));
  // on efface les éléments qui ne servent pas à la compta
  SetControlVisible('L_DDU', False);
  SetControlVisible('L_DAU', False);
  fOnSaveKeyDownEcran   := TFMul(Ecran).OnKeyDown;
  TFMul(Ecran).OnKeyDown := LaFormOnKeyDown;
{$ENDIF}

Calculer:= TToolbarButton97(GetControl('BCALCULER'));
if Calculer <> nil then
   begin
   Calculer.Visible:= True;
   Calculer.Enabled:= False;
   Calculer.OnClick:= CalculerClick;
   end;

Q_Mul:= THQuery(Ecran.FindComponent('Q'));

Validite:= THValComboBox(GetControl('ANNEE'));

//Gestion de la combo de l'année qui devient l'exercice
{$IFDEF COMPTA}  //PT-8
  LCptDu  := THEdit(GetControl('LCPTDU'));
  { FQ 20722 BVE 18.06.07 }
  LCptAu  := THEdit(GetControl('LCPTDU_'));
  { END FQ 20722 }

  LCptDu.Visible  := True;
  LCptAu.Visible  := True;
  // SetControlText('LANNEE', 'E&xercice');
  // Validite.DataType := 'TTEXERCICE';
  Validite.OnChange := DateChange;

  Exo := VH^.Precedent.Code;
  if Exo = '' then
    Exo := VH^.EnCours.Code;

  // Validite.Value := Exo;

  { FQ 20719 BVE 25.01.08
  JourJ:= VH^.Encours.Fin; }
  JourJ:= VH^.Encours.Deb;
  { END FQ 20719 }
  
  DecodeDate(JourJ, AnneeA, MoisM, Jour);
  AnneePrec:= IntToStr(AnneeA);

  if Validite <> nil then
  begin
     Validite.value:= copy(AnneePrec, 1, 1)+copy(AnneePrec, 3, 2);
     PGAnnee:= Validite.value;
     PGExercice:= AnneePrec;
     Validite.OnChange:= DateChange;
   end;

{$ELSE}

  JourJ:= Date;
  DecodeDate(JourJ, AnneeA, MoisM, Jour);
  if MoisM > 9 then
     AnneePrec:= IntToStr(AnneeA)
  else
     AnneePrec:= IntToStr(AnneeA-1);

  if Validite <> nil then
     begin
     Validite.value:= copy(AnneePrec, 1, 1)+copy(AnneePrec, 3, 2);
     PGAnnee:= Validite.value;
     PGExercice:= AnneePrec;
     Validite.OnChange:= DateChange;
     end;

  SetControlText('L_DDU', '01/01/'+AnneePrec);
  SetControlText('L_DAU', '31/12/'+AnneePrec);

  DebExer:= StrToDate(GetControlText('L_DDU'));
  FinExer:= StrToDate(GetControlText('L_DAU'));
{$ENDIF}

Pages:= TPageControl (GetControl ('Pages')); //PT5

ActiveWhere(nil);
end;

procedure TOF_MUL_DADS2TIERS.OnClose;
begin
inherited;
end;

procedure TOF_MUL_DADS2TIERS.OnDisplay;
begin
inherited;
end;

procedure TOF_MUL_DADS2TIERS.OnCancel;
begin
inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 06/09/2004
Modifié le ... :   /  /
Description .. : XX_WHERE
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOF_MUL_DADS2TIERS.ActiveWhere(Sender: TObject);
begin
if (Validite <> nil) then
   begin
   if Q_Mul <> nil then
      begin
      if (GetCheckBoxState('CBCALC') = cbChecked) then
         TFMul(Ecran).SetDBListe('PGMULHONORDADS2')
      else
         TFMul(Ecran).SetDBListe('PGMULTIERSDADS2');
      end;

   if ((DebExer = 0) or (FinExer = 0)) then
      SetControlText('XX_WHERE', '')
   else
      begin
      if (GetCheckBoxState('CBCALC') = cbChecked) then
         begin
         SetControlText ('T_EAN', '');
         SetControlText ('T_DEVISE', '');
         SetControlText ('T_LIBELLE', '');
         SetControlText ('T_NATUREAUXI', '');
         SetControlText ('T_COLLECTIF', '');
         SetControlText ('T_SECTEUR', '');
         SetControlText ('T_TARIFTIERS', '');
         SetControlText ('T_REGIMETVA', '');
         SetControlText ('T_FACTURE', '');
         SetControlText ('T_PAYEUR', '');
         SetControlProperty ('T_ISPAYEUR', 'State', cbGrayed);
         SetControlProperty ('T_LETTRABLE', 'State', cbGrayed);
         SetControlProperty ('T_MULTIDEVISE', 'State', cbGrayed);
         SetControlVisible ('bSelectAll', False);
         SetControlText ('XX_WHERE', ' PDH_VALIDITE="'+
                         RechDom('PGANNEE', PGAnnee, FALSE)+'"');
         end
      else
         begin
         SetControlText('T_NATUREAUXI', 'FOU');
         SetControlVisible('bSelectAll', True);
{$IFDEF COMPTA}
         SetControlText('XX_WHERE', ' YTC_DAS2="X" AND YTC_REMUNERATION <> ""');
{$ELSE}
         SetControlText('XX_WHERE', ' YTC_REMUNERATION <> ""');
{$ENDIF}
         end;
      end;

   SetControlEnabled('BCALCULER', False);
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 06/09/2004
Modifié le ... :   /  /
Description .. : Modification de la date
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOF_MUL_DADS2TIERS.DateChange(Sender: TObject);
begin
Parametrage;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 06/09/2004
Modifié le ... :   /  /
Description .. : Gestion de modification de la nature
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOF_MUL_DADS2TIERS.Parametrage;
{$IFDEF COMPTA}
var
  DateCalcul : TDateTime;
  i : integer;
  Exo : TExoDate;
{$ENDIF}
begin
{$IFDEF COMPTA}    //PT-8
(*
  Exo := Validite.Value;
  ExoToDates(Exo, LCptDu, LCptAu);
  DebExer := StrToDate(LCptDu.Text);
  FinExer := StrToDate(LCptAu.Text);
  PGAnnee := FormatDateTime('yyyy', DebExer);
*)
  { FQ 21015 BVE 11.07.07 }
  PGAnnee:= Validite.value;
  { END FQ 21015 }
  PGExercice:= RechDom('PGANNEE', Validite.Value, False);
  DateCalcul := EncodeDate (StrToInt(PGExercice),12,31);
  if CQuelExercice(DateCalcul, Exo) then
  begin
    if (Exo.Fin=DateCalcul) then
    begin
      SetControlText('LCPTDU', '01/01/'+PGExercice);
      { FQ 20722 BVE 18.06.07 }
      SetControlText('LCPTDU_', '31/12/'+PGExercice);
      { END FQ 20722 }
    end else
    begin
      i :=1;
      while ((VH^.Exercices[i].Code<>'') and (VH^.Exercices[i].Code<>Exo.Code)) do Inc (i,1);
      if ((i > 1 ) and (VH^.Exercices[i].Code<>'')) then
      begin
        { FQ 20719 BVE 18.06.07 }
        SetControlText('LCPTDU', '01/01/' + IntToStr(YearOf(VH^.Exercices[i].Deb)));
        { FQ 20722 BVE 18.06.07 }
        SetControlText('LCPTDU_', '31/12/' + IntToStr(YearOf(VH^.Exercices[i].Deb)));
        { END FQ 20722 }
        { END FQ 20719 }
      end else
      begin
        SetControlText('LCPTDU', '01/01/'+PGExercice);
        { FQ 20722 BVE 18.06.07 }
        SetControlText('LCPTDU_', '31/12/'+PGExercice);
        { END FQ 20722 }
      end;
    end;
  end else
  begin
    SetControlText('LCPTDU', '01/01/'+PGExercice);
    { FQ 20722 BVE 18.06.07 }
    SetControlText('LCPTDU_', '31/12/'+PGExercice);
    { END FQ 20722 }
  end;
  DebExer:= StrToDate(GetControlText('LCPTDU'));
  { FQ 20722 BVE 18.06.07 }
  FinExer:= StrToDate(GetControlText('LCPTDU_'));
  { END FQ 20722 }
  SetControlText('L_DDU', '01/01/'+PGExercice);
  SetControlText('L_DAU', '31/12/'+PGExercice);

{$ELSE}
  PGExercice:= RechDom('PGANNEE', Validite.Value, False);
  PGAnnee:= Validite.value;

  SetControlText('L_DDU', '01/01/'+PGExercice);
  SetControlText('L_DAU', '31/12/'+PGExercice);

  DebExer:= StrToDate(GetControlText('L_DDU'));
  FinExer:= StrToDate(GetControlText('L_DAU'));
{$ENDIF}
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 07/09/2004
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "Calculer"
Mots clefs ... : PAIE,PGDADSB
*****************************************************************}
procedure TOF_MUL_DADS2TIERS.CalculerClick(Sender: TObject);
var
BufDest : String;
i : integer;
{$IFNDEF COMPTA} //PT-8
NomFic: string;
reponse: integer;
{PT7-1
Maintenant: TDateTime;
}
{$ELSE}
  LaRep : string;
  bInitG  : Boolean;
  bRempl  : Boolean;
  StrSql  : String;
{$ENDIF}

begin
{$IFNDEF EAGLCLIENT}
Liste:= THDBGrid(GetControl('FListe'));
{$ELSE}
Liste:= THGrid(GetControl('FListe'));
{$ENDIF}

{$IFDEF COMPTA}    //PT-8
  LaRep := AGLLanceFiche('PAY', 'QU_DADS2TIERS', '', '', TFMul(Ecran).Caption + ';' + RechDom('PGANNEE', PGAnnee, FALSE));

  bInitG  := False;
  bRempl  := False;

  Case VALEURI(LaRep) of
    1 : //Init générale du fichier
      bInitG  := True;

    2 : //Rempl les honoraires existant
      bRempl  := True;

    3 : // 1 + 2
      begin
        bInitG  := True;
        bRempl  := True;
      end;
  end;
{$ELSE}
  reponse:= PGIAsk ('Cette commande supprimera les honoraires ayant comme#13#10'+
                    'code le numéro d''un compte auxiliaire sélectionné.#13#10'+
                    'Voulez-vous continuer ?', TFMul(Ecran).Caption);
  if (reponse <> mrYes) then
    exit;
{$ENDIF}


if Liste <> nil then
   begin
   if (Liste.NbSelected = 0) and (not Liste.AllSelected) then
      begin
      MessageAlerte('Aucun élément sélectionné');
      exit;
      end;

   TListeHono:= Tob.Create ('Les honoraires', nil, -1);  //PT6
{$IFDEF COMPTA}
  EpureChar(GetParamSoc('SO_SIRET'), BufDest);
  if not VerifSiret(BufDest) then
{$ELSE}
   ForceNumerique(GetParamSoc('SO_SIRET'), BufDest);
   if ControlSiret(BufDest) = False then
{$ENDIF}
{ FQ 22288 BVE 28.01.08 }
{$IFDEF COMPTA}
      PGIBox ('Le numéro SIRET de la société n''est pas valide. #13#10' +
              'Vous devez le vérifier par l''intermédiaire du Dossier Client.');
{$ELSE}
      PGIBox ('Le SIRET de la société n''est pas valide.#13#10'+
              'Vous devez le vérifier en y accédant par le module#13#10'+
              'Paramètres/menu Comptabilité/commande Paramètres comptables/Coordonnées.#13#10'+
              'Si vous travaillez en environnement multi-dossiers,#13#10'+
              'vous pouvez y accéder par le Bureau PGI/Annuaire',
              'Calcul TD Bilatéral');
{$ENDIF}
{ END FQ 22288 }
{$IFDEF EAGLCLIENT}
{PT7-1
   NomFic:= VH_Paie.PgCheminEagl+'\'+BufDest+'_TDB_PGI.log';
}
{$ELSE}
{PT7-1
    NomFic := V_PGI.DatPath + '\' + BufDest + '_TDB_PGI.log';
}
{$ENDIF}

{$IFDEF COMPTA}   //PT-8
  Insert := false;
  // si initialisation générale du fichier des honoraires pour l'année ....
  if bInitG then
  begin
    Try
      BEGINTRANS;
      StrSql  := 'DELETE FROM DADS2HONORAIRES WHERE PDH_VALIDITE="'+RechDom('PGANNEE', PGAnnee, FALSE)+'"';
      ExecuteSQL(StrSql);
      COMMITTRANS;
    Except
      on E:Exception do
      begin
        ROLLBACK;
        MessageAlerte('Erreur : ' + E.Message );
      end;
    end;
  end;
{$ENDIF}
   if (Liste.AllSelected = TRUE) then
      begin
      InitMoveProgressForm (nil, 'Calcul en cours',
                            'Veuillez patienter SVP ...',
                            TFmul(Ecran).Q.RecordCount, FALSE, TRUE);
      InitMove (TFmul(Ecran).Q.RecordCount, '');
{PT7-1
      InitCalcul(NomFic);
}
      InitCalcul;
//FIN PT7-1
      TFmul(Ecran).Q.First;
      while not TFmul(Ecran).Q.EOF do
            begin
            {$IFDEF COMPTA}  //PT-8
            Calcul_un(bRempl);
            {$ELSE}
            Calcul_un (True);
            {$ENDIF}
            TFmul(Ecran).Q.Next;
            end;
      Liste.AllSelected:= False;
      TFMul(Ecran).bSelectAll.Down:= Liste.AllSelected;
      end
   else
      begin
      InitMoveProgressForm (nil, 'Calcul en cours',
                            'Veuillez patienter SVP ...', Liste.NbSelected,
                            FALSE, TRUE);
      InitMove (Liste.NbSelected, '');
{PT7-1
      InitCalcul(NomFic);
}
      InitCalcul;
//FIN PT7-1
      for i := 0 to Liste.NbSelected - 1 do
          begin
          Liste.GotoLeBOOKMARK(i);
          {$IFDEF COMPTA} //PT-8
          Calcul_un (bRempl)
          {$ELSE}
          Calcul_un (True);
          {$ENDIF}
          end;
      Liste.ClearSelected;
      end;
{ FQ 20689 BVE 19.06.07 }
{$IFDEF COMPTA}
   if not(Insert) then
      PGIBox('Toutes les données ont déjà été importées')
   else
{$ENDIF}
{ END FQ 20689 }
   PGIBox ('Traitement terminé', 'Calcul TD Bilatéral');
{PT7-1
   Maintenant:= Now;
   Writeln (FRapport,
            'Calcul TD Bilatéral terminé : '+DateTimeToStr(Maintenant));
   CloseFile (FRapport);
}
   FiniMove;
   FiniMoveProgressForm;
{PT7-1
   ShellExecute (0, PCHAR('open'), PChar('WordPad'), PChar(NomFic), nil,
                 SW_RESTORE);
}
   end;

//PT5
{$IFDEF COMPTA}  //PT-8
  if PGIAsk('Voulez vous lancer l''édition ?', TFMul(Ecran).Caption) = mrYes then
{$ENDIF}
    LanceEtatTob ('E', 'PDB', 'PDI', TListeHono, True, False, False, Pages,
                 '', '', False);
//FIN PT5
FreeAndNil (TListeHono);      //PT6
{$IFNDEF COMPTA}
if Calculer <> nil then
   Calculer.Enabled:= False;
{$ENDIF}
end;



{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 07/09/2004
Modifié le ... :   /  /
Description .. : Initialisation du calcul de la TD Bilatéral
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOF_MUL_DADS2TIERS.InitCalcul({NomFic: string});
(*
var
DateCalcul, Maintenant: TDateTime;
FileAttrs: integer;
sr: TSearchRec;
*)
begin
{PT7-1
if FileExists(NomFic) then
   begin
   Maintenant:= Now;
   DateCalcul:= Now;
   FileAttrs:= 0;
   FileAttrs:= FileAttrs+faAnyFile;
   if FindFirst(NomFic, FileAttrs, sr) = 0 then
      begin
      if (sr.Attr and FileAttrs) = sr.Attr then
         DateCalcul:= FileDateToDateTime(sr.Time);
      sysutils.FindClose(sr);
      end;

   if (PlusMois(Maintenant, -6) > DateCalcul) then
      DeleteFile(PChar(NomFic));
   end;

AssignFile(FRapport, NomFic);
if FileExists(NomFic) then
   begin
   Append(FRapport);
   Writeln(FRapport, '');
   end
else
   begin
   ReWrite (FRapport);
   Writeln (FRapport,
            'Attention, Le dernier calcul se trouve en fin du fichier');
   end;

Writeln(FRapport, '_____________________________________');
Writeln(FRapport, 'Début de calcul : ' + DateTimeToStr(Now));
}
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 07/09/2004
Modifié le ... :   /  /
Description .. : Calcul d'un élément
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOF_MUL_DADS2TIERS.Calcul_un (bRemplace : boolean);
var
St: string;
SQL : string;
{PT7-1
Maintenant: TDateTime;
}
begin
St:= TFmul(Ecran).Q.FindField('T_AUXILIAIRE').asstring;
if St <> '' then
   begin
   try
      begintrans;
      if bRemplace then
        ExecuteSQL ('DELETE FROM DADS2HONORAIRES WHERE'+
                  ' PDH_HONORAIRE="'+St+'" AND'+
                  ' PDH_VALIDITE="'+RechDom('PGANNEE', PGAnnee, FALSE)+'"');
{$IFDEF COMPTA}
      { FQ 20689 BVE 19.06.07 }
      SQL := 'SELECT PDH_HONORAIRE FROM DADS2HONORAIRES ' +
             'WHERE PDH_HONORAIRE = "'+St+'" AND ' +
             'PDH_VALIDITE = "'+RechDom('PGANNEE', PGAnnee, FALSE) + '"';
      if not(ExisteSQL(SQL)) then
      begin
         Insert := true;
{$ENDIF}
         ChargeTOBHonB (St);
         DeleteErreur (St, 'PDH');	//PT7-1
         CalculHonDADSB (St);
         LibereTOBHonB;
         CommitTrans;
{$IFDEF COMPTA}
      end;
{$ENDIF}
      { END FQ 20689 }
   except
      Rollback;
{PT7-1
      Maintenant:= Now;
      Writeln (FRapport, 'Tiers '+St+' : Calcul TD Bilatéral annulé : '+
               DateTimeToStr(Maintenant));
}               
      end;
   end;
MoveCur(False);
MoveCurProgressForm(St);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 07/09/2004
Modifié le ... :   /  /
Description .. : ChargeZones
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOF_MUL_DADS2TIERS.LibereTOBHonB;
begin
  FreeAndNil(THonoCompta);
  TDADSUD.SetAllModifie(TRUE);
  try
     TDADSUD.InsertDB(nil, FALSE);
  except
     on E: Exception do
     begin
        PgiError ('Erreur InsertDb : ' + E.Message) ;
     end ;
  end ;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 07/09/2004
Modifié le ... :   /  /
Description .. : ChargeZones
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOF_MUL_DADS2TIERS.ChargeTOBHonB(Honoraire: string);
var
OD, StHonoCompta: string;
QRechHonoCompta: TQuery;
begin
//PT4
if (GetCheckBoxState('OD') = cbChecked) then
   OD:= ' J_NATUREJAL="OD" OR '
else
   OD:= '';
//FIN PT4
//Chargement de la TOB HONOCOMPTA
{ FQ 20702 BVE 19.06.07 }
StHonoCompta:= 'SELECT E_ETABLISSEMENT, E_DEBIT, E_CREDIT, E_DATECOMPTABLE,'+
               ' E_LIBELLE, E_JOURNAL, E_REFINTERNE, ET_SIRET,'+
               ' ETB_DADSSECTION, ETB_TYPDADSSECT, T_AUXILIAIRE, T_SIRET,'+
               ' T_LIBELLE, T_PRENOM, T_ADRESSE1, T_ADRESSE2, T_ADRESSE3,'+
               ' T_CODEPOSTAL, T_VILLE, YTC_PROFESSION, YTC_REMUNERATION,'+
               ' YTC_INDEMNITE, YTC_AVANTAGE'+
               ' FROM ECRITURE'+
               ' LEFT JOIN GENERAUX ON'+
               ' E_GENERAL=G_GENERAL'+
               ' LEFT JOIN TIERS ON'+
               ' E_AUXILIAIRE=T_AUXILIAIRE'+
               ' LEFT JOIN TIERSCOMPL ON'+
               ' E_AUXILIAIRE=YTC_AUXILIAIRE'+
               ' LEFT JOIN JOURNAL ON'+
               ' E_JOURNAL=J_JOURNAL'+
               ' LEFT JOIN ETABLISS ON'+
               ' E_ETABLISSEMENT=ET_ETABLISSEMENT'+
               ' LEFT JOIN ETABCOMPL ON'+
               ' ETB_ETABLISSEMENT=ET_ETABLISSEMENT WHERE';
{ END FQ 20702 }
{$IFDEF COMPTA} //PT-8
  StHonoCompta := StHonoCompta + ' T_AUXILIAIRE="'+Honoraire+'" AND'+
               ' T_NATUREAUXI="FOU" AND'+
               ' ('+OD+'J_NATUREJAL="CAI" OR J_NATUREJAL="BQE") AND'+
               ' (E_NATUREPIECE="OD" OR E_NATUREPIECE="OF" OR E_NATUREPIECE="RF") AND'+
               ' E_DATECOMPTABLE>="'+USDateTime(StrToDate(LCptDu.Text))+'" AND'+
               ' E_DATECOMPTABLE<="'+USDateTime(StrToDate(LCptAu.Text))+'" AND'+
               ' E_QUALIFPIECE="N" AND'+
               ' G_NATUREGENE="COF" AND'+
               ' YTC_REMUNERATION<>""' +
               ' AND YTC_DAS2="X"';

{$ELSE}
  StHonoCompta := StHonoCompta + ' T_AUXILIAIRE="'+Honoraire+'" AND'+
               ' T_NATUREAUXI="FOU" AND'+
               ' ('+OD+'J_NATUREJAL="CAI" OR J_NATUREJAL="BQE") AND'+
               ' (E_NATUREPIECE="OD" OR E_NATUREPIECE="OF" OR E_NATUREPIECE="RF") AND'+
               ' E_DATECOMPTABLE>="'+UsDateTime(DebExer)+'" AND'+
               ' E_DATECOMPTABLE<="'+UsDateTime(FinExer)+'" AND'+
               ' E_QUALIFPIECE="N" AND'+
               ' G_NATUREGENE="COF" AND'+
               ' YTC_REMUNERATION<>""';
{$ENDIF}




QRechHonoCompta := OpenSql(StHonoCompta, TRUE);
THonoCompta:= TOB.Create('Les Honoraires', nil, -1);
THonoCompta.LoadDetailDB ('Honoraires comptables', '', '', QRechHonoCompta,
                          False);
Ferme(QRechHonoCompta);



//PT6
QRechHonoCompta := OpenSql(StHonoCompta, TRUE);
TListeHono.LoadDetailDB ('Honoraires comptables', '', '', QRechHonoCompta,
                         True);
Ferme(QRechHonoCompta);
//FIN PT6

TDADSUD:= TOB.Create('Mère DADSB Détail', nil, -1);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 07/09/2004
Modifié le ... :   /  /
Description .. : ChargeZones
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOF_MUL_DADS2TIERS.CalculHonDADSB(Honoraire: string);
var
TDADSUDetail, THonoComptaD: TOB;
Avantage, BufChamp, Indemnite, Remuneration: string;
Montant: double;
Controle : TControle;
begin
//PT7-2
Controle.Salarie:= Honoraire;
Controle.TypeD:= 'TDH';
Controle.Num:= 1;
Controle.DateDeb:= IDate1900;
Controle.DateFin:= IDate1900;
Controle.Exercice:= PGExercice;
//FIN PT7-2
Montant:= 0;

//Recherche du salarié concerné dans la TOB Salariés
THonoComptaD:= THonoCompta.FindFirst(['T_AUXILIAIRE'], [Honoraire], TRUE);

if THonoComptaD <> nil then
   begin
{Ecriture dans le fichier de contrôle du libellé, code postal ville du compte
tiers concerné}
{PT7-1
   Writeln(FRapport, '');
   Writeln(FRapport, '');
   Writeln (FRapport, 'Compte auxiliaire : '+Honoraire+' du tiers '+
            THonoComptaD.GetValue('T_LIBELLE')+' '+
            THonoComptaD.GetValue('T_CODEPOSTAL')+' '+
            THonoComptaD.GetValue('T_VILLE'));
}
   TDADSUDetail:= TOB.Create ('DADS2HONORAIRES', TDADSUD, -1);
   TDADSUDetail.PutValue ('PDH_HONORAIRE', Honoraire);
{$IFDEF COMPTA} //PT-8
   TDADSUDetail.PutValue ('PDH_VALIDITE', RechDom('PGANNEE', PGAnnee, FALSE));
{$ELSE}
   TDADSUDetail.PutValue ('PDH_VALIDITE', PGExercice);
{$ENDIF}
   TDADSUDetail.PutValue ('PDH_SIRET', THonoComptaD.GetValue('ET_SIRET'));
   TDADSUDetail.PutValue ('PDH_ETABLISSEMENT', THonoComptaD.GetValue('E_ETABLISSEMENT'));
   TDADSUDetail.PutValue ('PDH_SECTIONETAB',
                          THonoComptaD.GetValue('ETB_DADSSECTION'));
   TDADSUDetail.PutValue ('PDH_TYPEDADS',
                          THonoComptaD.GetValue('ETB_TYPDADSSECT'));

   TDADSUDetail.PutValue ('PDH_PROFESSIONBEN', THonoComptaD.GetValue('YTC_PROFESSION'));
   { BVE 12.09.07 : Transfert des données Tiers vers Dads2}
   TiersToDads2 (THonoComptaD, TDADSUDetail, Controle);

//PT1
   Remuneration:= THonoComptaD.GetValue('YTC_REMUNERATION');
   Indemnite:= THonoComptaD.GetValue('YTC_INDEMNITE');
   Avantage:= THonoComptaD.GetValue('YTC_AVANTAGE');
   while (THonoComptaD <> nil) do
         begin
         Montant:= Montant+THonoComptaD.GetValue('E_DEBIT') -
                   THonoComptaD.GetValue('E_CREDIT');
         THonoComptaD:= THonoCompta.FindNext (['T_AUXILIAIRE'],
                                              [Honoraire], TRUE);
         end;
   if (Montant > 0) then
      begin
      if (Remuneration = 'HON') then
         TDADSUDetail.PutValue('PDH_REMHONOR', Montant)
      else
      if (Remuneration = 'COM') then
         TDADSUDetail.PutValue('PDH_REMCOMMISS', Montant)
      else
      if (Remuneration = 'COU') then
         TDADSUDetail.PutValue('PDH_REMCOURTAGE', Montant)
      else
      if (Remuneration = 'RIS') then
         TDADSUDetail.PutValue('PDH_REMRISTOURNE', Montant)
      else
      if (Remuneration = 'JET') then
         TDADSUDetail.PutValue('PDH_REMJETON', Montant)
      else
      if (Remuneration = 'DRA') then
         TDADSUDetail.PutValue('PDH_REMAUTEUR', Montant)
      else
      if (Remuneration = 'DRI') then
         TDADSUDetail.PutValue('PDH_REMINVENT', Montant)
      else
      if (Remuneration = 'AUT') then
         TDADSUDetail.PutValue('PDH_REMAUTRE', Montant)
      else
      if (Remuneration = 'IND') then
         begin
         TDADSUDetail.PutValue('PDH_REMINDEMNITE', Montant);
         if (Indemnite = 'ALL') then
            BufChamp:= 'F  '
         else
         if (Indemnite = 'REM') then
            BufChamp:= ' R '
         else
         if (Indemnite = 'EMP') then
            BufChamp:= '  P';
         TDADSUDetail.PutValue('PDH_CHARGEINDEMN', BufChamp);
         end
      else
      if (Remuneration = 'AVA') then
         begin
         TDADSUDetail.PutValue('PDH_REMAVANTAGE', Montant);
         if (Avantage = 'NOU') then
            BufChamp:= 'N   '
         else
         if (Avantage = 'LOG') then
            BufChamp:= ' L  '
         else
         if (Avantage = 'VOI') then
            BufChamp:= '  V '
         else
         if (Avantage = 'AUT') then
            BufChamp:= '   A'
         else
            BufChamp:= '    ';
         TDADSUDetail.PutValue('PDH_AVANTAGENATN', BufChamp);
         if (Avantage = 'NTI') then
            BufChamp:= 'T'
         else
            BufChamp:= ' ';
         TDADSUDetail.PutValue('PDH_NTIC', BufChamp);
         end
      else
         if (Remuneration = 'RET') then
            TDADSUDetail.PutValue('PDH_RETENUESOURC', Montant);
      end;
   end;
end;

{$IFDEF COMPTA}      //PT-8
procedure TOF_MUL_DADS2TIERS.LaFormOnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  Case Key of
   VK_F9 :
      begin
        key := 0;
        TFMul(Ecran).BChercheClick(nil);
      end;

    VK_F10 :
      begin
        key := 0;
        CalculerClick(nil);
      end;
     else fOnSaveKeyDownEcran(Sender, Key, Shift );      
  end;
end;
{$ENDIF}



{ TOF_QU_DADS2TIERS }
{$IFDEF COMPTA}     //PT-8
procedure TOF_QU_DADS2TIERS.BValideOnClick(Sender: Tobject);
var
  ValRetour   : Integer;

begin
  ValRetour := 0;
  if QCalcul1.Checked then
    Inc(ValRetour, 1);
  if QCalcul2.Checked then
    Inc(ValRetour, 2);

  TFVierge(Ecran).Retour  := IntToStr(ValRetour);

end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 19/06/2007
Modifié le ... :   /  /
Description .. : Permet de griser la coche du bas si celle du haut est 
Suite ........ : cochée.
Mots clefs ... : 
*****************************************************************}
procedure TOF_QU_DADS2TIERS.CaseChecked (Sender : TObject);
begin
   if THCheckBox(Sender).Checked then
      QCalcul2.Enabled := false
   else
      QCalcul2.Enabled := true;
end;

procedure TOF_QU_DADS2TIERS.OnArgument(S: string);
var
  LeTitre : string;
  Lannee  : String;

begin
  inherited;
  LeTitre := READTOKENST(S);
  Lannee  := READTOKENST(S);
  TForm(Ecran).Caption := LeTitre;
  UpdateCaption(Tform(Ecran));
  QCalcul1  := THCheckbox(GetControl('QCALCUL1'));
  QCalcul2  := THCheckbox(GetControl('QCALCUL2'));
  BValide   := TToolbarButton97(Getcontrol('BVALIDER'));
  BValide.OnClick := BValideOnClick;
  QCalcul1.Caption  := QCalcul1.Caption + ' ' + Lannee;
  QCalcul2.Checked  := True;
  QCalcul1.OnClick  := CaseChecked;
end;

procedure TOF_QU_DADS2TIERS.OnClose;
begin
  inherited;

end;
{$ENDIF}

initialization
registerclasses([TOF_MUL_DADS2TIERS]);
{$IFDEF COMPTA}
registerclasses([TOF_QU_DADS2TIERS]);
{$ENDIF}
end.

