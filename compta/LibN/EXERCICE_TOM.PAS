{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 06/03/2002
Modifié le ... : 16/05/2003
Description .. : Source TOM de la TABLE : EXERCICE (EXERCICE)
Suite ........ : - CA - 22/04/2003 - FQ 12259 - Contrôle durée d'exercice
Suite ........ : - CA - 16/05/2003 - FQ 11686 - Ecritures sur le dernier jour
Suite ........ : de l'exercice à ne pas prendre en compte.
Suite ........ : - GCO - 28/06/2004 - Ajout de la modification de la date de
Suite ........ : fin sur l'exerice en cours
Suite ........ : - GCO - 05/07/2004 - Prise en compte du E_NUMGRPECR
Mots clefs ... : TOM;EXERCICE
*****************************************************************}
unit EXERCICE_TOM;

interface

uses Windows, SysUtils, Controls, Classes, Forms, HTB97,
{$IFDEF EAGLCLIENT}
  eFiche, MaineAGL, eMul, eFichList, eTablette, eFichGrid, HPdfPrev,
{$ELSE}
  db,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Fiche, Fe_main, Mul, FichList, Tablette, FichGrid, EdtEtat,
{$ENDIF}
{$IFDEF AMORTISSEMENT}
  ImPlan,
  ImEnt,
{$ENDIF}
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  UtilPgi, Graphics, HEnt1, HCtrls, HMsgBox, UTOM, UTob, Ent1, ParamSoc, EntGC;

procedure YYLanceFiche_Exercice(Qui: string);
//function ControleDureeExercice(DateDebut, DateFin: TDateTime): boolean;

type
  TOM_EXERCICE = class(TOM)
    procedure OnDeleteRecord ; override;
    procedure OnNewRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  private
    Qui: Integer; { 0 : Menu , 1 : depuis assistant création société }
    NombreExo: Byte;
    Comment: TActionFiche;
    gszDateFin: string;

    // GCO - 29/06/2004 - Changement de Code de Fin d'Exercice
    FExerciceDepart  : string;    // Code Exercice de départ
    FExerciceArrivee : string;    // Code Exercice à mettre dans les écritures
    FNouvelleDate    : TDateTime; // Nouvelle Date de Fin d'exercice
    FDateSuperieure  : TDateTime; // Date de test opérateur Supérieur
    FDateInferieure  : TDateTime; // Date de test opérateur Inférieur ou =
    FbDateDebutModifie : boolean; // Flag indiquant si la date de début a été modifiée
    {$IFDEF AMORTISSEMENT}
    FMajImmo         : boolean;   // Flag de recalcul des immos
    {$ENDIF}
    FOkModifDate     : boolean;
    function CombienExo: Integer;
    function EcrExist(Exercice: String3; Table: string): Boolean;
    function AutoriseNouvelExo(Exercice: String3): Boolean;
    function CtrlDebFinExo      : Boolean;
    function CtrlDureeMinMaxExo : Boolean;
    function CtrlEuroDeb        : Boolean;
    function AutoriseNonOuvert  : Boolean;
    function LastExo(var Exo: TExoDate): string;
    function TabVarexist (Exercice : String) : Boolean;         //test présence tableau variation sur l'exercice
    function TabNoteTavailexist (Exercice : String) : Boolean;  //test la présence d'une note de travail sur l'exercice

    // GCO - 29/06/2004 - Changement de Code de Fin d'Exercice
    procedure MAJExercice;
    procedure MAJAnalytiquePur;
    procedure MAJBalanceSituation;
    procedure MAJEcriture;

    // Mise à jour du pièce ainsi que l'analytique concerné
    procedure MAJModePiece;

    procedure MAJAutreTable( vStNomTable : string ; vBoavecDateModif : Boolean );
{$IFDEF AMORTISSEMENT}
    procedure MajAmortissement;
{$ENDIF}
    function  ExisteEcriturePosterieure : Boolean;
    function  ExisteBalanceSituation    : Boolean;
    function  ExisteAnalytique          : Boolean;
    function  ExisteEcritureTreso       : Boolean;

    function  VerifIntegriteBordereau   : Boolean;
    procedure ChangementFinExercice;

  end;

  // libellés des messages
const
  HMessage: array[1..8] of string = (
    {1}'0;Exercices;La date de fin d''exercice saisie est incohérente.;W;O;O;O;',
    {2}'1;Exercices;Vous devez indiquer une durée d''exercice comprise entre 1 et 23 mois.;W;O;O;O;',
    {3}'2;Exercices;Des écritures ont déjà été saisies au delà de cette date.;W;O;O;O;',
    {4}'3;Exercices;La date de fin d''exercice doit être une fin de mois.;W;O;O;O;',
    {5}'4;Exercices;Etat "Non Ouvert" inaccessible : cet exercice fait référence à des mouvements.;W;O;O;O;',
    {6}'5;Exercices;La date de début de l''exercice doit être un début de mois.;W;O;O;O;',
    {7}'6;Exercices;Suppression imposible : cet exercice fait référence à des mouvements.;W;O;O;O;',
    {8}'7;Exercices;La date de début de l''exercice doit être supérieure à celle d''entrée en vigueur de l''Euro.;W;O;O;O;');

const
  Confirm: array[1..5] of string = (
    {1}'6;Définition d''un nouvel exercice.;',
    {2}'ATTENTION : Les dates saisies sont les suivantes :',
    {3}'Début d''exercice :',
    {4}'Fin   d''exercice :',
    {5}'Confirmez-vous ?;E;YN;Y;N');

implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  {$ENDIF MODENT1}
  uLibExercice,   // CRelatifVersExercice, CControleDureeExercice
  SoldeCpt,       // MajTotTousComptes
  uLibEcriture,   // CNumeroPiece
  ED_TOOLS,       // InitMoveProgressForm
  uLibAnalytique; // CChargeAna

{=======================================================================================}

procedure YYLanceFiche_Exercice(Qui: string);
begin
  // GCO 14/10/2003 FB 12637
  if _Blocage(['nrCloture', 'nrBatch', 'nrSaisieCreat', 'nrSaisieModif',
      'nrPointage', 'nrLettrage', 'nrRelance', 'nrEnca', ',nrDeca'], True,
      'nrCloture') then
  begin
    Exit;

  end
  else
    AGLLanceFiche('CP', 'CPEXERCICE', '', '', Qui);
end;

{=======================================================================================}
// Lors de la suppression d'un exercice

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2004
Modifié le ... :   /  /
Description .. : Ajout de test avant suppression !
Mots clefs ... :
*****************************************************************}
procedure TOM_EXERCICE.OnDeleteRecord;
begin
    inherited;

    LastError := 1;

    if ((EcrExist(GetControlText('EX_EXERCICE'),'ECRITURE')) or (EcrExist(GetControlText('EX_EXERCICE'),'BUDECR'))) then
    begin
        PGIBox ('Présence d''écritures sur cet exercice. Suppression impossible',Ecran.Caption);
        exit;
    end;

    //controle présence tableau des variations
    if TabVarexist(GetControlText('EX_EXERCICE')) then
    begin
      PGIBox('Présence de tableau des variations sur cet exercice. Suppression impossible', Ecran.Caption);
      exit;
    end;

    if TabNoteTavailexist(GetControlText('EX_EXERCICE')) then
    begin
      PGIBox('Présence d''une note de travail sur cet exercice. Suppression impossible', Ecran.Caption);
      exit;
    end;

    LastError := 0;
end;



{=======================================================================================}
// Lors de la création d'un nouvel exercice

procedure TOM_EXERCICE.OnNewRecord;
var
  NumExo: integer;
  DatDebut: TDateTime;
  DatFin: TDateTime;
  a, m, j: Word;
  Exo: tExoDate;
begin
  inherited;
  // Affecte les valeurs par défaut
  SetField('EX_ETATCPTA', 'NON');
  SetField('EX_ETATADV', 'NON');
  SetField('EX_ETATAPPRO', 'NON');
  SetField('EX_ETATPROD', 'NON');
  SetField('EX_ETATBUDGET', 'OUV');

  // Verrouille le bouton Nouveau
  SetControlEnabled('BInsert', False);

  case NombreExo of
    0:
      begin // Création du premier exercice
        SetControlText('EX_EXERCICE', FormatFloat('000', 001));
        SetField('EX_EXERCICE', FormatFloat('000', 001));
        SetControlEnabled('EX_LIBELLE', True);
        SetControlEnabled('EX_ABREGE', True);
        SetControlEnabled('EX_DATEDEBUT', True);
        // Si lancé par l'assistant société
        if Qui = 1 then
          SetField('EX_ETATCPTA', 'OUV');
      end
  else
    begin
      LastExo(Exo);
      if Exo.Code <> '' then
      begin
        DatDebut := Exo.Fin + 1;
        DatFin := Exo.Fin;
        DecodeDate(DatFin, a, m, j);
        Inc(a);
        DatFin := EncodeDate(a, m, j);
        NumExo := StrToInt(Exo.Code);
        Inc(NumExo);
        SetControlText('EX_EXERCICE', FormatFloat('000', NumExo));
        SetControlEnabled('EX_DATEDEBUT', False);
        SetControlText('EX_DATEDEBUT', DateToStr(DatDebut));
        SetControlText('EX_DATEFIN', DateToStr(DatFin));
        SetField('EX_EXERCICE', FormatFloat('000', NumExo));
        SetField('EX_DATEDEBUT', DatDebut);
        SetField('EX_DATEFIN', DatFin);
      end;
    end;
  end;

  SetFocusControl('EX_LIBELLE');
end;

{=======================================================================================}
// Lors de la validation et avant celle-ci

procedure TOM_EXERCICE.OnUpdateRecord;
var Reponse: Word;
    lIoResult : TIoErr;
begin
  inherited;
  {$IFDEF AMORTISSEMENT}
  FMajImmo := False;
  {$ENDIF}
  FOkModifDate := False;
  // Si lancé par l'assistant société
  if (Qui = 1) then
  begin
    // Demande confirmation
    HMessage[6] := Confirm[1] + Confirm[2] + chr(13) + chr(10) + chr(13) +
      chr(10) +
      Confirm[3] + GetControlText('EX_DATEDEBUT') + chr(13) + chr(10) +
      Confirm[4] + GetControlText('EX_DATEFIN') + chr(13) + chr(10) + chr(13) +
        chr(10) +
      Confirm[5];
    Reponse := HShowMessage(HMessage[6], '', '');
    // Si on ne confirme pas : Sort
    case Reponse of
      mrYes: ;
      mrNo, mrCancel: begin
                        Lasterror := -1;
                        Exit;
                      end;
    end;
  end;

  if (DS.state in [dsEdit, dsInsert]) then
  begin
    if StrToDate(GetControlText('EX_DATEFIN')) < StrToDate(GetControlText('EX_DATEDEBUT')) then
    begin
      PgiInfo('La date de fin d''exercice doit être supérieure ou égale à la date de début.', Ecran.Caption);
      SetFocusControl('EX_DATEDEFIN');
      Lasterror := -1;
      Exit;
    end;
    
    // La date de début de l'exercice doit être supérieure à celle d'entrée en vigueur de l'Euro
    if not CtrlEuroDeb then
    begin
      HShowMessage(HMessage[8], '', '');
      SetFocusControl('EX_DATEDEBUT');
      Lasterror := -1;
      Exit;
    end;

    // La date de fin d'exercice saisie est incohérente
    if not CtrlDebFinExo then
    begin
      HShowMessage(HMessage[1], '', '');
      SetFocusControl('EX_DATEFIN');
      Lasterror := -1;
      Exit;
    end;

    // Vous devez indiquer une durée d'exercice comprise entre 1 et 23 mois
    if not CtrlDureeMinMaxExo then
    begin
      PGIInfo('La durée d''exercice ne doit pas excéder 24 mois.', Ecran.Caption);
      SetFocusControl('EX_DATEFIN');
      LastError := -1;
      Exit;
    end;

    // GCO - 15/05/2006 - FQ 13795
    if (VH^.Precedent.Code = '') and (GetControlEnabled('EX_DATEDEBUT')) and
       (VH^.Encours.Deb <> GetField('EX_DATEDEBUT')) then
    begin
      if ExisteSQL('SELECT E_EXERCICE FROM ECRITURE WHERE E_DATECOMPTABLE < "' +
                   UsDateTime(GetField('EX_DATEDEBUT')) + '" ORDER BY E_EXERCICE') then
      begin
        PgiInfo('Traitement impossible, des écritures ont été saisies avant le ' +
               GetControlText('EX_DATEDEBUT') + '.', Ecran.Caption);
        LastError := -1;
        Exit;
      end;
    end;

    // Des écritures comptables ou budgétaires existent sur la période de l'exercice
    if AutoriseNonOuvert then
    begin
      Lasterror := -1;
      Exit;
    end;

{$IFDEF AMORTISSEMENT}
    // GCO - 22/06/2006 - FQ 18469
    if ExisteSql('SELECT I_IMMO FROM IMMO') then
    begin
      if not((GetParamSocSecur('SO_EXOCLOIMMO','') = '') or
             (GetParamSocSecur('SO_EXOCLOIMMO','') = VHImmo^.Precedent.Code)) then
      begin
        PGIInfo('Traitement impossible, l''exercice comptable en cours ne ' +
                'correspond pas à celui des immobilisations. #10#13' +
                'Veuillez effectuer la clôture comptable ou annuler la ' +
                'clôture des immobilisations.',Ecran.Caption);
        LastError := -1;
        Exit;
      end;
    end;
    // FIN GCO
{$ENDIF}

    // Ajout du changement de date de fin d'exercice pour l'exercice en cours
    if ((GetField('EX_EXERCICE') = VH^.Encours.Code) and (VH^.Encours.Fin <> StrToDate(GetControlText('EX_DATEFIN')))) or
       ((GetField('EX_EXERCICE') = VH^.Suivant.Code) and (VH^.Suivant.Fin <> StrToDate(GetControlText('EX_DATEFIN')))) or
       ((GetField('EX_EXERCICE') = VH^.EnCours.Code) and (VH^.EnCours.Deb <> StrToDate(GetControlText('EX_DATEDEBUT')))) then
    begin
      if PgiAsk('Attention, vous venez de modifier les dates d''exercice. ' + #13 + #10 +
                'Différents traitements vont être effectués, voulez-vous continuer ?') = MrYes then
      begin
        // Init des variables necessaires au traitement
        FNouvelleDate := StrToDate(GetControlText('EX_DATEFIN'));
        FbDateDebutModifie := (GetField('EX_EXERCICE') = VH^.EnCours.Code) and (VH^.EnCours.Deb <> StrToDate(GetControlText('EX_DATEDEBUT')));

        // Test si la Date de Cloture Périodique n'est pas supérieure à la nouvelle Date
        if FNouvelleDate < GetParamSocSecur('SO_DATECLOTUREPER',iDate1900) then
        begin
          PgiInfo('Traitement impossible, la nouvelle date de fin doit être ' + #13 + #10 +
                  'supérieure ou égale à la date de cloture périodique', Ecran.Caption);
          LastError := -1;
          Exit;
        end;

        // Interdiction de modifier la date si on trouve un exercice non ouvert
        if ExisteSQL('SELECT EX_EXERCICE FROM EXERCICE WHERE EX_ETATCPTA = "NON" ORDER BY EX_EXERCICE') then
        begin
          PgiInfo('Traitement impossible, il existe un exercice non ouvert.', Ecran.Caption);
          LastError := -1;
          Exit;
        end;

        // Test de la postérieurité des écritures
        if ((GetField('EX_EXERCICE') = VH^.Encours.Code) and (VH^.Suivant.Code = '') and
            (VH^.Encours.Fin > StrToDate(GetControlText('EX_DATEFIN')))) OR
           ((GetField('EX_EXERCICE') = VH^.Suivant.Code) and
            (VH^.Suivant.Fin > StrToDate(GetControlText('EX_DATEFIN')))) then
        begin
          if (ExisteEcriturePosterieure) or (ExisteBalanceSituation) or
             (ExisteAnalytique) or (ExisteEcritureTreso) then
          begin
            PgiInfo('Traitement impossible, vous devez créer l''exercice suivant ' +
                    'avant de faire ce traitement.', Ecran.Caption);
            LastError := -1;
            Exit;
          end;
        end;

        // GCO - 11/10/2007 - FQ 21597
        // Controlle de la durée de l'exercice suivant suite au changement de
        // date de fin de l'exercice en cours
        if (GetField('EX_EXERCICE') = VH^.Encours.Code) and (VH^.Suivant.Code <> '') and
           (not CControleDureeExercice( GetField('EX_DATEFIN')+1 , VH^.Suivant.Fin)) then
        begin
          PgiInfo('Traitement impossible, La durée de l''exercice suivant ne doit pas excéder 24 mois.', Ecran.Caption);
          LastError := -1;
          Exit;
        end;

        if (GetField('EX_EXERCICE') = VH^.EnCours.Code) and (VH^.Suivant.Code <> '') then
        begin
          // Init des variables necessaires au Traitement
          if FNouvelleDate > VH^.Encours.Fin then
          begin
            FExerciceDepart  := VH^.Suivant.Code;
            FExerciceArrivee := VH^.Encours.Code;
            FDateSuperieure  := VH^.Encours.Fin;
            FDateInferieure  := FNouvelleDate;
          end
          else
          begin
            FExerciceDepart  := VH^.EnCours.Code;
            FExerciceArrivee := VH^.Suivant.Code;
            FDateSuperieure  := FNouvelleDate;
            FDateInferieure  := VH^.Encours.Fin;
          end;

          // Vérification de l'equilibre des bordereaux
          if GetField('EX_EXERCICE') = VH^.Encours.Code then
          begin
            if not VerifIntegriteBordereau then
            begin
              PgiInfo('Traitement impossible, un ou plusieurs bordereaux se ' +
                      'retrouveraient non équilibrés.', Ecran.Caption);
              LastError := -1;
              Exit;
            end;
          end;
        end;

        // Vérification de l'absence d'opérations sur les immobilisations sur l'exercice suivant
        if ExisteSQL ('SELECT * FROM IMMOLOG WHERE IL_DATEOP>"'+USDateTime(StrToDate(GetControlText('EX_DATEFIN')))+'" AND IL_TYPEOP<>"ACQ"') then
        begin
          PgiInfo('Traitement impossible, des opérations sur les immobilisations '+
                  'ont été pratiquées sur l''exercice suivant.', Ecran.Caption);
          LastError := -1;
          Exit;
        end;

        if BlocageMonoPoste(True) then
        begin
          lIoResult := Transactions(ChangementFinExercice,0);
          DeBlocageMonoPoste(True);

          if lIoResult <> oeOk then
          begin
            LastError := -1;
            PGIInfo('Traitement annulé. Une erreur est survenue lors de la mise ' + #13 + #10 +
                    'à jour des données.', Ecran.caption);
          end
          else
          begin
            // Message de Fin de traitement
            FOkModifDate := True;
            {$IFDEF AMORTISSEMENT}
            FMajImmo := True;
            {$ENDIF}
          end;
        end;
      end
      else
      begin
        LastError := -1;
        Exit;
      end;
    end;
  end;
end;

{=======================================================================================}

procedure TOM_EXERCICE.OnAfterUpdateRecord;
begin
  inherited;
  AvertirMultiTable('TTEXERCICE');
{$IFDEF EAGLCLIENT}
  // GCO - 25/06/2007 - FQ 20829
  AvertirCacheServer('EXERCICE');
{$ENDIF}
  ChargeMagExo(False); // Chargement des ParamSOC
  {$IFDEF AMORTISSEMENT}
  if FMajImmo then
  begin
    InitMoveProgressForm(Ecran, 'Changement de date de fin d''exercice', '', 1, True, True);
    MoveCurProgressForm('Recalcul des immobilisations...');
    MajAmortissement;
    FiniMoveProgressForm;
  end;
  {$ENDIF}

  // Déverrouille le bouton de suppression si l'exercice est non ouvert
  { FQ 18348 BVE 12.04.07
  SetControlEnabled('BDelete', GetField('EX_ETATCPTA') = 'NON');
    END FQ 18348 }
  SetControlVisible('BDelete', GetField('EX_ETATCPTA') = 'NON');

  // Récupère le nombre d'exercice
  NombreExo := CombienExo;

  // Affiche le combo Etat Comptabilité
  SetControlVisible('TEX_EtatCpta', True);
  SetControlVisible('EX_EtatCpta2', False);

  if FOkModifDate then
  begin
    PGIInfo('Traitement terminé. Aucune Anomalie détectée.' + #13 + #10 +
            'Cependant les filtres associés aux anciennes dates d''exercice' + #13 + #10 +
            'sont à revoir. Suite au traitement, cet écran va se fermer.', Ecran.Caption);

    // GCO - 20/07/2007 - FQ 20324
    TToolBarbutton97(GetControl('BFERME', True)).Click;
  end
  else
  begin
    RefreshDB;
  end;

end;

{=======================================================================================}
// Au chargement des données

procedure TOM_EXERCICE.OnLoadRecord;
{$IFDEF CCS3}
var
  cControle: TControl;
{$ENDIF}
 // bEnabled: boolean;
begin
  inherited;
  // Si l'exercice est clôturé définitivement : Verrouille le contrôle
  if GetField('EX_ETATCPTA') <> 'NON' then
  begin
    SetControlEnabled('EX_ETATCPTA', False);
  end;

  SetControlEnabled('EX_DATEFIN', True);

  // Impossible de modifier la date de fin si l'exercice = SO_EXOCLOIMMO
  if (GetParamSocSecur('SO_EXOCLOIMMO','') = GetField('EX_EXERCICE')) or
     (GetField('EX_ETATCPTA') <> 'OUV') then
  begin
    SetControlEnabled('EX_DATEFIN', False);
  end;

  // GCO - 15/05/2006 - FQ 13795
  SetControlEnabled('EX_DATEDEBUT', False);
  if GetField('EX_ETATCPTA') = 'OUV' then
  begin // Autorisation de modification  de la date début
    if (GetField('EX_EXERCICE') = VH^.EnCours.Code) and
       (VH^.Precedent.Code = '') and
       (not ExisteSQL('SELECT E_JOURNAL FROM ECRITURE WHERE ' +
        'E_EXERCICE = "' + VH^.EnCours.Code + '" AND ' +
        '(E_ECRANOUVEAU = "H" OR E_ECRANOUVEAU = "OAN") ORDER BY E_JOURNAL')) then
      SetControlEnabled('EX_DATEDEBUT', True);
  end;

  // Modifie l'état du bouton d'insertion
  if (NombreExo > 1) then
    SetControlEnabled('BInsert', AutoriseNouvelExo(VH.Suivant.Code));

  // Si la feuille est lancé depuis l'assistant création société et qu'il y a au moins 1 exercice
  // Alors on ne peut pas en créer un (parce qu'on est déjà en mode création)
  if (Qui = 1) and (NombreExo > 0) then
    SetControlEnabled('BInsert', False);

  // Déverrouille le bouton de suppression si l'exercice est non ouvert
  { FQ 18348 BVE 12.04.07
  SetControlEnabled('BDelete', GetField('EX_ETATCPTA') = 'NON');
    END FQ 18348 }
  SetControlVisible('BDelete', GetField('EX_ETATCPTA') = 'NON');

  // Si aucun exercice
  if (NombreExo = 0) then
  begin
    // Affiche le combo Etat Comptabilité
    SetControlVisible('TEX_EtatCpta', False);
    SetControlVisible('EX_EtatCpta2', True);
    // Verrouille les autres contrôles d'affichage d'état
    SetControlEnabled('EEX_ETATBUDGET', False);
  end;

  // GCO - 15/05/2006 - Case obsolète semble t il !!
  (*
  case NombreExo of
    0:
      begin // Aucun exercice
        SetControlEnabled('EX_DATEDEBUT', True);
        SetControlProperty('EX_DATEDEBUT', 'Color', clWindow);
      end;

    1:
      begin // Un seul exercice
        bEnabled := not EcrExist('001', 'ECRITURE');
        SetControlEnabled('EX_DATEDEBUT', bEnabled);
        if (bEnabled = True) then
          SetControlProperty('EX_DATEDEBUT', 'Color', clWindow)
        else
          SetControlProperty('EX_DATEDEBUT', 'Color', clBtnFace);
      end;
  end;*)

  // Si lancé par l'assistant société : On peut créer un exercice uniquement s'il n'y en a pas
  if Qui = 1 then
    SetControlEnabled('BInsert', (NombreExo = 0));

  // Si lancé par l'assistant société et aucun exercice
  if (Qui = 1) and (NombreExo = 0) then
  begin
    // L'exercice est forcément ouvert
    SetField('EX_ETATCPTA', 'OUV');
  end;

  gszDateFin := GetField('EX_DATEFIN');

{$IFDEF CCS3}
  // Si S3 : Pas de budget visible
  if Ecran <> nil then
  begin
    cControle := GetControl('TEX_ETATBUDGET');
    if (cControle <> nil) then
      cControle.Visible := False;
    cControle := GetControl('TEX_ETATBUDGET2');
    if (cControle <> nil) then
      cControle.Visible := False;
  end;
{$ENDIF}
end;

{=======================================================================================}

procedure TOM_EXERCICE.OnArgument(S: string);
begin
  inherited;
  // FQ 11737
  SetControlVisible('BINSERT', False);
  
  // GCO - 14/10/2003 Code de blocage reporté dans le YYlanceFiche
  Comment := TFFicheListe(Ecran).TypeAction;
  Qui := ReadTokenI(S);
  // Récupère le nombre d'exercice
  NombreExo := CombienExo;
end;

{=======================================================================================}

procedure TOM_EXERCICE.OnClose;
begin
  inherited;
  //Si différent du mode consultation
  if Comment <> taConsult then
  begin
    // Déverrouille la fonction
    _Bloqueur('nrCloture', False);
  end;

  // Recharge les données des exercices
  ChargemagExo(False);
end;

{=======================================================================================}
// Renvoi le nombre d'exercice

function TOM_EXERCICE.CombienExo: Integer;
var
  Q: TQuery;
begin
  Q := OpenSQL('SELECT COUNT(EX_EXERCICE) FROM EXERCICE', TRUE);
  Result := Q.Fields[0].AsInteger;
  Ferme(Q);
end;

{=======================================================================================}
// Regarde s'il y a des écritures pour l'exercice passé en paramètres

function TOM_EXERCICE.EcrExist(Exercice: String3; Table: string): Boolean;
var
  Pref: string;
begin
  Result := False;
  if (Exercice = '') then
    Exit;
  Pref := 'E_';
  if Table = 'BUDECR' then
    Pref := 'BE_';
  Result := ExisteSQL('Select EX_EXERCICE From EXERCICE Where ' +
            'Exists(Select ' + Pref + 'EXERCICE From ' + Table + ' Where ' + Pref +
            'EXERCICE="' + Exercice + '") AND EX_EXERCICE="' + Exercice + '"');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/06/2004
Modifié le ... :   /  /
Description .. : Renvoi True si l'état de l'exercice suivant est égal à ouvert
Mots clefs ... : 
*****************************************************************}
function TOM_EXERCICE.AutoriseNouvelExo(Exercice: String3): Boolean;
var
  Q: TQuery;
  St: string;
begin
  Result := False;
  if Exercice = '' then
    Exit;
  St := '';
  Q := OpenSql('Select EX_ETATCPTA From EXERCICE Where EX_EXERCICE="' + Exercice
    + '"', True);
  if not Q.Eof then
    St := Q.Fields[0].AsString;
  Ferme(Q);
  if St <> 'OUV' then
    Exit;
  Result := True;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... :
Description .. : La date de fin d'exercice saisie est incohérente (> à celle de début)
Mots clefs ... :
*****************************************************************}
function TOM_EXERCICE.CtrlDebFinExo: Boolean;
begin
  Result := FinDeMois(GetField('EX_DATEFIN')) >= DebutDeMois(GetField('EX_DATEDEBUT'));
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... :
Description .. : Retourne True si la durée de l'exercice est inférieure à 24 mois
Mots clefs ... : 
*****************************************************************}
function TOM_EXERCICE.CtrlDureeMinMaxExo: Boolean;
var DateExo: TExoDate;
begin
  DateExo.Deb := DebutDeMois(GetField('EX_DATEDEBUT'));
  DateExo.Fin := FinDeMois(GetField('EX_DATEFIN'));
  Result := CControleDureeExercice(DateExo.Deb, DateExo.Fin);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... :
Description .. : Lors de la création du premier exercice et si on tient la
Suite ........ : comptabilité en Euro,  la date de début de l'exercice doit
Suite ........ : être supérieure à celle d'entrée en vigueur de l'Euro
Mots clefs ... :
*****************************************************************}
function TOM_EXERCICE.CtrlEuroDeb: Boolean;
begin
  Result := True;
  if ((DS.state in [dsInsert]) and (VH^.TenueEuro) and (NombreExo <= 0)) then
    Result := (GetField('EX_DATEDEBUT') >= V_PGI.DateDebutEuro);
end;

{=======================================================================================}
// Retourne True s'il y a des écritures comptables ou budgétaires

function TOM_EXERCICE.AutoriseNonOuvert: Boolean;
var
  EcrCpta: Boolean;
  //  EcrBud : Boolean ;
begin
  Result := False;
  // Vérifie s'il y a des écritures comptables et budgétaires
  EcrCpta := EcrExist(GetField('EX_EXERCICE'), 'ECRITURE');
  //	EcrBud:=EcrExist(GetField('EX_EXERCICE'),'BUDECR') ;
  if EcrCpta then
  begin
    if GetField('EX_ETATCPTA') = 'NON' then
    begin
      // Etat "Non Ouvert" inaccessible : cet exercice fait référence à des mouvements
      HShowMessage(HMessage[5], '', '');
      Exit;
    end;
  end;
  //	if EcrBud then Result:=True ;
end;

{=======================================================================================}
// Récupère le n°, la date de début et de fin du dernier exercice

function TOM_EXERCICE.LastExo(var Exo: tExoDate): string;
var
  Q: TQuery;
begin
  Fillchar(Exo, SizeOf(Exo), #0);
  Q :=
    OpenSQL('SELECT EX_EXERCICE,EX_DATEDEBUT,EX_DATEFIN FROM EXERCICE ORDER BY EX_DATEDEBUT DESC', TRUE);
  if not Q.Eof then
  begin
    Exo.Code := Q.FindField('EX_EXERCICE').AsString;
    Exo.Deb := Q.FindField('EX_DATEDEBUT').AsDateTime;
    Exo.Fin := Q.FindField('EX_DATEFIN').AsDateTime;
  end;
  Ferme(Q);
end;

function ControleDureeExercice(DateDebut, DateFin: TDateTime): boolean;
var
  AnneeDebut, AnneeFin: Word;
  MoisDebut, MoisFin, Jour: Word;
begin
  DecodeDate(DateDebut, AnneeDebut, MoisDebut, Jour);
  DecodeDate(DateFin, AnneeFin, MoisFin, Jour);
  Result := ((AnneeDebut = AnneeFin) and (MoisFin > MoisDebut)) or
    (AnneeFin = AnneeDebut + 1);
  if not Result then
    PGIInfo('La durée d''un exercice doit être comprise entre 1 et 24 mois,#10#13et ne peut s''étendre au delà de 2 années civiles.', 'Durée d''exercice');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/06/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_EXERCICE.ExisteEcriturePosterieure : Boolean;
begin
  Result := ExecuteSQL('SELECT E_JOURNAL FROM ECRITURE WHERE E_DATECOMPTABLE > "' +
            UsDateTime( FNouvelleDate ) + '" AND ' +
            'E_EXERCICE = "' + GetField('EX_EXERCICE') + '" ORDER BY E_JOURNAL') <> 0 ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/01/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_EXERCICE.ExisteBalanceSituation : Boolean;
begin
  Result := ExisteSQL('SELECT BSI_CODEBAL FROM CBALSIT WHERE ' +
                      'BSI_DATE1 < "' + UsDateTime( FNouvelleDate ) + '" AND ' +
                      'BSI_DATE2 > "' + UsDateTime( FNouvelleDate ) + '" AND ' +
                      'BSI_EXERCICE = "' + GetField('EX_EXERCICE') + '"');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/01/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_EXERCICE.ExisteEcritureTreso : Boolean;
begin
  Result := ExisteSql('SELECT CRL_COMPTEUR,CRL_NUMLIGNE FROM CRELBQE WHERE ' +
                      'CRL_DATECOMPTABLE > "' + UsDateTime( FNouvelleDate )+ '"');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/01/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_EXERCICE.ExisteAnalytique : Boolean;
begin
  Result := ExisteSql('SELECT Y_JOURNAL FROM ANALYTIQ WHERE ' +
                      'Y_DATECOMPTABLE > "' + UsDateTime( FNouvelleDate ) + '" AND ' +
                      'Y_EXERCICE = "' + GetField('EX_EXERCICE') + '"');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/06/2004
Modifié le ... : 28/09/2004
Description .. : Vérifie que l'équilibre Folio d'un bordereau de type libre
Suite ........ : n'est pas perdu lors du changement du changement de la Date
Mots clefs ... :
*****************************************************************}
function TOM_EXERCICE.VerifIntegriteBordereau: Boolean;
var lSt    : string;
    lQuery : TQuery;
begin
  Result := True;
  lSt := 'SELECT SUM(E_DEBIT) - SUM(E_CREDIT) TOTAL FROM ECRITURE WHERE '+
         '(E_MODESAISIE = "LIB" OR E_MODESAISIE = "BOR") '+
         'AND E_EXERCICE = "' + GetField('EX_EXERCICE') + '" '+
         'AND E_DATECOMPTABLE > "' + UsDateTime( FDateSuperieure ) + '" ' +
         'AND E_DATECOMPTABLE <= "' + UsDateTime( FDateInferieure ) + '" ' +
         'GROUP BY E_EXERCICE, E_JOURNAL, E_QUALIFPIECE, E_PERIODE, E_NUMEROPIECE';
  try
    lQuery := OpenSQL( lSt, True);
    if not lQuery.Eof then
    begin
      while not lQuery.Eof do
      begin
        if lQuery.FindField('TOTAL').AsFloat <> 0 then
        begin
          Result := False;
          Exit;
        end;
        lQuery.Next;
      end;
    end;
  finally
    Ferme( lQuery );
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/06/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_EXERCICE.MAJExercice;
begin
  try
    MoveCurProgressForm('Mise à jour de la table : EXERCICE');

    if (GetField('EX_EXERCICE') = VH^.EnCours.Code) and (VH^.Suivant.Code <> '') then
    begin
      if FNouvelleDate >= VH^.Suivant.Fin then
      begin
        // On supprime l'exercice suivant car il n'a plus lieu d'être
        ExecuteSQL('DELETE FROM EXERCICE WHERE EX_EXERCICE = "' + VH^.Suivant.Code + '"');
      end
      else
      begin
        // On met à jour la date de début de l'exercice suivant
        ExecuteSQL('UPDATE EXERCICE SET EX_DATEDEBUT = "' +
                   UsDateTime( FNouvelleDate + 1 ) + '" WHERE ' +
                   'EX_EXERCICE = "' + VH^.Suivant.Code + '"');
      end;
    end;
    // Rechargement du VH^.Encours et VH^.Suivant.Code
    ChargeMagExo(False);
  except
    on E: Exception do
    begin
      PgiError('Erreur de requête SQL : ' + E.Message, 'Procedure : MAJExercice');
      V_Pgi.IOError := OeUnknown;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/06/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_EXERCICE.MAJEcriture;
var lQuery         : TQuery;
    lQueryMereEcr  : TQuery;
    lQueryMEreEcr2 : TQuery;
    lTobMereEcr    : Tob;
    lTobMEreEcr2   : Tob;
    lTobEcriture   : Tob;
    lTobAxe        : Tob;
    lTobAna        : Tob;
    lStSql         : string;
    lStSqlMereEcr  : string;
    lStSqlMereEcr2 : string;
    i,j,k          : integer;
    lIndice        : integer;

    function TrouvePiece( vPrefixe : string ; vCodeExercice : string ) : string;
    begin
      Result := vPrefixe + '_EXERCICE = "'+ vCodeExercice + '" AND ' +
                vPrefixe + '_JOURNAL = "' + lQuery.FindField('E_JOURNAL').AsString + '" AND ' +
                vPrefixe + '_PERIODE = ' + IntToStr(lQuery.FindField('E_PERIODE').AsInteger) + ' AND ' +
                vPrefixe + '_NUMEROPIECE = ' + IntToStr( lQuery.FindField('E_NUMEROPIECE').AsInteger) + ' AND ' +
                vPrefixe + '_QUALIFPIECE = "'+ lQuery.FindField('E_QUALIFPIECE').AsString + '" ';
    end;

begin
  // Mise à jour du mode Pièce
  MajModePiece;
  if V_Pgi.IOError <> OeOk then Exit;

  // Récupération des Entêtes des Borderaux concernés
  lStSql := 'SELECT E_EXERCICE, E_JOURNAL, E_PERIODE, E_NUMEROPIECE, E_QUALIFPIECE FROM ECRITURE WHERE ' +
            '(E_MODESAISIE = "LIB" OR E_MODESAISIE = "BOR")' +
            ' AND E_QUALIFPIECE <> "Z"' +
            ' AND E_EXERCICE = "' + FExerciceDepart + '"' +
            ' AND E_DATECOMPTABLE > "' + UsDateTime( FDateSuperieure ) + '"' +
            ' AND E_DATECOMPTABLE <= "' + UsDateTime( FDateInferieure ) + '"' +
            ' GROUP BY E_EXERCICE, E_JOURNAL, E_QUALIFPIECE, E_PERIODE, E_NUMEROPIECE';

  lQuery := nil;
  lTobMereEcr  := Tob.Create('', nil, -1);
  lTobMereEcr2 := Tob.Create('', nil, -1);
  try
    try
      lQuery := OpenSql( lStSql, True);
      while not lQuery.Eof do
      begin
        // Récupération complète du Bordereau
        lStSQLMereEcr := TrouvePiece( 'E', FExerciceDepart );

        lQueryMereEcr := OpenSql('SELECT * FROM ECRITURE WHERE ' + lStSQLMereEcr +
                                 ' ORDER BY E_EXERCICE, E_JOURNAL, E_QUALIFPIECE, E_PERIODE, E_NUMEROPIECE, E_NUMLIGNE', True);

        // Requête qui récupère le bordereau d'arrivé des écritures à traiter
        lStSQLMereEcr2 := TrouvePiece( 'E', FExerciceArrivee );

        lQueryMereEcr2 := OpenSql('SELECT * FROM ECRITURE WHERE ' + lStSQLMereEcr2 +
                                  ' ORDER BY E_EXERCICE, E_JOURNAL, E_QUALIFPIECE, E_PERIODE, E_NUMEROPIECE, E_NUMLIGNE', True);

        // Suppression des TobFilles
        lTobMereEcr.ClearDetail;
        lTobMereEcr2.ClearDetail;
        // Chargement du Bordereau
        lTobMereEcr.LoadDetailDB( 'ECRITURE', '', '', lQueryMereEcr,  False);
        lTobMereEcr2.LoadDetailDB('ECRITURE', '', '', lQueryMereEcr2, False);

        // Libération des requêtes
        Ferme( lQueryMereEcr  );
        Ferme( lQueryMereEcr2 );

        // Parcours du Bordereau pour changer l'exercice des écritures concernées
        i := lTobMereEcr.Detail.Count - 1;
        repeat
          lTobEcriture := lTobMereEcr.Detail[i];
          // Chargement de l'analytique de la pièce entière
          CChargeAna( lTobEcriture );

          if (lTobEcriture.GetValue('E_DATECOMPTABLE') > FDateSuperieure) and
             (lTobEcriture.GetValue('E_DATECOMPTABLE') <= FDateInferieure) then
          begin
            // Indice pour l'ajout dans le ChangeParent
            lIndice := lTobMereEcr2.Detail.Count;

            for j := 0 to lTobEcriture.Detail.Count - 1 do
            begin
              lTobAxe := lTobEcriture.Detail[j];
              // Parcours des Sections Analytiques
              for k := 0 to lTobAxe.Detail.Count - 1 do
              begin
                // Parcours des Lignes d'Analytiques
                lTobAna := lTobAxe.Detail[k];
                lTobAna.PutValue('Y_EXERCICE', FExerciceArrivee);
                lTobAna.PutValue('Y_DATEMODIF',  NowH);
              end;
            end;

            // Ecritures qui changent d'exercice
            lTobEcriture.PutValue('E_EXERCICE', FExerciceArrivee);
            lTobEcriture.PutValue('E_IO', 'X');
            lTobEcriture.PutValue('E_DATEMODIF', NowH);
            lTobEcriture.ChangeParent( lTobMereEcr2, lIndice );
          end;
          i := i - 1;
        until (i < 0);

        // Suppression en Base des Bordereaux avant ré-insertion
        ExecuteSQL('DELETE FROM ECRITURE WHERE ' + lStSQLMereEcr);
        ExecuteSQL('DELETE FROM ECRITURE WHERE ' + lStSQLMereEcr2);

        // Suppression en Base de l'analytique
        ExecuteSQL('DELETE FROM ANALYTIQ WHERE ' + TrouvePiece( 'Y', FExerciceDepart));
        ExecuteSQL('DELETE FROM ANALYTIQ WHERE ' + TrouvePiece( 'Y', FExerciceArrivee));

        // Re-numérotation de la pièce
        if lTobMereEcr.Detail.Count <> 0 then
        begin
          CNumeroLigneBor( lTobMereEcr  );
          //lTobMereEcr.SaveToFile('C:\LTOBMEREECR.TXT', False, True, True);
        end;

        if lTobMereEcr2.Detail.Count <> 0 then
        begin
          CNumeroLigneBor( lTobMereEcr2 );
          //lTobMereEcr2.SaveToFile('C:\LTOBMEREECR2.TXT', False, True, True);
        end;

        // Enregistrement en Base
        lTobMereEcr.InsertDb ( nil, False );
        lTobMereEcr2.InsertDb( nil, False );

        // Passage à l'entête de Bordereau suivante
        lQuery.Next;
      end;
    except
      on E: Exception do
      begin
        PgiError('Erreur de requête SQL : ' + E.Message, 'Procedure : MAJEcriture');
        V_PGI.IoError := oeUnKnown;
      end;
    end;

  finally
    Ferme( lQuery );
    // Libération Mémoire des Tob
    FreeAndNil( lTobMereEcr  );
    FreeAndNil( lTobMereEcr2 );
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/06/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_EXERCICE.ChangementFinExercice;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Lek
Créé le ...... : 26/09/2005
Modifié le ... :   /  /
Description .. : Mettre ajour le champ exercice début ou fin de la table
Suite ........ : BUDJAL selon le changement de duré exercie. FQ16520
Mots clefs ... :
*****************************************************************}
  Procedure MajExerciceBudJal;
  begin
    try
      ExecuteSQL('UPDATE BUDJAL SET BJ_EXOFIN = "'+FExerciceArrivee+'" '+
              'WHERE BJ_PERFIN > "' + UsDateTime( FDateSuperieure ) + '" ' +
              'AND BJ_PERFIN <= "' + UsDateTime( FDateInferieure ) + '" ');
      ExecuteSQL('UPDATE BUDJAL SET BJ_EXODEB = "'+FExerciceArrivee+'" '+
              'WHERE BJ_PERDEB > "' + UsDateTime( FDateSuperieure ) + '" ' +
              'AND BJ_PERDEB <= "' + UsDateTime( FDateInferieure ) + '" ');
    except
      on E : Exception do
      begin
        PgiError('Erreur de requête SQL : ' + E.Message, 'Procedure : MajExerciceBudJal');
        V_PGI.IoError := oeUnKnown;
      end;
    end;
  end;

begin
  // INIT des variables utilisées tout au long du traitement
  try
    InitMoveProgressForm(Ecran, 'Changement de date de fin d''exercice', '', 10, True, True);

    if ((GetField('EX_EXERCICE') = VH^.Encours.Code) and (VH^.Suivant.Code = '') and
      (VH^.Encours.Fin > StrToDate(GetControlText('EX_DATEFIN'))))
     or
     ((GetField('EX_EXERCICE') = VH^.Suivant.Code) and
      (VH^.Suivant.Fin > StrToDate(GetControlText('EX_DATEFIN')))) then
    begin
      // Mise à jour de l'exercice uniquement
      MAJExercice;
    end
    else
    begin
      // Traitement sur l'exercice en cours, si présence d'un exercice suivant
      if (GetField('EX_EXERCICE') = VH^.EnCours.Code) and (VH^.Suivant.Code <> '') then
      begin
        // Exercice Suivant existe
        if FNouvelleDate > VH^.Encours.Fin then
        begin
          FExerciceDepart  := VH^.Suivant.Code;
          FExerciceArrivee := VH^.Encours.Code;
          FDateSuperieure  := VH^.Encours.Fin;
          FDateInferieure  := FNouvelleDate;
        end
        else
        begin
          FExerciceDepart  := VH^.EnCours.Code;
          FExerciceArrivee := VH^.Suivant.Code;
          FDateSuperieure  := FNouvelleDate;
          FDateInferieure  := VH^.Encours.Fin;
        end;

        if V_Pgi.IoError = OeOk then
        begin
          MAJExercice;
          MoveCurProgressForm('Mise à jour de la table : EXERCICE');
        end;

        if V_Pgi.IoError = OeOk then
        begin
          MoveCurProgressForm('Mise à jour de la table : ECRITURE');
          MAJEcriture;
        end;

        if V_Pgi.IoError = OeOk then
        begin
          MoveCurProgressForm('Mise à jour de la table : ANALYTIQ');
          MAJAnalytiquePur;
        end;

        if V_Pgi.IoError = OeOk then
        begin
          MoveCurProgressForm('Mise à jour de la table : CBALSIT');
          MAJBalanceSituation;
        end;

        if V_Pgi.IoError = OeOk then
        begin
          MoveCurProgressForm('Mise à jour de la table : BUDECR');
          MAJAutreTable( 'BUDECR', True );
        end;

        if V_Pgi.IoError = OeOk then {Lek 26/09/05 FA16520}
        begin
          MoveCurProgressForm('Mise à jour de la table : BUDJAL');
          MajExerciceBudJal;
        end;

        if V_Pgi.IoError = OeOk then
        begin
          MoveCurProgressForm('Mise à jour de la table : CRELBQEANALYTIQ');
          MAJAutreTable( 'CRELBQEANALYTIQ', True);
        end;

        if V_Pgi.IoError = OeOk then
        begin
          MoveCurProgressForm('Mise à jour de la table : RBLIGECR');
          MAJAutreTable( 'RBLIGECR', False );
        end;

        if V_Pgi.IoError = OeOk then
        begin
          MoveCurProgressForm('Mise à jour de la table : TRECRITURE');
          MAJAutreTable( 'TRECRITURE', True );
        end;

        if V_Pgi.IoError = OeOk then
        begin
          // Recalcul des cumuls des comptes, journaux et sections
          MoveCurProgressForm('Recalcul des comptes, journaux et sections...');
          MajTotTousComptes(False, '');
        end;

        // Recalcul des immobilisations
  {$IFDEF AMORTISSEMENT}
        if V_Pgi.IoError = OeOk then
          FMajImmo := True;
  {$ENDIF}
      end;
    end;

  finally
    FiniMoveProgressForm;
  end;
end;

{$IFDEF AMORTISSEMENT}
procedure TOM_EXERCICE.MajAmortissement;
begin
  if ExisteSQL ('SELECT I_IMMO FROM IMMO') then
  begin
    ChargeMagExo(False); // Chargement des ParamSOC
    ChargeVHImmo;
    { FQ 15673 - CA - 29/06/2005  - si la date de début d'exercice a changé, on force le recalcul des reprises }
    UpdateBaseImmo( False , FbDateDebutModifie );
  end;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/08/2004
Modifié le ... :   /  /
Description .. : Mise à jour des écritures analytiques pures
Mots clefs ... : Y_TYPEANALYTIQUE;
*****************************************************************}
procedure TOM_EXERCICE.MAJAnalytiquePur;
var lSt : string;
begin
  try
    lSt := 'UPDATE ANALYTIQ SET Y_EXERCICE = "' + FExerciceArrivee + '", ' +
           'Y_DATEMODIF = "' + UsTime(NowH) + '" WHERE ' +
           'Y_EXERCICE = "' + FExerciceDepart + '" AND ' +
           'Y_DATECOMPTABLE > "' + UsDateTime( FDateSuperieure ) + '" AND ' +
           'Y_DATECOMPTABLE <= "' + UsDateTime( FDateInferieure ) + '" AND ' +
           'Y_TYPEANALYTIQUE = "X"';

    ExecuteSQL( lSt );
  except
    on E: Exception do
      begin
        PgiError('Erreur de requête SQL : ' + E.Message, 'Procedure : MAJAnalytiquePur');
        V_Pgi.IOError := OeUnknown;
      end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/06/2004
Modifié le ... : 23/09/2005
Description .. : Lek FQ16520 et FQ 16484
Suite ........ : La Maj  _DateModif n'est pas la clause Where mais c'est 
Suite ........ : dans l'intruction SET.
Mots clefs ... :
*****************************************************************}
procedure TOM_EXERCICE.MAJAutreTable( vStNomTable : string ; vBoavecDateModif : Boolean );
var lStPrefixe : string;
    lSt        : string;
begin
  try
    if vStNomTable = '' then Exit;
    lStPrefixe := TableToPrefixe( vStNomTable );
    lSt := 'UPDATE ' + vstNomTable + ' SET ' +
           lStPrefixe + '_EXERCICE = "' + FExerciceArrivee + '"';
    if vBoAvecDateModif then
      lSt := lSt + ', ' + lStPrefixe + '_DATEMODIF = "' + UsTime(Now) + '"';
    lSt := lSt +' WHERE ' +
           lStPrefixe + '_EXERCICE = "' + FExerciceDepart + '" AND ' +
           lStPrefixe + '_DATECOMPTABLE > "' + UsDateTime( FDateSuperieure ) + '" AND ' +
           lStPrefixe + '_DATECOMPTABLE <= "' + UsDateTime( FDateInferieure ) + '"';
    ExecuteSQL( lSt );
  except
    on E: Exception do
      begin
        PgiError('Erreur de requête SQL : ' + E.Message, 'Procedure : MAJAutreTable');
        V_Pgi.IOError := OeUnknown;
      end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_EXERCICE.MAJBalanceSituation;
var lQuery     : TQuery;
    lTobBalSit : Tob;
    i          : integer;
begin
  lQuery     := nil;
  lTobBalSit := nil;
  try
    try
      lQuery := OpenSql('SELECT * FROM CBALSIT WHERE ' +
                      '(BSI_DATE1 >= "' + UsDateTime(FDateSuperieure) + '" AND ' +
                      ' BSI_DATE1 <= "' + UsDateTime(FDateInferieure) + '" ) OR ' +
                      '(BSI_DATE2 >= "' + UsDateTime(FDateSuperieure) + '" AND ' +
                      ' BSI_DATE2 <= "' + UsDateTime(FDateInferieure) + '" ) ' +
                      'ORDER BY BSI_CODEBAL', True);

      if not lQuery.Eof then
      begin
      lTobBalSit := Tob.Create('', nil, -1);
      lTobBalSit.LoadDetailDB('CBALSIT', '', '', lQuery,  False);
      Ferme( lQuery );

      for i := 0 to lTobBalSit.Detail.Count - 1 do
      begin
        if (lTobBalSit.Detail[i].GetValue('BSI_DATE2') <= FNouvelleDate) then
        begin
          if (lTobBalSit.Detail[i].GetValue('BSI_DATE1') >= VH^.EnCours.Deb) then
            lTobBalSit.Detail[i].PutValue('BSI_EXERCICE', VH^.EnCours.Code)
          else
            lTobBalSit.Detail[i].PutValue('BSI_EXERCICE', '---');
        end
        else
        begin
          if (lTobBalSit.Detail[i].GetValue('BSI_DATE1') >= VH^.EnCours.Deb) and
             (lTobBalSit.Detail[i].GetValue('BSI_DATE1') <= FNouvelleDate) then
            lTobBalSit.Detail[i].PutValue('BSI_EXERCICE', '---')
          else
            lTobBalSit.Detail[i].PutValue('BSI_EXERCICE', VH^.Suivant.Code);
        end;
      end;
      lTobBalSit.UpdateDB;
      end;
    except
      on E: Exception do
      begin
        PgiError('Erreur de requête SQL : ' + E.Message, 'Procedure : MAJBalanceSituation');
        V_Pgi.IOError := OeUnknown;
      end;
    end;
  finally
    FreeAndNil(lTobBalSit);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/09/2004
Modifié le ... :   /  /
Description .. : Mise à jour du Mode Pièce
Mots clefs ... :
*****************************************************************}
procedure TOM_EXERCICE.MAJModePiece;
var lQuery : TQuery;
    lSt    : string;
    lTob   : Tob;
    i      : integer;
begin
  lQuery := nil;
  lTob   := nil;
  try
    try
      lQuery := OpenSql('SELECT Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, ' +
                        'Y_NUMEROPIECE, Y_NUMLIGNE, Y_AXE, Y_NUMVENTIL, ' +
                        'Y_QUALIFPIECE, Y_DATEMODIF FROM ANALYTIQ ' +
                        'LEFT JOIN ECRITURE ON ' +
                        'E_JOURNAL = Y_JOURNAL AND ' +
                        'E_EXERCICE = Y_EXERCICE AND ' +
                        'E_DATECOMPTABLE = Y_DATECOMPTABLE AND ' +
                        'E_NUMEROPIECE = Y_NUMEROPIECE AND ' +
                        'E_NUMLIGNE = Y_NUMLIGNE AND ' +
                        'E_QUALIFPIECE = Y_QUALIFPIECE WHERE ' +
                        '(E_MODESAISIE = "" OR E_MODESAISIE = "-")' +
                        ' AND E_QUALIFPIECE <> "Z"' +
                        ' AND E_EXERCICE = "' + FExerciceDepart + '"' +
                        ' AND E_DATECOMPTABLE > "' + UsDateTime( FDateSuperieure ) + '"' +
                        ' AND E_DATECOMPTABLE <= "' + UsDateTime( FDateInferieure ) + '" ' +
                        'ORDER BY Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE', True);

      lTob := Tob.Create('', nil, -1);
      lTob.LoadDetailDB('ANALYTIQ', '', '', lQuery, False);
      Ferme( lQuery );
      for i := 0 to lTob.Detail.Count -1 do
      begin
        lTob.Detail[i].PutValue('Y_EXERCICE', FExerciceArrivee );
        lTob.Detail[i].PutValue('Y_DATEMODIF', NowH );
      end;

      if lTob.Detail.Count > 0 then
        lTob.UpdateDb;

      // Mise à jour du Mode Pièce
      lSt := 'UPDATE ECRITURE SET E_EXERCICE = "' + FExerciceArrivee + '", ' +
             'E_IO = "X", ' +
             'E_DATEMODIF = "' + UsTime( NowH ) + '" WHERE ' +
             '(E_MODESAISIE = "" OR E_MODESAISIE = "-")' +
             ' AND E_QUALIFPIECE <> "Z"' +
             ' AND E_EXERCICE = "' + FExerciceDepart + '"' +
             ' AND E_DATECOMPTABLE > "' + UsDateTime( FDateSuperieure ) + '"' +
             ' AND E_DATECOMPTABLE <= "' + UsDateTime( FDateInferieure ) + '"';

      ExecuteSQL(lSt);
    except
      on E: Exception do
      begin
        PgiError('Erreur de requête SQL : ' + E.Message, 'Procedure : MAJModePiece');
        V_PGI.IOError := OeUnKnown;
      end;
    end;

  finally
    FreeAndNil( lTob );
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 09/11/2006
Modifié le ... :   /  /
Description .. : regarde la présence de tableau des variations sur l'exercice
Mots clefs ... :
*****************************************************************}
function TOM_EXERCICE.TabVarexist(Exercice: String): Boolean;
begin
  Result                := ExisteSQL('SELECT 1 FROM CPTABLEAUVAR WHERE CTV_EXERCICE="'+Exercice+'"');
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 05/02/2007
Modifié le ... :   /  /
Description .. : Test la présence d'une note de travail sur un exercice
Mots clefs ... :
*****************************************************************}
function TOM_EXERCICE.TabNoteTavailexist(Exercice: String): Boolean;
begin
  Result                := ExisteSQL('SELECT 1 FROM CPNOTETRAVAIL WHERE CNO_EXERCICE="'+Exercice+'"');
end;



////////////////////////////////////////////////////////////////////////////////


initialization
  registerclasses([TOM_EXERCICE]);
end.


