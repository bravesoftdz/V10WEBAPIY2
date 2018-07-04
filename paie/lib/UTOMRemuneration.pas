{***********UNITE*************************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 16/08/2001
Modifié le ... :   /  /
Description .. : TOM de la table REMUNERATION
Mots clefs ... : PAIE
*****************************************************************}
{
PT1   : 02/08/2001 PH V547 Le champ PRM_BASEMTQTE passe de boolean en COMBO
PT2   : 10/08/2001 PH V547 Rend tjrs accessible le champ PRM_DECMONTANT
PT3   : 14/11/2001 SB V562 Duplication affectation des predefinis et no dossier
PT4   : 14/11/2001 SB V562 Test si enr existe ds PROFILRUB
                           Ajout des champs PREDEFINI Et NODOSSIER
PT5   : 29/11/2001 SB V563 On force la validation avant la Duplication
PT6   : 30/11/2001 SB V563 Dysfonctionnement test code existant sur Duplication
PT7   : 13/12/2001 SB V570 Fiche de bug n° 279
                           Test code existant ne test pas bon numéro de dossier
PT8   : 18/01/2002 PH V571 Si Rem de code calcul = Montant alors le taux non
                           géré et on ne force pas un montant mais une quantité
PT9   : 19/04/2002 SB V571 Fiche de bug n°369 Duplication d'une rub. Dossier en
                           Standard : L'affectation des prédéfini est erronée..
PT10  : 12/06/2002 VG V582 Version S3 + Affichage en fonction de VH_Paie.PGBTP
PT11  : 18/12/2002 PH V591 PORTAGECWAS controle des types de champs sur les
                           setfield
PT12-1: 18/02/2003 SB V595 FQ 10202 Suppression : Contrôle existence rubrique
                           dans Masque de Saisie et motif absence
PT12-2: 18/02/2003 SB V595 FQ 10531 Anomalie sur controle existence rubrique
                           dans table commune
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
PT13  : 19/06/2003 SB V595 Force validation pour rechargement contexte paie
PT14  : 16/07/2003 SB V595 FQ 10202 Validation controle de vraisemblance sur le
                           paramétrage du type de calcul
PT15  : 08/08/2003 PH V595 FQ 10757  si on a renseigné une date de début de
                           validité, obliger la saisie d'une date de fin de
                           validité
PT16  : 01/09/2003 PH V_421 Non RAZ type de base et base dans le cas d'un calcul
                           de type montant
PT17-1: 11/09/2003 SB V_42 FQ 10649 duplication du paramétrage comptable lié à
                           la remunération
PT17-2: 11/09/2003 SB V_42 FQ 10202 Anomalie controle element type de calcul
PT18-1: 16/09/2003 SB V_42 FQ 10820 Message duplication effectuée
PT18-2: 10/10/2003 VG V_42 FQ 10852 BTP en S3
PT19  : 14/10/2003 SB V_42 Portage CWAS : Impression fiche
PT20  : 23/02/2004 SB V_50 Mise à jour date modif en duplication
PT21  : 02/09/2004 PH V_50 FQ 11573 Validité de calcul de la rémunération
PT22  : 01/10/2004 PH V_50 Changement Caption Nom court dans le cas où le thème
                           est net à payer
PT23  : 14/12/2004 JL V_60 FQ 11799 rémunération au lieu de rénumération
PT24  : 07/02/2005 SB V_60 FQ 11804 En cwas force validation après gestion
                           associé
PT25  : 26/05/2005 SB V_60 FQ 12307 Modification de l'ordre de présentation des
                           rubriques
PT26  : 10/06/2005 SB V_60 CWAS : Chargement tablette compatibilite CWAS
PT27  : 10/06/2005 SB V_60 Ergonomie CWAS
PT28  : 02/08/2005 PH V_60 FQ 12467 Ergonomie
PT29  : 11/08/2005 PH V_60 FQ 12502 Duplication de rubrique avec cumul alpha
                           associé
PT30  : 18/08/2005 SB V_60 FQ 11969 Mise à jour du libellé de la table PROFILRUB
PT31  : 12/09/2005 PH V_60 Affichage memo liste des champs personnalisés
PT32  : 06/01/2006 SB V_65 FQ 12804 suppression setfield en chargement fiche
PT33  : 09/02/2006 EP V_65 FQ 12781 Utilisateur non réviseur : interdire la
                           personnalisation
PT33  : 27/03/2006 PH V_65 Prise en compte des champs motifs absences dédoublés
                           (Heures,Jours)
PT34  : 24/04/2006 SB V_65 FQ 12613 Refonte contrôle ondeleterecord pour test
                           parametrage multidossier
PT35  : 25/04/2006 SB V_65 FQ 12917 Erreur SQL : Refonte execute sql sur libellé
                           suite FQ 11969
PT36  : 01/09/2006 PH V_65 erreur dans la table courrier à cause du rechargement
                           des tablettes
PT37  : 10/01/2007 GGS V_72 FQ 12694 journal d'événement
PT38  : 08/03/2007 VG V_72 Ajout bouton gestion spécifique
PT39  : 29/03/2007 PH V_72 FQ 13628 Rajout mois de validité
PT40  : 26/04/2007 GGS V_80 modification de gestion de la trace
PT41  : 03/05/2007 MF V_72 FQ 14088 : paramétrage pour calcul des IJSS sur
                           clique droit saisie bulletin
PT42  : 07/05/2007 PH V_72 Concept Paie
PT43  : 28/04/2007 GGS V_72 re-modification de gestion de la trace (dupli)
PT44  : 05/06/2007 MF V_72 FQ 14338 : modification pamétrage calcul des IJSS sur
                           clique droit saisie bulletin
PT45  : 06/09/2007 VG V_80 Gestion spécifique grisée en prédéfini dossier
                           FQ N°14601
PT46  : 06/11/2007 FC V_80 Correction bug journal des évènement quand création (lib Modification au lieu de Création)
PT49  : 20/08/2008 JPP V_80 FQ15450 On empeche la duplication de la personnalisation lors de la duplication d'une rubrique CEGID en dossier
}
unit UTOMRemuneration;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
{$IFDEF EAGLCLIENT}
  UtileAGL, eFiche, MaineAgl,
{$ELSE}
  db, {$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}HDB, Fiche, DBCtrls, FE_Main, EdtREtat,
{$ENDIF$}
  HCtrls, HEnt1, HMsgBox, UTOM, UTOB, HTB97,
  P5Def,
  EntPaie,
  PAIETOM; //PT37


type
  TOM_Remuneration = class(PGTOM) //PT37 class TOM devient PGTOM
    procedure OnChangeField(F: TField); override;
    procedure OnNewRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnArgument(stArgument: string); override;
    procedure OnAfterUpdateRecord; override;
    procedure OnClose; override;
  private
    ExisteCod, LectureSeule, CEG, STD, DOS, OnFerme: boolean;
    mode, DerniereCreate: string;
    NomChamp: array[1..4] of Hstring;
    ValChamp: array[1..4] of variant;
    Trace: TStringList; //PT37
    LeStatut:TDataSetState; //PT46
//    TobRem : Tob;

    procedure AffecteLookupChampRem(Champ, TypeChamp, libelle: string);
    procedure DupliquerRemuneration(Sender: TObject);
    function ControlChampValeur: integer; //PT14
    procedure CBAccesFicheClick(Sender: TObject);
    procedure ImprimerClick(Sender: TObject); //PT19
    procedure PgUpdateDateAgl(Rub: string); //PT20
    procedure GestionSpec_OnClick(Sender: TObject);
// d PT41
    procedure CBIJSSClick(Sender: TObject);
    procedure TYPRUBClick(Sender: TObject);
// f PT41
  end;

implementation
uses PgOutils2, Pgoutils;
//============================================================================================//
//                                 PROCEDURE OnArgument
//============================================================================================//

procedure TOM_Remuneration.OnArgument;
var
  Btn: TToolBarButton97;
{DEB PT33}
  Q: TQuery;
  Perso: Boolean;
{FIN PT33}
// d PT41
  CHB: TCheckbox;
  CMB: TComboBox;
// f PT41
begin
  inherited;
  Btn := TToolBarButton97(GetControl('BDUPLIQUER'));
  if (btn <> nil) then
    Btn.OnClick := DupliquerRemuneration;

  if (VH_Paie.PGBTP = False) then
  begin
    SetControlVisible('LABEL_BTP', FALSE);
    SetControlVisible('BLVBTP', FALSE);
    SetControlVisible('TPRM_BTPARRET', FALSE);
    SetControlVisible('PRM_BTPARRET', FALSE);
    SetControlVisible('PRM_BTPSALAIRE', FALSE);
  end;

  if TCheckBox(GetControl('CBACCESFICHE')) <> nil then
    TCheckBox(GetControl('CBACCESFICHE')).OnClick := CBAccesFicheClick;

  Btn := TToolBarButton97(GetControl('BIMPRIMER'));
  if (btn <> nil) then
  begin
    setControlvisible('BIMPRIMER', TRUE);
    Btn.OnClick := ImprimerClick;
  end;
{DEB PT33}
  Perso := False;
  Q := OpenSql('SELECT US_CONTROLEUR' +
    ' FROM UTILISAT WHERE' +
    ' US_UTILISATEUR="' + V_PGI.FUser + '"', True);
  if (not Q.EOF) then
  begin
    if Q.FindField('US_CONTROLEUR').AsString = 'X' then
      Perso := True;
  end;
  Ferme(Q);
  SetControlEnabled('BTNPERSO', Perso);
{ FIN PT33 }
//PT43  Trace := TStringList.Create; //PT37

//PT38
  Btn := TToolBarButton97(GetControl('BGESTIONSPEC'));
  if (Btn <> nil) then
    Btn.OnClick := GestionSpec_OnClick;
//FIN PT38

//d PT41
  CHB := TCheckbox(GetControl('CBIJSS'));
  if (CHB <> nil) then
    CHB.Onclick := CBIJSSClick;

  CMB := TComboBox(GetControl('TYPRUB'));
  if (CMB <> nil) then
    CMB.Onclick := TYPRUBClick;

// d PT44
// Ces contrôles ne sont visibles que pour les rubriques CEGID
  SetControlVisible('CBIJSS', FALSE);
  SetControlVisible('TYPRUB', FALSE);
//f PT44
// f PT41
  PaieConceptPlanPaie(Ecran); // PT42
end;

//============================================================================================//
//                                 PROCEDURE On Load Record
//============================================================================================//
procedure TOM_Remuneration.OnLoadRecord;
var Q: TQuery;
MonMemo: TMemo;
LabPers: ThLabel;
Personnalise : Boolean;
begin
inherited;
Personnalise:= False;  //PT38
if not (DS.State in [dsInsert]) then
   DerniereCreate:= '';

if (mode='DUPLICATION') then
   exit;
SetControlProperty ('ACCES', 'Text', 'FALSE');
AccesPredefini ('TOUS', CEG, STD, DOS);
LectureSeule:= FALSE;
if (Getfield ('PRM_PREDEFINI')='CEG') then
   begin
   LectureSeule:= (CEG = False);
   PaieLectureSeule (TFFiche(Ecran), (CEG = False));
   SetControlEnabled ('BDelete', CEG);
// d PT44
// Ces contrôles ne sont visibles que pour les rubriques CEGID
   SetControlVisible ('CBIJSS', True);
   SetControlVisible ('TYPRUB', True);
//f PT44

   if (CEG=True) then
      SetControlProperty ('ACCES', 'Text', 'TRUE')
   else
      SetControlProperty ('ACCES', 'Text', 'FALSE');
   SetControlEnabled ('BGESTIONSPEC', True);        //PT45
   end
else
if (Getfield ('PRM_PREDEFINI')='STD') then
   begin
   LectureSeule:= (STD = False);
   PaieLectureSeule (TFFiche(Ecran), (STD = False));
   SetControlEnabled ('BDelete', STD);
   if (STD=True) then
      SetControlProperty ('ACCES', 'Text', 'TRUE')
   else
      SetControlProperty ('ACCES', 'Text', 'FALSE');
   SetControlEnabled ('BGESTIONSPEC', True);        //PT45
   end
else
if (Getfield ('PRM_PREDEFINI')='DOS') then
   begin
   LectureSeule:= (DOS = False);
   PaieLectureSeule (TFFiche(Ecran), LectureSeule);
   SetControlEnabled ('BDelete', DOS);
   if (DOS=True) then
      SetControlProperty ('ACCES', 'Text', 'TRUE')
   else
      SetControlProperty ('ACCES', 'Text', 'FALSE');
   SetControlEnabled ('BGESTIONSPEC', False);       //PT45
   end;
SetControlEnabled ('BInsert', True);
SetControlEnabled ('PRM_PREDEFINI', False);
SetControlEnabled ('PRM_RUBRIQUE', False);
SetControlEnabled ('BDUPLIQUER', True);
if (DS.State in [dsInsert]) then
   begin
   LectureSeule:= FALSE;
   PaieLectureSeule (TFFiche(Ecran), False);
   SetControlEnabled ('PRM_PREDEFINI', True);
   SetControlEnabled ('PRM_RUBRIQUE', True);
   SetControlEnabled ('BGESTIONASS', False);
   SetControlEnabled ('BVENTIL', False);
   SetControlEnabled ('BInsert', False);
   SetControlEnabled ('BDUPLIQUER', False);
   SetControlEnabled ('BDelete', False);
   end;
SetControlVisible ('BDUPLIQUER', True);
PaieConceptPlanPaie (Ecran); // PT42

Q:= OpenSql ('SELECT *'+
             ' FROM PGEXCEPTIONS WHERE'+
             ' PEN_RUBRIQUE="'+GetField ('PRM_RUBRIQUE')+'" AND'+
             ' ##PEN_PREDEFINI## AND'+
             ' PEN_NATURERUB="AAA"'+
             ' ORDER BY PEN_PREDEFINI', false);
if (not Q.EOF) then
   begin
   Personnalise:= True;
   Q.FIRST;
   LabPers:= THLabel (GetControl ('LBLPERSO'));
   if (LabPers <> nil) then
      begin
      LabPers.Caption:= 'Attention, cette rémunération est personnalisée';
      if (Q.FindField ('PEN_PREDEFINI').AsString='DOS') then
         LabPers.Caption:= LabPers.Caption+' dossier'
      else
         LabPers.Caption:= LabPers.Caption+' standard';
      end;
   SetControlVisible ('LBLPERSO', TRUE);
   MonMemo:= TMemo (GetControl ('MEMOPERSO'));
   if (MonMemo <> nil) then
      begin
      SetControlVisible ('MEMOPERSO', TRUE);
      MonMemo.Enabled:= TRUE;
      MonMemo.Clear;
      MonMemo.Lines.Add ('Les champs suivants ont été modifiés :');
      if (GetField ('PRM_LIBELLE')<>Q.FindField ('PEN_LIBELLE').AsString) then
         MonMemo.Lines.Add ('-Libellé');
      if (GetField ('PRM_IMPRIMABLE')<>Q.FindField ('PEN_IMPRIMABLE').AsString) then
         MonMemo.Lines.Add ('-Rubrique imprimable');
      if (GetField ('PRM_BASEIMPRIMABLE')<>Q.FindField ('PEN_BASEIMPRIMABLE').AsString) then
         MonMemo.Lines.Add ('-Base imprimable');
      if (GetField ('PRM_TAUXIMPRIMABLE')<>Q.FindField ('PEN_TAUXIMPRIMABLE').AsString) then
         MonMemo.Lines.Add ('-Taux imprimable');
      if (GetField ('PRM_COEFFIMPRIM')<>Q.FindField ('PEN_COEFFIMPRIM').AsString) then
         MonMemo.Lines.Add ('-Coefficient imprimable');
      if (GetField ('PRM_DECBASE')<>Q.FindField ('PEN_DECBASE').AsInteger) then
         MonMemo.Lines.Add ('-Nombre décimales de la base');
      if (GetField ('PRM_DECTAUX')<>Q.FindField ('PEN_DECTAUX').AsInteger) then
         MonMemo.Lines.Add ('-Nombre décimales du taux');
      if (GetField ('PRM_DECCOEFF')<>Q.FindField ('PEN_DECCOEFF').AsInteger) then
         MonMemo.Lines.Add ('-Nombre décimales du coefficient');
      end
   else
      SetControlVisible ('MEMOPERSO', FALSE);
   end
else
   begin
   SetControlVisible ('LBLPERSO', FALSE);
   SetControlVisible ('MEMOPERSO', FALSE);
   end;
Ferme (Q);

//PT38
Q:= OpenSql ('SELECT *'+
             ' FROM PGEXCEPTIONS WHERE'+
             ' PEN_RUBRIQUE="'+GetField ('PRM_RUBRIQUE')+'" AND'+
             ' ##PEN_PREDEFINI## AND'+
             ' PEN_NATURERUB="AAA"'+
             ' ORDER BY PEN_PREDEFINI', false);
if (not Q.EOF) then
   begin
   Personnalise:= True;
   Q.FIRST;
   LabPers:= THLabel (GetControl ('LBLPERSO'));
   if (LabPers <> nil) then
      begin
      LabPers.Caption:= 'Attention, cette rémunération est personnalisée';
      if (Q.FindField ('PEN_PREDEFINI').AsString='DOS') then
         LabPers.Caption:= LabPers.Caption+' dossier'
      else
         LabPers.Caption:= LabPers.Caption+' standard';
      end;
   SetControlVisible ('LBLPERSO', TRUE);
   MonMemo:= TMemo (GetControl ('MEMOPERSO'));
   if (MonMemo <> nil) then
      begin
      SetControlVisible ('MEMOPERSO', TRUE);
      MonMemo.Enabled:= TRUE;
      MonMemo.Clear;
      MonMemo.Lines.Add ('Les champs suivants ont été modifiés :');
      if (GetField ('PRM_LIBELLE')<>Q.FindField ('PEN_LIBELLE').AsString) then
         MonMemo.Lines.Add ('-Libellé');
      if (GetField ('PRM_IMPRIMABLE')<>Q.FindField ('PEN_IMPRIMABLE').AsString) then
         MonMemo.Lines.Add ('-Rubrique imprimable');
      if (GetField ('PRM_BASEIMPRIMABLE')<>Q.FindField ('PEN_BASEIMPRIMABLE').AsString) then
         MonMemo.Lines.Add ('-Base imprimable');
      if (GetField ('PRM_TAUXIMPRIMABLE')<>Q.FindField ('PEN_TAUXIMPRIMABLE').AsString) then
         MonMemo.Lines.Add ('-Taux imprimable');
      if (GetField ('PRM_COEFFIMPRIM')<>Q.FindField ('PEN_COEFFIMPRIM').AsString) then
         MonMemo.Lines.Add ('-Coefficient imprimable');
      if (GetField ('PRM_DECBASE')<>Q.FindField ('PEN_DECBASE').AsInteger) then
         MonMemo.Lines.Add ('-Nombre décimales de la base');
      if (GetField ('PRM_DECTAUX')<>Q.FindField ('PEN_DECTAUX').AsInteger) then
         MonMemo.Lines.Add ('-Nombre décimales du taux');
      if (GetField ('PRM_DECCOEFF')<>Q.FindField ('PEN_DECCOEFF').AsInteger) then
         MonMemo.Lines.Add ('-Nombre décimales du coefficient');
      end
   else
      SetControlVisible ('MEMOPERSO', FALSE);
   end
else
   begin
   SetControlVisible ('LBLPERSO', FALSE);
   SetControlVisible ('MEMOPERSO', FALSE);
   end;
Ferme (Q);
//FIN PT38

// d PT41
// d PT44
// Ces contrôles ne sont visibles que pour les rubriques CEGID
if (Getfield ('PRM_PREDEFINI')='CEG') then
   begin
//f PT44
   if (copy (Getfield ('PRM_DADS'), 1, 1)<>'X') then
      begin
      SetControlProperty ('CBIJSS', 'Text', 'FALSE');
      SetControlVisible ('TYPRUB', False);
      SetControlText ('CBIJSS', '-');
      SetControlProperty ('TYPRUB', 'Value', '');
      end
   else
      begin
      SetControlProperty ('TYPRUB', 'Value', copy (Getfield ('PRM_DADS'), 2, 3));
      SetControlVisible ('TYPRUB', True);
      SetControlText ('CBIJSS', 'X');
      end;
   end; // PT44
// f PT41
end;

//============================================================================================//
//                                 PROCEDURE Dupliquer Remuneration
//============================================================================================//

procedure TOM_Remuneration.DupliquerRemuneration(Sender: TObject);
var
{$IFDEF EAGLCLIENT}
  Code: THEdit;
{$ELSE}
  Code: THDBEdit;
{$ENDIF}
  T, T_Cumul, TCumulSpec, TOB_GestAssoc, TOB_GestionSpec, T_Ventil: TOB;
  i: integer;
  AncValNat, AncValRub, ChampBug, ChampBug1, ChampBug2, ChampBug3: string;
  ChampBug4, St, NoDossier, Cumulpaie, OriginePred, PredExist: string;
  Champ: array[1..3] of Hstring;
  Valeur: array[1..3] of variant;
  Ok: Boolean;
  Q: TQuery;
begin
  ChampBug := GetField('PRM_THEMEREM');
  ChampBug1 := GetField('PRM_IMPRIMABLE');
  ChampBug2 := GetField('PRM_BASEIMPRIMABLE');
  ChampBug3 := GetField('PRM_TAUXIMPRIMABLE');
  ChampBug4 := GetField('PRM_CODECALCUL');
  AncValNat := GetField('PRM_NATURERUB');
  AncValRub := GetField('PRM_RUBRIQUE');
  TFFiche(Ecran).BValider.Click;
  mode := 'DUPLICATION';
  AglLanceFiche('PAY', 'CODE', '', '', 'AAA;' + AncValRub + '; ;4');
// d PT44
// Ces contrôles ne sont visibles que pour les rubriques CEGID
  if (PGCodePredefini <> 'CEG') then
  begin
    SetControlVisible('CBIJSS', FALSE);
    SetControlVisible('TYPRUB', FALSE);
  end;
// f PT44
    // DEB PT42
  if (PGCodePredefini = 'DOS') AND (NOT DOS) then
  begin
    PgiInfo ('Vous n''êtes pas autorisé à créer une cotisation de type dossier.', Ecran.Caption);
    exit;
  end;
  // FIN PT42
  if (PGCodeDupliquer <> '') then
  begin
    Champ[1] := 'PRM_PREDEFINI';
    Valeur[1] := PGCodePredefini;
    Champ[2] := 'PRM_NODOSSIER';
    if (PGCodePredefini = 'DOS') then
      Valeur[2] := PgRendNoDossier()
    else
      Valeur[2] := '000000';
    Champ[3] := 'PRM_RUBRIQUE';
    Valeur[3] := PGCodeDupliquer;
    Ok := RechEnrAssocier('REMUNERATION', Champ, Valeur);
    if (Ok = False) then //Test si code existe ou non
    begin
      TOB_GestAssoc := TOB.Create('La Remunération originale', nil, -1);
      st := 'SELECT *' +
        ' FROM CUMULRUBRIQUE WHERE' +
        ' PCR_NATURERUB="' + AncValNat + '" AND' +
        ' ##PCR_PREDEFINI## PCR_RUBRIQUE="' + AncValRub + '"';
      Q := OpenSql(st, TRUE);
      TOB_GestAssoc.LoadDetailDB('CUMULRUBRIQUE', '', '', Q, FALSE);
      FERME(Q);
      T_Cumul := TOB.Create('La Remunération dupliquée', nil, -1);
//PT38
      TOB_GestionSpec := TOB.Create('La gestion spéc orig', nil, -1);
      st := 'SELECT *' +
        ' FROM CUMULRUBDOSSIER WHERE' +
        ' PKC_NATURERUB="' + AncValNat + '" AND' +
        ' PKC_RUBRIQUE="' + AncValRub + '"';
      Q := OpenSql(st, TRUE);
      TOB_GestionSpec.LoadDetailDB('CUMULRUBDOSSIER', '', '', Q, FALSE);
      FERME(Q);
      TCumulSpec := TOB.Create('La gestion spéc dest', nil, -1);
//FIN PT38
{$IFDEF EAGLCLIENT}
      Code := THEdit(GetControl('PRM_RUBRIQUE'));
{$ELSE}
      Code := THDBEdit(GetControl('PRM_RUBRIQUE'));
{$ENDIF}
      if (code <> nil) then
        DupliquerPaie(TFFiche(Ecran).TableName, Ecran);
      SetField('PRM_RUBRIQUE', PGCodeDupliquer);
      try
        BeginTrans;
        for i := 0 to TOB_GestAssoc.Detail.Count - 1 do
        begin
          if ((TOB_GestAssoc.Detail[i].GetValue('PCR_NATURERUB')) = AncValNat) and
            ((TOB_GestAssoc.Detail[i].GetValue('PCR_RUBRIQUE')) = AncValRub) then
          begin
            T := TOB_GestAssoc.Detail[i];
            if (T <> nil) then
            begin
              T.PutValue('PCR_RUBRIQUE', PGCodeDupliquer);
              T.PutValue('PCR_PREDEFINI', PGCodePredefini);
              Cumulpaie := T.GetValue('PCR_CUMULPAIE');
              OriginePred := '';
              if (Cumulpaie <> '') and (Length(Cumulpaie) = 2) then
              begin
                if (IsNumeric(CumulPaie)) then
                  if (StrToInt(CumulPaie) > 50) and
                    ((Cumulpaie[2] = '5') or (Cumulpaie[2] = '7') or
                    (Cumulpaie[2] = '9')) then
                    OriginePred := 'DOS';
              end;
              if (OriginePred = 'DOS') then
              begin
                T.PutValue('PCR_PREDEFINI', 'DOS');
                T.PutValue('PCR_NODOSSIER', PgRendNoDossier());
              end
              else
                T.PutValue('PCR_NODOSSIER', '000000');

              if (PGCodePredefini = 'DOS') then
                T.PutValue('PCR_NODOSSIER', PgRendNoDossier());
            end;
          end;
        end;
        T_Cumul.Dupliquer(Tob_GestAssoc, TRUE, TRUE, FALSE);
        T_Cumul.InsertDB(nil, False);
        TOB_GestAssoc.free;
        T_Cumul.free;

//PT38
//DEB PT49
        if (PGCodePredefini <> 'DOS') then
        begin
            for i := 0 to TOB_GestionSpec.Detail.Count - 1 do
            begin
              if ((TOB_GestionSpec.Detail[i].GetValue('PKC_NATURERUB')) = AncValNat) and
                ((TOB_GestionSpec.Detail[i].GetValue('PKC_RUBRIQUE')) = AncValRub) then
              begin
                T := TOB_GestionSpec.Detail[i];
                if (T <> nil) then
                begin
                  T.PutValue('PKC_RUBRIQUE', PGCodeDupliquer);
                  Cumulpaie := T.GetValue('PKC_CUMULPAIE');
                end;
              end;
            end;
            TCumulSpec.Dupliquer(TOB_GestionSpec, TRUE, TRUE, FALSE);
            TCumulSpec.InsertDB(nil, False);
        end;
        //else
        //   PGIBox('Personnalisation non dupliquée');
//FIN PT49
        FreeAndNil(TOB_GestionSpec);
        FreeAndNil(TCumulSpec);
//FIN PT38

//Duplication de la Ventilation analytique
        PredExist := '';
        st := 'SELECT *' +
          ' FROM VENTIREMPAIE WHERE' +
          ' ##PVS_PREDEFINI## PVS_RUBRIQUE="' + AncValRub + '"';
        Q := OpenSql(st, TRUE);
        if (not Q.eof) then
        begin
          TOB_GestAssoc := TOB.Create('La ventilation comptable', nil, -1);
          TOB_GestAssoc.LoadDetailDB('VENTIREMPAIE', '', '', Q, FALSE);
          ferme(Q);
//Suppression des élements comptables non conservés pour ne pas créer de doublon
          if (TOB_GestAssoc.FindFirst(['PVS_PREDEFINI'], ['CEG'], False) <> nil) then
            PredExist := 'CEG';
          if (TOB_GestAssoc.FindFirst(['PVS_PREDEFINI'], ['STD'], False) <> nil) then
            PredExist := PredExist + ';STD';
          if (TOB_GestAssoc.FindFirst(['PVS_PREDEFINI'], ['DOS'], False) <> nil) then
            PredExist := PredExist + ';DOS';
          T := TOB_GestAssoc.FindFirst([''], [''], False);
          while (T <> nil) do
          begin
            if (PGCodePredefini <> 'CEG') and
              (T.GetValue('PVS_PREDEFINI') <> PGCodePredefini) and
              (Pos(PGCodePredefini, PredExist) > 0) then
              T.Free;
            if (Pos('DOS', PredExist) = 0) and (PGCodePredefini = 'DOS') and
              (Pos('CEG', PredExist) > 0) and
              (Pos('STD', PredExist) > 0) then
              if (T.GetValue('PVS_PREDEFINI') = 'CEG') then
                T.Free;
            T := TOB_GestAssoc.FindNext([''], [''], False);
          end;
//Duplication des éléments restants
          T_Ventil := TOB.Create('VENTIREMPAIE', nil, -1);
          for i := 0 to TOB_GestAssoc.Detail.Count - 1 do
          begin
            T := TOB_GestAssoc.Detail[i];
            T.PutValue('PVS_RUBRIQUE', PGCodeDupliquer);
            T.PutValue('PVS_PREDEFINI', PGCodePredefini);
            if (PGCodePredefini = 'DOS') then
              T.PutValue('PVS_NODOSSIER', PgRendNoDossier)
            else
              T.PutValue('PVS_NODOSSIER', '000000');
          end;
          T_Ventil.Dupliquer(Tob_GestAssoc, TRUE, TRUE, FALSE);
          if (T_Ventil.detail.count > 0) then
            T_Ventil.InsertDB(nil, False);
          TOB_GestAssoc.free;
          T_Ventil.free;
        end
        else
          Ferme(Q);
        CommitTrans;
      except
        Rollback;
        PGIBox('Une erreur est survenue lors de la duplication des éléments' +
          ' associés à la rubrique.', Ecran.caption);
      end;
      SetField('PRM_THEMEREM', ChampBug);
      SetField('PRM_IMPRIMABLE', ChampBug1);
      SetField('PRM_BASEIMPRIMABLE', ChampBug2);
      SetField('PRM_TAUXIMPRIMABLE', ChampBug3);
      SetField('PRM_CODECALCUL', ChampBug4);
      SetField('PRM_PREDEFINI', PGCodePredefini);
      AccesFicheDupliquer(TFFiche(Ecran), PGCodePredefini, NoDossier,
        LectureSeule);
      SetField('PRM_NODOSSIER', NoDossier);
      SetControlEnabled('BInsert', True);
      SetControlEnabled('PRM_PREDEFINI', False);
      SetControlEnabled('PRM_RUBRIQUE', False);
      SetControlEnabled('BDUPLIQUER', True);
      PgUpdateDateAgl(GetField('PRM_RUBRIQUE'));
//PT43
      if Assigned(Trace) then FreeAndNil(Trace);
      Trace := TStringList.Create; //PT 
      st := 'Duplication de la rubrique '+AncValRub;
      Trace.add (st);
      st := 'Création de la rubrique '+ GetField('PCT_RUBRIQUE');
      Trace.add (st);
      EnDupl := 'OUI';
      IsDifferent(dernierecreate, PrefixeToTable(TFFiche(Ecran).TableName), TFFiche(Ecran).CodeName, TFFiche(Ecran).LibelleName, Trace, TFFiche(Ecran));
      FreeAndNil(Trace);
      EnDupl := 'NON';
//Fin PT43
      TFFiche(Ecran).Bouge(nbPost); //Force enregistrement
{$IFNDEF EAGLCLIENT}
//PT36      ChargementTablette(TFFiche(Ecran).TableName, '');
{$ELSE}
      ChargementTablette(TableToPrefixe(TFFiche(Ecran).TableName), '');
{$ENDIF}
      PgiInfo('Duplication effectuée. Seule la ventilation analytique n''est' +
        ' pas reprise.', Ecran.caption);
    end
    else
      HShowMessage('5;Rémunération :;La duplication est impossible, la' +
        ' rubrique existe déjà.;W;O;O;O;;;', '', '');
  end;
  mode := '';
end;

//============================================================================================//
//                                 PROCEDURE On Delete Record
//============================================================================================//
{ PT34 Refonte de la proc. }

procedure TOM_Remuneration.OnDeleteRecord;
var
  i: integer;
begin
  inherited;
  NomChamp[3] := '';
  NomChamp[4] := '';
  ValChamp[3] := '';
  ValChamp[4] := '';

//Recherche sur la table Histobulletin
  NomChamp[1] := 'PHB_NATURERUB';
  NomChamp[2] := 'PHB_RUBRIQUE';
  ValChamp[1] := 'AAA';
  ValChamp[2] := GetField('PRM_RUBRIQUE');
  ExisteCod := RechEnrAssocier('HISTOBULLETIN', NomChamp, Valchamp);
  if (ExisteCod = TRUE) then
  begin
    LastError := 1;
    LastErrorMsg := 'Attention! Certains bulletins de paie utilisent cette' +
      ' rubrique#13#10Vous ne pouvez la supprimer!';
    exit;
  end;

//Recherche sur les profils
  NomChamp[1] := 'PPM_NATURERUB';
  NomChamp[2] := '##PPM_PREDEFINI## PPM_RUBRIQUE';
  ExisteCod := RechEnrAssocier('PROFILRUB', NomChamp, Valchamp);
  if (ExisteCod = TRUE) then
  begin
    LastError := 1;
    PgiBox('Suppression impossible. Des profils de paie utilisent cette' +
      ' rubrique.', Ecran.caption);
    Exit;
  end;

//Recherche enregistrement associé multidossier
  if ((V_PGI.ModePCL = '1') and (GetField('PRM_PREDEFINI') <> 'DOS')) then {PT34}
  begin
    NomChamp[2] := 'PPM_RUBRIQUE';
    ExisteCod := RechEnrAssocier('PROFILRUB', NomChamp, Valchamp);
    if (ExisteCod = TRUE) then
    begin
      LastError := 1;
      PgiBox('Suppression impossible. Des profils de paie d''autres dossiers' +
        ' utilisent cette rubrique.', Ecran.caption);
      Exit;
    end;
  end;

//Recherche sur les motifs d'absence
// Recherche en motif qui gère des heures
  NomChamp[1] := '##PMA_PREDEFINI## PMA_RUBRIQUE';
  ValChamp[1] := GetField('PRM_RUBRIQUE');
  NomChamp[2] := '';
  ValChamp[2] := '';
  ExisteCod := RechEnrAssocier('MOTIFABSENCE', NomChamp, Valchamp);
  if (not ExisteCod) then
  begin // Recherche en motif qui gère des jours
    NomChamp[1] := '##PMA_PREDEFINI## PMA_RUBRIQUEJ';
    ValChamp[1] := GetField('PRM_RUBRIQUE');
    ExisteCod := RechEnrAssocier('MOTIFABSENCE', NomChamp, Valchamp);
  end;
  if (ExisteCod = TRUE) then
  begin
    LastError := 1;
    PgiBox('Suppression impossible. Des motifs d''absence utilisent cette' +
      ' rubrique.', Ecran.caption);
    Exit;
  end;

//Recherche enregistrement associé multidossier
  if ((V_PGI.ModePCL = '1') and (GetField('PRM_PREDEFINI') <> 'DOS')) then {PT34}
  begin
    NomChamp[1] := 'PMA_RUBRIQUE';
    ValChamp[1] := GetField('PRM_RUBRIQUE');
    ExisteCod := RechEnrAssocier('MOTIFABSENCE', NomChamp, Valchamp);
    if (not ExisteCod) then
    begin // Recherche en motif qui gère des jours
      NomChamp[1] := 'PMA_RUBRIQUEJ';
      ExisteCod := RechEnrAssocier('MOTIFABSENCE', NomChamp, Valchamp);
    end;
    if (ExisteCod = TRUE) then
    begin
      LastError := 1;
      PgiBox('Suppression impossible. Des motifs d''absence d''autres' +
        ' dossiers utilisent cette rubrique.', Ecran.caption);
      Exit;
    end;
  end;

//Recherche sur les masques de saisie
  for i := 1 to 4 do
  begin
    NomChamp[i] := '';
    ValChamp[i] := '';
  end;
  ValChamp[1] := GetField('PRM_RUBRIQUE');
  for i := 1 to 7 do
  begin
    NomChamp[1] := 'PMR_COL' + IntToStr(i);
    ExisteCod := RechEnrAssocier('MASQUESAISRUB', NomChamp, Valchamp);
    if (ExisteCod = TRUE) then
    begin
      LastError := 1;
      PgiBox('Suppression impossible. Des masques de saisie utilisent cette' +
        ' rubrique.', Ecran.caption);
      exit;
    end;
  end;

  if (ExisteCod = FALSE) then
  begin
    ValChamp[1] := 'AAA';
    ValChamp[2] := GetField('PRM_RUBRIQUE');
    ExecuteSQL('DELETE FROM CUMULRUBRIQUE WHERE' +
      ' PCR_NATURERUB="' + ValChamp[1] + '" AND' +
      ' ##PCR_PREDEFINI## PCR_RUBRIQUE="' + ValChamp[2] + '"');

//PT38
    ExecuteSQL('DELETE FROM CUMULRUBDOSSIER WHERE' +
      ' PKC_NATURERUB="' + ValChamp[1] + '" AND' +
      ' PKC_RUBRIQUE="' + ValChamp[2] + '"');
//FIN PT38

    ExecuteSQL('DELETE FROM VENTIREMPAIE WHERE' +
      ' ##PVS_PREDEFINI## PVS_RUBRIQUE="' + ValChamp[2] + '"');
{$IFNDEF EAGLCLIENT}
    ChargementTablette(TFFiche(Ecran).TableName, '');
{$ELSE}
    ChargementTablette(TableToPrefixe(TFFiche(Ecran).TableName), '');
{$ENDIF}
{A AJOUTER L'ANALYTIQUE COMPTABILITE
DELETE FROM VENTIL WHERE V_NATURE="RC" AND V_IDENTIFIANT="Rubrique"
}
//  RR pour remuneration
//PT43
    if Assigned(Trace) then FreeAndNil(Trace);
    Trace := TStringList.Create;
//PT37
    Trace.Add('SUPPRESSION REMUNERATION ' + GetField('PRM_RUBRIQUE') + ' ' +
      GetField('PRM_LIBELLE'));
    CreeJnalEvt('003', '080', 'OK', nil, nil, Trace);
//PT43   Trace.free;   PT40 déplacé dans on_close
    FreeAndNil(Trace);
  end;
end;

//============================================================================================//
//                                 PROCEDURE On Update Record
//============================================================================================//

procedure TOM_Remuneration.OnUpdateRecord;
var
  Onglet: TPageControl;
  cumul, theme, rubrique, TypeCalcul, mess, libelle: string;
{$IFDEF EAGLCLIENT}
  imprim: TCheckBox;
{$ELSE}
  imprim: TDBCheckBox;
{$ENDIF}
  ExistCod: Boolean;
  ordre, i: integer;
  TOBGestAssoc: Tob;
  sens: string;
  st, Predef: string;
  Q: TQuery;
begin
  inherited;
  LeStatut := DS.State; //PT46
  OnFerme := False;
  if (DS.State in [dsInsert]) then
    DerniereCreate := GetField('PRM_RUBRIQUE')
  else
    if (DerniereCreate = GetField('PRM_RUBRIQUE')) then OnFerme := True; // le bug arrive on se casse !!!
  if mode = 'DUPLICATION' then exit;
  if Ecran is TFFiche then TFFiche(Ecran).Retour := '0';

  //   Affect du N° d'ordre de presentation en fonction des cumuls
  //   regle :
  //   1 : rubriques associès au cumul 01
  //   2 : rubriques associès au cumul 02 et non au 01
  //   3 : rubriques de cotisations à initialiser dans UTOMCOtisation
  //   4 => 6 : rubriques associès au cumul 09 et non au 01 et 02
  //   5 => 7 : rubriques associès au cumul 10 et non au 01,02 et 09
  //   6 => 9: rubriques associès à aucun des cumuls précédents

  TOBGestAssoc := TOB.Create('La Remunération originale', nil, -1);
  st := 'SELECT * FROM CUMULRUBRIQUE WHERE PCR_NATURERUB="AAA" AND ##PCR_PREDEFINI## PCR_RUBRIQUE="' + Getfield('PRM_RUBRIQUE') + '"'; //**//
  Q := OpenSql(st, TRUE);
  TOBGestAssoc.LoadDetailDB('CUMULRUBRIQUE', '', '', Q, FALSE);
  Ferme(Q);
  ordre := 0;
  sens := '';
  for i := 0 to TOBGestAssoc.Detail.Count - 1 do
  begin
    Cumul := TOBGestAssoc.Detail[i].GetValue('PCR_CUMULPAIE');
    if (Cumul = '01') then
    begin
      ordre := 1;
      sens := TOBGestAssoc.Detail[i].GetValue('PCR_SENS');
    end;
    if (Cumul = '02') and (ordre = 0) then
    begin
      ordre := 2;
      sens := TOBGestAssoc.Detail[i].GetValue('PCR_SENS');
    end;
    if (Cumul = '09') and (ordre = 0) then
    begin
      ordre := 6; { PT25 }
      sens := TOBGestAssoc.Detail[i].GetValue('PCR_SENS');
    end;
    if (Cumul = '10') and (ordre = 0) then
    begin
      ordre := 7; { PT25 }
      sens := TOBGestAssoc.Detail[i].GetValue('PCR_SENS');
    end;
  end;
  if (Ordre = 0) then ordre := 9; { PT25 }

  if sens = '+' then sens := 'P';
  if sens = '-' then sens := 'M';
  if sens = '' then sens := 'S';
  Setfield('PRM_ORDREETAT', ordre);
  Setfield('PRM_SENSBUL', sens);
  TOBGestAssoc.free;


  if (DS.State = dsinsert) then
  begin
    if (GetField('PRM_PREDEFINI') <> 'DOS') then
      SetField('PRM_NODOSSIER', '000000')
    else
      // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
      SetField('PRM_NODOSSIER', PgRendNoDossier());
  end;


  LastError := ControlChampValeur; //control des champs valeurs
  if LastError <> 0 then
  begin
    LastErrorMsg := '';
    exit;
  end; //PT14


  //Obligation de renseigner les champs Theme et Type de Profil
  Onglet := TPageControl(GetControl('Pages'));
  theme := Getfield('PRM_THEMEREM');
  TypeCalcul := Getfield('PRM_CODECALCUL');
  rubrique := Getfield('PRM_RUBRIQUE');
  libelle := Getfield('PRM_LIBELLE');
  if (Onglet.ActivePage = Onglet.Pages[0]) and (rubrique <> '') and (libelle <> '') then
  begin
    if theme = '' then
    begin
      mess := 'le thème de la rémunération'; //PT23
      Onglet.ActivePage := Onglet.Pages[0];
      SetFocusControl('PRM_THEMEREM');
    end;
    if (TypeCalcul = '') and (theme <> '') and (rubrique <> '') then
    begin
      mess := 'le type de calcul';
      Onglet.ActivePage := Onglet.Pages[1];
      SetFocusControl('PRM_CODECALCUL');
    end;
  end;
  if (Onglet.ActivePage = Onglet.Pages[1]) then
  begin
    if (rubrique <> '') and (libelle <> '') then
      if (theme = '') and (TypeCalcul <> '') then
      begin
        mess := 'le thème de la rémunération'; //PT23
        Onglet.ActivePage := Onglet.Pages[0];
        SetFocusControl('PRM_THEMEREM');
      end;
    if (TypeCalcul = '') then
    begin
      mess := 'le type de calcul';
      Onglet.ActivePage := Onglet.Pages[1];
      SetFocusControl('PRM_CODECALCUL');
    end;
  end;

  if (mess <> '') then
    LastError := 2;
  LastErrorMsg := 'Veuillez renseigner ' + mess + '';

  // Recherche si Gestion Associe renseigné
  NomChamp[1] := 'PCR_NATURERUB';
  NomChamp[2] := '##PCR_PREDEFINI## PCR_RUBRIQUE'; //PT12-2
  ValChamp[1] := 'AAA';
  ValChamp[2] := GetField('PRM_RUBRIQUE');
  NomChamp[3] := '';
  NomChamp[4] := ''; //PT4
  ValChamp[3] := '';
  ValChamp[4] := '';
  ExistCod := RechEnrAssocier('CUMULRUBRIQUE', NomChamp, ValChamp);
  //Gestion du CheckBox "Rubrique imprimable...."
// A Cocher si Gestion Associé renseignée...
{$IFDEF EAGLCLIENT}
  imprim := TCheckBox(GetControl('PRM_IMPRIMABLE'));
{$ELSE}
  imprim := TDBCheckBox(GetControl('PRM_IMPRIMABLE'));
{$ENDIF}
  if (mess = '') and ((rubrique <> '0000') or (rubrique <> '')) and (libelle <> '') and (imprim <> nil) then
  begin
    Onglet.ActivePage := Onglet.Pages[0];
    { On change d'avis ==> possiblilté d'avoir une remuneration non imprimable qui alimente des cumuls 07/09/00
        if (imprim.Checked=True) and (ExistCod=False) then
            begin   LastError:=4;    SetFocusControl('PRM_IMPRIMABLE');
                   LastErrorMsg:='Veuillez #13#10    alimenter vos cumuls,#13#10 ou #13#10    décocher l''option " rubrique imprimable sur le bulletin "';
              end;}
    if (imprim.Checked = False) and (ExistCod = False) then
      HShowMessage('5;Rémunération :;Vous ne gérez aucun cumul !;E;O;O;O;;;', '', '');
    { if (ExistCod=True) and (imprim.checked=False) then  SetField('PRM_IMPRIMABLE','X'); }

  end;
  Predef := GetField('PRM_PREDEFINI');
  if (Predef <> 'CEG') and (Predef <> 'DOS') and (Predef <> 'STD') then
  begin
    LastError := 1;
    LastErrorMsg := 'Vous devez renseigner le champ prédéfini';
    SetFocusControl('PRM_PREDEFINI');
  end;
  // PT15  : 08/08/2003 PH V595 FQ 10757  si on a renseigné une date de début de validité, obliger la saisie d'une date de fin de validité
  // PT21 Rajout de la condition <> '00' pour prendr en compte exercice social
  if ((GetField('PRM_DU') <> '') and (GetField('PRM_DU') <> '00')) and (GetField('PRM_AU') = '') then
  begin
    LastError := 1;
    LastErrorMsg := 'Vous devez renseigner le champ validité au';
    SetFocusControl('PRM_AU');
  end;
  if (GetField('PRM_DU') = '') and (GetField('PRM_AU') <> '') then SetField('PRM_AU', '');
  // FIN PT15
  { DEB PT13 Validation inutile si pas acces modification gestion associée }
  if (CEG = False) and (Predef = 'CEG') then SetControlChecked('CBACCESFICHE', False);
  if (STD = False) and (Predef = 'STD') then SetControlChecked('CBACCESFICHE', False);
  if (DOS = False) and (Predef = 'DOS') then SetControlChecked('CBACCESFICHE', False);
  if LastError = 0 then
  begin
    PgUpdateDateAgl(GetField('PRM_RUBRIQUE')); //PT20 déplacé dans procedure
  end;
  { FIN PT13 }

// d PT41
  if (copy(GetField('PRM_DADS'), 1, 1) = 'X') then
  begin
    if (copy(GetField('PRM_DADS'), 2, 3) = '') then
    begin
      LastError := 1;
      LastErrorMsg := 'Vous devez renseigner le type de la rubrique';
      SetFocusControl('TYPRUB');
    end;
  end;
// f  PT41
end;

//============================================================================================//
//                                 PROCEDURE On Change Field
//============================================================================================//

procedure TOM_Remuneration.OnChangeField(F: TField);
var
  TypeChamp, Champ, vide, ValidDu, Rubrique, TempRub, Mes, Pred, libelle: string;
{$IFDEF EAGLCLIENT}
  ValidAu, Mois1, Mois2, Mois3, Mois4: THValComboBox;
  ControlChamp: THEdit;
{$ELSE}
//  Mois1, Mois2, Mois3, Mois4: THDBValComboBox;
  ControlChamp: THDBEdit;
{$ENDIF}
  iCode: integer;
  OKRub: boolean;

begin
  inherited;
  //Libelle := GetField('PRM_BASEREM');
  if mode = 'DUPLICATION' then exit;
  Champ := '';
  if (F.FieldName = 'PRM_RUBRIQUE') then
  begin
    Rubrique := Getfield('PRM_RUBRIQUE');
    if Rubrique = '' then exit;
    if ((isnumeric(Rubrique)) and (Rubrique <> '    ')) then
    begin
      iCode := strtoint(trim(Rubrique));
      TempRub := ColleZeroDevant(iCode, 4);
      if (DS.State = dsinsert) and (TempRub <> '') and (GetField('PRM_PREDEFINI') <> '') then
      begin
        OKRub := TestRubrique(GetField('PRM_PREDEFINI'), TempRub, 0);
        if (OkRub = False) or (rubrique = '0000') then
        begin
          Mes := MesTestRubrique('AAA', GetField('PRM_PREDEFINI'), 0);
          HShowMessage('2;Code Erroné: ' + TempRub + ' ;' + Mes + ';W;O;O;;;', '', '');
          TempRub := '';
        end;
      end;
      if TempRub <> Rubrique then
      begin
        SetField('PRM_RUBRIQUE', TempRub);
        SetFocusControl('PRM_RUBRIQUE');
      end;
    end;
  end;

  if (F.FieldName = 'PRM_PREDEFINI') and (DS.State = dsinsert) then
  begin
    Pred := GetField('PRM_PREDEFINI');
    Rubrique := (GetField('PRM_RUBRIQUE'));
    if Pred = '' then exit;
    AccesPredefini('TOUS', CEG, STD, DOS);
    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer de rubrique prédéfinie CEGID', 'Accès refusé'); // PT28
      Pred := 'DOS';
      SetControlProperty('PRM_PREDEFINI', 'Value', Pred);
    end;
    if (Pred = 'STD') and (STD = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer de rubrique prédéfinie Standard', 'Accès refusé'); // PT28
      Pred := 'DOS';
      SetControlProperty('PRM_PREDEFINI', 'Value', Pred);
    end;
// d PT44
// Ces contrôles ne sont visibles que pour les rubriques CEGID
    if (Pred = 'CEG') and (CEG = TRUE) then
    begin
      SetControlVisible('CBIJSS', True);
      SetControlVisible('TYPRUB', True);
    end;
//f PT44

    if (rubrique <> '') and (Pred <> '') then
    begin
      OKRub := TestRubrique(pred, rubrique, 0);
      if (OkRub = False) or (rubrique = '0000') then
      begin
        Mes := MesTestRubrique('AAA', pred, 0);
        HShowMessage('2;Code Erroné: ' + Rubrique + ' ;' + Mes + ';W;O;O;;;', '', '');
        SetField('PRM_RUBRIQUE', vide);
        if Pred <> GetField('PRM_PREDEFINI') then SetField('PRM_PREDEFINI', pred);
        SetFocusControl('PRM_RUBRIQUE');
        exit;
      end;
    end;
    if Pred <> GetField('PRM_PREDEFINI') then SetField('PRM_PREDEFINI', pred);
  end;



  if (F.FieldName = 'PRM_TYPEBASE') then
  begin
    TypeChamp := GetField('PRM_TYPEBASE');
    Champ := 'PRM_BASEREM';
    libelle := 'LBLBASEREM';
  end;

  if (F.FieldName = 'PRM_TYPETAUX') then
  begin
    TypeChamp := GetField('PRM_TYPETAUX');
    Champ := 'PRM_TAUXREM';
    libelle := 'LBLTAUXREM';
  end;

  if (F.FieldName = 'PRM_TYPECOEFF') then
  begin
    TypeChamp := GetField('PRM_TYPECOEFF');
    Champ := 'PRM_COEFFREM';
    libelle := 'LBLCOEFFREM';
  end;

  if (F.FieldName = 'PRM_TYPEMONTANT') then
  begin
    TypeChamp := GetField('PRM_TYPEMONTANT');
    Champ := 'PRM_MONTANT';
    libelle := 'LBLMONTANTREM';
  end;
  if (F.FieldName = 'PRM_TYPEMINI') then
  begin
    TypeChamp := GetField('PRM_TYPEMINI');
    Champ := 'PRM_VALEURMINI';
    libelle := 'LBLMINI';
  end;
  if (F.FieldName = 'PRM_TYPEMAXI') then
  begin
    TypeChamp := GetField('PRM_TYPEMAXI');
    Champ := 'PRM_VALEURMAXI';
    libelle := 'LBLMAXI';
  end;

  if Champ <> '' then AffecteLookupChampRem(Champ, TypeChamp, libelle);

  // En fonction du code choisi on réinit et rend enabled ou non les Champs associés ou non

  if (F.Fieldname = 'PRM_CODECALCUL') then
  begin
    SetControlEnabled('PRM_TYPEBASE', TRUE);
    SetControlEnabled('PRM_BASEREM', TRUE);
    SetControlEnabled('PRM_BASEMTQTE', TRUE);
    //SetControlChecked('PRM_BASEIMPRIMABLE',True);
    SetControlEnabled('PRM_BASEIMPRIMABLE', TRUE);
    SetControlEnabled('PRM_TYPETAUX', TRUE);
    SetControlEnabled('PRM_TAUXREM', TRUE);
    SetControlEnabled('PRM_TAUXMTQTE', TRUE);
    SetControlEnabled('PRM_TAUXIMPRIMABLE', TRUE);
    //SetControlChecked('PRM_TAUXIMPRIMABLE',TRUE);
    SetControlEnabled('PRM_TYPECOEFF', TRUE);
    SetControlEnabled('PRM_COEFFREM', TRUE);
    SetControlEnabled('PRM_COEFFMTQTE', TRUE);
    SetControlEnabled('PRM_COEFFIMPRIM', TRUE);
    // SetControlChecked('PRM_COEFFIMPRIM',TRUE);
    SetControlEnabled('PRM_TYPEMONTANT', TRUE);
    SetControlEnabled('PRM_MONTANT', TRUE);
    SetControlEnabled('PRM_MONTANTQTE', TRUE);

    SetControlEnabled('PRM_DECBASE', TRUE);
    SetControlEnabled('PRM_DECTAUX', TRUE);
    SetControlEnabled('PRM_DECCOEFF', TRUE);
    SetControlEnabled('PRM_DECMONTANT', TRUE);


    vide := '';

    if (GetField('PRM_CODECALCUL') = '01') then //MONTANT
    begin
      // PT16  : 01/09/2003 PH V_421 Non RAZ type de base et base dans le cas d'un calcul de type montant
      {       SetField('PRM_TYPEBASE',vide);
             SetField('PRM_BASEREM',vide);}
      if GetField('PRM_TYPETAUX') <> vide then SetField('PRM_TYPETAUX', vide); //PT27
      SetControlEnabled('PRM_TYPETAUX', FALSE);
      if GetField('PRM_TAUXREM') <> vide then SetField('PRM_TAUXREM', vide); //PT27
      SetControlEnabled('PRM_TAUXREM', FALSE);
      { PT8 PH 18/01/2002 Taux non géré donc une quantité car si Le montant est une quantité
           alors tous les qualifiants du code calcul doivent être en qté pour que le montant
           de la rémunération ne soit pas un montant mais une qté ==> Pb conversion euro qui
           analyse tous les champs pour déterminer si le montant est bien un montant
      }
      //       SetField('PRM_TAUXMTQTE','-');  On laisse la valeur tq
      SetControlEnabled('PRM_TAUXMTQTE', FALSE);
      SetControlChecked('PRM_TAUXIMPRIMABLE', FALSE);
      SetControlEnabled('PRM_TAUXIMPRIMABLE', FALSE);
      SetControlEnabled('PRM_DECTAUX', FALSE);
      if GetField('PRM_DECTAUX') <> '2' then SetField('PRM_DECTAUX', 2); //PT27// PT11  : 18/12/2002 PH V591 PORTAGECWAS controle des types de champs sur les setfield
      if GetField('PRM_TYPECOEFF') <> vide then SetField('PRM_TYPECOEFF', vide); //PT27
      SetControlEnabled('PRM_TYPECOEFF', FALSE);
      if GetField('PRM_COEFFREM') <> vide then SetField('PRM_COEFFREM', vide); //PT27
      SetControlEnabled('PRM_COEFFREM', FALSE);
      SetControlChecked('PRM_COEFFIMPRIM', FALSE);
      SetControlEnabled('PRM_COEFFIMPRIM', FALSE);
      if GetField('PRM_COEFFMTQTE') <> '-' then SetField('PRM_COEFFMTQTE', '-'); //PT27
      SetControlEnabled('PRM_COEFFMTQTE', FALSE);
      if GetField('PRM_DECCOEFF') <> 2 then SetField('PRM_DECCOEFF', 2); //PT27 // PT11  : 18/12/2002 PH V591 PORTAGECWAS controle des types de champs sur les setfield
      SetControlEnabled('PRM_DECCOEFF', FALSE);
    end;

    //BASE TAUX COEFF
    if (GetField('PRM_CODECALCUL') = '02') or (GetField('PRM_CODECALCUL') = '03') or (GetField('PRM_CODECALCUL') = '06') or (GetField('PRM_CODECALCUL') = '07') then
    begin
      if GetField('PRM_TYPEMONTANT') <> vide then SetField('PRM_TYPEMONTANT', vide); //PT27
      SetControlEnabled('PRM_TYPEMONTANT', FALSE);
      if GetField('PRM_MONTANT') <> vide then SetField('PRM_MONTANT', vide); //PT27
      SetControlEnabled('PRM_MONTANT', FALSE);
      if GetField('PRM_MONTANTMTQTE') <> 'X' then SetField('PRM_MONTANTMTQTE', 'X'); //PT27
      SetControlEnabled('PRM_MONTANTQTE', FALSE);
      // PT2 : 10/08/2001 : V547  PH  Rend tjrs accessible le champ PRM_DECMONTANT
      //          SetField('PRM_DECMONTANT','2');
      if (Pred = 'CEG') and (CEG = FALSE) then
        SetControlEnabled('PRM_DECMONTANT', FALSE);
      if (Pred = 'STD') and (STD = FALSE) then
        SetControlEnabled('PRM_DECMONTANT', FALSE);

    end;

    // BASE TAUX
    if (GetField('PRM_CODECALCUL') = '04') or (GetField('PRM_CODECALCUL') = '05') then
    begin
      if GetField('PRM_TYPECOEFF') <> vide then SetField('PRM_TYPECOEFF', vide); //PT27
      SetControlEnabled('PRM_TYPECOEFF', FALSE);
      if GetField('PRM_COEFFREM') <> vide then SetField('PRM_COEFFREM', vide); //PT27
      SetControlEnabled('PRM_COEFFREM', FALSE);
      if GetField('PRM_COEFFMTQTE') <> '-' then SetField('PRM_COEFFMTQTE', '-'); //PT27
      SetControlEnabled('PRM_COEFFMTQTE', FALSE);
      SetControlChecked('PRM_COEFFIMPRIM', FALSE);
      SetControlEnabled('PRM_COEFFIMPRIM', FALSE);
      SetControlEnabled('PRM_DECCOEFF', FALSE);
      // PT11  : 18/12/2002 PH V591 PORTAGECWAS controle des types de champs sur les setfield
      if GetField('PRM_DECCOEFF') <> 2 then SetField('PRM_DECCOEFF', 2); //PT27
      if GetField('PRM_TYPEMONTANT') <> vide then SetField('PRM_TYPEMONTANT', vide); //PT27
      SetControlEnabled('PRM_TYPEMONTANT', FALSE);
      if GetField('PRM_MONTANT') <> vide then SetField('PRM_MONTANT', vide); //PT27
      SetControlEnabled('PRM_MONTANT', FALSE);
      if GetField('PRM_MONTANTMTQTE') <> 'X' then SetField('PRM_MONTANTMTQTE', 'X'); //PT27
      SetControlEnabled('PRM_MONTANTQTE', FALSE);
      //PT2 : 10/08/2001 : V547  PH  Rend tjrs accessible le champ PRM_DECMONTANT
      //     SetControlEnabled('PRM_DECMONTANT',FALSE);    SetField('PRM_DECMONTANT','2');
      if (Pred = 'CEG') and (CEG = FALSE) then
        SetControlEnabled('PRM_DECMONTANT', FALSE);
      if (Pred = 'STD') and (STD = FALSE) then
        SetControlEnabled('PRM_DECMONTANT', FALSE);
    end;

    //BASE COEFF
    if (GetField('PRM_CODECALCUL') = '08') then
    begin
      if GetField('PRM_TYPETAUX') <> vide then SetField('PRM_TYPETAUX', vide); //PT27
      SetControlEnabled('PRM_TYPETAUX', FALSE);
      if GetField('PRM_TAUXREM') <> vide then SetField('PRM_TAUXREM', vide); //PT27
      SetControlEnabled('PRM_TAUXREM', FALSE);
      SetControlChecked('PRM_TAUXIMPRIMABLE', FALSE);
      SetControlEnabled('PRM_TAUXIMPRIMABLE', FALSE);
      if GetField('PRM_TAUXMTQTE') <> 'X' then SetField('PRM_TAUXMTQTE', 'X'); //PT27
      SetControlEnabled('PRM_TAUXMTQTE', FALSE);
      SetControlEnabled('PRM_DECTAUX', FALSE);
      // PT11  : 18/12/2002 PH V591 PORTAGECWAS controle des types de champs sur les setfield
      if GetField('PRM_DECTAUX') <> 2 then SetField('PRM_DECTAUX', 2); //PT27
      if GetField('PRM_TYPEMONTANT') <> vide then SetField('PRM_TYPEMONTANT', vide); //PT27
      SetControlEnabled('PRM_TYPEMONTANT', FALSE);
      if GetField('PRM_MONTANT') <> vide then SetField('PRM_MONTANT', vide); //PT27
      SetControlEnabled('PRM_MONTANT', FALSE);
      if GetField('PRM_MONTANTMTQTE') <> 'X' then SetField('PRM_MONTANTMTQTE', 'X'); //PT27
      SetControlEnabled('PRM_MONTANTQTE', FALSE);
      // PT2 : 10/08/2001 : V547  PH  Rend tjrs accessible le champ PRM_DECMONTANT
      //     SetControlEnabled('PRM_DECMONTANT',FALSE);SetField('PRM_DECMONTANT','2');
      if (Pred = 'CEG') and (CEG = FALSE) then
        SetControlEnabled('PRM_DECMONTANT', FALSE);
      if (Pred = 'STD') and (STD = FALSE) then
        SetControlEnabled('PRM_DECMONTANT', FALSE);

    end;

  end;

  //Test Valeur Saisie par rapport au tablettes associées ElipsisButton

  if (F.FieldName = 'PRM_BASEREM') then
  begin
{$IFDEF EAGLCLIENT}
    ControlChamp := THEdit(GetControl('PRM_BASEREM'));
{$ELSE}
    ControlChamp := THDBEdit(GetControl('PRM_BASEREM'));
{$ENDIF}
    if ControlChamp <> nil then
    begin
      LastError := TestValLookupChamp(ControlChamp, GetField('PRM_BASEREM'));
      if LastError = 1 then
      begin
        LastErrorMsg := '"' + GetField('PRM_BASEREM') + '" :#13#10  Veuiller saisir une valeur existante';
        SetField('PRM_TYPEBASE', GetField('PRM_TYPEBASE'));
        SetFocusControl('PRM_TYPEBASE');
        LastError := 1;
      end;
    end;
  end;

  if (F.FieldName = 'PRM_TAUXREM') then
  begin
{$IFDEF EAGLCLIENT}
    ControlChamp := THEdit(GetControl('PRM_TAUXREM'));
{$ELSE}
    ControlChamp := THDBEdit(GetControl('PRM_TAUXREM'));
{$ENDIF}
    if ControlChamp <> nil then
    begin
      LastError := TestValLookupChamp(ControlChamp, GetField('PRM_TAUXREM'));
      if LastError = 1 then
      begin
        LastErrorMsg := '"' + GetField('PRM_TAUXREM') + '" :#13#10  Veuiller saisir une valeur existante';
        SetField('PRM_TYPETAUX', GetField('PRM_TYPETAUX'));
        SetFocusControl('PRM_TYPETAUX');
        LastError := 1;
      end;
    end;
  end;


  if (F.FieldName = 'PRM_COEFFREM') then
  begin
{$IFDEF EAGLCLIENT}
    ControlChamp := THEdit(GetControl('PRM_COEFFREM'));
{$ELSE}
    ControlChamp := THDBEdit(GetControl('PRM_COEFFREM'));
{$ENDIF}
    if ControlChamp <> nil then
    begin
      LastError := TestValLookupChamp(ControlChamp, GetField('PRM_COEFFREM'));
      if LastError = 1 then
      begin
        LastErrorMsg := '"' + GetField('PRM_COEFFREM') + '" :#13#10  Veuiller saisir une valeur existante';
        SetField('PRM_TYPECOEFF', GetField('PRM_TYPECOEFF'));
        SetFocusControl('PRM_TYPECOEFF');
        LastError := 1;
      end;
    end;
  end;

  if (F.FieldName = 'PRM_MONTANT') then
  begin
{$IFDEF EAGLCLIENT}
    ControlChamp := THEdit(GetControl('PRM_MONTANT'));
{$ELSE}
    ControlChamp := THDBEdit(GetControl('PRM_MONTANT'));
{$ENDIF}
    if ControlChamp <> nil then
    begin
      LastError := TestValLookupChamp(ControlChamp, GetField('PRM_MONTANT'));
      if LastError = 1 then
      begin
        LastErrorMsg := '"' + GetField('PRM_MONTANT') + '" :#13#10  Veuiller saisir une valeur existante';
        SetField('PRM_TYPEMONTANT', GetField('PRM_TYPEMONTANT'));
        SetFocusControl('PRM_TYPEMONTANT');
        LastError := 1;
      end;
    end;
  end;

  if (F.FieldName = 'PRM_VALEURMINI') then
  begin
{$IFDEF EAGLCLIENT}
    ControlChamp := THEdit(GetControl('PRM_VALEURMINI'));
{$ELSE}
    ControlChamp := THDBEdit(GetControl('PRM_VALEURMINI'));
{$ENDIF}
    if ControlChamp <> nil then
    begin
      LastError := TestValLookupChamp(ControlChamp, GetField('PRM_VALEURMINI'));
      if LastError = 1 then
      begin
        LastErrorMsg := '"' + GetField('PRM_VALEURMINI') + '" :#13#10  Veuiller saisir une valeur existante';
        SetField('PRM_TYPEMINI', GetField('PRM_TYPEMINI'));
        SetFocusControl('PRM_TYPEMINI');
        LastError := 1;
      end;
    end;
  end;

  if (F.FieldName = 'PRM_VALEURMAXI') then
  begin
{$IFDEF EAGLCLIENT}
    ControlChamp := THEdit(GetControl('PRM_VALEURMAXI'));
{$ELSE}
    ControlChamp := THDBEdit(GetControl('PRM_VALEURMAXI'));
{$ENDIF}
    if ControlChamp <> nil then
    begin
      LastError := TestValLookupChamp(ControlChamp, GetField('PRM_VALEURMAXI'));
      if LastError = 1 then
      begin
        LastErrorMsg := '"' + GetField('PRM_VALEURMAXI') + '" :#13#10  Veuiller saisir une valeur existante';
        SetField('PRM_TYPEMAXI', GetField('PRM_TYPEMAXI'));
        SetFocusControl('PRM_TYPEMAXI');
        LastError := 1;
      end;
    end;
  end;

  if (F.Fieldname = 'PRM_DU') then
  begin
    ValidDu := Getfield('PRM_DU');
    if ValidDu <> '' then
    begin
      SetField('PRM_MOIS1', Vide);
      SetField('PRM_MOIS2', Vide);
      SetField('PRM_MOIS3', Vide);
      SetField('PRM_MOIS4', Vide);
      SetField('PRM_MOIS5', Vide); // PT39
      SetField('PRM_MOIS6', Vide); // PT39
      SetControlEnabled('PRM_MOIS1', False);
      SetControlEnabled('PRM_MOIS2', False);
      SetControlEnabled('PRM_MOIS3', False);
      SetControlEnabled('PRM_MOIS4', False);
      SetControlEnabled('PRM_MOIS5', False); // PT39
      SetControlEnabled('PRM_MOIS6', False); // PT39
    end;
    if (ValidDu = '00') then
    begin
      SetControlProperty('PRM_AU', 'DataType', 'PGVALIDITERUB');
      if GetField('PRM_AU') <> ValidDu then SetField('PRM_AU', ValidDu); //PT32
      SetControlEnabled('PRM_AU', False); //PT32
    end;

    if (ValidDu <> '00') and (ValidDu <> '') then
    begin
      if GetField('PRM_AU') <> vide then SetField('PRM_AU', Vide); //PT32
      SetControlProperty('PRM_AU', 'DataType', 'PGVALIDITEMOISRUB');
      SetControlEnabled('PRM_AU', True); // PT32
    end;

    if (ValidDu = '') then
    begin
      SetControlEnabled('PRM_AU', False);
      SetControlEnabled('PRM_MOIS1', True);
      SetControlEnabled('PRM_MOIS2', True);
      SetControlEnabled('PRM_MOIS3', True);
      SetControlEnabled('PRM_MOIS4', True);
      SetControlEnabled('PRM_MOIS5', True); // PT39
      SetControlEnabled('PRM_MOIS6', True); // PT39
    end;
  end;

  if (F.Fieldname = 'PRM_MOIS1') then
  begin
    if (Getfield('PRM_MOIS1') = '') and (Getfield('PRM_MOIS2') = '') and (Getfield('PRM_MOIS3') = '') and (Getfield('PRM_MOIS4') = '')
      and (Getfield('PRM_MOIS5') = '') and (Getfield('PRM_MOIS6') = '') then SetControlEnabled('PRM_DU', True); // PT39
    if (GetField('PRM_MOIS1') <> '') then SetControlEnabled('PRM_DU', False);
      //if (Getfield('PRM_MOIS1')='') then  ReinitMoisValidite(Mois2,Mois3,Mois4);
  end;

  if (F.Fieldname = 'PRM_MOIS2') then
  begin
    if (Getfield('PRM_MOIS1') = '') and (Getfield('PRM_MOIS2') = '') and (Getfield('PRM_MOIS3') = '') and (Getfield('PRM_MOIS4') = '')
      and (Getfield('PRM_MOIS5') = '') and (Getfield('PRM_MOIS6') = '') then SetControlEnabled('PRM_DU', True); // PT39
    if (GetField('PRM_MOIS2') <> '') then SetControlEnabled('PRM_DU', False);
      //if (Getfield('PRM_MOIS2')='')then   ReinitMoisValidite(Mois1,Mois3,Mois4);
  end;

  if (F.Fieldname = 'PRM_MOIS3') then
  begin
    if (Getfield('PRM_MOIS1') = '') and (Getfield('PRM_MOIS2') = '') and (Getfield('PRM_MOIS3') = '') and (Getfield('PRM_MOIS4') = '')
      and (Getfield('PRM_MOIS5') = '') and (Getfield('PRM_MOIS6') = '') then SetControlEnabled('PRM_DU', True); // PT39
    if (GetField('PRM_MOIS3') <> '') then SetControlEnabled('PRM_DU', False);
      //if (Getfield('PRM_MOIS3')='') then  ReinitMoisValidite(Mois1,Mois2,Mois4)  ;
  end;

  if (F.Fieldname = 'PRM_MOIS4') then
  begin
    if (Getfield('PRM_MOIS1') = '') and (Getfield('PRM_MOIS2') = '') and (Getfield('PRM_MOIS3') = '') and (Getfield('PRM_MOIS4') = '')
      and (Getfield('PRM_MOIS5') = '') and (Getfield('PRM_MOIS6') = '') then SetControlEnabled('PRM_DU', True); // PT39
    if (GetField('PRM_MOIS4') <> '') then SetControlEnabled('PRM_DU', False);
      //if (Getfield('PRM_MOIS4')='') then  ReinitMoisValidite(Mois1,Mois2,Mois3)  ;
  end;
  // DEB PT22
  if (F.Fieldname = 'PRM_THEMEREM') then
  begin
    if GetField('PRM_THEMEREM') <> 'ZZZ' then SetControlProperty('TPRM_ABREGE', 'Caption', 'Nom court')
    else SetControlProperty('TPRM_ABREGE', 'Caption', 'Rémunération associée');
  end;
  // FIN PT22
  if (ds.state in [dsBrowse]) then
  begin
    if LectureSeule then
    begin
      PaieLectureSeule(TFFiche(Ecran), True);
      SetControlEnabled('BDelete', False);
      SetControlEnabled('MEMOPERSO', True); // PT31
    end;
    SetControlEnabled('BInsert', True);
    SetControlEnabled('PRM_PREDEFINI', False);
    SetControlEnabled('PRM_RUBRIQUE', False);
    SetControlEnabled('BDUPLIQUER', True);
  end;
end;


//============================================================================================//
//                                 PROCEDURE Affecte Lookup Champ Rem
//============================================================================================//

procedure TOM_Remuneration.AffecteLookupChampRem(Champ, TypeChamp, libelle: string);
var
  vide: string;
  ControlLib: THLabel;
begin
  //SetField(champ,vide);
  if (TypeChamp = '00') or (TypeChamp = '01') then
  begin
    SetControlProperty(Champ, 'DataType', vide);
    SetControlProperty(Champ, 'ElipsisButton', FALSE);
    SetControlEnabled(champ, False);
    if GetField(champ) <> vide then SetField(champ, vide); //PT27
    ControlLib := THLabel(GetControl(libelle));
    if ControlLib <> nil then ControlLib.caption := '';
  end;
  if (TypeChamp = '08') then
  begin
    SetControlProperty(Champ, 'DataType', vide);
    SetControlProperty(Champ, 'ElipsisButton', FALSE);
    SetControlEnabled(champ, True);
{$IFNDEF EAGLCLIENT}
    if GetField(champ) <> vide then SetField(champ, vide); //PT27
{$ELSE}
    SetFocusControl(Champ);
{$ENDIF}
    ControlLib := THLabel(GetControl(libelle));
    if ControlLib <> nil then ControlLib.caption := '';
  end;
  if (TypeChamp = '02') then
  begin
    SetControlProperty(Champ, 'DataType', 'PGELEMENTNAT');
    SetControlProperty(Champ, 'ElipsisButton', TRUE);
    SetControlEnabled(champ, True);
    ControlLib := THLabel(GetControl(libelle));
    if (ControlLib <> nil) then
      ControlLib.caption := RechDom('PGELEMENTNAT', GetControlText(Champ), FALSE);
  end;
  if (TypeChamp = '03') then
  begin
    SetControlProperty(Champ, 'DataType', 'PGVARIABLE');
    SetControlProperty(Champ, 'ElipsisButton', TRUE);
    SetControlEnabled(champ, True);
    ControlLib := THLabel(GetControl(libelle));
    if (ControlLib <> nil) then
      ControlLib.caption := RechDom('PGVARIABLE', GetControlText(Champ), FALSE);
  end;
  if (TypeChamp > '03') and (TypeChamp < '08') then
  begin
    SetControlProperty(Champ, 'DataType', 'PGREMUNERATION');
    SetControlProperty(Champ, 'ElipsisButton', TRUE);
    SetControlEnabled(champ, True);
    ControlLib := THLabel(GetControl(libelle));
    if (ControlLib <> nil) then
      ControlLib.caption := RechDom('PGREMUNERATION', GetControlText(Champ), FALSE);
  end;
  if (TypeChamp = '09') then
  begin
    SetControlProperty(Champ, 'DataType', 'PGCUMULPAIE');
    SetControlProperty(Champ, 'ElipsisButton', TRUE);
    SetControlEnabled(champ, True);
    ControlLib := THLabel(GetControl(libelle));
    if (ControlLib <> nil) then
      ControlLib.caption := RechDom('PGCUMULPAIE', GetControlText(Champ), FALSE);
  end;

end;


//============================================================================================//
//                                 PROCEDURE On New Record
//============================================================================================//

procedure TOM_Remuneration.OnNewRecord;

begin
  inherited;

  if mode = 'DUPLICATION' then exit;
  //Check box et Bouton radio par défault
// PT1 : 02/08/2001 : V547  Initialisation de PRM_BASEMTQTE
  SetField('PRM_BASEMTQTE', '004');
  SetField('PRM_PREDEFINI', 'DOS');
  SetField('PRM_TAUXMTQTE', 'X');
  SetField('PRM_COEFFMTQTE', '-');
  SetField('PRM_MONTANTMTQTE', 'X');
  // PT11  : 18/12/2002 PH V591 PORTAGECWAS controle des types de champs sur les setfield
  SetField('PRM_DECBASE', 2);
  SetField('PRM_DECTAUX', 2);
  SetField('PRM_DECCOEFF', 2);
  SetField('PRM_DECMONTANT', 2);
  SetField('PRM_NATURERUB', 'AAA');
  SetField('PRM_DADS', '-'); // PT41
end;

procedure TOM_Remuneration.OnAfterUpdateRecord;
var StSql: string;
  Tob_Prof: Tob;
  i: integer;
  Q: TQuery;
  even: boolean; //PT37
{PT37
    titre,LaTablette,TitreLabel,Ch10,Ch11,Ch20,Ch21 : String;
    LeLabel : THLabel;
    LeControl: TControl;
    Edititre: boolean;
    even: boolean;
//FIN PT37 }
begin
  inherited;
  SetControlEnabled('BDUPLIQUER', True);
  SetControlEnabled('BInsert', True);
  SetControlEnabled('BDelete', True);
  SetControlEnabled('BGESTIONASS', True);
  SetControlEnabled('BVENTIL', True);
  SetControlEnabled('PRM_PREDEFINI', False);
  SetControlEnabled('PRM_RUBRIQUE', False);
//PT37  titre := 'MODIFICATION REMUNERATION ';
//PT37  Edititre:= True;
  { DEB PT35 }
  { DEB PT30 }
  StSql := 'SELECT * FROM PROFILRUB ' +
    'WHERE ##PPM_PREDEFINI## PPM_NATURERUB="' + GetField('PRM_NATURERUB') + '" ' +
    'AND PPM_RUBRIQUE="' + GetField('PRM_RUBRIQUE') + '" ';
  try
    Q := OpenSql(StSql, True);
    if not Q.eof then
    begin
      Tob_Prof := TOB.Create('Les profils', nil, -1);
      Tob_Prof.LoadDetailDB('PROFILRUB', '', '', Q, FAlse);
    end;
    Ferme(Q);
    if Assigned(Tob_Prof) and (Tob_Prof.detail.count > 0) then
    begin
      for i := 0 to Tob_Prof.detail.count - 1 do
        Tob_Prof.detail[i].PutValue('PPM_LIBELLE', GetField('PRM_LIBELLE'));
      Tob_Prof.SetAllModifie(True);
      Tob_Prof.UpdateDB;
    end;
  finally
    FreeAndNil(Tob_Prof);
  end;
  { FIN PT30 }
  { FIN PT35 }
  //Rechargement des tablettes
//PT43  if (not Onferme) then begin
  if Assigned(Trace) then FreeAndNil(Trace);
  Trace := TStringList.Create;
  even := IsDifferent(dernierecreate, PrefixeToTable(TFFiche(Ecran).TableName), TFFiche(Ecran).CodeName, TFFiche(Ecran).LibelleName, Trace, TFFiche(Ecran),LeStatut); //PT37  //PT46
//PT43    Trace.free;
  FreeAndNil(Trace);
{$IFNDEF EAGLCLIENT}
  ChargementTablette(TFFiche(Ecran).TableName, '');
{$ELSE}
  ChargementTablette(TableToPrefixe(TFFiche(Ecran).TableName), ''); //PT26
{$ENDIF}
  if OnFerme then begin
    Ecran.Close;
  end;
end;

function TOM_Remuneration.ControlChampValeur: Integer;
var
  Mes, ValChamp: string;
begin
  Result := 0;
  Mes := 'Vous devez renseigner la valeur : ';
  if (GetField('PRM_TYPEBASE') = '') and (GetField('PRM_TYPETAUX') = '') and
    (GetField('PRM_TYPECOEFF') = '') and (GetField('PRM_TYPEMONTANT') = '') then
  begin
    Mes := mes + '#13#10- du type de la base, taux, coefficient ou montant';
    Result := 1;
  end
  else { DEB PT14 }
    if (GetField('PRM_CODECALCUL') = '01') and (GetField('PRM_TYPEBASE') = '') and (GetField('PRM_TYPEMONTANT') = '') then
    begin
      Mes := mes + '#13#10- du type de la base ou montant';
      Result := 1;
    end
    else
      if ((GetField('PRM_CODECALCUL') = '02') or (GetField('PRM_CODECALCUL') = '03') or (GetField('PRM_CODECALCUL') = '06') or (GetField('PRM_CODECALCUL') = '07'))
        and ((GetField('PRM_TYPEBASE') = '') or (GetField('PRM_TYPETAUX') = '') or (GetField('PRM_TYPECOEFF') = '')) then
      begin
        Mes := mes + '#13#10- du type de la base, taux et coefficient';
        Result := 1;
      end
      else
        if ((GetField('PRM_CODECALCUL') = '04') or (GetField('PRM_CODECALCUL') = '05'))
          and ((GetField('PRM_TYPEBASE') = '') or (GetField('PRM_TYPETAUX') = '')) then
        begin
          Mes := mes + '#13#10- du type de la base et taux';
          Result := 1;
        end { FIN PT14 }
        else { DEB PT17-2 }
          if (GetField('PRM_CODECALCUL') = '08') and
            ((GetField('PRM_TYPEBASE') = '') or (GetField('PRM_TYPECOEFF') = '')) then
          begin
            Mes := mes + '#13#10- du type de la base et Coefficient';
            Result := 1;
          end; { FIN PT17-2 }

  ValChamp := GetField('PRM_TYPEBASE');
  if ((ValChamp <> '') and (ValChamp <> '00') and (ValChamp <> '01')) and (GetField('PRM_BASEREM') = '') then
  begin
    Mes := mes + '#13#10- de la base de rémunération';
    Result := 1;
  end;
  ValChamp := GetField('PRM_TYPETAUX');
  if ((ValChamp <> '') and (ValChamp <> '00') and (ValChamp <> '01')) and (GetField('PRM_TAUXREM') = '') then
  begin
    Mes := mes + '#13#10- du taux';
    Result := 1;
  end;
  ValChamp := GetField('PRM_TYPECOEFF');
  if ((ValChamp <> '') and (ValChamp <> '00') and (ValChamp <> '01')) and (GetField('PRM_COEFFREM') = '') then
  begin
    Mes := mes + '#13#10- du coefficient';
    Result := 1;
  end;
  ValChamp := GetField('PRM_TYPEMONTANT');
  if ((ValChamp <> '') and (ValChamp <> '00') and (ValChamp <> '01')) and (GetField('PRM_MONTANT') = '') then
  begin
    Mes := mes + '#13#10- du montant';
    Result := 1;
  end;
  ValChamp := GetField('PRM_TYPEMINI');
  if ((ValChamp <> '') and (ValChamp <> '00') and (ValChamp <> '01')) and (GetField('PRM_VALEURMINI') = '') then
  begin
    Mes := mes + '#13#10- du minimum';
    Result := 1;
  end;
  ValChamp := GetField('PRM_TYPEMAXI');
  if ((ValChamp <> '') and (ValChamp <> '00') and (ValChamp <> '01')) and (GetField('PRM_VALEURMAXI') = '') then
  begin
    Mes := mes + '#13#10- du maximum';
    Result := 1;
  end;

  if Result = 1 then
    PgiBox(Mes + '.#13#10 Sinon le calcul de la paie sera erroné!', Ecran.caption);
  //HShowMessage('5;Rémunération :;'+mes+'.#13#10 Sinon le calcul de la paie sera erroné!;E;O;O;O;;;','','');
end;

procedure TOM_Remuneration.OnClose;
begin
  inherited;
  { DEB PT13 }
  if GetControlText('CBACCESFICHE') = 'X' then
  begin
    LastError := 1;
    LastErrorMsg := '';
    PgiBox('Suite à la modification de la gestion associée, vous devez valider votre fiche.', Ecran.caption);
  end;
  { FIN PT13 }
  Trace.free; //PT40
end;

procedure TOM_Remuneration.CBAccesFicheClick(Sender: TObject);
begin
  { DEB PT13 disabled si acces fiche gestion associée }
  SetControlEnabled('BFirst', (GetControlText('CBACCESFICHE') = '-'));
  SetControlEnabled('BPrev', (GetControlText('CBACCESFICHE') = '-'));
  SetControlEnabled('BNext', (GetControlText('CBACCESFICHE') = '-'));
  SetControlEnabled('BLast', (GetControlText('CBACCESFICHE') = '-'));
  SetControlEnabled('bInsert', (GetControlText('CBACCESFICHE') = '-'));
  SetControlEnabled('bDefaire', (GetControlText('CBACCESFICHE') = '-'));
  SetControlEnabled('BDUPLIQUER', (GetControlText('CBACCESFICHE') = '-'));
  { FIN PT13 }
{$IFDEF EAGLCLIENT}
  DS.Edit; { PT24 }
{$ENDIF}
end;
{ DEB PT19 }

procedure TOM_Remuneration.ImprimerClick(Sender: TObject);
begin
  LanceEtat('E', 'PAY', 'PRM', True, False, False, nil, 'AND ##PRM_PREDEFINI## PRM_RUBRIQUE="' + GetField('PRM_RUBRIQUE') + '"', '', False);
end;
{ FIN PT19 }
{ PT20 Mise à jour date modif table associée }

procedure TOM_Remuneration.PgUpdateDateAgl(Rub: string);
begin
  SetControlChecked('CBACCESFICHE', False);
  ExecuteSql('UPDATE CUMULRUBRIQUE SET' +
    ' PCR_DATEMODIF="' + UsTime(Now) + '" WHERE' +
    ' ##PCR_PREDEFINI## PCR_RUBRIQUE="' + Rub + '"');

//PT38
  ExecuteSql('UPDATE CUMULRUBDOSSIER SET' +
    ' PKC_DATEMODIF="' + UsTime(Now) + '" WHERE' +
    ' PKC_RUBRIQUE="' + Rub + '"');
//FIN PT38
end;

//PT38
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 08/03/2007
Modifié le ... :   /  /
Description .. : Click sur le bouton de la gestion spécifique
Mots clefs ... : PAIE
*****************************************************************}

procedure TOM_Remuneration.GestionSpec_OnClick(Sender: TObject);
var
  Quoi: string;
begin
  if (GetField('PRM_PREDEFINI') = 'CEG') then
    Quoi := 'C'
  else
    if (GetField('PRM_PREDEFINI') = 'STD') then
      Quoi := 'S'
    else
      Quoi := 'D';
  AglLanceFiche('PAY', 'CUMULGESTIONSPEC', '', '', GetField('PRM_NATURERUB') +
    ';' + GetField('PRM_RUBRIQUE') + ';' + GetField('PRM_LIBELLE') + ';' +
    Quoi);
{
if (GetControlText ('acces')='TRUE') then
   begin
   SetField ('PRM_ORDREETAT', '6');
   SetControlChecked ('CBACCESFICHE', True);
   end;
}
end;
//FIN PT38
// d PT41
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Monique FAUDEL
Créé le ...... : 08/03/2007
Modifié le ... :   /  /
Description .. : Click sur la case à cocher "Participe au calcul
Suite .........: des IOJSS"
Mots clefs ... : PAIE
*****************************************************************}

procedure TOM_Remuneration.CBIJSSClick(Sender: TObject);
begin
  if (GetControlText('CBIJSS') = 'X') then
    SetControlVisible('TYPRUB', True)
  else
  begin
    SetControlVisible('TYPRUB', False);
    SetControlText('TYPRUB', '');
  end;

  SetField('PRM_DADS', GetControlText('CBIJSS') + GetControlText('TYPRUB'));

end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Monique FAUDEL
Créé le ...... : 08/03/2007
Modifié le ... :   /  /
Description .. : Click sur le champ type de rubrique (calcul IJSS)
Mots clefs ... : PAIE
*****************************************************************}

procedure TOM_Remuneration.TYPRUBClick(Sender: TObject);
begin
  SetField('PRM_DADS', GetControlText('CBIJSS') + GetControlText('TYPRUB'));
end;
// f PT41
initialization
  registerclasses([TOM_Remuneration]);

end.

