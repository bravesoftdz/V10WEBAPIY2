{***********UNITE*************************************************
Auteur  ...... : Philippe Dumet
Créé le ...... : 18/02/2000
Modifié le ... : 08/11/2001
Description .. : Multi critère de lancement, de sélection des fichiers à
Suite ........ : émettre
Suite ........ : Confectionne une clause XX_WHERE
Mots clefs ... : PAIE;ENVOISOCIAL;PGDADSU
*****************************************************************}
{
PT1    : 08/01/2002 PH V571 Gestion déclaration des Prud'hom
PT2    : 06/02/2002 PH V571 Activation du filtre millesime pour les Prud'hom
PT3    : 27/03/2002 VG V571 Purge des envois DADS-U et prud'hommes
PT4    : 13/06/2002 MF V582 Gestion des Ducs EDI
PT5    : 02/07/2002 VG V585 Version S3
PT6    : 15/07/2002 VG V585 On replace le bouton bSelectAll dans sa position
                            initiale en fin de traitement
PT7    : 08/08/2002 MF V585 1- traitement des champs Code application et serveur
                            unique
                            2- Limitation des critères de sélection au type de
                            destinataire et à la période.
                            3- Pour les envois DUCS utilisation de la liste
                            PGENVOIDUCS
PT8    : 27/09/2002 MF V585 1- Nouvelle fiche PREP_COPYAM car infos. inutiles
                            sur PRE_DADSU
PT9    : 07/10/2002 VG V585 FQ N° 10258 et 10259 - Ergonomie
PT10   : 10/10/2002 MF V585 1- Code appli et serveur unique : 2 critères
                            supplémentaires sur la fiche
                            2- Traitement du cas où la traduction n'a pas abouti
PT11   : 26/11/2002 VG V591 Suite à la modification de la tablette
                            PGINSTITUTION, Les organismes ne s'affichaient plus
                            dans le look-up list
PT12   : 06/01/2003 MF V591 1- Correction des avertissement de compile
PT13   : 25/03/2003 VG V_42 Utilisation du nouveau paramètre société
                            "Emetteur du dossier" - FQ N°10600
PT14   : 31/03/2003 JL V_42 Gestion envoi congés spectacles pour intermittents
PT15   : 21/08/2003 VG V_42 Traitement de la DADS Bilatérale
PT16   : 18/09/2003 MF V_421
                            1- le champ PES_CODAPPLI n'existe pas sur la fiche
                            MUL_CONSULTENVSOC
                            2- PGIBox remplace MessageAlerte
                            3- Mise au point Cwas utilisation des répertoires
PT17   : 23/09/2003 MF V_421 FQ 10829 : pour les envois DUCS sélection possible
                            de l'émetteur.
PT18   : 26/09/2003 JL V_421 congés spectacles  : support par défaut = disquette
PT19   : 02/10/2003 VG V_42 Changement de l'exercice par défaut : 2003 -
                            FQ N°10854
PT20   : 28/04/2004 JL V_50 Gestion MSA EDI
PT21   : 08/10/2004 VG V_50 Exercice par défaut modifié automatiquement
PT22   : 11/10/2004 VG V_50 En consultation des envois, enlever le champ
                            fraction - FQ N°11661
PT23   : 13/12/2004 JL V_60 FQ 11086 Message si envoi test
PT24   : 26/01/2005 VG V_60 Pour les purges, désactivation du bouton [Ouvrir]
                            FQ N°11896
PT25   : 07/02/2005 VG V_60 Raffraichir la liste après la purge - FQ N°11961
PT26   : 18/02/2005 MF V_60 pour la ducs suupression des fichiers d'envoi
                            possible.
PT27   : 25/07/2005 MF V_604 compatibilité avec Socref 703 (prépa ducs edi v4.2)
PT28   : 30/09/2005 MF V_610 FQ 12611 : on rend le champ code appli invisible
PT29   : 20/10/2005 VG V_60 Suppression de l'emetteur pour TD bilatéral
PT30   : 21/10/2005 VG V_60 Remodelage
PT31   : 08/11/2005 VG V_65 Si PES_FICHIEREMIS ne fait pas partie de la liste,
                            erreur de transaction. on limite ce problème à la
                            purge de la DUCS - FQ N°11961
PT32   : 26/01/2006 VG V_65 Nouvelle liste pour TD bilatéral - FQ N°11661
PT33   : 09/02/2006 MF V_65 DUCS EDI V 4.2
PT34   : 02/05/2006 MF V_65 FQ 12028 : on remplace EMETSOC par EMETSOC2
PT35-1 : 12/05/2006 VG V_65 Possibilité d'afficher plusieurs fractions dans la
                            gestion des envois - FQ N°12872
PT35-2 : 12/05/2006 VG V_65 L'affichage par défaut ne se limite pas aux
                            destinataires ZTDS
PT36   : 16/06/2006 MF V_70 FQ 11997 : Purge DUCS correction pour annul du
                            fichier .cop qd "répertoire de travail choisi".
PT37   : 11/09/2006 MF V_70 FQ 13144 : On désactive le filtre par séfaut
PT38   : 27/09/2006 VG V_70 Affichage du libellé de l'emetteur dans un label
PT39   : 25/04/2007 MF V_72 Modifs Mise en base des fichiers Ducs EDI : traitement de purge
                            la modif concerne tous les types de fichier
PT40   : 08/08/2007 JL V_80 FQ 14477 Ajout filtre émetteur pour MSA
PT41   : 03/10/2007 VG V_80 Retour en arrière pour gestion correcte des fichiers en base
PT42   : 03/04/2007 MF V_80 Mise en place des WebServices Jedeclare
PT43   : 16/10/2007 MF V_80 FQ 14864 le support JeDeclare uniquement pour les envois social
                            + Jedeclare uniquement dans un environnement PCL 
}
unit UTofPG_MulEnvoiSocial;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97, Hqry,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, Mul, Fe_Main,
{$ELSE}
  MaineAgl, emul,
{$ENDIF}
{$IFNDEF DADSUSEULE}
{$IFNDEF COMPTA}
{$IFNDEF DADSB}
  PgEnvoiDucsEdi, // PT4  DUCSEDI
{$ENDIF}
{$ENDIF}
{$ENDIF}
  Grids, HCtrls, HEnt1, EntPaie, HMsgBox, UTOF, UTOB, UTOM, Vierge, AGLInit, ed_tools,
  HStatus,
{$IFNDEF DADSUSEULE}
{$IFNDEF COMPTA}
{$IFNDEF DADSB}
  PGOutils,
{$ENDIF}
{$ENDIF}
{$ENDIF}
{$IFNDEF COMPTA}
  PGoutils2,
{$ENDIF}
  uYFILESTD,  // PT39
{$IFDEF PAIEGRH}
  Paie_Cjdc,      // PT42 JDC
  Paie_cjdc_lib,  // PT42 JDC
{$ENDIF}
  paramsoc;
type
  TOF_PGMULEnvoiSocial = class(TOF)
  private
    DD, DF: THEdit;
{$IFDEF PAIEGRH}
    SUPPORTEMIS : THValComboBox; //PT42
    ENVOIREEL : TCheckBox; //PT42
{$ENDIF}
    Emetteur, LeType, TypeEnvoi : string; // concerne le type d'envoi
    Q_Mul: THQuery;
    BCherche: TToolbarButton97;
    procedure LanceFicheDADSU(Sender: TObject);
    procedure ActiveWhere(Sender: TObject);
    procedure SupprimerClick(Sender: TObject);
    procedure Supprimer_un;
{$IFDEF PAIEGRH}
    procedure FListe_OnDblClick(Sender: TObject); //PT42
{$ENDIF}
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
  end;

var // Definition de la grille
{$IFNDEF EAGLCLIENT}
Grille: THDBGrid;
{$ELSE}
Grille: THGrid;
{$ENDIF}
FRapport: TextFile;
{$IFDEF PAIEGRH}
SiretEmett : string; //PT42
{$ENDIF}
  procedure LanceMul_DADS2Envoi (LaFiche, LeParam : String);
implementation

procedure LanceMul_DADS2Envoi (LaFiche, LeParam : String);
begin
  AGLLanceFiche('PAY', LaFiche, '' ,'', LeParam);
end;


procedure TOF_PGMULEnvoiSocial.ActiveWhere(Sender: TObject);
var
WW: THEdit;
Dat1, Dat2, StrEmet : string;
begin
WW := THEdit(GetControl('XX_WHERE'));
WW.Text := '';

if (TypeEnvoi = 'DUCS') then
   begin
   if (Q_Mul <> nil) then
      TFMul(Ecran).SetDBListe('PGENVOIDUCS');

   if LeType = 'DUCS' then
      begin
      if (DD <> nil) and (WW <> nil) then
         begin
         Dat1:= UsDateTime(StrToDate(DD.Text));
         Dat2:= UsDateTime(StrToDate(DF.Text));
         StrEmet:='';
         if (Emetteur <> '') then
            StrEmet:= ' PES_EMETSOC="'+Emetteur+'" AND';
         WW.Text:= ' PES_DATEDEBUT >="'+Dat1+'"'+' AND'+
                   ' PES_DATEFIN <="'+Dat2+'"'+' AND'+StrEmet+
                   ' (PES_MONNAIEDECL="EUR" OR PES_MONNAIEDECL="FRF")';
         end;
      end;
// d PT42
{$IFDEF PAIEGRH}
   if (Ecran.name = 'MUL_CONSULTENVSOC') or (LeType = 'DUCSP') then
   begin
     if (SUPPORTEMIS <> nil) then
     begin
       if WW.Text <> '' then
         WW.Text := WW.Text+' AND';
       WW.Text:= WW.text + ' PES_SUPPORTEMIS = "'+ SUPPORTEMIS.value+'" '+
                           ' OR PES_SUPPORTEMIS = "" ';
     end;
     if (LeType = 'DUCSP') and (ENVOIREEL <> nil) then
     begin
       if WW.Text <> '' then
         WW.Text := WW.Text+' AND';
       if (ENVOIREEL.checked) = true  then
         WW.Text:= WW.text + ' PES_ENVOIREEL ="X"'
       else
         WW.Text:= WW.text + ' PES_ENVOIREEL ="-"';
     end;
   end;
{$ENDIF}
// f PT42
   end;

if ((TypeEnvoi = 'DADSU') or (TypeEnvoi = 'DADSB') or
   (TypeEnvoi = 'PRUDH')) then
   begin
//PT32
   if (TypeEnvoi = 'DADSB') then
      begin
      if (Q_Mul <> nil) then
         TFMul(Ecran).SetDBListe('PGENVOITDB');
      end;
//FIN PT32

   DD := THEdit(GetControl('MILLESSOC'));
   if (DD <> nil) and (WW <> nil) then
      WW.Text:= ' PES_MILLESSOC = "'+GetControlText('MILLESSOC')+'"';
{PT35-1
   if (LeType = 'DADSU') then
}
   if ((LeType = 'DADSU') and (GetControlText('FRACTIONDADS') <> '')) then
      WW.Text:= WW.text+' AND PES_FRACTIONDADS="'+GetControlText('FRACTIONDADS')+'"';
   end;
if (TypeEnvoi = 'MSA') then //PT40
   begin
         if (Emetteur <> '') then
         WW.Text:= ' PES_EMETSOC="'+Emetteur+'"';
   end;
end;

procedure TOF_PGMULEnvoiSocial.LanceFicheDADSU(Sender: TObject);
var
{$IFNDEF COMPTA}
  FichierCopaym         : String;
  resultat              : Integer;
{$ENDIF}
st: string;
rep : Integer;
begin
// On sensibilise les utilsateurs sur la mouette bleue car avec la multiples sélection
// ???????? rafraichessement de la liste  et désélection
Rep:= PGIAsk ('Avez-vous bien appliqué vos critères avant de lancer le traitement?',
             Ecran.Caption);
if rep <> 6 then
   exit;

if GetCheckBoxState('ENVOIREEL') <> CbChecked then
   PGIBox ('Vous n''avez pas coché envoi réel, il s''agit donc d''un fichier test',
          Ecran.Caption); // PT23

{$IFDEF EAGLCLIENT}
TheMulQ := TOB(Ecran.FindComponent('Q'));
Grille := THGrid(GetControl('FLISTE'));
{$ELSE}
TheMulQ := THQuery(Ecran.FindComponent('Q'));
Grille := THDBGrid(GetControl('FLISTE'));
{$ENDIF}
st:= GetControlText('PES_TYPEMESSAGE')+';'+GetControlText('MILLESSOC')+';'+
     GetControlText('PES_PGPERIODICITE')+';'+
//PT34     GetControlText('PES_INSTITUTION')+';'+GetControlText('EMETSOC')+';'+
     GetControlText('PES_INSTITUTION')+';'+GetControlText('EMETSOC2')+';'+
     GetControlText('SUPPORTEMIS')+';'+GetControlText('ENVOIREEL')+';'+
     GetControlText('PES_MONNAIEDECL');

if Grille <> nil then
   if (Grille.NbSelected = 0) and (not Grille.AllSelected) then
      begin
      PgiBox('Aucun élément sélectionné', Ecran.Caption);
      exit;
      end;

{PT29
if (LeType <> 'CONGSPEC') and (LeType <> 'MSA') and
   ((GetControlText('PES_INSTITUTION') = '') or (((LeType <> 'DUCS') and
   (LeType <> 'DUCSP')) and (GetControlText('EMETSOC') = '')) or
   (GetControlText('SUPPORTEMIS') = '')) then
   PgiBox('Vous devez renseigner entièrement votre écran', Ecran.caption)
}
if ((LeType <> 'DADSB') and (LeType <> 'DUCS') and
//PT34   (GetControlText('EMETSOC') = '')) then
   (GetControlText('EMETSOC2') = '')) then
   PgiBox('Vous devez renseigner l''émetteur', Ecran.caption)
else
   begin
   if LeType = 'DADSU' then
      AglLanceFiche('PAY', 'PREP_DADSU', '', '', st)
   else
   if LeType = 'DADSB' then
      AglLanceFiche('PAY', 'PREP_DADSB', '', '', st)
   else
   if LeType = 'PRUDH' then
      AglLanceFiche('PAY', 'ENVOI_PRUDH', '', '', st);
{$IFNDEF DADSUSEULE}
{$IFNDEF COMPTA}
{$IFNDEF DADSB}
   if LeType = 'DUCS' then
      begin
// Sélection des DUCS
      if Grille <> nil then
         begin
// d PT33 DUCS EDI V 4.2
//         st := st +';'+ GetControlText ('PES_CODAPPLI')+';'+GetControlText ('PES_SERVUNIQ');
         st := st +';'+GetControlText ('PES_SERVUNIQ');
// f PT33 DUCS EDI V 4.2
         LanceEnvoiDucsEdi(Grille, TFMul(Ecran), st, FichierCopaym, resultat);
         if (resultat = 0) then
            begin
            st:= st+';'+FichierCopaym+';'+IntToStr(resultat);
            AglLanceFiche ('PAY', 'PREP_COPAYM', '', '', st);
            end;
         end;
      end;
{$ENDIF}
{$ENDIF}
{$ENDIF}
//  PT33 DUCS EDI V 4.2

   if LeType = 'CONGSPEC' then
{PT30
      begin
      if (GetControlText('EMETSOC') = '') or
         (GetControlText('SUPPORTEMIS') = '') then
         begin
         PgiBox('Vous devez renseigner entièrement votre écran', Ecran.caption);
         Exit;
         end;
      AglLanceFiche('PAY', 'PREP_CONGESSPEC', '', '', st);
      end;
}
      AglLanceFiche('PAY', 'PREP_CONGESSPEC', '', '', st);

   if LeType = 'MSA' then
{PT30
      begin
      if (GetControlText('EMETSOC') = '') or
         (GetControlText('SUPPORTEMIS') = '') then
         begin
         PgiBox('Vous devez renseigner entièrement votre écran', Ecran.caption);
         Exit;
         end;
      AglLanceFiche('PAY', 'PREP_MSA', '', '', st);
      end;
}
      AglLanceFiche('PAY', 'PREP_MSA', '', '', st);

   if Grille <> nil then
      begin
      if (Grille.AllSelected = TRUE) then
         Grille.AllSelected := FALSE
      else
         Grille.ClearSelected;
      end;
   end;
TheMulQ := nil;

TFMul(Ecran).bSelectAll.Down := Grille.AllSelected;
end;

procedure TOF_PGMULEnvoiSocial.OnArgument(Arguments: string);
var
BtnDeleteMul, BtnValidMul: TToolbarButton97;
{$IFNDEF COMPTA}
  AnneeE                : String;
  DebPer                : String;
  ExerPerEncours        : String;
  FinPer                : String;
  MoisE                 : String;
  DebTrim               : TdateTime;
  FinTrim               : TdateTime;
  DebSem                : TdateTime;
  FinSem                : TdateTime;
  DebExer               : TdateTime;
  FinExer               : TdateTime;
{$ENDIF}
AnneePrec, Purge : string;
Jour: TDateTime;
TT: TFMul;
Emett: THValComboBox;
AnneeA, JourJ, MoisM: Word;
{$IFDEF PAIEGRH}
sSql  : string; // PT42
Q : TQuery;     // PT42
JDC : boolean;  // PT42
{$ENDIF}
begin
inherited;
TFMul(Ecran).FiltreDisabled:= True; // PT37
Q_Mul := THQuery(Ecran.FindComponent('Q'));
LeType := Trim(Arguments);
Purge := Copy (LeType, Length(LeType), 1);
if (Purge='P') then
   TypeEnvoi:= Copy (LeType, 1, Length(LeType)-1)
else
   TypeEnvoi:= Copy (LeType, 1, Length(LeType));

//On rend inaccessible le type de fichier qui dépend du menu choisi
SetControlEnabled('PES_TYPEMESSAGE', FALSE);

{3 types de menus
Envoi        => Fiche MUL_ENVOISOCIAL   + Type  (Purge<>'P')
Consultation => Fiche MUL_CONSULTENVSOC + Type  (Purge<>'P')
Purge        => Fiche MUL_ENVOISOCIAL   + TypeP (Purge='P')
}

//Menu Envoi et Purge
if Ecran.Name = 'MUL_ENVOISOCIAL' then
   begin
//  SetControlText('EMETSOC', GetParamSoc('SO_PGEMETTEUR'));
   SetControlText('EMETSOC2', GetParamSocSecur('SO_PGEMETTEUR',''));
//PT38
   SetControlText ('LEMETSOC2', RechDom ('PGEMETTEURSOC',
                                         GetControlText ('EMETSOC2'), False));

//Menu purge
   if (Purge='P') then
      begin
      SetControlVisible ('BOuvrir', FALSE);
      SetControlVisible ('BDelete', TRUE);
      BtnDeleteMul:= TToolbarButton97(GetControl('BDelete'));
      if BtnDeleteMul <> nil then
         BtnDeleteMul.OnClick:= SupprimerClick;
      TFMul(Ecran).Caption:= 'Purge ';
      end
   else
//Menu Envoi
      TFMul(Ecran).Caption:= 'Envois ';
   end
else
//Menu Consultation
   begin
   TFMul(Ecran).Caption:= 'Consultation ';

   setcontrolvisible('BOUVRIR', FALSE);
   SetControlText('PES_EMETSOC', GetParamSocSecur('SO_PGEMETTEUR','')); 
   end;

if (TypeEnvoi = 'DUCS') then
   TFMul(Ecran).Caption:= TFMul(Ecran).Caption+'DUCS'
else
if (TypeEnvoi = 'DADSU') then
   TFMul(Ecran).Caption:= TFMul(Ecran).Caption+'DADS-U'
else
if (TypeEnvoi = 'DADSB') then
   TFMul(Ecran).Caption:= TFMul(Ecran).Caption+'TD Bilatéral'
else
if (TypeEnvoi = 'PRUDH') then
   TFMul(Ecran).Caption:= TFMul(Ecran).Caption+'Prud''hom'
else
if (TypeEnvoi = 'CONGSPEC') then
   TFMul(Ecran).Caption:= TFMul(Ecran).Caption+'Congés Spectacles'
else
if (TypeEnvoi = 'MSA') then
   TFMul(Ecran).Caption:= TFMul(Ecran).Caption+'Mutualité Sociale Agricole';

TT := TFMul(Ecran);
if TT <> nil then
   UpdateCaption(TT);

// d PT42
{$IFDEF PAIEGRH}

JDC := false;
if (V_PGI.ModePCL='1') then
begin

  if ((Ecran.Name = 'MUL_ENVOISOCIAL') and (GetControlText('EMETSOC2') <> '')) or
     ((Ecran.Name <> 'MUL_ENVOISOCIAL') and (GetControlText('PES_EMETSOC') <> '')) then
  begin
    sSql := 'SELECT  * FROM EMETTEURSOCIAL, NJDC_COMPTES ';

    if Ecran.Name = 'MUL_ENVOISOCIAL' then
      sSql := sSql + ' WHERE PET_EMETTSOC="' + GetControlText('EMETSOC2')+ '" '
    else
      sSql := sSql + ' WHERE PET_EMETTSOC="' + GetControlText('PES_EMETSOC')+ '" ';

    sSql := sSql + 'AND PET_SIRET = NJC_INTITULE';
    Q := OpenSQL(sSql, TRUE);
    if (not Q.Eof) then
    begin
      SiretEmett := Q.FindField('PET_SIRET').AsString;
// d PT43
      //  NJC_LISTEPROCEDURE = 1 si social NJC_LISTEPROCEDURE = 2 si fiscal & social
      if (Q.FindField('NJC_LISTEPROCEDURE').AsString ='1') or
         (Q.FindField('NJC_LISTEPROCEDURE').AsString ='2') then
// f PT43
        JDC := true;
    end;
    Ferme(Q);
  end;
end;
if (JDC) and (Ecran.Name <> 'MUL_ENVOISOCIAL')  then
begin
   TFMul(Ecran).FListe.OnDblClick := FListe_OnDblClick;
end;
{$ENDIF}
// f PT42

SetControlText('SUPPORTEMIS', 'DIS');
//d PT34
    if (Ecran.Name = 'MUL_CONSULTENVSOC') then
    begin
      Emett := THValCombobox(GetControl('PES_EMETSOC'));
      if (Emett <> nil) and (Emett.VideString = '') then
         begin
         Emett.Items.Add(Traduirememoire('Tous'));
         Emett.Values.Add('');
         end;
    end;
// f PT34

//Spécificités DADS-U
if (TypeEnvoi = 'DADSU') then
   begin
   SetControlText('PES_TYPEMESSAGE', 'DAD');

//PT35-2
   if (Purge = 'P') then
      SetControlText('PES_INSTITUTION', '')  ;
{
   else
      SetControlText('PES_INSTITUTION', 'ZTDS');
FIN PT35-2}
   SetControlProperty('PES_INSTITUTION', 'Plus', ' AND PIP_INSTITUTION NOT LIKE "ZD%"');

   if VH_Paie.PGBTP = False then
      SetControlEnabled('PES_PGPERIODICITE', FALSE)
   else
      SetControlEnabled('PES_PGPERIODICITE', TRUE);

   SetControlText('FRACTIONDADS', '1');
   end
else
//Spécificités DUCS
if (TypeEnvoi = 'DUCS') then
   begin
   DD := THEdit(GetControl('DATEDEBUT'));
   DF := THEdit(GetControl('DATEFIN'));
// d PT42
{$IFDEF PAIEGRH}

   if (Ecran.name = 'MUL_CONSULTENVOI') or (Purge = 'P') then
   begin
     SUPPORTEMIS := THValComboBox(GetControl('SUPPORTEMIS'));
     if (Purge = 'P') then
       ENVOIREEL := TCheckBox(GetControl('ENVOIREEL'));
   end;
{$ENDIF}
// f PT42
// Dates par défaut :
//   La période proposées est le trimestre en cours si la période
//   en cours correspond à une fin de trimestre, sinon c'est le mois
//   en cours.
{$IFNDEF DADSUSEULE}
{$IFNDEF COMPTA}
{$IFNDEF DADSB}
   if RendExerSocialEnCours(MoisE, AnneeE, ExerPerEncours, DebExer, FinExer) = True then
      begin
      RendPeriodeEnCours(ExerPerEnCours, DebPer, FinPer);
      RendTrimestreEnCours(StrToDate(DebPer), DebExer, FinExer, DebTrim, FinTrim, DebSem, FinSem);
      if FindeMois(StrToDate(DebPer)) <> FinTrim then
// mois en cours
         begin
         if DD <> nil then
            DD.text := DateToStr(StrToDate(DebPer));
         if DF <> nil then
            DF.text := DateToStr(FindeMois(StrToDate(DebPer)));
         end
      else
// trimestre en cours
         begin
         if DD <> nil then
            DD.text := DateToStr(DebTrim);
         if DF <> nil then
            DF.text := DateToStr(FinTrim);
         end;
      end;
{$ENDIF}
{$ENDIF}
{$ENDIF}
   if (Ecran.Name = 'MUL_ENVOISOCIAL') then
      begin
//PT28      SetControlVisible('PES_CODAPPLI', TRUE);
//PT28      SetControlVisible('TPES_CODAPPLI', TRUE);
      SetControlEnabled('PES_SERVUNIQ', TRUE);
      SetControlVisible('PES_SERVUNIQ', TRUE);
{PT34
      Emett := THValCombobox(GetControl('EMETSOC'));
      if (Emett <> nil) and (Emett.VideString = '') then
         begin
         Emett.Items.Add(Traduirememoire('Tous'));
         Emett.Values.Add('');
         end;}
//FQ 12658      SetControlText('PES_CODAPPLI', '013');
      SetControlText('PES_CODAPPLI', '');
      end;
   SetControlText('PES_TYPEMESSAGE', 'DUC');
   SetControlText('PES_INSTITUTION', 'ZDAC');
   SetControlProperty('PES_INSTITUTION','Plus',' AND PIP_INSTITUTION LIKE "ZD%"');
   SetControlText('SUPPORTEMIS', 'TEL');
   SetControlText('TPES_DATEDEBUT', 'Période de');
   SetControlText('TPES_DATEFIN', 'à');
// d PT42
{$IFDEF PAIEGRH}
   if (JDC) then
     SetControlText('SUPPORTEMIS', 'JDC');
{$ENDIF}
// f PT42
   end
else
//Spécificités TD bilatéral
if (TypeEnvoi = 'DADSB') then
   begin
   SetControlText('PES_TYPEMESSAGE', 'DA2');
   SetControlText('PES_INSTITUTION', 'ZIMP');
   SetControlEnabled('PES_PGPERIODICITE', FALSE);
//PT29
   if Ecran.Name = 'MUL_ENVOISOCIAL' then
      begin
{PT38
      SetControlVisible('EMETSOC', False);
}
      SetControlVisible('LEMETSOC2', False);
//FIN PT38
      SetControlVisible('EMETSOC2', False);
      end
   else
      SetControlVisible('PES_EMETSOC', False);
   SetControlVisible('TPES_EMETSOC', False);
//FIN PT29
   end
else
//Spécificités Congés Spectacles
if (TypeEnvoi = 'CONGSPEC') then
   begin
   SetControlText('PES_TYPEMESSAGE', 'PCS');
   SetControlText('SUPPORTEMIS', 'DTK');
   SetControlVisible('ENVOIREEL', False);
   SetControlVisible('PES_ENVOIREEL', False);
   end
else
//Spécificités Mutualité Sociale Agricole
if (TypeEnvoi = 'MSA') then
   SetControlText('PES_TYPEMESSAGE', 'MSA')
else
//Spécificités Envois Prud'hom
if (TypeEnvoi = 'PRUDH') then
   begin
   SetControlText('PES_TYPEMESSAGE', 'PRH');
   SetControlText('PES_INSTITUTION', 'ZPRU');
   SetControlText('FRACTIONDADS', 'Z');
   SetControlVisible ('MILLESSOC', FALSE);
   SetControlVisible ('TPES_MILLESSOC', FALSE);
   end;

if ((TypeEnvoi <> 'DADSU') and (TypeEnvoi <> 'DUCS')) then
   begin
   SetControlVisible('PES_INSTITUTION', FALSE);
   SetControlVisible('TPES_INSTITUTION', FALSE);
   end;

if (TypeEnvoi <> 'DADSU') then
   begin
   SetControlVisible('FRACTIONDADS', FALSE);
   if (Ecran.Name = 'MUL_CONSULTENVSOC') then
      SetControlVisible('LBLFRACTION', FALSE)
   else
      SetControlVisible('TFRACTIONDADS', FALSE);
   end;

if (TypeEnvoi <> 'DUCS') then
   begin
   SetControlVisible('DATEDEBUT', FALSE);
   SetControlVisible('DATEFIN', FALSE);
   SetControlVisible('TPES_DATEDEBUT', FALSE);
   SetControlVisible('TPES_DATEFIN', FALSE);
   SetControlText('PES_MONNAIEDECL', 'EUR');
   end;

if ((TypeEnvoi='DUCS') or (TypeEnvoi='CONGSPEC') or (TypeEnvoi='MSA')) then
   begin
   SetControlVisible('MILLESSOC', FALSE);
   SetControlVisible('TPES_MILLESSOC', FALSE);
   if ((TypeEnvoi='CONGSPEC') or (TypeEnvoi='MSA')) then
      SetControlText('PES_PGPERIODICITE', 'T');
   end
else
   begin
   SetControlText('PES_PGPERIODICITE', 'A');
   Jour := Date;
   DecodeDate(Jour, AnneeA, MoisM, JourJ);
   if MoisM > 9 then
      AnneePrec := IntToStr(AnneeA)
   else
      AnneePrec := IntToStr(AnneeA - 1);

   SetControlText('MILLESSOC', copy(AnneePrec, 1, 1) + copy(AnneePrec, 3, 2));
   end;

if ((TypeEnvoi = 'DUCS') or (TypeEnvoi = 'PRUDH')) then
   begin
   SetControlVisible('PES_PGPERIODICITE', FALSE);
   SetControlVisible('TPES_PGPERIODICITE', FALSE);
   SetControlVisible('PES_MONNAIEDECL', FALSE);
   SetControlVisible('TPES_MONNAIEDECL', FALSE);
   end;

//PT29
//PT33 DUCS EDI V 4.2
{if (TypeEnvoi <> 'DADSB') then
   begin
   SetControlVisible('EMETSOC', TRUE);
   SetControlEnabled('EMETSOC', TRUE);
   SetControlVisible('EMETSOC2', FALSE);
   end;}

if ((TypeEnvoi = 'CONGSPEC') or (TypeEnvoi = 'MSA')) then
   SetControlEnabled('PES_MONNAIEDECL', False);

BtnValidMul := TToolbarButton97(GetControl('BOuvrir'));
if BtnValidMul <> nil then
   BtnValidMul.OnClick := LanceFicheDADSU;

BCherche := TToolbarButton97(GetControl('BCherche'));
end;

procedure TOF_PGMULEnvoiSocial.OnLoad;
begin
inherited;
if (Ecran.Name <> 'MUL_CONSULTENVSOC') then
//PT34   Emetteur := GetControlText('EMETSOC');
Emetteur:= GetControlText('EMETSOC2');
ActiveWhere(nil);
end;

//PT3
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 27/03/2002
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "Supprimer"
Mots clefs ... : PAIE;ENVOISOCIAL;PGDADSU
*****************************************************************}

procedure TOF_PGMULEnvoiSocial.SupprimerClick(Sender: TObject);
var
Ficlog: string;
i: integer;
begin
{$IFNDEF EAGLCLIENT}
Grille := THDBGrid(GetControl('FListe'));
{$ELSE}
Grille := THGrid(GetControl('FListe'));
{$ENDIF}
if Grille <> nil then
   begin
   if (Grille.NbSelected = 0) and (not Grille.AllSelected) then
      begin
      PgiBox('Aucun élément sélectionné', Ecran.Caption); // PT16-2
      exit;
      end;
   end;
// d PT16-3
{$IFNDEF EAGLCLIENT}
Ficlog := V_PGI.DatPath + '\PENVOI.log';
{$ELSE}
Ficlog := VH_Paie.PGCheminEagl + '\PENVOI.log';
{$ENDIF}
// f PT16-3

if (Grille.AllSelected = TRUE) then
   begin
   InitMoveProgressForm (nil, 'Suppression en cours',
                        'Veuillez patienter SVP ...',
                        TFmul(Ecran).Q.RecordCount, FALSE, TRUE);
   InitMove(TFmul(Ecran).Q.RecordCount, '');

   AssignFile(FRapport, Ficlog);
   if FileExists(Ficlog) then
      Append(FRapport)
   else
      ReWrite(FRapport);

   TFmul(Ecran).Q.First;
   while not TFmul(Ecran).Q.EOF do
         begin
         Supprimer_un;
         TFmul(Ecran).Q.Next;
         end;
   Grille.AllSelected := False;
   TFMul(Ecran).bSelectAll.Down := Grille.AllSelected;
   end
else
   begin
   InitMoveProgressForm (nil, 'Suppression en cours',
                        'Veuillez patienter SVP ...', Grille.NbSelected, FALSE,
                        TRUE);
   InitMove(Grille.NbSelected, '');

   AssignFile(FRapport, Ficlog);
   if FileExists(Ficlog) then
      Append(FRapport)
   else
      ReWrite(FRapport);

   for i := 0 to Grille.NbSelected - 1 do
       begin
       Grille.GotoLeBOOKMARK(i);
       Supprimer_un;
       end;

   Grille.ClearSelected;
   end;

PGIBox('Traitement terminé', 'Purge des envois');
CloseFile(FRapport);
FiniMove;
FiniMoveProgressForm;

//PT25
if BCherche <> nil then
   BCherche.click;
//FIN PT25
end;
//FIN PT3


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 12/05/2004
Modifié le ... :   /  /
Description .. : Suppression d'un seul fichier
Mots clefs ... : PAIE;ENVOISOCIAL;PGDADSU
*****************************************************************}
procedure TOF_PGMULEnvoiSocial.Supprimer_un;
var
FicSuppr, St: string;
FichierRecu: string; // PT26
CodeRetour : integer; // PT10
Crit2 : string; // PT10

begin
MoveCur(False);
Writeln(FRapport, '');

St := TFmul(Ecran).Q.FindField('PES_FICHIERRECU').asstring;
FichierRecu := St; // PT26

//d PT39
Crit2 := '';
  // Récupération du fichier Ducs EDI dans la base, FileM = Chemin et nom du fichier
{ FQ 20737 BVE 18.06.07 }
{PT41     }
{$IFDEF COMPTA}
if (TFmul(Ecran).DBListe = 'PGENVOIDUCS') and
   (TFmul(Ecran).Q.findfield('PES_GUID1').AsString <> '') then
{$ELSE}
if (TFmul(Ecran).Q.findfield('PES_GUID1').AsString <> '') then
{$ENDIF}
{ END FQ 20737 }
  begin
  Crit2 := Copy(TFmul(Ecran).Q.findfield('PES_FICHIERRECU').AsString, 1, 3);
  CodeRetour := AGL_YFILESTD_EXTRACT (FicSuppr,
                                      'PAIE',
                                      TFmul(Ecran).Q.findfield('PES_FICHIERRECU').AsString ,
                                      'DUC', Crit2, '', '','',
                                      False, 'FRA', 'DOS');
  if (CodeRetour <> -1 ) then
    PGIInfo(AGL_YFILESTD_GET_ERR(CodeRetour) +
                                 #13#10 +
                                 TFmul(Ecran).Q.findfield('PES_FICHIERRECU').AsString);

  end;
// f PT39

if St <> '' then
   begin
   try
      begintrans;
// d PT39
      { FQ 20737 BVE 18.06.07 }
{PT41           }
{$IFDEF COMPTA}
      if (TFmul(Ecran).DBListe = 'PGENVOIDUCS') and
         (TFmul(Ecran).Q.findfield('PES_GUID1').AsString = '') then
{$ELSE}
      if (TFmul(Ecran).Q.findfield('PES_GUID1').AsString = '') then
{$ENDIF}
      { END FQ 20737 }
      begin
// f PT39
{$IFNDEF EAGLCLIENT}
// d PT33 DUCS EDI V 4.2
{$IFNDEF DUCS41}
// DUCS 4.2
      if (Q_Mul.Liste  = 'PGENVOIDUCS') and
         (TFmul(Ecran).Q.FindField('PES_CODAPPLI').asstring = 'V42') then
        FicSuppr := VH_Paie.PGCheminEagl+'\'+St
      else
      FicSuppr:= V_PGI.DatPath+'\'+St;
{$ELSE}
// DUCS 4.1
      FicSuppr:= V_PGI.DatPath+'\'+St;
{$ENDIF}
// f PT33 DUCS EDI V 4.2

{$ELSE}
      FicSuppr:= VH_Paie.PGCheminEagl+'\'+St;
{$ENDIF}
      end; //PT39

      if FileExists(FicSuppr) then
         begin
         DeleteFile(PChar(FicSuppr));
         Writeln(FRapport, 'Fichier '+St+' supprimé : '+DateTimeToStr(Now));
         end
      else
         Writeln(FRapport, 'Fichier '+St+' absent : '+DateTimeToStr(Now));
// d PT26
// d PT39
//    suppression fichier dans la base
      { FQ 20737 BVE 18.06.07 }
{PT41           }
{$IFDEF COMPTA}
      if (TFmul(Ecran).DBListe = 'PGENVOIDUCS') and
         (TFmul(Ecran).Q.findfield('PES_GUID1').AsString <> '') then
{$ELSE}
      if (TFmul(Ecran).Q.findfield('PES_GUID1').AsString <> '') then
{$ENDIF}
      { END FQ 20737 }
      begin
         ExecuteSQL('DELETE FROM YFILESTD WHERE YFS_FILEGUID="'+
                    TFmul(Ecran).Q.findfield('PES_GUID1').AsString+'"');
         ExecuteSQL('DELETE FROM NFILES WHERE NFI_FILEGUID="'+
                    TFmul(Ecran).Q.findfield('PES_GUID1').AsString+'"');
         ExecuteSQL('DELETE FROM NFILEPARTS WHERE NFS_FILEGUID="'+
                    TFmul(Ecran).Q.findfield('PES_GUID1').AsString+'"');
      end;
// f PT39
{PT31
      St :=TFmul(Ecran).Q.FindField('PES_FICHIEREMIS').asstring;
      if (LeType = 'DUCSP') and (st <> '') then
}
      if (LeType = 'DUCSP') then
         begin
         St :=TFmul(Ecran).Q.FindField('PES_FICHIEREMIS').asstring;
         if (st <> '') then
//FIN PT31
            begin
            if (PGIAsk ('Voulez-vous supprimer le fichier envoyé? '+st,
                       Ecran.Caption) = mrYes) then
               begin
{$IFNDEF EAGLCLIENT}
// d PT33 DUCS EDI V 4.2
{$IFNDEF DUCS41}
// DUCS 4.2
          if (Q_Mul.Liste  = 'PGENVOIDUCS') and
             (TFmul(Ecran).Q.FindField('PES_CODAPPLI').asstring = 'V42') then
            FicSuppr := VH_Paie.PGCheminEagl+'\'+St
          else
// d PT36
             if (TFmul(Ecran).Q.FindField('PES_SUPPORTEMIS').asstring = 'TRA') then
               FicSuppr := VH_Paie.PGCheminEagl+'\'+St
             else
               FicSuppr:= V_PGI.DatPath+'\'+St;
{$ELSE}
// DUCS 4.1
             if (TFmul(Ecran).Q.FindField('PES_SUPPORTEMIS').asstring = 'TRA') then
               FicSuppr := VH_Paie.PGCheminEagl+'\'+St
             else
// f PT36
               FicSuppr:= V_PGI.DatPath+'\'+St;
{$ENDIF}
// f PT33 DUCS EDI V 4.2
{$ELSE}
               FicSuppr:= VH_Paie.PGCheminEagl+'\'+St;
{$ENDIF}
               if FileExists(FicSuppr) then
                  begin
                  DeleteFile(PChar(FicSuppr));
                  Writeln(FRapport, 'Fichier '+St+' supprimé : '+DateTimeToStr(Now));
                  end
               else
                  Writeln(FRapport, 'Fichier '+St+' absent : '+DateTimeToStr(Now));
               end;
            end;
         end;   //PT31
// f PT26
//PT26      ExecuteSQL('DELETE FROM ENVOISOCIAL WHERE PES_FICHIERRECU = "'+St+'"');
      ExecuteSQL('DELETE FROM ENVOISOCIAL WHERE PES_FICHIERRECU = "'+FichierRecu+'"');
      CommitTrans;
   except
      Rollback;
      Writeln(FRapport, 'Fichier '+St+' : traitement annulé : '+DateTimeToStr(Now));
      end;
   end;
MoveCurProgressForm(St);
end;
// d PT42
{$IFDEF PAIEGRH}
procedure TOF_PGMULEnvoiSocial.FListe_OnDblClick(Sender: TObject);
var
//    NEnvoi:integer;
    Transmission: TTRANSMISSION;
begin

//   NEnvoi := 0;
{$IFDEF EAGLCLIENT}
   TheMulQ:=TOB (TFMul(Ecran).Q.TQ);
   if TheMulQ.RecordCount=0 then exit; //grille vide
   TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row-1) ;

//   NEnvoi :=TheMulQ.detail[TFMul(Ecran).FListe.Row-1].GetValue('PES_CHRONOMESS');
 {$ELSE}
   TheMulQ := THQuery(Ecran.FindComponent('Q'));
   if TheMulQ.RecordCount=0 then exit; //grille vide
//   NEnvoi:=TheMulQ.findfield('PES_CHRONOMESS').asInteger;
{$ENDIF}
  Transmission.ModeEnvoi := RECUP;
  if (LeType = 'DUCS') then
    Transmission.teleprocedure := DUCS
  else
  if (LeType = 'DADSU') then
    Transmission.teleprocedure := DADSU;

  PConfig.RecupWS := 'True';
  Aff_RecepEnvoi_Web(Transmission,True,SiretEmett);

end;
{$ENDIF}
// f PT42
initialization
  registerclasses([TOF_PGMULEnvoiSocial]);
end.

