{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 30/10/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : PGEDIT_DADSPER ()
Mots clefs ... : TOF;PGEDIT_DADSPER
*****************************************************************
PT1   : 17/01/2002 JL V571 Edition récapitulative effectuée par établissements
PT2-1 : 25/01/2002 SB V571 Erreur dans la condition de selection des etabs. de
                           la requête
PT2-2 : 25/01/2002 SB V571 Récupération des critères Etab pour Fn @
PT3   : 31/01/2002 JL V571 Suppression des procédures inutilisées
PT3-1 : 31/01/2002 JL V571 Ajout de la gestion des organismes.
PT3-2 : 31/01/2002 JL V571 Modification requête pour édition récapitulative : le
                           calcul du nbre de période est effectué dans la
                           requête au lieu d'un champ SQL dans l'état.
PT3-3 : 31/01/2002 JL V571 Modification condition établissement
PT4   : 05/02/2002 JL V571 Ajout du champ pour libellé du dossier
PT4-2 : 18/04/2002 SB V571 Fiche de bug n°10087 : V_PGI_env.LibDossier non
                           renseigné en Mono
PT5   : 26/04/2002 VG V571 Par défaut, année précédente jusqu'au mois de
                           septembre. A partir du mois d'octobre, année en cours
                           (pour tests)
PT6   : 29/09/2002 VG V585 Edition de la DADS-U BTP
PT7   : 14/11/2002 VG V585 Edition des traitements et salaires DADSU
PT8   : 26/11/2002 JL V591 Ajout checkbox rupture par organisme
PT9   : 14/01/2003 VG V591 Edition des traitements et salaires par établissement
PT10  : 29/01/2003 VG V591 Affichage d'un message si l'exercice par défaut
                           n'existe pas - FQ N°10469
PT11  : 26/02/2003 VG V_42 Edition des périodes d'inactivité DADSU - FQ N° 10463
PT12  : 24/10/2003 VG V_42 Ajout de critères pour les états DADS-U
PT13  : 02/02/2004 JL V_50 Lancement de l'état avec TOB **** MODIFS EN COURS,
                           lancement mis en commentaire***********
PT14  : 17/03/2004 JL V_50 FQ 11097 Ajout Organisme de ... à ...
PT15  : 06/04/2004 VG V_50 Correction édition DADS-U BTP
PT16  : 14/04/2004 VG V_50 Ajout de l'année de référence au niveau de l'état
PT17  : 13/05/2004 JL V_50 FQ 11269 Edition récap : ne prenais que le montant de
                           la dernière période
PT18  : 18/05/2004 VG V_50 On désactive la liste d'exportation dans le cas d'une
                           DADS-U complète ou normale mais pas récapitulative
                           FQ N°10680
PT19  : 27/05/2004 VG V_50 Classement par ordre alphabétique des salariés pour
                           l'édition des traitements et salaires ainsi que pour
                           l'édition des périodes d'inactivité - FQ N°11098
PT20  : 15/06/2004 JL V_50 FQ 11269 tri alpha + libellé organisme pour édition
                           récapitulative
PT21  : 14/09/2004 JL V_50 FQ 11603 Tri alpha pour édition complète ou normal
PT22  : 17/09/2004 VG V_50 FQ N°11269 - Calcul du nombre de salarié pour
                           l'édition récapitulative
PT23  : 05/05/2004 JL V_60 Corrections null et numérique sous oracle
PT24  : 05/01/2005 JL V_60 Gestion raccourci F9 pour lancer etat
PT25  : 02/02/2005 VG V_60 "Travail à l'étranger ou frontalier" mal renseigné
                           FQ N°11957
PT26  : 01/03/2005 VG V_60 Remodelage complet de l'unité - FQ N°11778
PT27  : 07/10/2005 VG V_60 Adaptation cahier des charges DADS-U V8R02
PT28  : 12/10/2005 VG V_60 Correction éditions
PT29  : 17/10/2005 VG V_60 Ajout édition des données prud'hommales - FQ N°11816
PT30  : 02/11/2005 VG V_60 Lancement de l'état du dernier exercice connu même si
                           on se positionne sur un exercice posterieur
                           FQ N°12181
PT31-1: 09/11/2005 VG V_65 En période d'inactivité, le montant versé était
                           toujours à zéro - FQ N°11778
PT31-2: 09/11/2005 VG V_65 Mauvais formatage des dates pour l'édition des
                           périodes d'inactivité - FQ N°12679
PT32  : 22/11/2005 JL V_65 FQ 12714 Correction order by si tri alpha
PT33  : 30/11/2005 JL V_65 FQ 12717 Ajout clause établissement et correction
                           affichage adresse
PT34  : 27/12/2005 JL V_65 FQ 12771 Corrections requête segment pour traitement
                           et salaires. (= au lieu de <=)
PT35  : 03/01/2006 PH V_651 FQ 12281 Multi-périodes DADS-U
PT36-1: 16/01/2006 VG V_65 Edition du libellé au lieu du code - FQ 12788
PT36-2: 16/01/2006 VG V_65 Corrections multiples - FQ N°11778
PT37  : 15/03/2006 VG V_65 Edition des périodes d'inactivité pour la DADS-U BTP
PT38  : 10/04/2006 VG V_65 Correction PT30 - FQ N°12991
PT39-1: 13/10/2006 VG V_70 Remplacement des ellipsis et combo par des multi-val
                           combo
PT39-2: 13/10/2006 VG V_70 Adaptation cahier des charges DADS-U V8R03 
PT39-3: 13/10/2006 VG V_70 Edition du NIR sur l'édition récapitulative
                           FQ N°12883
PT39-4: 13/10/2006 VG V_70 Lorsque l'édition se fait à partir d'un fichier, le
                           libellé du coefficient édité est celui inscrit dans
                           le fichier - FQ N°12863
PT39-5: 13/10/2006 VG V_70 Edition des périodes d'inactivité par orde croissant
                           de date - FQ N°12945
PT40  : 31/10/2006 VG V_70 Inactivation de liste d'exportation pour l'édition
                           des périodes d'inactivité
PT41  : 15/11/2006 VG V_70 Possibilité de personnaliser l'état "Traitement et
                           salaires payés" - FQ N°13659
PT42  : 11/01/2007 VG V_72 Suppression des enregistrements ayant le champ
                           exercicedads mal alimenté - FQ N°13827
PT43  : 30/01/2007 FC V_72 Mise en place filtrage des
                           habilitations/poupulations
PT44  : 13/04/2007 MF V_72 FQ 13970 - Ajout impression des DADSU récapitulatives
PT45  : 05/11/2007 VG V_80 Adaptation cahier des charges V8R06 
PT46  : 28/11/2007 VG V_80 Gestion du champ ET_FICTIF - FQ N°13925
PT47  : 29/11/2007 VG V_80 Adaptation cahier des charges V8R06 - FQ N°14980
PT48  : 03/12/2007 VG V_80 Prise en compte du trimestre civil dans le cas d'une
                           déclaration trimestrielle - FQ N°13245
}
Unit PGEDIT_DADSPER_TOF ;

Interface
uses
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     StdCtrls,
     Controls,
     Classes,
     sysutils,
{$IFDEF EAGLCLIENT}
     eQRS1,
     MaineAgl,
{$ELSE}
     FE_Main,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     QRS1,
{$ENDIF}
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     ParamSoc,
     UTOB,
     EntPaie,
     PGoutils2,
     HStatus,
     UTobDebug,
     ed_tools,
     PgDADSCommun,
     HTB97,
     PgDADSOutils,
     UTOFPGEtats,// PT44
     P5Def //MonHabilitation
     ;


Type
  TOF_PGEDIT_DADSPER = Class (TOF_PGEtats)   // PT44
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnNew                    ; override ;
    procedure OnClose                  ; override ;
    private
    Car, THAnneeRef:THValCombobox;
    BTP, Complet, Honor, Inact, Recap:TCheckBox;
    Arg : String;
    TobEditRecap, TobPeriodes : Tob;

    function ChoisiCodeEtat(TypeEtat, NumEtat : string) : string;
    procedure ChoixEdition (Sender:TObject);
    procedure ChoixEditionFichier (Sender:TObject);
    procedure AccesRuptOrg(Sender:TObject);
    function ConstruireRequete(RuptOrganisme,RuptEtab,OrdreAlpha:Boolean):String;
    procedure ChangeRuptOrg(Sender:TObject);
    procedure ChangeRuptEtab(Sender:TObject);
    procedure ChangeDetailSal(Sender:TObject);
    function LancerEtatAvecTob(TypeEdition:string):string;
    function EditionFichierDads(TypeEdition:string):string;
    procedure CreationSegments(ExerciceDads:string; var T:tob; TobEntete, TobLexique:tob);
    procedure EnregistreSegment(TL : tob; ExerciceDads, Segment, ValeurSeg : string;var T : tob);
    procedure ChargerLexiqueDads (var TobLexique:tob; Annee:STring;
                                  Fichier:Boolean=False;TypeEdition:string='');
    procedure DateDecaleChange(Sender: TObject);
    procedure Parametrage(Sender: TObject);
  end ;

Implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 02/11/2005
Modifié le ... :   /  /    
Description .. : Cette fonction sélectionne le bon code état en fonction du 
Suite ........ : contenu de la base
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
function TOF_PGEDIT_DADSPER.ChoisiCodeEtat(TypeEtat, NumEtat : string) : string;
var
CodeEtat, MaxModele, StModele, StTypeEtat : string;
QModele : TQuery;
begin
CodeEtat:= TypeEtat+NumEtat;
//PT38
if (TypeEtat='PR') then
   StTypeEtat:= ' AND MO_CODE <> "PRU"'
else
if (TypeEtat='PC') then
   StTypeEtat:= ' AND MO_CODE <> "PCD" AND MO_CODE <> "PCF"'
else
   StTypeEtat:= '';
//FIN PT38
StModele:= 'SELECT MAX(MO_CODE) AS CODEETAT'+
           ' FROM MODELES WHERE'+
           ' MO_TYPE="E" AND'+
           ' MO_NATURE="PDA" AND'+
           ' MO_CODE LIKE "'+TypeEtat+'%"'+StTypeEtat;
QModele:= OpenSQL(StModele, True);
if not QModele.eof then
   begin
   MaxModele:= QModele.FindField ('CODEETAT').AsString;
   if (CodeEtat>=MaxModele) then
      CodeEtat:= MaxModele;
   end;
Ferme (QModele);
Result:= CodeEtat;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : JL
Créé le ...... : 30/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDIT_DADSPER.OnUpdate ;
var
ExerciceDads, NumEtat, TypeEdition : String;
begin
Inherited ;
FreeAndNil (TobEditRecap);
FreeAndNil (TobPeriodes);
if (arg='I') then
   SetControlChecked('INACT', True);
TypeEdition:= '';
if (arg='S') then
   TypeEdition:= 'S'
else
if (GetCheckBoxState('RECAP')=CbChecked) then
   TypeEdition:= 'R'
else
if (GetCheckBoxState('CBTP')=CbChecked) then
   TypeEdition:= 'B'
else
if (GetCheckBoxState('INACT')=CbChecked) then
   TypeEdition:= 'I'
else
if (GetCheckBoxState('HONOR')=CbChecked) then
   TypeEdition:= 'H';

If (arg = 'FICHIER') then
   ExerciceDads:= EditionFichierDads(TypeEdition)
else
   ExerciceDads:= LancerEtatAvecTob(TypeEdition);

if (ExerciceDads<>'') then
   begin
   NumEtat := Copy(ExerciceDads,4,1);

   TFQRS1(Ecran).TypeEtat:= 'E';
   if (TypeEdition='') then
      begin
      TFQRS1(Ecran).LaTob:= TobPeriodes;
      If GetCheckBoxState('CCOMPLET') = CbChecked then
         TFQRS1(Ecran).CodeEtat:= ChoisiCodeEtat('PC', NumEtat)
      else
         TFQRS1(Ecran).CodeEtat:= ChoisiCodeEtat('PP', NumEtat);
      end
   else
   if (TypeEdition='R') then
      begin
      TFQRS1(Ecran).CodeEtat:= ChoisiCodeEtat('PR', NumEtat);
      If (arg = 'FICHIER') then
         TFQRS1(Ecran).LaTob:= TobEditRecap
      else
         TFQRS1(Ecran).LaTob:= TobPeriodes;
      end
   else
   if (TypeEdition='B') then
      begin
      TFQRS1(Ecran).CodeEtat:= ChoisiCodeEtat('PB', NumEtat);
      TFQRS1(Ecran).LaTob:= TobPeriodes;
      end
   else
   if (TypeEdition='I') then
      begin
      TFQRS1(Ecran).CodeEtat:= ChoisiCodeEtat('PI', NumEtat);
      TFQRS1(Ecran).LaTob:= TobPeriodes;
      end
   else
   if (TypeEdition='H') then
      begin
      TFQRS1(Ecran).CodeEtat:= ChoisiCodeEtat('PH', NumEtat);
      TFQRS1(Ecran).LaTob:= TobPeriodes;
      end
   else
   if (TypeEdition='S') then
{PT41
      begin
      TFQRS1 (Ecran).CodeEtat:= ChoisiCodeEtat('PS', NumEtat);
      TFQRS1 (Ecran).LaTob:= TobPeriodes;
      end;
}
      TFQRS1 (Ecran).LaTob:= TobPeriodes;
//FIN PT41
   end
else
   PGIBox('Exercice non renseigné', Ecran.Caption);
//TobDebug (TobEditRecap);
//TobDebug (TobPeriodes);
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JL
Créé le ...... : 30/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDIT_DADSPER.OnArgument (S : String ) ;
var
AnneeE, AnneePrec, MoisE, ComboExer, StPlus : string;
JourJ : TDateTime;
AnneeA, Jour, MoisM : Word;
CDetailSal, CRupt:TCheckBox;
ModifDateDecale : TToolbarButton97;
Annee : THValComboBox;
begin
Inherited ;
// d PT44
if (Pos('CHAINES', S) > 0) then
begin
  Arg :='N';
  SetControlChecked('RECAP', True);
end
else
  Arg := S;
// f PT44

Recap:=TCheckBox(GetControl('RECAP'));
If Recap<>NIL Then
   Recap.OnClick:= ChoixEdition;
Complet:=TCheckBox(GetControl('CCOMPLET'));
If Complet<>NIL Then Complet.OnClick:=ChoixEdition;

BTP:=TCheckBox(GetControl('CBTP'));
If BTP<>NIL Then
   BTP.OnClick:=ChoixEdition;

if (Ecran.Name<>'EDITFICHIERDADS') then
   begin
   SetControlText('DOSSIER', GetParamSoc ('SO_LIBELLE'));
{PT39-1
   RecupMinMaxTablette ('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
   SetControlText ('ETAB1', Min);
   SetControlText ('ETAB2', Max);
   RecupMinMaxTablette ('PG', 'SALARIES', 'PSA_SALARIE', Min, Max);
   SetControlText ('Sal1', Min);
   SetControlText ('Sal2', Max);
   THSal1:= ThEdit (getcontrol('SAL1'));
   If THSal1<>nil then
      THSal1.OnExit:= ExitEdit;
   THSal2:= ThEdit (getcontrol('SAL2'));
   If THSal2<>nil then
      THSal2.OnExit:= ExitEdit;
}      
   Recap:= TCheckBox (GetControl('RECAP'));
   If Recap<>Nil Then
      Recap.OnClick:= AccesRuptOrg;
   CRupt:= TCheckBox (GetControl('CRUPTORG'));
   If CRupt<>Nil then
      CRupt.Onclick:= ChangeRuptOrg;
   CRupt:= TCheckBox (GetControl('CRUPTETAB'));
   If CRupt<>Nil then
      CRupt.OnClick:= ChangeRuptEtab;
   CDetailSal:= TCheckBox (GetControl('CDETAILSAL'));
   If CDetailSal<>Nil then
      CDetailSal.OnClick:= ChangeDetailSal;
{PT39-1
   Defaut:= THEdit (GetControl('ORGANISME'));
   If Defaut<>Nil then
      Defaut.OnElipsisClick:= OrgElipsisClick;
   Defaut:= THEdit (GetControl('ORGANISME1'));
   If Defaut<>Nil then
      Defaut.OnElipsisClick:= OrgElipsisClick;
}      
   Car:= THValComboBox(GetControl('CCAR'));
//FIN PT39-1
   THAnneeRef:= THValCombobox(GetControl('ANNEE'));
   JourJ:= Date;
   DecodeDate (JourJ, AnneeA, MoisM, Jour);
   if (Arg='B') then
      AnneePrec:= IntToStr(AnneeA)
   else
      begin
      if MoisM>9 then
         AnneePrec:= IntToStr(AnneeA)
      else
         AnneePrec:= IntToStr(AnneeA-1);
      end;

{PT39-2
   if RendExerSocialPrec (MoisE, AnneeE, ComboExer, Deb, Fin, AnneePrec)=TRUE then
      begin
      THAnneeRef.Value:= ComboExer;
      SetControlText ('DATEDEB', DateTostr(deb));
      SetControlText ('DATEFIN', DateTostr(fin));
}
   if RendExerSocialPrec (MoisE, AnneeE, ComboExer, DebExer, FinExer, AnneePrec)=TRUE then
      begin
      THAnneeRef.Value:= ComboExer;
// DEB PT35
      PGExercice := AnneeE;
      if ((Arg='B') and (PGExercice <> '')) then
         begin
         DebExer:= StrToDate('01/04/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer:= StrToDate('31/03/'+PGExercice);
         end;
      SetControlText ('DATEDEB', DateTostr(DebExer));
      SetControlText ('DATEFIN', DateTostr(FinExer));
//FIN PT39-2
// FIN PT35
      end
   else
      PGIBox('L''exercice '+AnneePrec+' n''existe pas', Ecran.Caption);
   end
else
   begin
   Honor:= TCheckBox (GetControl('HONOR'));
   If (Honor<>NIL) Then
      Honor.OnClick:= ChoixEditionFichier;

   Inact:= TCheckBox (GetControl('INACT'));
   If (Inact<>NIL) Then
      Inact.OnClick:= ChoixEditionFichier;
   end;

if ((Arg = 'S') or (Arg='I')) then
   begin
   SetControlVisible('CRUPTETAB',False);
   SetControlVisible('CRUPTORG',False);
   SetControlVisible('TORGANISME',FALSE);
//   SetControlVisible('TORGANISME1',FALSE);      PT39-1
   SetControlVisible('ORGANISME',FALSE);
//   SetControlVisible('ORGANISME1',FALSE);       PT39-1
   SetControlVisible('RECAP',FALSE);
   SetControlVisible('CCOMPLET',FALSE);
   SetControlVisible('CDETAILSAL',False);
   end
Else
   begin
   SetControlEnabled('CRUPTORG',False);
   SetControlEnabled('CRUPTETAB',False);
   SetControlEnabled('CDETAILSAL',False);
   end;

if (Arg<>'B') then
   begin
   if (Arg<>'FICHIER') then
      SetControlVisible ('CBTP', False);
   TFQRS1(Ecran).Caption:='Edition DADS-U';
//PT39-1
   if Car <> nil then
      begin
      Car.Enabled:= False;
      Car.Visible:= False;
      Car.OnChange:= Parametrage;
      Car.Value:= 'A00';
      end;
   SetControlVisible ('LCAR', False);
//FIN PT39-1
   end
else
//PT37
   begin
   SetControlVisible ('INACT', True);
   SetControlEnabled ('INACT', True);
   TFQRS1(Ecran).Caption:='Edition DADS-U BTP';
//PT39-1
   if Car <> nil then
      begin
      Car.Enabled:= True;
      Car.OnChange:= Parametrage;
      Car.Value:= 'A00';
      end;
   SetControlVisible ('LCAR', True);
//FIN PT39-1
   end;
//FIN PT37
UpdateCaption(TFQRS1(Ecran)) ;

{PT40
if ((arg='N') or (arg='B') or (arg='FICHIER')) then
}
if ((arg='B')  or (arg='I') or (arg='N') or (arg='FICHIER')) then
   begin
   SetControlEnabled('FListe',False);
   SetControlChecked('FListe',False);
   end
else
   SetControlEnabled('FListe',True);

//PT39-2
// Gestion du bouton de modification des dates
ModifDateDecale:= TToolbarButton97(GetControl('BMODIFDATE'));
if ModifDateDecale<>nil then
   ModifDateDecale.OnClick:= DateDecaleChange;

Annee:= THValComboBox (GetControl ('ANNEE'));
If Annee <> NIL then
   Annee.OnChange := Parametrage;
//FIN PT39-2
//PT41
if (Arg='S') then
   begin
   TFQRS1 (Ecran).ChoixEtat:= True;
   TFQRS1 (Ecran).ParamEtat:= True;
   end;
//FIN PT41

MajExercice;                  //PT42

//PT46
StPlus:= ' WHERE ET_FICTIF<>"X"';
SetControlProperty ('ETAB', 'Plus', StPlus);
//FIN PT46
end;

//PT41
{***********A.G.L.Privé.*****************************************
Auteur  ...... : VG
Créé le ...... : 15/11/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDIT_DADSPER.OnNew;
var
NumEtat : string;
begin
Inherited ;
if (Arg='S') then
   begin
   NumEtat:= Copy(PGExercice, 4 ,1);
   TFQRS1 (Ecran).CodeEtat:= ChoisiCodeEtat('PS', NumEtat);
   SetControlText ('FEtat', TFQRS1 (Ecran).CodeEtat);
   end;
end;
//FIN PT41


procedure TOF_PGEDIT_DADSPER.OnClose ;
begin
Inherited ;
FreeAndNil (TobEditRecap);
FreeAndNil (TobPeriodes);
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JL
Créé le ...... : 30/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDIT_DADSPER.ChoixEdition (Sender:TObject);
begin
If (TCheckBox(Sender)=Complet) Then
   begin
   If (Complet.Checked=True) Then
      begin
      SetControlEnabled('CBTP',False);
      SetControlEnabled('RECAP',False);
      SetControlChecked('CBTP', False);
      SetControlChecked('RECAP', False);
      if (Ecran.Name='EDITFICHIERDADS') then
         begin
         SetControlEnabled('INACT',False);
         SetControlEnabled('HONOR',False);
         SetControlChecked('INACT', False);
         SetControlChecked('HONOR', False);
         end;
      end
   else
      begin
      SetControlEnabled('CBTP',True);
      SetControlEnabled('RECAP',True);
      if (Ecran.Name='EDITFICHIERDADS') then
         begin
         SetControlEnabled('INACT',True);
         SetControlEnabled('HONOR',True);
         end;
      end;
   end;

If (TCheckBox(Sender)=Recap) Then
   begin
   If (Recap.Checked=True) Then
      begin
      SetControlEnabled('CBTP',False);
      SetControlEnabled('CCOMPLET',False);
      SetControlChecked('CBTP', False);
      SetControlChecked('CCOMPLET', False);
      if (Ecran.Name='EDITFICHIERDADS') then
         begin
         SetControlEnabled('INACT',False);
         SetControlEnabled('HONOR',False);
         SetControlChecked('INACT', False);
         SetControlChecked('HONOR', False);
         end;
      end
   else
      begin
      SetControlEnabled('CBTP',True);
      SetControlEnabled('CCOMPLET',True);
      if (Ecran.Name='EDITFICHIERDADS') then
         begin
         SetControlEnabled('INACT',True);
         SetControlEnabled('HONOR',True);
         end;
      end;
   end;

If (TCheckBox(Sender)=BTP) Then
   begin
   If (BTP.Checked=True) Then
      begin
      SetControlEnabled('CCOMPLET',False);
      SetControlEnabled('RECAP',False);
      if (Ecran.Name='EDITFICHIERDADS') then
         begin
         SetControlEnabled('INACT',False);
         SetControlEnabled('HONOR',False);
         SetControlChecked('INACT', False);
         SetControlChecked('HONOR', False);
         end;
      SetControlEnabled('FListe',True);
      end
   else
      begin
      SetControlEnabled('CCOMPLET',True);
      SetControlEnabled('RECAP',True);
      if (Ecran.Name='EDITFICHIERDADS') then
         begin
         SetControlEnabled('INACT',True);
         SetControlEnabled('HONOR',True);
         end;
      SetControlEnabled('FListe',False);
      SetControlChecked('FListe',False);
      end;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 08/02/2005
Modifié le ... :   /  /
Description .. : Gestion des coches propres à la fiche EDITFICHIERDADS
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGEDIT_DADSPER.ChoixEditionFichier (Sender:TObject);
begin
If (TCheckBox(Sender)=Honor) Then
   begin
   If (Honor.Checked=True) Then
      begin
      SetControlEnabled('CCOMPLET',False);
      SetControlEnabled('RECAP',False);
      SetControlEnabled('CBTP',False);
      SetControlEnabled('INACT',False);
      SetControlChecked('CCOMPLET', False);
      SetControlChecked('RECAP', False);
      SetControlChecked('CBTP', False);
      SetControlChecked('INACT', False);
      end
   else
      begin
      SetControlEnabled('CCOMPLET',True);
      SetControlEnabled('RECAP',True);
      SetControlEnabled('CBTP',True);
      SetControlEnabled('INACT',True);
      end;
   end;

If (TCheckBox(Sender)=Inact) Then
   begin
   If (Inact.Checked=True) Then
      begin
      SetControlEnabled('CCOMPLET',False);
      SetControlEnabled('RECAP',False);
      SetControlEnabled('CBTP',False);
      SetControlEnabled('HONOR',False);
      SetControlChecked('CCOMPLET', False);
      SetControlChecked('RECAP', False);
      SetControlChecked('CBTP', False);
      SetControlChecked('HONOR', False);
      end
   else
      begin
      SetControlEnabled('CCOMPLET',True);
      SetControlEnabled('RECAP',True);
      SetControlEnabled('CBTP',True);
      SetControlEnabled('HONOR',True);
      end;
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : JL
Créé le ...... : 30/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDIT_DADSPER.AccesRuptOrg(Sender:TObject);
begin
if TCheckBox(Sender).Checked=True then
   begin
   SetControlEnabled('CRUPTORG',True);
   SetControlEnabled('CRUPTETAB',True);
   SetControlEnabled('CCOMPLET',False);
   SetControlEnabled('CBTP',False);
   SetControlEnabled('ORGANISME',False);
{PT39-1
   SetControlEnabled('ORGANISME1',False);
}
   if (arg<>'FICHIER') then
      SetControlText ('ORGANISME','');
{PT39-1
      SetControlText('ORGANISME1','');
}
   SetControlEnabled('CDETAILSAL',True);
{PT39-1
   SetControlenabled('ETAB1',False);
   SetControlenabled('ETAB2',False);
   RecupMinMaxTablette('PG','ETABLISS','ET_ETABLISSEMENT',Min,Max);
   SetControlText('ETAB1',Min);
   SetControlText('ETAB2',Max);
}
   SetControlenabled('ETAB',False);
   SetControlText('ETAB', '');
//FIN PT39-1
   SetControlEnabled('FListe',True);
   end
Else
   begin
   SetControlEnabled('CRUPTORG',False);
   SetControlChecked('CRUPTORG',False);
   SetControlEnabled('CRUPTETAB',False);
   SetControlChecked('CRUPTETAB',False);
   SetControlEnabled('CCOMPLET',True);
   SetControlEnabled('CBTP',True);
   SetControlChecked('CDETAILSAL',False);
   SetControlEnabled('CDETAILSAL',False);
{PT39-1
   SetControlEnabled('ORGANISME1',True);
   SetControlEnabled('ETAB1',True);
   SetControlEnabled('ETAB2',True);
}
   SetControlEnabled('ORGANISME',True);
   SetControlEnabled('ETAB',True);
//FIN PT39-1
   SetControlChecked('FListe',False);
   SetControlEnabled('FListe',False);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JL
Créé le ...... : 30/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Function TOF_PGEDIT_DADSPER.ConstruireRequete(RuptOrganisme,RuptEtab,OrdreAlpha:Boolean):String;
var Requete:String;
begin
If RuptOrganisme=true then
   begin
   If RuptEtab=True then
      begin
      If OrdreAlpha=True then
         begin
         requete:='SELECT DISTINCT (D2.PDS_DONNEEAFFICH) as DONNEE,'+
                  ' COUNT(PDE_ORDRE) ORDRE, ET_ETABLISSEMENT, PDE_SALARIE,'+
                  ' PSA_LIBELLE, PSA_PRENOM'+
                  ' FROM ETABLISS'+
                  ' LEFT JOIN DADSDETAIL D1 ON'+
                  ' D1.PDS_DONNEEAFFICH=ET_ETABLISSEMENT AND'+
                  ' D1.PDS_SEGMENT="S41.G01.00.005"'+
                  ' LEFT JOIN DADSDETAIL D2 ON'+
                  ' D1.PDS_SALARIE=D2.PDS_SALARIE AND'+
                  ' D1.PDS_ORDRE=D2.PDS_ORDRE AND'+
                  ' D1.PDS_TYPE=D2.PDS_TYPE AND'+
                  ' D1.PDS_DATEDEBUT=D2.PDS_DATEDEBUT AND'+
                  ' D1.PDS_DATEFIN=D2.PDS_DATEFIN'+
                  ' LEFT JOIN DADSPERIODES ON'+
                  ' PDE_SALARIE=D1.PDS_SALARIE AND'+
                  ' PDE_ORDRE=D1.PDS_ORDRE AND'+
                  ' PDE_TYPE=D1.PDS_TYPE AND'+
                  ' PDE_DATEDEBUT=D1.PDS_DATEDEBUT AND'+
                  ' PDE_DATEFIN=D1.PDS_DATEFIN'+
                  ' LEFT JOIN SALARIES ON'+
                  ' D1.PDS_SALARIE=PSA_SALARIE WHERE'+
                  ' PDE_SALARIE NOT LIKE "--H%" AND'+
                  ' D2.PDS_SEGMENT="S41.G01.01.001"';
         if (Arg='B') then
            requete:=requete+' AND PDE_TYPE<>"001" AND PDE_TYPE<>"201"'
         else
            requete:=requete+' AND (PDE_TYPE="001" OR PDE_TYPE="201")';
         end
      Else
         begin
         requete:='SELECT DISTINCT (D2.PDS_DONNEEAFFICH) as DONNEE,'+
                  ' COUNT(PDE_ORDRE) ORDRE, ET_ETABLISSEMENT, PDE_SALARIE,'+
                  ' D1.PDS_DONNEEAFFICH'+
                  ' FROM ETABLISS'+
                  ' LEFT JOIN DADSDETAIL D1 ON'+
                  ' D1.PDS_DONNEEAFFICH=ET_ETABLISSEMENT AND'+
                  ' D1.PDS_SEGMENT="S41.G01.00.005"'+
                  ' LEFT JOIN DADSDETAIL D2 ON'+
                  ' D1.PDS_SALARIE=D2.PDS_SALARIE AND'+
                  ' D1.PDS_ORDRE=D2.PDS_ORDRE AND'+
                  ' D1.PDS_TYPE=D2.PDS_TYPE AND'+
                  ' D1.PDS_DATEDEBUT=D2.PDS_DATEDEBUT AND'+
                  ' D1.PDS_DATEFIN=D2.PDS_DATEFIN'+
                  ' LEFT JOIN DADSPERIODES ON'+
                  ' PDE_SALARIE=D1.PDS_SALARIE AND'+
                  ' PDE_ORDRE=D1.PDS_ORDRE AND'+
                  ' PDE_TYPE=D1.PDS_TYPE AND'+
                  ' PDE_DATEDEBUT=D1.PDS_DATEDEBUT AND'+
                  ' PDE_DATEFIN=D1.PDS_DATEFIN WHERE'+
                  ' PDE_SALARIE NOT LIKE "--H%" AND'+
                  ' D2.PDS_SEGMENT="S41.G01.01.001"';
         if (Arg='B') then
            requete:=requete+' AND PDE_TYPE<>"001" AND PDE_TYPE<>"201"'
         else
            requete:=requete+' AND (PDE_TYPE="001" OR PDE_TYPE="201")';
         end;
      end
   Else
      begin
      If OrdreAlpha=True then
         begin
         requete:='SELECT DISTINCT (D2.PDS_DONNEEAFFICH) as DONNEE,'+
                  ' COUNT(PDE_ORDRE) ORDRE, PDE_SALARIE,'+
                  ' PSA_LIBELLE, PSA_PRENOM'+
                  ' FROM DADSDETAIL D2'+
                  ' LEFT JOIN DADSPERIODES ON'+
                  ' PDE_SALARIE=D2.PDS_SALARIE AND'+
                  ' PDE_ORDRE=D2.PDS_ORDRE AND'+
                  ' PDE_TYPE=D2.PDS_TYPE AND'+
                  ' PDE_DATEDEBUT=D2.PDS_DATEDEBUT AND'+
                  ' PDE_DATEFIN=D2.PDS_DATEFIN AND'+
                  ' PDE_EXERCICEDADS=D2.PDS_EXERCICEDADS'+
                  ' LEFT JOIN SALARIES ON'+
                  ' D2.PDS_SALARIE=PSA_SALARIE WHERE'+
                  ' PDE_SALARIE NOT LIKE "--H%" AND'+
                  ' D2.PDS_SEGMENT="S41.G01.01.001"';
         if (Arg='B') then
            requete:=requete+' AND PDE_TYPE<>"001" AND PDE_TYPE<>"201"'
         else
            requete:=requete+' AND (PDE_TYPE="001" OR PDE_TYPE="201")';
         end
      Else
         begin
         requete:='SELECT DISTINCT (D2.PDS_DONNEEAFFICH) as DONNEE,'+
                  ' COUNT(PDE_ORDRE) ORDRE, PDE_SALARIE'+
                  ' FROM DADSDETAIL D2'+
                  ' LEFT JOIN DADSPERIODES ON'+
                  ' PDE_SALARIE=D2.PDS_SALARIE AND'+
                  ' PDE_ORDRE=D2.PDS_ORDRE AND'+
                  ' PDE_TYPE=D2.PDS_TYPE AND'+
                  ' PDE_DATEDEBUT=D2.PDS_DATEDEBUT AND'+
                  ' PDE_DATEFIN=D2.PDS_DATEFIN AND'+
                  ' PDE_EXERCICEDADS=D2.PDS_EXERCICEDADS WHERE'+
                  ' PDE_SALARIE NOT LIKE "--H%" AND'+
                  ' D2.PDS_SEGMENT="S41.G01.01.001"';
         if (Arg='B') then
            requete:=requete+' AND PDE_TYPE<>"001" AND PDE_TYPE<>"201"'
         else
            requete:=requete+' AND (PDE_TYPE="001" OR PDE_TYPE="201")';
         end;
      end;
   end
Else
    begin
    If RuptEtab=True then
      begin
      If OrdreAlpha=True then
         begin
         requete:='SELECT DISTINCT '+
                  ' COUNT(PDE_ORDRE) ORDRE, ET_ETABLISSEMENT, PDE_SALARIE,'+
                  ' PSA_LIBELLE, PSA_PRENOM'+
                  ' FROM ETABLISS'+
                  ' LEFT JOIN DADSDETAIL D1 ON'+
                  ' D1.PDS_DONNEEAFFICH=ET_ETABLISSEMENT AND'+
                  ' D1.PDS_SEGMENT="S41.G01.00.005"'+
                  ' LEFT JOIN DADSPERIODES ON'+
                  ' PDE_SALARIE=D1.PDS_SALARIE AND'+
                  ' PDE_ORDRE=D1.PDS_ORDRE AND'+
                  ' PDE_TYPE=D1.PDS_TYPE AND'+
                  ' PDE_DATEDEBUT=D1.PDS_DATEDEBUT AND'+
                  ' PDE_DATEFIN=D1.PDS_DATEFIN'+
                  ' LEFT JOIN SALARIES ON'+
                  ' D1.PDS_SALARIE=PSA_SALARIE WHERE'+
                  ' PDE_SALARIE NOT LIKE "--H%"';
         if (Arg='B') then
            requete:=requete+' AND PDE_TYPE<>"001" AND PDE_TYPE<>"201"'
         else
            requete:=requete+' AND (PDE_TYPE="001" OR PDE_TYPE="201")';
         end
      Else
         begin
         requete:='SELECT DISTINCT '+
                  ' COUNT(PDE_ORDRE) ORDRE, ET_ETABLISSEMENT, PDE_SALARIE,'+
                  ' D1.PDS_DONNEEAFFICH'+
                  ' FROM ETABLISS'+
                  ' LEFT JOIN DADSDETAIL D1 ON'+
                  ' D1.PDS_DONNEEAFFICH=ET_ETABLISSEMENT AND'+
                  ' D1.PDS_SEGMENT="S41.G01.00.005"'+
                  ' LEFT JOIN DADSPERIODES ON'+
                  ' PDE_SALARIE=D1.PDS_SALARIE AND'+
                  ' PDE_ORDRE=D1.PDS_ORDRE AND'+
                  ' PDE_TYPE=D1.PDS_TYPE AND'+
                  ' PDE_DATEDEBUT=D1.PDS_DATEDEBUT AND'+
                  ' PDE_DATEFIN=D1.PDS_DATEFIN WHERE'+
                  ' PDE_SALARIE NOT LIKE "--H%"';
         if (Arg='B') then
            requete:=requete+' AND PDE_TYPE<>"001" AND PDE_TYPE<>"201"'
         else
            requete:=requete+' AND (PDE_TYPE="001" OR PDE_TYPE="201")';
         end;
      end
    Else
      begin
      If OrdreAlpha=True then
         begin
         requete:='SELECT DISTINCT '+
                  ' COUNT(PDE_ORDRE) ORDRE,  PDE_SALARIE,'+
                  ' PSA_LIBELLE, PSA_PRENOM'+
                  ' FROM DADSDETAIL D1'+
                  ' LEFT JOIN DADSPERIODES ON'+
                  ' PDE_SALARIE=D1.PDS_SALARIE AND'+
                  ' PDE_ORDRE=D1.PDS_ORDRE AND'+
                  ' PDE_TYPE=D1.PDS_TYPE AND'+
                  ' PDE_DATEDEBUT=D1.PDS_DATEDEBUT AND'+
                  ' PDE_DATEFIN=D1.PDS_DATEFIN AND'+
                  ' PDE_EXERCICEDADS=D1.PDS_EXERCICEDADS'+
                  ' LEFT JOIN SALARIES ON'+
                  ' D1.PDS_SALARIE=PSA_SALARIE WHERE'+
                  ' D1.PDS_SEGMENT="S41.G01.00.005" AND'+
                  ' PDE_SALARIE NOT LIKE "--H%"';
         if (Arg='B') then
            requete:=requete+' AND PDE_TYPE<>"001" AND PDE_TYPE<>"201"'
         else
            requete:=requete+' AND (PDE_TYPE="001" OR PDE_TYPE="201")';
         end
      Else
         begin
         requete:='SELECT DISTINCT '+
                  ' COUNT(PDE_ORDRE) ORDRE, PDE_SALARIE'+
                  ' FROM DADSDETAIL D1'+
                  ' LEFT JOIN DADSPERIODES ON'+
                  ' PDE_SALARIE=D1.PDS_SALARIE AND'+
                  ' PDE_ORDRE=D1.PDS_ORDRE AND'+
                  ' PDE_TYPE=D1.PDS_TYPE AND'+
                  ' PDE_DATEDEBUT=D1.PDS_DATEDEBUT AND'+
                  ' PDE_DATEFIN=D1.PDS_DATEFIN AND'+
                  ' PDE_EXERCICEDADS=D1.PDS_EXERCICEDADS WHERE'+
                  ' D1.PDS_SEGMENT="S41.G01.00.005" AND'+
                  ' PDE_SALARIE NOT LIKE "--H%"';
         if (Arg='B') then
            requete:=requete+' AND PDE_TYPE<>"001" AND PDE_TYPE<>"201"'
         else
            requete:=requete+' AND (PDE_TYPE="001" OR PDE_TYPE="201")';
         end;
      end;
    end;
Result:= Requete;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JL
Créé le ...... : 30/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDIT_DADSPER.ChangeRuptOrg(Sender:TObject);
begin
If GetCheckBoxState('CRUPTORG')=CbChecked then
   begin
   SetControlEnabled('CDETAILSAL',True);
   SetControlEnabled('ORGANISME',True);
//   SetControlEnabled('ORGANISME1',True);    PT39-1
   end
Else
   begin
   If GetCheckBoxState('CRUPTETAB')=CbChecked then SetControlEnabled('CDETAILSAL',True)
   Else
       begin
       SetControlEnabled('CDETAILSAL',False);
       SetControlChecked('CDETAILSAL',False);
       end;
   SetControlEnabled('ORGANISME',False);
//   SetControlEnabled('ORGANISME1',False);  PT39-1
   SetControlText('ORGANISME','');
//   SetControlText('ORGANISME1','');        PT39-1
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JL
Créé le ...... : 30/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDIT_DADSPER.ChangeRuptEtab(Sender:TObject);
begin
if GetCheckBoxState('CRUPTETAB')=CbChecked then
   begin
   SetControlEnabled('CDETAILSAL',True);
{PT39-1
   SetControlenabled('ETAB1',True);
   SetControlenabled('ETAB2',True);
}
   SetControlenabled('ETAB',True);
   end
Else
   begin
   If GetCheckBoxState('CRUPTORG')=CbChecked then
      SetControlEnabled('CDETAILSAL',True)
   Else
      begin
      SetControlEnabled('CDETAILSAL',False);
      SetControlChecked('CDETAILSAL',False);
      end;
{PT39-1
   SetControlenabled('ETAB1',False);
   SetControlenabled('ETAB2',False);
   RecupMinMaxTablette('PG','ETABLISS','ET_ETABLISSEMENT',Min,Max);
   SetControlText('ETAB1',Min);
   SetControlText('ETAB2',Max);
}
   SetControlenabled('ETAB',False);
   SetControlText('ETAB', '');
//FIN PT39-1
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JL
Créé le ...... : 15/06/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDIT_DADSPER.ChangeDetailSal(Sender:TObject);
begin
if GetCheckBoxState('CDETAILSAL')=CbChecked then
   begin
   SetControlEnabled('TRISAL', False);
   SetControlChecked('TRISAL', False);
   end
Else
   SetControlEnabled('TRISAL', True);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JL
Créé le ...... : 30/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_PGEDIT_DADSPER.LancerEtatAvecTob(TypeEdition:string):string;
var
Annee, AnneeLexique, CDate, CEtab, COrganisme, CSal1, CSal2 : String;
Nature, OrderBy, OrdreSeg, OrdreSegS, Organisme : String;
Req, Requete, Salarie, Segment, SQL, SQLRecap : String;
SQLRecapEtab, SQLRecapOrgJoin, SQLrecapOrgWhere, Tablette, ValeurSeg : String;
EtabSelect, MCEtab, MCSal : String;
Q : TQuery;
T, TD, TEtabDADS, TEtabDADSD, TListeSalarie, TobDetail, TobLexique, TSal : Tob;
TSalD : Tob;
i,NumOrdre,OrdreSegI,x : Integer;
TriSal : TCheckBox;
MtSegment : Double;
Habilitation:String;//PT43
begin
result:= '';
OrderBy:= '';
SetControlText ('XX_RUPTURE1','');
SetControlText ('XX_RUPTURE2','');
{PT39-1
PGGlbOrg:= '';
Etab1:= GetControlText('ETAB1');
PGGlbEtabDe:= Etab1;
Etab2:= GetControlText('ETAB2');
PGGlbEtabA:= Etab2;
Sal1:= GetControlText('SAL1');
PGGlbSalDe:= Sal1;
Sal2:= GetControlText('SAL2');
PGGlbSalA:= Sal2;
}
TriSal:= TCheckBox(GetControl('TRISAL'));
Organisme:= GetControlText('ORGANISME');
//Organisme1:= GetControlText('ORGANISME1'); PT39-1
if (THAnneeRef<>nil) then
   Annee:= THAnneeRef.Value;
AnneeLexique:= RechDom ('PGANNEESOCIALE', GetControlText('ANNEE'), False);

{PT39-2
DateDeb:= IDate1900;
DateFin:= IDate1900;
if (Arg='B') then
   begin
   AnneeLexique:= IntToStr(StrToInt(AnneeLexique)-1);
   QExer:= OpenSQL('SELECT PEX_ANNEEREFER FROM EXERSOCIAL WHERE PEX_EXERCICE="'+Annee+'"',True);
   if not QExer.eof then
      begin
      PGExercice:= QExer.FindField ('PEX_ANNEEREFER').AsString;
      DateDeb:= StrToDate('01/04/'+IntToStr(StrToInt(PGExercice)-1));
      DateFin:= StrToDate('31/03/'+PGExercice);
      end
   Else
      PGIBox ('Il n''existe aucun exercice pour l''année choisie',Ecran.Caption);
   PGGlbTypeDADS:= '002';
   end
else
   begin
   QExer:= OpenSQL('SELECT PEX_DATEDEBUT,PEX_DATEFIN FROM EXERSOCIAL WHERE PEX_EXERCICE="'+Annee+'"',True);
   if not QExer.eof then
      begin
      DateDeb:= QExer.FindField('PEX_DATEDEBUT').AsDateTime;
      DateFin:= QExer.FindField('PEX_DATEFIN').AsDateTime;
      if (DateDeb <> StrToDate (GetControltext ('DATEDEB'))) then DateDeb := StrToDate (GetControltext ('DATEDEB'));
      if (DateFin <> StrToDate (GetControltext ('DATEFIN'))) then DateFin := StrToDate (GetControltext ('DATEFIN'));
      end
   Else
      PGIBox ('Il n''existe aucun exercice pour l''année choisie',Ecran.Caption);
   PGGlbTypeDADS:= '001';
   end;
if QExer <> NIL then Ferme (QExer);
}
result:= AnneeLexique;

CDate:= ' AND PDE_EXERCICEDADS="'+PGExercice+'"';

TEtabDADS:= TOB.Create ('Les etablissements', nil, -1); //PT46

if (TypeEdition='I') then
   begin
{PT39-1
   If Sal1<>'' Then
      CSal1:= ' AND PDE_SALARIE>="'+Sal1+'"';
   If Sal2<>'' Then
      CSal2:= ' AND PDE_SALARIE<="'+Sal2+'"';
}
   if (THMultiValCombobox(GetControl('SAL')).Tous) then
      CSal1:= ''
   else
      begin
      MCSal:= GetControlText('SAL');
      Salarie:= ReadTokenpipe(MCSal, ';');
      CSal1:= ' AND (';
      While (Salarie <> '') do
            begin
            CSal1:= CSal1+' PDE_SALARIE="'+Salarie+'"';
            Salarie := ReadTokenpipe(MCSal,';');
            if (Salarie <> '') then
               CSal1:= CSal1+' OR'
            else
               CSal1:= CSal1+')';
            end;
      end;
//FIN PT39-1
   end
else
if (GetCheckBoxState('CBTP')=cbChecked) then
   begin
{PT39-1
   If Sal1<>'' Then
      CSal1:= ' AND PDS_SALARIE>="'+Sal1+'"';
   If Sal2<>'' Then
      CSal2:= ' AND PDS_SALARIE<="'+Sal2+'"';
}
   if (THMultiValCombobox(GetControl('SAL')).Tous) then
      CSal1:= ''
   else
      begin
      MCSal:= GetControlText('SAL');
      Salarie:= ReadTokenpipe(MCSal, ';');
      CSal1:= ' AND (';
      While (Salarie <> '') do
            begin
            CSal1:= CSal1+' PDS_SALARIE="'+Salarie+'"';
            Salarie := ReadTokenpipe(MCSal,';');
            if (Salarie <> '') then
               CSal1:= CSal1+' OR'
            else
               CSal1:= CSal1+')';
            end;
      end;
//FIN PT39-1
   end
else
   begin
{PT39-1
   If Sal1<>'' Then
      CSal1:= ' AND D1.PDS_SALARIE>="'+Sal1+'"';
   If Sal2<>'' Then
      CSal2:= ' AND D1.PDS_SALARIE<="'+Sal2+'"';
}
   if (THMultiValCombobox(GetControl('SAL')).Tous) then
      CSal1:= ''
   else
      begin
      MCSal:= GetControlText('SAL');
      Salarie:= ReadTokenpipe(MCSal, ';');
      CSal1:= ' AND (';
      While (Salarie <> '') do
            begin
            CSal1:= CSal1+' D1.PDS_SALARIE="'+Salarie+'"';
            Salarie := ReadTokenpipe(MCSal,';');
            if (Salarie <> '') then
               CSal1:= CSal1+' OR'
            else
               CSal1:= CSal1+')';
            end;
      end;
//FIN PT39-1
   end;

if (GetCheckBoxState('CBTP')=cbChecked) then
   begin
   requete:= 'SELECT PDE_SALARIE, PDE_DATEDEBUT, PDE_DATEFIN, PDE_ORDRE,'+
             ' PDS_ORDRE, PDS_SALARIE'+
             ' FROM DADSPERIODES'+
             ' LEFT JOIN DADSDETAIL D1 ON'+
             ' PDE_SALARIE=PDS_SALARIE AND'+
             ' PDE_ORDRE=PDS_ORDRE AND'+
             ' PDE_TYPE=PDS_TYPE AND'+
             ' PDE_DATEDEBUT=PDS_DATEDEBUT AND'+
             ' PDE_DATEFIN=PDS_DATEFIN AND'+
             ' PDE_EXERCICEDADS=PDS_EXERCICEDADS WHERE'+
             ' PDE_SALARIE NOT LIKE "--H%" AND'+
             ' PDE_TYPE<>"001" AND'+
             ' PDE_TYPE<>"201"';
{PT39-1
   CEtab:= ' AND PDS_SEGMENT="S41.G01.00.005" AND'+
           ' PDS_DONNEEAFFICH>="'+Etab1+'" AND'+
           ' PDS_DONNEEAFFICH<="'+Etab2+'"';
}
   if (THMultiValCombobox(GetControl('ETAB')).Tous) then
{PT46
      CEtab:= ' AND PDS_SEGMENT="S41.G01.00.005"'
}
      begin
      ChargeEtabNonFictif (TEtabDADS);
      if (TEtabDADS<>nil) then
         begin
         TEtabDADSD:= TEtabDADS.FindFirst ([''], [''], False);
         if (TEtabDADSD<>nil) then
            begin
            EtabSelect:= TEtabDADSD.GetValue ('ET_ETABLISSEMENT');
            CEtab:= ' AND PDS_SEGMENT="S41.G01.00.005" AND (';
            end
         else
            CEtab:= ' AND PDS_SEGMENT="S41.G01.00.005"';
         While (EtabSelect <> '') do
               begin
               CEtab:= CEtab+' PDS_DONNEEAFFICH="'+EtabSelect+'"';
               TEtabDADSD:= TEtabDADS.FindNext ([''], [''], False);
               if (TEtabDADSD<>nil) then
                  EtabSelect:= TEtabDADSD.GetValue ('ET_ETABLISSEMENT')
               else
                  EtabSelect:= '';
               if (EtabSelect <> '') then
                  CEtab:= CEtab+' OR'
               else
                  CEtab:= CEtab+')';
               end;
         end
      else
         CEtab:= ' AND PDS_SEGMENT="S41.G01.00.005"';
      end
//FIN PT46
   else
      begin
      MCEtab:= GetControlText('ETAB');
      EtabSelect:= ReadTokenpipe(MCEtab, ';');
      CEtab:= ' AND PDS_SEGMENT="S41.G01.00.005" AND (';
      While (EtabSelect <> '') do
            begin
            CEtab:= CEtab+' PDS_DONNEEAFFICH="'+EtabSelect+'"';
            EtabSelect := ReadTokenpipe(MCEtab,';');
            if (EtabSelect <> '') then
               CEtab:= CEtab+' OR'
            else
               CEtab:= CEtab+')';
            end;
      end;
//FIN PT39-1      

   OrderBy:= ' GROUP BY PDE_SALARIE, PDE_DATEDEBUT, PDE_DATEFIN, PDE_ORDRE,'+
             ' PDS_ORDRE, PDS_SALARIE'+
             ' ORDER BY PDE_SALARIE, PDS_ORDRE';
   end
else
if (TypeEdition='I') then
   begin
   If (GetCheckBoxState('TRISAL')=CbChecked) then
      begin
      requete:= 'SELECT PDE_SALARIE, PDE_ORDRE, PDE_DATEDEBUT, PSA_LIBELLE,'+
                ' PSA_PRENOM'+
                ' FROM DADSPERIODES'+
                ' LEFT JOIN SALARIES ON'+
                ' PDE_SALARIE=PSA_SALARIE WHERE'+
                ' PDE_SALARIE NOT LIKE "--H%" AND'+
                ' PDE_ORDRE < 0 AND'+
                ' PDE_EXERCICEDADS = "'+PGExercice+'"';
{PT37
                ' PDE_TYPE="001"';
}
      if (Arg='B') then
         requete:= requete+' AND PDE_TYPE<>"001" AND PDE_TYPE<>"201"'
      else
         requete:= requete+' AND (PDE_TYPE="001" OR PDE_TYPE="201")';
//FIN PT37
      OrderBy:= ' GROUP BY PSA_LIBELLE, PSA_PRENOM, PDE_SALARIE,'+
                ' PDE_DATEDEBUT, PDE_ORDRE'+
                ' ORDER BY PSA_LIBELLE, PSA_PRENOM, PDE_SALARIE,'+
                ' PDE_DATEDEBUT, PDE_ORDRE';
      CEtab:= '';
      end
   else
      begin
      requete:= 'SELECT PDE_SALARIE, PDE_ORDRE, PDE_DATEDEBUT'+
                ' FROM DADSPERIODES WHERE'+
                ' PDE_SALARIE NOT LIKE "--H%" AND'+
                ' PDE_ORDRE < 0 AND'+
                ' PDE_EXERCICEDADS = "'+PGExercice+'"';
{PT37
                ' PDE_TYPE="001"';
}
      if (Arg='B') then
         requete:= requete+' AND PDE_TYPE<>"001" AND PDE_TYPE<>"201"'
      else
         requete:= requete+' AND (PDE_TYPE="001" OR PDE_TYPE="201")';
//FIN PT37
      OrderBy:= ' GROUP BY PDE_SALARIE, PDE_DATEDEBUT, PDE_ORDRE'+
                ' ORDER BY PDE_SALARIE, PDE_DATEDEBUT, PDE_ORDRE';
      CEtab:= '';
      end;
   end
else
if ((TypeEdition='R') OR (TypeEdition='S')) then
   begin
   If (Organisme<>'') Then
{PT39-1
      COrganisme:= ' AND D2.PDS_DONNEEAFFICH>="'+Organisme+'" AND'+
                   ' D2.PDS_DONNEEAFFICH<="'+Organisme1+'"';
   CEtab:= ' AND ET_ETABLISSEMENT>="'+Etab1+'" AND'+
           ' ET_ETABLISSEMENT<="'+Etab2+'"';
}
      begin
      if (THMultiValCombobox(GetControl('ORGANISME')).Tous) then
         COrganisme:= ''
      else
         begin
         MCEtab:= GetControlText('ORGANISME');
         EtabSelect:= ReadTokenpipe(MCEtab, ';');
         COrganisme:= ' AND (';
         While (EtabSelect <> '') do
               begin
               COrganisme:= COrganisme+' D2.PDS_DONNEEAFFICH="'+EtabSelect+'"';
               EtabSelect := ReadTokenpipe(MCEtab,';');
               if (EtabSelect <> '') then
                  COrganisme:= COrganisme+' OR'
               else
                  COrganisme:= COrganisme+')';
               end;
         end;
      end;

   if (THMultiValCombobox(GetControl('ETAB')).Tous) then
{PT46
      CEtab:= ''
}
      begin
      ChargeEtabNonFictif (TEtabDADS);
      if (TEtabDADS<>nil) then
         begin
         TEtabDADSD:= TEtabDADS.FindFirst ([''], [''], False);
         if (TEtabDADSD<>nil) then
            begin
            EtabSelect:= TEtabDADSD.GetValue ('ET_ETABLISSEMENT');
            CEtab:= ' AND (';
            end
         else
            CEtab:= '';
         While (EtabSelect <> '') do
               begin
               CEtab:= CEtab+' ET_ETABLISSEMENT="'+EtabSelect+'"';
               TEtabDADSD:= TEtabDADS.FindNext ([''], [''], False);
               if (TEtabDADSD<>nil) then
                  EtabSelect:= TEtabDADSD.GetValue ('ET_ETABLISSEMENT')
               else
                  EtabSelect:= '';
               if (EtabSelect <> '') then
                  CEtab:= CEtab+' OR'
               else
                  CEtab:= CEtab+')';
               end;
         end
      else
         CEtab:= '';
      end
//FIN PT46
   else
      begin
      MCEtab:= GetControlText('ETAB');
      EtabSelect:= ReadTokenpipe(MCEtab, ';');
      CEtab:= ' AND (';
      While (EtabSelect <> '') do
            begin
            CEtab:= CEtab+' ET_ETABLISSEMENT="'+EtabSelect+'"';
            EtabSelect := ReadTokenpipe(MCEtab,';');
            if (EtabSelect <> '') then
               CEtab:= CEtab+' OR'
            else
               CEtab:= CEtab+')';
            end;
      end;
//FIN PT39-1
   If (GetControlText('CRUPTORG')='X') or (organisme<>'') Then
      begin
      if GetControlText('CRUPTETAB')='X' then
         begin
         If TriSal.Checked=True Then
            begin
            Requete:= ConstruireRequete(True,True,True);
            SetControlText('XX_RUPTURE1','DONNEE');
            SetControlText('XX_RUPTURE2','ET_ETABLISSEMENT');
            OrderBy:= ' GROUP BY ET_ETABLISSEMENT, PDE_SALARIE, PSA_LIBELLE,'+
                      ' PSA_PRENOM, D2.PDS_DONNEEAFFICH'+
                      ' ORDER BY DONNEE, ET_ETABLISSEMENT, PSA_LIBELLE,'+
                      ' PSA_PRENOM';
            end
         Else
            begin
            SetControlText('XX_RUPTURE1','DONNEE');
            SetControlText('XX_RUPTURE2','ET_ETABLISSEMENT');
            Requete:= ConstruireRequete(True,True,False);
            OrderBy:= ' GROUP BY ET_ETABLISSEMENT, PDE_SALARIE,'+
                      ' D1.PDS_DONNEEAFFICH, D2.PDS_DONNEEAFFICH'+
                      ' ORDER BY DONNEE,ET_ETABLISSEMENT, PDE_SALARIE';
            end;
         end
      Else
         begin
         CEtab:= '';
{PT39-1
         If Sal1<>'' Then
            CSal1:= ' AND D2.PDS_SALARIE>="'+Sal1+'"';
         If Sal2<>'' Then
            CSal2:= ' AND D2.PDS_SALARIE<="'+Sal2+'"';
}
         if (THMultiValCombobox(GetControl('SAL')).Tous) then
            CSal1:= ''
         else
            begin
            MCSal:= GetControlText('SAL');
            Salarie:= ReadTokenpipe(MCSal, ';');
            CSal1:= ' AND (';
            While (Salarie <> '') do
                  begin
                  CSal1:= CSal1+' D2.PDS_SALARIE="'+Salarie+'"';
                  Salarie := ReadTokenpipe(MCSal,';');
                  if (Salarie <> '') then
                     CSal1:= CSal1+' OR'
                  else
                     CSal1:= CSal1+')';
                  end;
            end;
//FIN PT39-1
         If TriSal.Checked=True Then
            begin
            SetControlText('XX_RUPTURE1','DONNEE');
            Requete:= ConstruireRequete(True,False,True);
            OrderBy:= ' GROUP BY  PDE_SALARIE, PSA_LIBELLE, PSA_PRENOM,'+
                      ' D2.PDS_DONNEEAFFICH'+
                      ' ORDER BY DONNEE, PSA_LIBELLE, PSA_PRENOM';
            end
         Else
            begin
            SetControlText('XX_RUPTURE1','DONNEE');
            Requete:= ConstruireRequete(True,False,False);
            OrderBy:= ' GROUP BY PDE_SALARIE, D2.PDS_DONNEEAFFICH'+
                      ' ORDER BY DONNEE,PDE_SALARIE';
            end;
         end;
      end
   Else
      begin
      If GetControlText('CRUPTETAB')='X' then
         begin
         If TriSal.Checked=True Then
            begin
            SetControlText('XX_RUPTURE2','ET_ETABLISSEMENT');
            Requete:= ConstruireRequete(False,True,True);
            OrderBy:= ' GROUP BY ET_ETABLISSEMENT, PDE_SALARIE, PSA_LIBELLE,'+
                      ' PSA_PRENOM'+
                      ' ORDER BY ET_ETABLISSEMENT,PSA_LIBELLE, PSA_PRENOM';
            end
         Else
            begin
            SetControlText('XX_RUPTURE2','ET_ETABLISSEMENT');
            Requete:= ConstruireRequete(False,True,False);
            OrderBy:= ' GROUP BY ET_ETABLISSEMENT, PDE_SALARIE,'+
                      ' D1.PDS_DONNEEAFFICH'+
                      ' ORDER BY  ET_ETABLISSEMENT, PDE_SALARIE';
            end;
         end
      Else
         begin
{PT39-1
         CEtab:= ' AND'+ //PT33
                 ' D1.PDS_DONNEEAFFICH>="'+Etab1+'" AND'+
                 ' D1.PDS_DONNEEAFFICH<="'+Etab2+'"';
}
         if (THMultiValCombobox(GetControl('ETAB')).Tous) then
{PT46
            CEtab:= ''
}
            begin
            if (TEtabDADS<>nil) then
               begin
               TEtabDADSD:= TEtabDADS.FindFirst ([''], [''], False);
               if (TEtabDADSD<>nil) then
                  begin
                  EtabSelect:= TEtabDADSD.GetValue ('ET_ETABLISSEMENT');
                  CEtab:= ' AND (';
                  end
               else
                  CEtab:= '';
               While (EtabSelect <> '') do
                     begin
                     CEtab:= CEtab+' D1.PDS_DONNEEAFFICH="'+EtabSelect+'"';
                     TEtabDADSD:= TEtabDADS.FindNext ([''], [''], False);
                     if (TEtabDADSD<>nil) then
                        EtabSelect:= TEtabDADSD.GetValue ('ET_ETABLISSEMENT')
                     else
                        EtabSelect:= '';
                     if (EtabSelect <> '') then
                        CEtab:= CEtab+' OR'
                     else
                        CEtab:= CEtab+')';
                     end;
               end
            else
               CEtab:= '';
            end
//FIN PT46
         else
            begin
            MCEtab:= GetControlText('ETAB');
            EtabSelect:= ReadTokenpipe(MCEtab, ';');
            CEtab:= ' AND (';
            While (EtabSelect <> '') do
                  begin
                  CEtab:= CEtab+' D1.PDS_DONNEEAFFICH="'+EtabSelect+'"';
                  EtabSelect:= ReadTokenpipe(MCEtab,';');
                  if (EtabSelect <> '') then
                     CEtab:= CEtab+' OR'
                  else
                     CEtab:= CEtab+')';
                  end;
            end;
//FIN PT39-1
         If TriSal.Checked=True Then
            begin
            Requete:= ConstruireRequete(False,False,True);
            OrderBy:= ' GROUP BY PDE_SALARIE, PSA_LIBELLE, PSA_PRENOM'+
                      ' ORDER BY PSA_LIBELLE, PSA_PRENOM';
            end
         Else
            begin
            Requete:= ConstruireRequete(False,False,False);
            OrderBy:= ' GROUP BY PDE_SALARIE'+
                      ' ORDER BY PDE_SALARIE';
            end;
         end;
      end;
   end
Else
   begin
   If (Organisme<>'') Then
      begin
      COrganisme:= ' AND PDS_SALARIE IN'+
                   ' (SELECT PDS_SALARIE'+
                   ' FROM DADSDETAIL D2 WHERE'+
                   ' D2.PDS_ORDRE<>0 AND'+
                   ' D2.PDS_SEGMENT="S41.G01.01.001" AND'+
                   ' D2.PDS_ORDRE=D1.PDS_ORDRE AND'+
                   ' D2.PDS_TYPE=D1.PDS_TYPE AND'+
                   ' D2.PDS_DATEDEBUT=D1.PDS_DATEDEBUT AND'+
                   ' D2.PDS_DATEFIN=D1.PDS_DATEFIN AND'+
                   ' D2.PDS_EXERCICEDADS=D1.PDS_EXERCICEDADS';
{PT39-1
                   ' AND D2.PDS_DONNEEAFFICH>="'+Organisme+'" AND'+
                   ' D2.PDS_DONNEEAFFICH<="'+Organisme1+'")';
      end;
   CEtab:= ' AND PDS_SEGMENT="S41.G01.00.005" AND'+
           ' PDS_DONNEEAFFICH>="'+Etab1+'" AND'+
           ' PDS_DONNEEAFFICH<="'+Etab2+'"';
}
      if (THMultiValCombobox(GetControl('ORGANISME')).Tous) then
         COrganisme:= COrganisme+')'
      else
         begin
         MCEtab:= GetControlText('ORGANISME');
         EtabSelect:= ReadTokenpipe(MCEtab, ';');
         COrganisme:= COrganisme+' AND (';
         While (EtabSelect <> '') do
               begin
               COrganisme:= COrganisme+' D2.PDS_DONNEEAFFICH="'+EtabSelect+'"';
               EtabSelect:= ReadTokenpipe(MCEtab,';');
               if (EtabSelect <> '') then
                  COrganisme:= COrganisme+' OR'
               else
                  COrganisme:= COrganisme+'))';
               end;
         end;
      end;

   if (THMultiValCombobox(GetControl('ETAB')).Tous) then
{PT46
      CEtab:= ' AND PDS_SEGMENT="S41.G01.00.005"'
}
      begin
      ChargeEtabNonFictif (TEtabDADS);
      if (TEtabDADS<>nil) then
         begin
         TEtabDADSD:= TEtabDADS.FindFirst ([''], [''], False);
         if (TEtabDADSD<>nil) then
            begin
            EtabSelect:= TEtabDADSD.GetValue ('ET_ETABLISSEMENT');
            CEtab:= ' AND PDS_SEGMENT="S41.G01.00.005" AND (';
            end
         else
            CEtab:= ' AND PDS_SEGMENT="S41.G01.00.005"';
         While (EtabSelect <> '') do
               begin
               CEtab:= CEtab+' D1.PDS_DONNEEAFFICH="'+EtabSelect+'"';
               TEtabDADSD:= TEtabDADS.FindNext ([''], [''], False);
               if (TEtabDADSD<>nil) then
                  EtabSelect:= TEtabDADSD.GetValue ('ET_ETABLISSEMENT')
               else
                  EtabSelect:= '';
               if (EtabSelect <> '') then
                  CEtab:= CEtab+' OR'
               else
                  CEtab:= CEtab+')';
               end;
         end
      else
         CEtab:= ' AND PDS_SEGMENT="S41.G01.00.005"';
      end
//FIN PT46
   else
      begin
      MCEtab:= GetControlText('ETAB');
      EtabSelect:= ReadTokenpipe(MCEtab, ';');
      CEtab:= ' AND PDS_SEGMENT="S41.G01.00.005" AND (';
      While (EtabSelect <> '') do
            begin
            CEtab:= CEtab+' D1.PDS_DONNEEAFFICH="'+EtabSelect+'"';
            EtabSelect:= ReadTokenpipe(MCEtab,';');
            if (EtabSelect <> '') then
               CEtab:= CEtab+' OR'
            else
               CEtab:= CEtab+')';
            end;
      end;
//FIN PT39-1
   If TriSal.Checked=True Then
      begin
      requete:= 'SELECT PDE_SALARIE, PDE_DATEDEBUT, PDE_DATEFIN, PDE_ORDRE,'+
                ' PDS_ORDRE, PDS_SALARIE, PSA_LIBELLE, PSA_PRENOM'+
                ' FROM DADSPERIODES'+
                ' LEFT JOIN DADSDETAIL D1 ON'+
                ' PDE_SALARIE=PDS_SALARIE AND'+
                ' PDE_ORDRE=PDS_ORDRE AND'+
                ' PDE_TYPE=PDS_TYPE AND'+
                ' PDE_DATEDEBUT=PDS_DATEDEBUT AND'+
                ' PDE_DATEFIN=PDS_DATEFIN AND'+
                ' PDE_EXERCICEDADS=PDS_EXERCICEDADS'+
                ' LEFT JOIN SALARIES ON'+
                ' PDE_SALARIE=PSA_SALARIE WHERE'+
                ' PDE_SALARIE NOT LIKE "--H%"';
      if (Arg='B') then
         requete:= requete+' AND PDE_TYPE<>"001" AND PDE_TYPE<>"201"'
      else
         requete:= requete+' AND (PDE_TYPE="001" OR PDE_TYPE="201")';
      OrderBy:= ' GROUP BY PSA_LIBELLE, PSA_PRENOM, PDE_SALARIE, PDS_ORDRE,'+
                ' PDE_DATEDEBUT, PDE_DATEFIN, PDE_ORDRE, PDS_SALARIE '+  //PT32
                ' ORDER BY PSA_LIBELLE, PSA_PRENOM, PDE_SALARIE, PDS_ORDRE';
      end
   Else
      begin
      requete:= 'SELECT PDE_SALARIE, PDE_DATEDEBUT, PDE_DATEFIN, PDE_ORDRE,'+
                ' PDS_ORDRE, PDS_SALARIE'+
                ' FROM DADSPERIODES'+
                ' LEFT JOIN DADSDETAIL D1 ON'+
                ' PDE_SALARIE=PDS_SALARIE AND'+
                ' PDE_ORDRE=PDS_ORDRE AND'+
                ' PDE_TYPE=PDS_TYPE AND'+
                ' PDE_DATEDEBUT=PDS_DATEDEBUT AND'+
                ' PDE_DATEFIN=PDS_DATEFIN AND'+
                ' PDE_EXERCICEDADS=PDS_EXERCICEDADS'+
                ' WHERE PDE_SALARIE NOT LIKE "--H%"';
      if (Arg='B') then
         requete:= requete+' AND PDE_TYPE<>"001" AND PDE_TYPE<>"201"'
      else
         requete:= requete+' AND (PDE_TYPE="001" OR PDE_TYPE="201")';
      OrderBy:= ' GROUP BY PDE_SALARIE, PDE_DATEDEBUT, PDE_DATEFIN, PDE_ORDRE,'+
                ' PDS_ORDRE, PDS_SALARIE ORDER BY PDE_SALARIE, PDS_ORDRE';
      end;
   end;

  //DEB PT43
  Habilitation := '';
  if (Assigned(MonHabilitation)) and (MonHabilitation.LeSQL<>'') then
  begin
    Habilitation := ' AND PDE_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE ' + MonHabilitation.LeSQL + ')';
  end;
  //FIN PT43

SQL:= requete+COrganisme+CDate+CSal1+CSal2+CEtab+Habilitation+OrderBy; //PT43
TobLexique:= Tob.Create('LesSegments',Nil,-1);
ChargerLexiqueDads(TobLexique, AnneeLexique, False, TypeEdition);
Q:= OpenSQL(SQL,True);
TobPeriodes:= Tob.Create('LesPeriodes',Nil,-1);
TobPeriodes.LoadDetailDB('LesPeriodes','','',Q,False);
Ferme(Q);
InitMoveProgressForm (NIL,'Chargement des données',
                     'Veuillez patienter SVP ...', TobPeriodes.Detail.Count,
                     False,True);
InitMove(TobPeriodes.Detail.Count,'');
If ((TypeEdition<>'R') and (TypeEdition<>'S')) then
   begin
   For i:=0 to TobPeriodes.Detail.Count-1 do
       begin
       T:= TobPeriodes.Detail[i];
       T.AddchampSupValeur('DOSSIER', GetControlText('DOSSIER'), False);
       T.AddchampSupValeur('DATEDEB', GetControlText('DATEDEB'), False);
       T.AddchampSupValeur('DATEFIN', GetControlText('DATEFIN'), False);
       if (Arg='B') then
          T.AddchampSupValeur('BTP', 'BTP', False)
       else
          T.AddchampSupValeur('BTP', '', False);
       Salarie:= T.GetValue('PDE_SALARIE');
       NumOrdre:= T.GetValue('PDE_ORDRE');
       Req:= 'SELECT *'+
             ' FROM DADSDETAIL WHERE'+
             ' PDS_SALARIE="'+Salarie+'" AND'+
             ' (PDS_ORDRE=0 OR PDS_ORDRE='+IntToStr(NumOrdre)+') AND'+
{PT39-1
             ' PDS_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND'+
             ' PDS_DATEDEBUT<="'+UsDateTime(DateFin)+'"';
}
             ' PDS_EXERCICEDADS = "'+PGExercice+'"';
       if (Arg='B') then
          Req:= Req+' AND PDS_TYPE<>"001" AND PDS_TYPE<>"201"'
       else
          Req:= Req+' AND (PDS_TYPE="001" OR PDS_TYPE="201")';
       Req:= req+' ORDER BY PDS_ORDRE';
       Q:= OpenSQL(Req,True);
       TobDetail:= Tob.Create('MonDetail',Nil,-1);
       TobDetail.LoadDetailDB('MonDetail','','',Q,False);
       Ferme(Q);
       For x:=0 to TobLexique.Detail.Count-1 do
           begin
           Segment:= TobLexique.Detail[x].GetValue('PDL_DADSSEGMENT');
           Nature:= TobLexique.Detail[x].GetValue('PDL_DADSNATURE');
           OrdreSeg:= TobLexique.Detail[x].GetValue('ORDRESEG');
           Tablette:= TobLexique.Detail[x].GetValue('TABLETTE');
           If OrdreSeg<>'' then
              begin
              OrdreSegS:= OrdreSeg;
              OrdreSegI:= StrToInt(ReadTokenPipe(OrdreSegS,'L'));
              TD:= TobDetail.FindFirst(['PDS_SEGMENT','PDS_ORDRESEG'],
                                       [Segment,OrdreSegI],False);
              If TD<>Nil then
                 begin
                 ValeurSeg:= TD.GetValue('PDS_DONNEEAFFICH');
                 If Tablette<>'' then
                    ValeurSeg:= RechDom(Tablette,ValeurSeg,False);
                 end
              else
                 ValeurSeg:= '';
              end
           else
              begin
              TD:= TobDetail.FindFirst(['PDS_SEGMENT'],[Segment],False);
              If TD<>nil then
                 begin
                 If TD.GetValue('PDS_DONNEEAFFICH')<>null then
                    ValeurSeg:= TD.GetValue('PDS_DONNEEAFFICH')
                 else
                    ValeurSeg:= '';
//PT36-2
                 If ((Segment = 'S41.G01.00.012') or
                    (Segment = 'S41.G01.00.012.001')) Then
                    begin
                    Q:= OpenSQL ('SELECT CO_LIBELLE FROM COMMUN WHERE'+
                                 ' CO_TYPE="PCT" AND'+
                                 ' CO_CODE="'+ValeurSeg+'"',
                                 True);
                    If Not Q.Eof then
                       ValeurSeg := Q.FindField('CO_LIBELLE').AsString
                    else
                       ValeurSeg := '';
                    Ferme(Q);
                    end;
//FIN PT36-2

                 if (TypeEdition='I') then
                    begin
                    If (Segment='S46.G01.00.001') Then
                       begin
                       Q:= OpenSQL ('SELECT CO_LIBELLE FROM COMMUN WHERE'+
                                    ' CO_TYPE="PTX" AND'+
                                    ' CO_CODE="'+ValeurSeg+'"', True);
                       If Not Q.Eof then
                          ValeurSeg := Q.FindField('CO_LIBELLE').AsString
                       else
                          ValeurSeg := '';
                       Ferme(Q);
                       end;
                    end;

                 If (Tablette<>'') and (ValeurSeg<>'') then
                    ValeurSeg:= RechDom (Tablette,
                                         TD.GetValue('PDS_DONNEEAFFICH'), False);
                 end
              else
                 ValeurSeg:= '';
              end;
           If OrdreSeg<>'' then
              Segment:= Segment+'.'+OrdreSeg;
           If Nature='N' then
              begin
              If IsNumeric (ValeurSeg) then
                 T.AddchampSupValeur(Segment,Valeur(ValeurSeg),False)
              else
                 T.AddchampSupValeur(Segment,0,False);
              end
           else
              T.AddchampSupValeur(Segment,ValeurSeg,False)
           end;
       FreeAndNil (TobDetail);
       MoveCurProgressForm ('Salarié : '+
                           TobPeriodes.Detail[i].GetValue('PDE_SALARIE'));
       end;
   FiniMoveProgressForm;
//PT39-5
   if (TypeEdition = 'I') then
      begin
      If GetCheckBoxState('TRISAL')=CbChecked then
         TobPeriodes.Detail.Sort('PSA_LIBELLE;PSA_PRENOM;PDE_SALARIE;PDE_DATEDEBUT')
      else
         TobPeriodes.Detail.Sort('PDE_SALARIE;PDE_DATEDEBUT');
      end
   else
      begin
//FIN PT39-5
      If GetCheckBoxState('TRISAL')=CbChecked then
         TobPeriodes.Detail.Sort('PSA_LIBELLE;PSA_PRENOM;PDE_SALARIE;PDE_ORDRE')
      else
         TobPeriodes.Detail.Sort('PDE_SALARIE;PDE_ORDRE');
      end;
//TobDebug (TobPeriodes);
   end
else
   begin
   TListeSalarie:= Tob.Create('Liste des salaries', nil, -1);
   For i := 0 to TobPeriodes.Detail.Count -1 do
       begin
       T:= TobPeriodes.Detail[i];
       If Not T.FieldExists('DONNEE') then
          T.AddchampSupValeur('DONNEE','');
       If T.GetValue('DONNEE') <> '' then
          T.AddchampSupValeur ('ORGANISME',RechDom('PGINSTITUTION',
                              T.GetValue('DONNEE'),False))
       else
          T.AddchampSupValeur('ORGANISME','');
       If Not T.FieldExists('ET_ETABLISSEMENT') then
          T.AddchampSupValeur('ET_ETABLISSEMENT','');
       T.AddchampSupValeur('RUPTURE','');
       T.AddchampSupValeur('NBSAL', 1);
       T.AddchampSupValeur('DOSSIER', GetControlText('DOSSIER'), False);
       T.AddchampSupValeur('DATEDEB', GetControlText('DATEDEB'), False);
       T.AddchampSupValeur('DATEFIN', GetControlText('DATEFIN'), False);
       if (Arg='B') then
          T.AddchampSupValeur('BTP', 'BTP', False)
       else
          T.AddchampSupValeur('BTP', '', False);
       Salarie:= T.GetValue('PDE_SALARIE');

       TSalD:= TListeSalarie.FindFirst(['MATRICULE'], [Salarie], TRUE);
       if (TSalD=nil) then
          begin
          TSal:= Tob.Create('Le salarie', TListeSalarie, -1);
          TSal.AddchampSupValeur('MATRICULE', Salarie);
          end;

       If GetCheckBoxState('CRUPTETAB') = CbChecked then
          SQLRecapEtab:= ' AND D2.PDS_DONNEEAFFICH="'+
                         T.GetValue('ET_ETABLISSEMENT')+'"'
       else
{PT39-1
          SQLRecapEtab:= ' AND D2.PDS_DONNEEAFFICH>="'+GetControlText('ETAB1')+'" AND'+
                         ' D2.PDS_DONNEEAFFICH<="'+GetControlText('ETAB2')+'"';
}
          begin
          if (THMultiValCombobox(GetControl('ETAB')).Tous) then
{PT46
             SQLRecapEtab:= ''
}
             begin
             if (TEtabDADS<>nil) then
                begin
                TEtabDADSD:= TEtabDADS.FindFirst ([''], [''], False);
                if (TEtabDADSD<>nil) then
                   begin
                   EtabSelect:= TEtabDADSD.GetValue ('ET_ETABLISSEMENT');
                   SQLRecapEtab:= ' AND (';
                   end
                else
                   SQLRecapEtab:= '';
                While (EtabSelect <> '') do
                      begin
                      SQLRecapEtab:= SQLRecapEtab+
                                     ' D2.PDS_DONNEEAFFICH="'+EtabSelect+'"';
                      TEtabDADSD:= TEtabDADS.FindNext ([''], [''], False);
                      if (TEtabDADSD<>nil) then
                         EtabSelect:= TEtabDADSD.GetValue ('ET_ETABLISSEMENT')
                      else
                         EtabSelect:= '';
                      if (EtabSelect <> '') then
                         SQLRecapEtab:= SQLRecapEtab+' OR'
                      else
                         SQLRecapEtab:= SQLRecapEtab+')';
                      end;
                end
             else
                SQLRecapEtab:= '';
             end
//FIN PT46
          else
             begin
             MCEtab:= GetControlText('ETAB');
             EtabSelect:= ReadTokenpipe(MCEtab, ';');
             SQLRecapEtab:= ' AND (';
             While (EtabSelect <> '') do
                   begin
                   SQLRecapEtab:= SQLRecapEtab+
                                  ' D2.PDS_DONNEEAFFICH="'+EtabSelect+'"';
                   EtabSelect:= ReadTokenpipe(MCEtab,';');
                   if (EtabSelect <> '') then
                      SQLRecapEtab:= SQLRecapEtab+' OR'
                   else
                      SQLRecapEtab:= SQLRecapEtab+')';
                   end;
             end;
          end;
//FIN PT39-1
       SQLRecapOrgJoin:= ' LEFT JOIN DADSDETAIL D3 ON'+
                         ' D1.PDS_SALARIE=D3.PDS_SALARIE AND'+
                         ' D1.PDS_ORDRE=D3.PDS_ORDRE AND'+
                         ' D1.PDS_TYPE=D3.PDS_TYPE AND'+
                         ' D1.PDS_DATEDEBUT=D3.PDS_DATEDEBUT AND'+
                         ' D1.PDS_DATEFIN=D3.PDS_DATEFIN';
       SQLrecapOrgWhere:= '';
       If GetCheckBoxState('CRUPTORG') = CbChecked then
          begin
          SQLrecapOrgWhere:=  ' AND D3.PDS_SEGMENT="S41.G01.01.001" AND'+
                              ' D3.PDS_DONNEEAFFICH="'+T.GetValue('DONNEE')+'"';
          end
       else
          if GetControlText('ORGANISME') <> '' then
{PT39-1
             SQLrecapOrgWhere:= ' AND D3.PDS_SEGMENT="S41.G01.01.001" AND'+
                                ' D3.PDS_DONNEEAFFICH>="'+GetControlText('ORGANISME')+'" AND'+
                                ' D3.PDS_DONNEEAFFICH<="'+GetControlText('ORGANISME1')+'"';
}
             begin
             if (THMultiValCombobox(GetControl('ORGANISME')).Tous) then
                SQLrecapOrgWhere:= ' AND D3.PDS_SEGMENT="S41.G01.01.001"'
             else
                begin
                MCEtab:= GetControlText('ORGANISME');
                EtabSelect:= ReadTokenpipe(MCEtab, ';');
                SQLrecapOrgWhere:= ' AND D3.PDS_SEGMENT="S41.G01.01.001" AND (';
                While (EtabSelect <> '') do
                      begin
                      SQLrecapOrgWhere:= SQLrecapOrgWhere+
                                     ' D3.PDS_DONNEEAFFICH="'+EtabSelect+'"';
                      EtabSelect:= ReadTokenpipe(MCEtab,';');
                      if (EtabSelect <> '') then
                         SQLrecapOrgWhere:= SQLrecapOrgWhere+' OR'
                      else
                         SQLrecapOrgWhere:= SQLrecapOrgWhere+')';
                      end;
                end;
          end;
//FIN PT39-1
       SQLRecap:= 'SELECT D1.PDS_SALARIE, D1.PDS_ORDRE, D1.PDS_ORDRESEG,'+
                  ' D1.PDS_SEGMENT, D1.PDS_DONNEEAFFICH'+
                  ' FROM DADSDETAIL D1'+
                  ' LEFT JOIN DADSDETAIL D2 ON'+
                  ' D1.PDS_SALARIE=D2.PDS_SALARIE AND'+
                  ' D1.PDS_ORDRE=D2.PDS_ORDRE AND'+
                  ' D1.PDS_TYPE=D2.PDS_TYPE AND'+
                  ' D1.PDS_DATEDEBUT=D2.PDS_DATEDEBUT AND'+
                  ' D1.PDS_DATEFIN=D2.PDS_DATEFIN AND'+
                  ' D1.PDS_EXERCICEDADS=D2.PDS_EXERCICEDADS'+
                  SQLRecapOrgJoin+' WHERE'+
                  ' D1.PDS_ORDRE<>0 AND'+
                  ' D1.PDS_EXERCICEDADS="'+PGExercice+'" AND';
       if (TypeEdition='R') then
          SQLRecap:= SQLRecap+ ' D1.PDS_SEGMENT>="S41.G01.00.022" AND'+
                     ' D1.PDS_SEGMENT<="S41.G30.35.004.002" AND'+
                     ' D1.PDS_SALARIE="'+Salarie+'"'
       else
          SQLRecap:= SQLRecap+ ' (D1.PDS_SEGMENT="S41.G01.00.021" OR'+
                     ' D1.PDS_SEGMENT="S41.G01.00.022" OR'+
                     ' D1.PDS_SEGMENT="S41.G01.00.035.001" OR'+
                     ' D1.PDS_SEGMENT="S41.G01.00.044.001" OR'+
                     ' D1.PDS_SEGMENT="S41.G01.00.063.001" OR'+
                     ' D1.PDS_SEGMENT="S41.G01.00.066.001" OR'+
                     ' D1.PDS_SEGMENT="S41.G01.00.067.001" OR'+   //PT34
                     ' D1.PDS_SEGMENT="S41.G30.35.001.002") AND'+
                     ' D1.PDS_SALARIE="'+Salarie+'"';

       if (Arg='B') then
          SQLRecap:= SQLRecap+' AND D1.PDS_TYPE<>"001" AND D1.PDS_TYPE<>"201"'
       else
          SQLRecap:= SQLRecap+' AND (D1.PDS_TYPE="001" OR D1.PDS_TYPE="201")';
       SQLRecap:= SQLRecap+' AND D2.PDS_SEGMENT="S41.G01.00.005"'+SQLRecapEtab+
                  SQLRecapOrgWhere+' GROUP BY D1.PDS_SALARIE, D1.PDS_ORDRE,'+
                                   ' D1.PDS_ORDRESEG, D1.PDS_SEGMENT,'+
                                   ' D1.PDS_DONNEEAFFICH';
       Q:= OpenSQL(SQLRecap,True);
       TobDetail:= Tob.Create('MonDetail',Nil,-1);
       TobDetail.LoadDetailDB('MonDetail','','',Q,False);
       Ferme(Q);
       For x := 0 to TobLexique.Detail.Count - 1 do
           begin
           Segment:= TobLexique.Detail[x].GetValue('PDL_DADSSEGMENT');
           MtSegment:= 0;
           TD:= TobDetail.FindFirst(['PDS_SEGMENT'],[Segment],False);
           While TD <> Nil do
                 begin
                 MtSegment:= MtSegment+StrToFloat(TD.GetValue('PDS_DONNEEAFFICH'));
                 TD:= TobDetail.FindNext(['PDS_SEGMENT'], [Segment],False);
                 end;
           T.AddchampSupValeur(Segment,MtSegment);
           end;
       FreeAndNil(TobDetail);
       SQLRecap:= 'SELECT D1.PDS_SALARIE, D1.PDS_ORDRE, D1.PDS_ORDRESEG,'+
                  ' D1.PDS_SEGMENT, D1.PDS_DONNEEAFFICH'+
                  ' FROM DADSDETAIL D1'+
                  ' LEFT JOIN DADSDETAIL D2 ON'+
                  ' D1.PDS_SALARIE=D2.PDS_SALARIE AND'+
                  ' D1.PDS_TYPE=D2.PDS_TYPE'+SQLRecapOrgJoin+' WHERE'+
{PT39-2
                  ' D1.PDS_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND'+
                  ' D1.PDS_DATEDEBUT<="'+UsDateTime(DateFin)+'" AND';
}
                  ' D1.PDS_EXERCICEDADS="'+PGExercice+'" AND';
       if (TypeEdition='R') then
          SQLRecap:= SQLRecap+ ' D1.PDS_SEGMENT="S41.G01.00.019" AND'+
                     ' D1.PDS_SALARIE="'+Salarie+'"'
       else
          SQLRecap:= SQLRecap+' (D1.PDS_SEGMENT LIKE "S30.G01.00.008%") AND'+
                     ' D1.PDS_SALARIE="'+Salarie+'"';
       if (Arg='B') then
          SQLRecap:= SQLRecap+' AND D1.PDS_TYPE<>"001" AND D1.PDS_TYPE<>"201"'
       else
          SQLRecap:= SQLRecap+' AND (D1.PDS_TYPE="001" OR D1.PDS_TYPE="201")';
       SQLRecap:= SQLRecap+' AND D2.PDS_SEGMENT="S41.G01.00.005"'+SQLRecapEtab+
                  SQLRecapOrgWhere;
       Q:= OpenSQL(SQLRecap,True);
       TobDetail:= Tob.Create('MonDetail',Nil,-1);
       TobDetail.LoadDetailDB('MonDetail','','',Q,False);
       Ferme(Q);
       SQLRecap:= 'SELECT PDS_SALARIE, PDS_ORDRE, PDS_ORDRESEG, PDS_SEGMENT,'+
                  ' PDS_DONNEEAFFICH FROM DADSDETAIL WHERE'+
                  ' PDS_SALARIE="'+Salarie+'" AND'+
{PT39-2
                  ' PDS_DATEDEBUT>="'+UsDateTime(DateDeb)+'" AND'+
                  ' PDS_DATEDEBUT<="'+UsDateTime(DateFin)+'" AND'+
}
                  ' PDS_EXERCICEDADS="'+PGExercice+'" AND'+
                  ' (PDS_SEGMENT="S30.G01.00.001" OR'+
                  ' PDS_SEGMENT="S30.G01.00.002" OR'+
                  ' PDS_SEGMENT="S30.G01.00.003")';
       Q:= OpenSQL(SQLRecap,True);
       TobDetail.LoadDetailDB('MonDetail','','',Q,True);
       Ferme(Q);
       For x := 0 to TobLexique.Detail.Count - 1 do
           begin
           Segment:= TobLexique.Detail[x].GetValue('PDL_DADSSEGMENT');
           TD:= TobDetail.FindFirst(['PDS_SEGMENT'],[Segment],False);
           If (TD<> nil) then
              T.PutValue(Segment,TD.GetValue('PDS_DONNEEAFFICH'))
              else If TobLexique.Detail[x].GetValue('PDL_DADSNATURE') <> 'N' then T.PutValue(Segment,'');//PT33
           end;

       MoveCurProgressForm ('Salarié : '+TobPeriodes.Detail[i].GetValue('PDE_SALARIE'));
       FreeAndNil(TobDetail);
       end;
   FiniMoveProgressForm;
   If GetCheckBoxState('CRUPTORG') = CbChecked then
      begin
      If GetCheckBoxState('CRUPTETAB') = CbChecked then
         begin
         If GetCheckBoxState('TRISAL') = CbChecked then
            TobPeriodes.Detail.Sort('DONNEE;ET_ETABLISSEMENT;PSA_LIBELLE;PSA_PRENOM;PDE_SALARIE')
         else
            TobPeriodes.Detail.Sort('DONNEE;ET_ETABLISSEMENT;PDE_SALARIE');
         end
      else
         begin
         If GetCheckBoxState('TRISAL') = CbChecked then
            TobPeriodes.Detail.Sort('DONNEE;PSA_LIBELLE;PSA_PRENOM;PDE_SALARIE')
         else
            TobPeriodes.Detail.Sort('DONNEE;PDE_SALARIE');
         end;
      end
   else
      begin
      If GetCheckBoxState('CRUPTETAB') = CbChecked then
         begin
         If GetCheckBoxState('TRISAL') = CbChecked then
            TobPeriodes.Detail.Sort('ET_ETABLISSEMENT;PSA_LIBELLE;PSA_PRENOM;PDE_SALARIE')
         else
            TobPeriodes.Detail.Sort('ET_ETABLISSEMENT;PDE_SALARIE');
         end
      else
         begin
         If GetCheckBoxState('TRISAL') = CbChecked then
            TobPeriodes.Detail.Sort('PSA_LIBELLE;PSA_PRENOM;PDE_SALARIE')
         else
            TobPeriodes.Detail.Sort('PDE_SALARIE');
         end;
      end;
   end;
FreeAndNil (TobLexique);
If (TypeEdition='R') then
   SetControlText('NBSAL', IntToStr(TListeSalarie.FillesCount(1)));
FreeAndNil (TListeSalarie);
FreeAndNil (TEtabDADS);    //PT46
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : JL
Créé le ...... : 30/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_PGEDIT_DADSPER.EditionFichierDads(TypeEdition:string):string;
var
DateDeb, DateFin, Dossier, ExerciceDads, FileN, NIC, S, Salarie : String;
Segment, Structure, ValeurSeg : String;
Q : TQuery;
T, TE, TL, TListeSalarie, TobEntete, TobEtabliss, TobLexique, TSal, TSalD: Tob;
i, IncBBS, IncBPS, IncExo, IncPrev, IncRC, IncTSI, LL, MaxBBS, MaxBPS : integer;
MaxExo, MaxPrev, MaxRC, MaxTSI, NbAG, NbBBS, NbBPS, NbExo, NbPrev : integer;
NbRC, NbTSI, Ordre, OrdreBBS, OrdreBPS, OrdreExo, OrdreG30, OrdrePrev : integer;
OrdreRC, OrdreTSI, Virgule, x : Integer;
FLect : TextFile;
CreationTob, EditionBTP, Structure2003 : Boolean;
begin
result:= '';
//Récupération du fichier
FileN:= GetControlText ('FICHIER');

if (FileN='') then
   begin
   PGIBox ('Aucun fichier sélectionné', Ecran.Caption);
   exit;
   end;

AssignFile (FLect, FileN);
Reset (FLect);
Readln (FLect,S);
ExerciceDads:= '';
TobPeriodes:= Tob.Create('LesPeriodes',Nil,-1);
TobLexique:= Tob.Create('LeLexique',Nil,-1);
TobEtabliss := Tob.Create('LesEtab',Nil,-1);

CreationTob:= False;
EditionBTP:= False;
Structure2003:= False;

while (not(eof(FLect))) do
      begin
      if (Structure='S90') then
         break;
      while (Segment<'S20.G01.00.002') do
            begin
            Readln (FLect,S);  // Recherche de la première entreprise
            Structure:= Copy (S, 1, 3);
            Segment:= Copy (S,1,14);
            end;
      if (Segment='S20.G01.00.002') then
         begin
         LL:= Length (S);
         SetControlText ('DOSSIER', Copy (S, 17, LL-17));
         end;
      Readln (FLect,S);
      Structure:= Copy (S, 1, 3);
      Segment:= Copy (S,1,14);
      ExerciceDads:= '';
      If Copy (S,1,18)='S20.G01.00.003.001' then
         begin
         DateDeb:= Copy (S, 21, 2)+'/'+Copy (S, 23, 2)+'/'+Copy (S, 25, 4);
         Readln (FLect,S);
         end;
      if Copy (S,1,18)='S20.G01.00.003.002' then
         begin
         DateFin:= Copy (S, 21, 2)+'/'+Copy (S, 23, 2)+'/'+Copy (S, 25, 4);
         ExerciceDads:= Copy (S, 25, 4);
         end
      else
         begin
         If Copy (S,1,14)='S20.G01.00.003' then
         // Recherche de l'année de la DADS
            begin
            LL:= Length (S);
            ExerciceDads:= Copy (S, 17, LL-17);
            Structure2003:= True;
            end;
         end;
      Readln (FLect,S);
      if (Copy (S,1,18)='S20.G01.00.004.001') then
      //Recherche du type de DADS
         begin
         LL:= Length (S);
         if (Copy(S,21, LL-21)='04') then
            EditionBTP:= True;
         end;
      if (Structure2003=True) then
         begin
         if (EditionBTP=False) then
            begin
            DateDeb:= '01/01/'+ExerciceDads;
            DateFin:= '31/12/'+ExerciceDads;
            end
         else
            begin
            DateDeb:= '01/04/'+IntToStr(StrToInt(ExerciceDads)-1);
            DateFin:= '31/03/'+ExerciceDads;
            end;
         end;
      if (EditionBTP=True) then
         ExerciceDads:= IntToStr(StrToInt(ExerciceDads)-1);
      SetControlText ('DATEDEB', DateDeb);
      SetControlText ('DATEFIN', DateFin);

{Construction tob qui contient les définitions des champs correspondant aux
segments à partir table DADSLEXIQUE}
      ChargerLexiqueDads(TobLexique, ExerciceDads, True, TypeEdition);

      if (TypeEdition<>'H') then
         begin
         Structure:= Copy (S, 1, 3);
         while (Structure<'S30') do
               begin
               Readln (FLect,S);  // Recherche du premier salarié
               Structure:= Copy (S, 1, 3);
               end;
         If (Structure='S90') then
            break;
         Segment:= Copy (S,1,14);
         while (Structure<'S70') do
               begin
               if (Segment='S30.G01.00.001') then
                  //Si nouveau salarié
                  begin
                  FreeAndNil (TobEntete);
                  TobEntete:= Tob.Create('MonEntete',Nil,-1);
                  end;
               //Si nouveau salarié mémorisation des segments S30
               while (Structure='S30') do
                     begin
                     LL:= Length (S);
                     If (Segment='S30.G01.00.008') then
                        begin
                        Segment:= Copy (S, 1, 18);
                        ValeurSeg:= Copy (S, 21, LL-21);
                        end
                     else
                        ValeurSeg:= Copy (S, 17, LL-17);
                     TobEntete.AddchampSupValeur(Segment, ValeurSeg, False);
                     Readln (FLect,S);
                     Structure:= Copy (S, 1, 3);
                     Segment:= Copy (S,1,14);
                     end;

               if (((Segment='S41.G01.00.001') and (TypeEdition<>'I')) or
                  ((Segment='S46.G01.00.001') and (TypeEdition='I'))) then
               //Création d'une période
                  begin
                  T:= Tob.create('Filles',TobPeriodes,-1);
                  T.AddchampSupValeur('DOSSIER', GetControlText('DOSSIER'), False);
                  T.AddchampSupValeur('DATEDEB', GetControlText('DATEDEB'), False);
                  T.AddchampSupValeur('DATEFIN', GetControlText('DATEFIN'), False);
                  if (EditionBTP=True) then
                     T.AddchampSupValeur('BTP', 'BTP', False)
                  else
                     T.AddchampSupValeur('BTP', '', False);
                  if (TypeEdition='I') then
                     T.AddchampSupValeur('S41.G01.00.019', Salarie, False);
                  CreationSegments(ExerciceDads, T, TobEntete, TobLexique);
                  CreationTob:= True;
                  end;

               if (CreationTob=True) then
                  begin
                  Segment:= Copy (S,1,14);
                  LL:= Length (S);
                  ValeurSeg:= Copy (S, 17, LL-17);
                  if (TypeEdition<>'I') then
                     ValeurSeg:= ValeurSeg+ExerciceDads;
                  if (TypeEdition<>'R') then
                     begin
                     if (TypeEdition='I') then
                        begin
                        If (Segment='S46.G01.00.001') Then
                           begin
                           Q:= OpenSQL ('SELECT CO_LIBELLE FROM COMMUN WHERE'+
                                        ' CO_TYPE="PTX" AND'+
                                        ' CO_ABREGE="'+ValeurSeg+'"',
                                        True);
                           If Not Q.Eof then
                              ValeurSeg := Q.FindField('CO_LIBELLE').AsString
                           else
                              ValeurSeg := '';
                           Ferme(Q);
                           end;
                        T.PutValue (Segment, ValeurSeg);
                        end
                     else
//Mise à jour date début période
                        T.PutValue (Segment,StrToDate(Copy (ValeurSeg,1,2)+'/'+
                                    Copy (ValeurSeg,3,2)+'/'+
                                    Copy (ValeurSeg,5,4)))
                     end;
                  Readln (FLect,S);
                  Structure:= Copy (S, 1, 3);
{PT36-2
                  Segment:= Copy (S,1,14);
}
                  Virgule:= Pos(',', S);
                  Segment:= Copy (S,1,Virgule-1);
//FIN PT36-2
//PT47
                  If (ExerciceDads>='2007') then
                     OrdreRC:= 111
                  else
                     OrdreRC:= 101;
//FIN PT47
                  NbRC:= 0;
                  IncRC:= 2;
//PT47
                  If (ExerciceDads>='2007') then
                     MaxRC:= 119
                  else
                     MaxRC:= 109;
//FIN PT47
                  OrdreBBS:= 201;
                  NbBBS:= 0;
                  IncBBS:= 3;
                  MaxBBS:= 204;
                  OrdreBPS:= 301;
                  NbBPS:= 0;
                  IncBPS:= 3;
                  MaxBPS:= 310;
                  OrdreTSI:= 401;
                  NbTSI:= 0;
                  IncTSI:= 6;
                  MaxTSI:= 419;
                  OrdreExo:= 601;
                  NbExo:= 0;
                  IncExo:= 5;
                  MaxExo:= 646;
//PT47
                  If (ExerciceDads>='2007') then
                     OrdrePrev:= 922
                  else
                     OrdrePrev:= 802;
//FIN PT47
                  NbPrev:= 0;
                  IncPrev:= 11;
//PT47
                  If (ExerciceDads>='2007') then
                     MaxPrev:= 969
                  else
                     MaxPrev:= 835;
//FIN PT47
                  While ((Segment<>'') and (Segment<>'S30.G01.00.001') and
                        (Segment<>'S41.G01.00.001') and
                        (Segment<>'S46.G01.00.001')) do
                        begin
                        LL:= Length (S);
{PT36-2
                        If (ExerciceDads >= '2003') and
                           ((Segment='S30.G01.00.008') or
                           (Segment='S41.G01.00.006') or
                           (Segment='S41.G01.00.007') or
                           (Segment='S41.G01.00.018') or
                           (Segment='S41.G01.00.029') or
                           (Segment='S41.G01.00.030') or
                           (Segment='S41.G01.00.032') or
                           (Segment='S41.G01.00.033') or
                           (Segment='S41.G01.00.035') or
                           (Segment='S41.G01.00.060') or
                           (Segment='S41.G01.00.063') or
                           (Segment='S41.G01.00.064') or
                           (Segment='S41.G01.00.065') or
                           (Segment='S41.G01.02.002') or
                           (Segment='S41.G01.03.002') or
                           (Segment='S41.G01.04.003') or
                           (Segment='S41.G01.04.004') or
                           (Segment='S41.G01.05.002') or
                           (Segment='S41.G01.06.002') or
                           (Segment='S41.G01.06.003') or
                           (Segment='S45.G01.01.007') or
                           (Segment='S46.G01.02.003')) then
                           begin
                           Segment:= Copy (S, 1, 18);
                           ValeurSeg:= Copy (S, 21, LL-21);
                           end
                        else
                           ValeurSeg:= Copy (S, 17, LL-17);
}
                        Virgule:= Pos(',', S);
                        Segment:= Copy (S,1,Virgule-1);
                        ValeurSeg:= Copy (S, Virgule+2, LL-(Virgule+2));
//FIN PT36-2
                        TL:= TobLexique.FindFirst (['PDL_DADSSEGMENT'],
                                                   [Segment], False);

                        if (TypeEdition<>'I') then
                           begin
                           if ((Segment='S41.G01.00.012') or
                              (Segment='S41.G01.00.012.001')) then
                              begin
                              Q:= OpenSQL ('SELECT CO_LIBELLE FROM COMMUN WHERE'+
                                           ' CO_TYPE="PCT" AND'+
                                           ' CO_ABREGE="'+ValeurSeg+'"',
                                           True);
                              If Not Q.Eof then
                                 ValeurSeg := Q.FindField('CO_LIBELLE').AsString
                              else
                                 ValeurSeg := '';
                              Ferme(Q);
                              end;

                           if Segment='S41.G01.00.013' then
                              begin
                              If (ValeurSeg = '90') then
                                 ValeurSeg := 'W'
                              else
                              If (ValeurSeg = '01') then
                                 ValeurSeg := 'C'
                              else
                              If (ValeurSeg = '02') then
                                 ValeurSeg := 'P'
                              else
                              If (ValeurSeg = '04') then
                                 ValeurSeg := 'I'
                              else
                              If (ValeurSeg = '05') then
                                 ValeurSeg := 'D'
                              else
                              If (ValeurSeg = '06') then
                                 ValeurSeg := 'S'
                              else
                              If (ValeurSeg = '07') then
                                 ValeurSeg := 'V'
                              else
                              If (ValeurSeg = '08') then
                                 ValeurSeg := 'O'
                              else
                              If (ValeurSeg = '09') then
                                 ValeurSeg := 'N'
                              else
                              If (ValeurSeg = '10') then
                                 ValeurSeg := 'F';
                              end
                           else
                           If (Segment='S41.G01.00.020') then
                              begin
                              If IsNumeric (ValeurSeg) then
                                 ValeurSeg:= FloatToStr (StrToFloat(ValeurSeg)/100)
                              else
                                 ValeurSeg:= '0';
                              end
                           else
                           If Segment = 'S41.G01.00.028' then
                              begin
                              If IsNumeric(ValeurSeg) then
                                 ValeurSeg := FloatToStr(StrToFloat(ValeurSeg)/100)
                              else
                                 ValeurSeg := '0';
                              end
                           else
                           If (Segment='S41.G01.00.034') Then
                              begin
                              If ValeurSeg='01' Then
                                 ValeurSeg:='F';
                              if ValeurSeg='02' Then
                                 ValeurSeg:='E';
                              end
                           else
                           If Copy(Segment,1,10) = 'S41.G01.01' then
                              begin
                              If Segment = 'S41.G01.01.001' then
                                 begin
                                 TL:= TobLexique.FindFirst(['PDL_DADSSEGMENT','ORDRESEG'],
                                                           [Segment,IntToStr(OrdreRC+NbRC*IncRC)+'L'],
                                                           False);
                                 Segment:= Segment+'.'+IntToStr(OrdreRC+NbRC*IncRC);
                                 If TL <> Nil then
                                    T.PutValue(Segment+'L',
                                              RechDom (TL.GetValue('TABLETTE'),ValeurSeg,False));
                                 NbRC:= NbRC + 1;
                                 end;

                              If Segment='S41.G01.01.002' then
                                 Segment:= Segment+'.'+
                                           IntToStr(OrdreRC+1+(NbRC-1)*IncRC);
                              end
                           else
                           If Copy(Segment,1,10) = 'S41.G01.02' then
                              begin
                              If Segment = 'S41.G01.02.001' then
                                 begin
                                 TL:= TobLexique.FindFirst(['PDL_DADSSEGMENT','ORDRESEG'],
                                                           [Segment,IntToStr(OrdreBBS+NbBBS*IncBBS)+'L'],
                                                           False);
                                 Segment:= Segment+'.'+IntToStr(OrdreBBS+NbBBS*IncBBS);
                                 if TL <> Nil then
                                    T.PutValue (Segment+'L',
                                                RechDom (TL.GetValue('TABLETTE'),
                                                         ValeurSeg,False));
                                 NbBBS := NbBBS + 1;
                                 end;

                              if ((ExerciceDads>='2003') and
                                 (Segment = 'S41.G01.02.002.001')) or
                                 ((ExerciceDads<'2003') and
                                 (Segment = 'S41.G01.02.002')) then
                                 Segment:= Segment+'.'+
                                           IntToStr(OrdreBBS+1+(NbBBS-1)*IncBBS);

                              If Segment = 'S41.G01.02.002.002' then
                                 Segment:= Segment+'.'+
                                           IntToStr(OrdreBBS+1+(NbBBS-1)*IncBBS);
                              end
                           else
                           If Copy(Segment,1,10) = 'S41.G01.03' then
                              begin
                              If Segment = 'S41.G01.03.001' then
                                 begin
                                 TL:= TobLexique.FindFirst(['PDL_DADSSEGMENT','ORDRESEG'],
                                                           [Segment,IntToStr(OrdreBPS+NbBPS*IncBPS)+'L'],
                                                           False);
                                 Segment:= Segment+'.'+IntToStr(OrdreBPS+NbBPS*IncBPS);
                                 If TL <> Nil then
                                    T.PutValue(Segment+'L',
                                               RechDom (TL.GetValue('TABLETTE'),
                                                        ValeurSeg,False));
                                 NbBPS:= NbBPS + 1;
                                 end;

                              If ((ExerciceDads>='2003') and
                                 (Segment = 'S41.G01.03.002.001')) or
                                 ((ExerciceDads<'2003') and
                                 (Segment = 'S41.G01.03.002')) then
                                 Segment:= Segment+'.'+
                                           IntToStr(OrdreBPS+1+(NbBPS-1)*IncBPS);

                              If Segment = 'S41.G01.03.002.002' then
                                 Segment:= Segment+'.'+
                                           IntToStr(OrdreBPS+1+(NbBPS-1)*IncBPS);
                              end
                           else
                           If Copy(Segment,1,10) = 'S41.G01.04' then
                              begin
                              If Segment = 'S41.G01.04.001' then
                                 begin
                                 TL:= TobLexique.FindFirst(['PDL_DADSSEGMENT','ORDRESEG'],
                                                           [Segment,IntToStr(OrdreTSI+NbTSI*IncTSI)+'L'],
                                                           False);
                                 Segment:= Segment+'.'+IntToStr(OrdreTSI+NbTSI*IncTSI);
                                 If TL <> Nil then
                                    T.PutValue (Segment+'L',
                                                RechDom (TL.GetValue('TABLETTE'),
                                                         ValeurSeg,False));
                                 NbTSI := NbTSI + 1;
                                 end;

                              If Segment = 'S41.G01.04.002' then
                                 Segment:= Segment+'.'+
                                           IntToStr(OrdreTSI+1+(NbTSI-1)*IncTSI);

                              If ((ExerciceDads>='2003') and
                                 (Segment = 'S41.G01.04.003.001')) or
                                 ((ExerciceDads<'2003') and
                                 (Segment = 'S41.G01.04.003')) then
                                 Segment:= Segment+'.'+
                                           IntToStr(OrdreTSI+2+(NbTSI-1)*IncTSI);

                              If Segment = 'S41.G01.04.003.002' then
                                 Segment:= Segment+'.'+
                                           IntToStr(OrdreTSI+2+(NbTSI-1)*IncTSI);

                              If ((ExerciceDads>='2003') and
                                 (Segment = 'S41.G01.04.004.001')) or
                                 ((ExerciceDads<'2003') and
                                 (Segment = 'S41.G01.04.004')) then
                                 Segment:= Segment+'.'+
                                           IntToStr(OrdreTSI+4+(NbTSI-1)*IncTSI);

                              If Segment = 'S41.G01.04.004.002' then
                                 Segment:= Segment+'.'+
                                           IntToStr(OrdreTSI+4+(NbTSI-1)*IncTSI);
                              end
                           else
                           If Copy(Segment,1,10) = 'S41.G01.06' then
                              begin
                              If Segment = 'S41.G01.06.001' then
                                 begin
                                 TL:= TobLexique.FindFirst(['PDL_DADSSEGMENT','ORDRESEG'],
                                                           [Segment,IntToStr(OrdreExo+NbExo*IncExo)+'L'],
                                                           False);
                                 Segment:= Segment+'.'+
                                           IntToStr(OrdreExo+NbExo*IncExo);
                                 If TL <> nil then
                                    T.PutValue (Segment+'L',
                                                RechDom (TL.GetValue('TABLETTE'), ValeurSeg,False));
                                 NbExo := NbExo + 1;
                                 end;

                              If ((ExerciceDads>='2003') and
                                 (Segment = 'S41.G01.06.002.001')) or
                                 ((ExerciceDads<'2003') and
                                 (Segment = 'S41.G01.06.002')) then
                                 Segment:= Segment+'.'+
                                           IntToStr(OrdreExo+1+(NbExo-1)*IncExo);

                              If Segment = 'S41.G01.06.002.002' then
                                 Segment:= Segment+'.'+
                                           IntToStr(OrdreExo+1+(NbExo-1)*IncExo);

                              If ((ExerciceDads>='2003') and
                                 (Segment = 'S41.G01.06.003.001')) or
                                 ((ExerciceDads<'2003') and
                                 (Segment = 'S41.G01.06.003')) then
                                 Segment:= Segment+'.'+
                                           IntToStr(OrdreExo+3+(NbExo-1)*IncExo);

                              If Segment = 'S41.G01.06.003.002' then
                                 Segment:= Segment+'.'+
                                           IntToStr(OrdreExo+3+(NbExo-1)*IncExo);
                              end
                           else
                           if ((Segment='S41.G02.00.009') OR
                              (Segment='S41.G02.00.010')) then
{PT36-2
                              ValeurSeg := '0'+ValeurSeg;
}
                              ValeurSeg := Copy(ValeurSeg, 2,1)
                           else
//PT47
                           If (Copy (Segment,1,10)='S41.G30.10') then
                              begin
                              If (Segment='S41.G30.10.001') then
                                 OrdreG30:= 799+StrToInt (ValeurSeg)*3
                              else
                              If (Segment='S41.G30.10.002.001') then
                                 Segment:= Segment+'.'+IntToStr (OrdreG30);
                              end
                           else
                           If (Copy (Segment,1,10)='S41.G30.11') then
                              begin
                              If (Segment='S41.G30.11.001.001') then
                                 Segment:= Segment+'.'+IntToStr (821+NbAG*2)
                              else
                              If (Segment='S41.G30.11.001.002') then
                                 begin
                                 Segment:= Segment+'.'+IntToStr (822+NbAG*2);
                                 NbAG:= NbAG+1;
                                 end;
                              end
                           else
                           If (Copy (Segment,1,10)='S41.G30.15') then
                              begin
                              If (Segment='S41.G30.15.001') then
                                 OrdreG30:= 829+StrToInt (ValeurSeg)*3
                              else
                              If (Segment='S41.G30.15.002.001') then
                                 Segment:= Segment+'.'+IntToStr (OrdreG30);
                              end
                           else
                           If (Copy (Segment,1,10)='S41.G30.20') then
                              begin
                              If (Segment='S41.G30.20.001') then
                                 OrdreG30:= 839+StrToInt (ValeurSeg)*3
                              else
                              If (Segment='S41.G30.20.002.001') then
                                 Segment:= Segment+'.'+IntToStr (OrdreG30);
                              end
                           else
                           If (Copy (Segment,1,10)='S41.G30.25') then
                              begin
                              If (Segment='S41.G30.25.001') then
                                 OrdreG30:= 859+StrToInt (ValeurSeg)*3
                              else
                              If (Segment='S41.G30.25.002.001') then
                                 Segment:= Segment+'.'+IntToStr (OrdreG30);
                              end
                           else
//FIN PT47
//PT36-2
                           If (Segment='S44.G01.00.002') then
                              begin
                              If IsNumeric (ValeurSeg) then
                                 ValeurSeg:= FloatToStr (StrToFloat(ValeurSeg)/100)
                              else
                                 ValeurSeg:= '0';
                              end
                           else
//FIN PT36-2
                           If Copy(Segment,1,10) = 'S45.G01.01' then
                              begin
                              If Segment = 'S45.G01.01.001' then
                                 begin
                                 Segment:= Segment+'.'+IntToStr(OrdrePrev+NbPrev*IncPrev);
                                 NbPrev:= NbPrev + 1;
                                 end;

                              If Segment = 'S45.G01.01.002' then
                                 Segment:= Segment+'.'+
                                           IntToStr(Ordreprev+1+(NbPrev-1)*IncPrev);
                              If Segment = 'S45.G01.01.003' then
                                 Segment:= Segment+'.'+
                                           IntToStr(Ordreprev+2+(NbPrev-1)*IncPrev);
{PT45
                              If Segment = 'S45.G01.01.004' then
}
                              If (ExerciceDads>='2007') then
                                 begin
                                 If (Segment='S45.G01.01.004.001') then
                                    Segment:= Segment+'.'+
                                              IntToStr(Ordreprev+3+(NbPrev-1)*IncPrev);
                                 end
                              else
                                 begin
                                 If Segment = 'S45.G01.01.004' then
                                    Segment:= Segment+'.'+
                                              IntToStr(Ordreprev+3+(NbPrev-1)*IncPrev);
                                 end;
//FIN PT45
                              If Segment = 'S45.G01.01.005' then
                                 Segment:= Segment+'.'+
                                           IntToStr(Ordreprev+4+(NbPrev-1)*IncPrev);
                              If Segment = 'S45.G01.01.006' then
                                 Segment:= Segment+'.'+
                                           IntToStr(Ordreprev+5+(NbPrev-1)*IncPrev);
                              If Segment = 'S45.G01.01.007.001' then
                                 Segment:= Segment+'.'+
                                           IntToStr(Ordreprev+7+(NbPrev-1)*IncPrev);
                              If Segment = 'S45.G01.01.008' then
                                 Segment:= Segment+'.'+
                                           IntToStr(Ordreprev+8+(NbPrev-1)*IncPrev);
                              If Segment = 'S45.G01.01.009' then
                                 Segment:= Segment+'.'+
                                           IntToStr(Ordreprev+9+(NbPrev-1)*IncPrev);
                              If Segment = 'S45.G01.01.010' then
                                 Segment:= Segment+'.'+
                                           IntToStr(Ordreprev+10+(NbPrev-1)*IncPrev);
                              end
                           else
                           If (Segment='S66.G01.00.003') then
                              begin
                              If IsNumeric (ValeurSeg) then
                                 ValeurSeg:= FloatToStr (StrToFloat (ValeurSeg)/100)
                              else
                                 ValeurSeg:= '0';
                              end
                           else
                           If Segment = 'S66.G01.00.009' then
                              begin
                              If ValeurSeg = '01' then
                                 ValeurSeg := 'HEB'
                              else
                                 ValeurSeg := 'MEN';
                              end
                           else
                           If (Segment='S66.G01.00.010') then
                              begin
                              If IsNumeric (ValeurSeg) then
                                 ValeurSeg:= FloatToStr (StrToFloat (ValeurSeg)/100)
                              else
                                 ValeurSeg:= '0';
                              end
                           else
                           If (Segment='S66.G01.00.011') then
                              begin
                              If IsNumeric (ValeurSeg) then
                                 ValeurSeg:= FloatToStr (StrToFloat (ValeurSeg)/100)
                              else
                                 ValeurSeg:= '0';
                              end
                           else
                           If Segment = 'S66.G01.00.012' then
                              begin
                              Q:= OpenSQL ('SELECT CO_LIBELLE FROM COMMUN WHERE'+
                                           ' CO_TYPE="PSF" AND'+
                                           ' CO_ABREGE="'+ValeurSeg+'"',
                                           True);
                              If Not Q.Eof then
                                 ValeurSeg:= Q.FindField('CO_LIBELLE').AsString
                              else
                                 ValeurSeg:= '';
                              Ferme(Q);
                              end;
                           end
                        else
                           //Edition des périodes d'inactivité
                           begin
                           if (Segment='S46.G01.02.002') then
                              ValeurSeg:= FloatToStr(StrToFloat(ValeurSeg)/100);
                           end;

                        EnregistreSegment(TL, ExerciceDads, Segment, ValeurSeg, T);

                        Readln (FLect,S);
                        Structure:= Copy (S, 1, 3);
{PT36-2
                        Segment:= Copy (S,1,14);
}
                        Virgule:= Pos(',', S);
                        Segment:= Copy (S,1,Virgule-1);
//FIN PT36-2
                        if (Structure>='S70') then
                           break;
                        end;
                  CreationTob:= False;
                  end
               else
                  begin
                  Readln (FLect,S);
                  Structure:= Copy (S, 1, 3);
{PT36-2
                  Segment:= Copy (S,1,14);
}
                  Virgule:= Pos(',', S);
                  Segment:= Copy (S,1,Virgule-1);
//FIN PT36-2
                  if (Structure>='S70') then
                     break;
                  While ((Segment<>'') and (Segment<>'S30.G01.00.001') and
                        (Segment<>'S41.G01.00.001') and
                        (Segment<>'S46.G01.00.001')) do
                        begin
                        Readln (FLect,S);
                        Structure:= Copy (S, 1, 3);
{PT36-2
                        Segment:= Copy (S,1,14);
}
                        Virgule:= Pos(',', S);
                        Segment:= Copy (S,1,Virgule-1);
//FIN PT36-2
                        if ((Segment='S41.G01.00.019') and
                           (TypeEdition='I')) then
                           begin
                           LL:= Length (S);
                           Salarie:= Copy (S, 17, LL-17);
                           end;

                        if (Structure>='S70') then
                           break;
                        end;
                  end;
               end;
         If (Structure='S90') then
            break;
         while (Structure<'S80') do
               begin
               Readln (FLect,S);
               Structure:= Copy (S, 1, 3);
{PT36-2
               Segment:= Copy (S,1,14);
}
               Virgule:= Pos(',', S);
               Segment:= Copy (S,1,Virgule-1);
//FIN PT36-2
               end;
         if (Structure='S90') then
            break;
         end
      else
         //Edition des honoraires
         begin
         Structure:= Copy (S, 1, 3);
         while (Structure<'S70') do
               begin
               Readln (FLect,S);  // Recherche du premier honoraire
               Structure:= Copy (S, 1, 3);
               end;
         If (Structure='S90') then
            break;
         Structure:= Copy (S, 1, 3);
{PT36-2
         Segment:= Copy (S,1,14);
}
         Virgule:= Pos(',', S);
         Segment:= Copy (S,1,Virgule-1);
//FIN PT36-2
         while (Structure='S70') do
               begin
               if ((Segment='S70.G01.00.001') or
{PT36-2
                  ((Segment='S70.G01.00.002') and (Copy(S,15,4)='.001'))) then
}
                  ((Segment='S70.G01.00.002.001'))) then
                  //Si nouveau honoraire
                  begin
                  FreeAndNil (TobEntete);
                  TobEntete:= Tob.Create('MonEntete',Nil,-1);
                  T:= Tob.create('Filles',TobPeriodes,-1);
                  T.AddchampSupValeur('DOSSIER', GetControlText('DOSSIER'), False);
                  T.AddchampSupValeur('DATEDEB', GetControlText('DATEDEB'), False);
                  T.AddchampSupValeur('DATEFIN', GetControlText('DATEFIN'), False);
                  CreationSegments (ExerciceDads, T, TobEntete, TobLexique);
                  OrdreRC:= 101;
                  NbRC:= 0;
                  IncRC:= 3;
                  MaxRC:= 113;
                  end;
               LL:= Length (S);
{PT36-2
               If ((Segment='S70.G01.00.002') or (Segment='S70.G01.00.003') or
                  (Segment='S70.G01.00.004')) then
                  begin
                  Segment:= Copy (S, 1, 18);
                  ValeurSeg:= Copy (S, 21, LL-21);
                  end
               else
                  ValeurSeg:= Copy (S, 17, LL-17);
}                  
               Virgule:= Pos(',', S);
               Segment:= Copy (S,1,Virgule-1);
               ValeurSeg:= Copy (S, Virgule+2, LL-(Virgule+2));
//FIN PT36-2
               TobEntete.AddchampSupValeur(Segment, ValeurSeg, False);
               TL:= TobLexique.FindFirst (['PDL_DADSSEGMENT'], [Segment],False);

               If Copy(Segment,1,10) = 'S70.G01.01' then
                  begin
                  If Segment = 'S70.G01.01.001' then
                     begin
                     TL:= TobLexique.FindFirst(['PDL_DADSSEGMENT','ORDRESEG'],
                                               [Segment,IntToStr(OrdreRC+NbRC*IncRC)],
                                                False);
                     Segment:= Segment+'.'+IntToStr(OrdreRC+NbRC*IncRC);
                     If TL <> Nil then
                        T.PutValue(Segment,
                                   RechDom (TL.GetValue('TABLETTE'),ValeurSeg,False));
                     NbRC:= NbRC + 1;
                     end;

                  If Segment='S70.G01.01.002' then
                     Segment:= Segment+'.'+IntToStr(OrdreRC+1+(NbRC-1)*IncRC);
                  end;

               EnregistreSegment(TL, ExerciceDads, Segment, ValeurSeg, T);
               Readln (FLect,S);  // Recherche des honoraires suivants
               Structure:= Copy (S, 1, 3);
{PT36-2
               Segment:= Copy (S, 1, 14);
}
               Virgule:= Pos(',', S);
               Segment:= Copy (S,1,Virgule-1);
//FIN PT36-2
               end;
         end;
      while (Structure='S80') do
            begin
{PT36-2
            if (Segment='S80.G01.00.001') then
               begin
               Segment:= Copy (S, 1, 18);
}
            if (Segment='S80.G01.00.001.002') then
               begin
               LL:= Length (S);
{PT36-2
               ValeurSeg:= Copy (S, 21, LL-21);
}
               ValeurSeg:= Copy (S, Virgule+2, LL-(Virgule+2));
//FIN PT36-2
               TE:= Tob.Create ('L établissement', TobEtabliss, -1);
               TE.AddchampSupValeur('NIC', ValeurSeg);
               end;
{PT36-2
               end;
}
            if (Segment='S80.G01.00.002') then
               begin
               LL:= Length (S);
{PT36-2
               ValeurSeg:= Copy (S, 17, LL-17);
}
               ValeurSeg:= Copy (S, Virgule+2, LL-(Virgule+2));
//FIN PT36-2
               TE.AddchampSupValeur('LIBELLE', ValeurSeg);
               end;
            Readln (FLect,S);
            Structure:= Copy (S, 1, 3);
{PT36-2
            Segment:= Copy (S,1,14);
}
            Virgule:= Pos(',', S);
            Segment:= Copy (S,1,Virgule-1);
//FIN PT36-2
            end;
      For i:=0 to TobEtabliss.Detail.Count-1 do
          begin
          NIC:= TobEtabliss.Detail[i].GetValue('NIC');
          if (TypeEdition<>'H') then
             begin
             T:= TobPeriodes.FindFirst(['S41.G01.00.005'], [NIC], False);
             while T<>nil do
                   begin
                   T.PutValue('S41.G01.00.005', TobEtabliss.Detail[i].GetValue('LIBELLE'));
                   T:= TobPeriodes.FindNext(['S41.G01.00.005'], [NIC], False);
                   end;
             end
          else
             begin
             T:= TobPeriodes.FindFirst(['S70.G01.00.014'], [NIC], False);
             while T<>nil do
                   begin
                   T.PutValue('S70.G01.00.014', TobEtabliss.Detail[i].GetValue('LIBELLE'));
                   T:= TobPeriodes.FindNext(['S70.G01.00.014'], [NIC], False);
                   end;
             end;
          end;
      end;

if (TypeEdition='R') then
   begin
   TobEditRecap:= Tob.Create('LeReacp',Nil,-1);
   TListeSalarie:= Tob.Create('Liste des salaries', nil, -1);
   For i := 0 to TobPeriodes.Detail.Count - 1 do
       begin
       Salarie:= TobPeriodes.Detail[i].GetValue('S41.G01.00.019');
       Dossier:= TobPeriodes.Detail[i].GetValue('DOSSIER');
       TSalD:= TListeSalarie.FindFirst (['MATRICULE', 'DOSSIER'],
                                        [Salarie, Dossier], TRUE);
       if (TSalD=nil) then
          begin
          TSal:= Tob.Create('Le salarie', TListeSalarie, -1);
          TSal.AddchampSupValeur('DOSSIER', Dossier);
          TSal.AddchampSupValeur('MATRICULE', Salarie);
          TobEditRecap.DelChampSup('PDE_SALARIE', False);
          end
       else
          TobEditRecap.AddchampSupValeur('PDE_SALARIE', Salarie);
       if TobEditRecap.FieldExists('PDE_SALARIE') then
          T:= TobEditRecap.FindFirst(['DOSSIER', 'PDE_SALARIE'],[Dossier, Salarie],False)
       else
          T:= Nil;

       If T <> Nil then
          begin
          For x := 0 to TobLexique.Detail.Count - 1 do
              begin
              Segment:= TobLexique.Detail[x].GetValue('PDL_DADSSEGMENT');
              If (Segment <> 'S30.G01.00.002') and
                 (Segment <> 'S30.G01.00.003') and
                 (Segment <> 'S41.G01.00.019') then
                 T.PutValue (Segment,
                             T.GetValue(Segment)+
                             TobPeriodes.Detail[i].GetValue(Segment));
              end;
          T.PutValue('ORDRE',T.GetValue('ORDRE')+1);
          end
       else
          begin
          T:= Tob.Create('UnePeriode',TobEditRecap,-1);
          For x := 0 to TobLexique.Detail.Count - 1 do
              begin
              Segment:= TobLexique.Detail[x].GetValue('PDL_DADSSEGMENT');
              T.AddchampSupValeur (Segment,
                                   TobPeriodes.Detail[i].GetValue(Segment));
              If Segment = 'S41.G01.00.019' then
                 T.AddchampSupValeur ('PDE_SALARIE',
                                      TobPeriodes.Detail[i].GetValue(Segment));
              T.AddchampSupValeur('DONNEE','');
              T.AddchampSupValeur('ET_ETABLISSEMENT','');
              T.AddchampSupValeur('RUPTURE','');
              T.AddchampSupValeur('ORDRE',1);
              T.AddchampSupValeur('NBSAL',1);
              T.AddChampSupValeur ('DOSSIER',
                                   TobPeriodes.Detail[i].GetValue('DOSSIER'));
              T.AddChampSupValeur ('DATEDEB',
                                   TobPeriodes.Detail[i].GetValue('DATEDEB'));
              T.AddChampSupValeur ('DATEFIN',
                                   TobPeriodes.Detail[i].GetValue('DATEFIN'));
//PT36-2
              if (Arg='B') then
                 T.AddchampSupValeur('BTP', 'BTP', False)
              else
                 T.AddchampSupValeur('BTP', '', False);
//FIN PT36-2
              end;
          end;
       end;
   If GetCheckBoxState('TRISAL') = CbChecked then
      TobEditRecap.Detail.Sort('DOSSIER;S30.G01.00.002;PDE_SALARIE')
   else
      TobEditRecap.Detail.Sort('DOSSIER;PDE_SALARIE');
   FreeAndNil (TobPeriodes);
   end
else
   begin
   if (TypeEdition<>'H') then
      begin
      Ordre:= 1;
      For i := 0 to TobPeriodes.Detail.Count-1 do
          begin
          Salarie:= TobPeriodes.Detail[i].GetValue('S41.G01.00.019');
          TobPeriodes.Detail[i].AddchampSupValeur('PDE_SALARIE',Salarie,False);
          TobPeriodes.Detail[i].AddchampSupValeur('PDE_ORDRE',Ordre,False);
          Ordre:= Ordre+1;
          end;
      If GetCheckBoxState('TRISAL') = CbChecked then
         TobPeriodes.Detail.Sort('DOSSIER;S30.G01.00.002;PDE_SALARIE;PDE_ORDRE')
      else
         TobPeriodes.Detail.Sort('DOSSIER;PDE_SALARIE;PDE_ORDRE');
      end
   else
      begin
      Ordre:= 1;
      For i := 0 to TobPeriodes.Detail.Count-1 do
          begin
          TobPeriodes.Detail[i].AddchampSupValeur('PDE_ORDRE',Ordre,False);
          Ordre:= Ordre+1;
          end;
      If GetCheckBoxState('TRISAL') = CbChecked then
         TobPeriodes.Detail.Sort('DOSSIER;S70.G01.00.001;S70.G01.00.002.001;PDE_ORDRE')
      else
         TobPeriodes.Detail.Sort('DOSSIER;PDE_ORDRE');
      end;
   end;

if (TypeEdition='R') then
   SetControlText('NBSAL', IntToStr(TListeSalarie.FillesCount(1)));

FreeAndNil (TobLexique);
FreeAndNil (TobEtabliss);
FreeAndNil (TListeSalarie);
FreeAndNil (TobEntete);
CloseFile(FLect);
result:= ExerciceDads;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/02/2005
Modifié le ... :   /  /
Description .. : Création des champs correspondants aux segments de la
Suite ........ : DADS
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGEDIT_DADSPER.CreationSegments(ExerciceDads:string; var T:tob; TobEntete, TobLexique:tob);
var
x : Integer;
Nature, OrdreSeg, Segment, Tablette, ValeurSeg : string;
begin
For x:=0 to TobLexique.Detail.Count-1 do
    begin
    Segment:= TobLexique.Detail[x].GetValue('PDL_DADSSEGMENT');
    Nature:= TobLexique.Detail[x].GetValue('PDL_DADSNATURE');
    OrdreSeg:= TobLexique.Detail[x].GetValue('ORDRESEG');
    Tablette:= TobLexique.Detail[x].GetValue('TABLETTE');
    If OrdreSeg <> '' then
       Segment:= Segment+'.'+OrdreSeg;
    If (TobEntete.FieldExists(Segment)) then
       ValeurSeg:= TobEntete.GetValue(Segment)
    else
       ValeurSeg:= '';

    If segment = 'S30.G01.00.007' then  // Civilité
       begin
       If ValeurSeg = '01' then
          ValeurSeg := 'Monsieur'
       else
       If ValeurSeg = '02' then
          ValeurSeg := 'Madame'
       else
       If ValeurSeg = '03' then
          ValeurSeg := 'Mademoiselle';
       end;

    If (Tablette <> '') and (ValeurSeg <> '') then
       //Recherhce tablette associé
       T.AddchampSupValeur(Segment+'L',RechDom(Tablette,ValeurSeg,False),False);
    If Nature = 'D' then  // Cas d'une date
       begin
       If Length(ValeurSeg) = 4 then
          ValeurSeg := ValeurSeg+ExerciceDads;
       If ValeurSeg <> '' then
          T.AddchampSupValeur (Segment, StrToDate (Copy(ValeurSeg,1,2)+'/'+
                                                   Copy(ValeurSeg,3,2)+'/'+
                                                   Copy(ValeurSeg,5,4)),False)
       else
          T.AddchampSupValeur(Segment,IDate1900,False);
       end
    else
    if nature = 'N' then //Cas d'un numérique
       begin
       If IsNumeric (ValeurSeg) then
          T.AddchampSupValeur (Segment, StrToFloat(ValeurSeg), False)
       else
          T.AddchampSupValeur(Segment,0,False);
       end
    else
          T.AddchampSupValeur(Segment,ValeurSeg,False);
    end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/02/2005
Modifié le ... :   /  /
Description .. : Enregistrement de la valeur du segment à éditer dans la tob
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGEDIT_DADSPER.EnregistreSegment(TL : tob; ExerciceDads, Segment, ValeurSeg : string;var T : tob);
var
Tablette, Nature : string;
begin
if (TL<>Nil) then
   begin
   Nature:= TL.GetValue('PDL_DADSNATURE');
   Tablette:= TL.GetValue('TABLETTE');
//PT36-2
{PT45
   If ((Segment = 'S41.G01.01.001.101') or (Segment = 'S41.G01.01.001.103') or
      (Segment = 'S41.G01.01.001.105') or (Segment = 'S41.G01.01.001.107') or
      (Segment = 'S41.G01.01.001.109') or (Segment = 'S41.G01.02.001.201') or
      (Segment = 'S41.G01.02.001.204') or (Segment = 'S41.G01.02.001.207') or
      (Segment = 'S41.G01.02.001.210') or (Segment = 'S41.G01.03.001.301') or
      (Segment = 'S41.G01.03.001.304') or (Segment = 'S41.G01.03.001.307') or
      (Segment = 'S41.G01.03.001.310') or (Segment = 'S41.G01.04.001.401') or
      (Segment = 'S41.G01.04.001.407') or (Segment = 'S41.G01.04.001.413') or
      (Segment = 'S41.G01.04.001.419') or (Segment = 'S41.G01.06.001.601') or
      (Segment = 'S41.G01.06.001.606') or (Segment = 'S41.G01.06.001.611') or
      (Segment = 'S41.G01.06.001.616') or (Segment = 'S41.G01.06.001.621') or
      (Segment = 'S41.G01.06.001.626') or (Segment = 'S41.G01.06.001.631') or
      (Segment = 'S41.G01.06.001.636') or (Segment = 'S41.G01.06.001.641') or
      (Segment = 'S41.G01.06.001.646')) then
      Tablette:= '';
}
   If (ExerciceDads>='2007') then
      begin
      If ((Segment = 'S41.G01.01.001.111') or (Segment = 'S41.G01.01.001.113') or
         (Segment = 'S41.G01.01.001.115') or (Segment = 'S41.G01.01.001.117') or
         (Segment = 'S41.G01.01.001.119') or (Segment = 'S41.G01.02.001.201') or
         (Segment = 'S41.G01.02.001.204') or (Segment = 'S41.G01.02.001.207') or
         (Segment = 'S41.G01.02.001.210') or (Segment = 'S41.G01.03.001.301') or
         (Segment = 'S41.G01.03.001.304') or (Segment = 'S41.G01.03.001.307') or
         (Segment = 'S41.G01.03.001.310') or (Segment = 'S41.G01.04.001.401') or
         (Segment = 'S41.G01.04.001.407') or (Segment = 'S41.G01.04.001.413') or
         (Segment = 'S41.G01.04.001.419') or (Segment = 'S41.G01.06.001.601') or
         (Segment = 'S41.G01.06.001.606') or (Segment = 'S41.G01.06.001.611') or
         (Segment = 'S41.G01.06.001.616') or (Segment = 'S41.G01.06.001.621') or
         (Segment = 'S41.G01.06.001.626') or (Segment = 'S41.G01.06.001.631') or
         (Segment = 'S41.G01.06.001.636') or (Segment = 'S41.G01.06.001.641') or
         (Segment = 'S41.G01.06.001.646')) then
         Tablette:= '';
      end
   else
      begin
      If ((Segment = 'S41.G01.01.001.101') or (Segment = 'S41.G01.01.001.103') or
         (Segment = 'S41.G01.01.001.105') or (Segment = 'S41.G01.01.001.107') or
         (Segment = 'S41.G01.01.001.109') or (Segment = 'S41.G01.02.001.201') or
         (Segment = 'S41.G01.02.001.204') or (Segment = 'S41.G01.02.001.207') or
         (Segment = 'S41.G01.02.001.210') or (Segment = 'S41.G01.03.001.301') or
         (Segment = 'S41.G01.03.001.304') or (Segment = 'S41.G01.03.001.307') or
         (Segment = 'S41.G01.03.001.310') or (Segment = 'S41.G01.04.001.401') or
         (Segment = 'S41.G01.04.001.407') or (Segment = 'S41.G01.04.001.413') or
         (Segment = 'S41.G01.04.001.419') or (Segment = 'S41.G01.06.001.601') or
         (Segment = 'S41.G01.06.001.606') or (Segment = 'S41.G01.06.001.611') or
         (Segment = 'S41.G01.06.001.616') or (Segment = 'S41.G01.06.001.621') or
         (Segment = 'S41.G01.06.001.626') or (Segment = 'S41.G01.06.001.631') or
         (Segment = 'S41.G01.06.001.636') or (Segment = 'S41.G01.06.001.641') or
         (Segment = 'S41.G01.06.001.646')) then
         Tablette:= '';
      end;
//FIN PT45
//FIN PT36-2
   If (Tablette<>'') and (ValeurSeg<>'') then
      T.PutValue (Segment, RechDom (Tablette, ValeurSeg, False))
   else
      If ((Nature='D') or (segment='S41.G01.00.001') or
         (segment='S41.G01.00.003') or (segment='S46.G01.00.002') or
         (segment='S46.G01.00.003')) then
         begin
         If Length(ValeurSeg) = 4 then
            ValeurSeg:= ValeurSeg+ExerciceDads;
         If (ValeurSeg<>'') then
            T.PutValue (Segment, StrToDate (Copy (ValeurSeg,1,2)+'/'+
                                            Copy (ValeurSeg,3,2)+'/'+
                                            Copy (ValeurSeg,5,4)))
         else
            T.PutValue(Segment,IDate1900);
         end
      else
         if nature = 'N' then
            begin
            If IsNumeric (ValeurSeg) then
               T.PutValue(Segment,StrToFloat(ValeurSeg))
            else
               T.PutValue(Segment,0);
            end
         else
            T.PutValue(Segment,ValeurSeg);
   end
else
   begin
   Nature:= '';
   Tablette:= '';
   end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : JL
Créé le ...... : 30/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDIT_DADSPER.ChargerLexiqueDads (var TobLexique:tob;
                                                 Annee:String;
                                                 Fichier:Boolean=False;
                                                 TypeEdition:string='');
var
Segment,Nature, StQ, DadsValeur : String;
TL : Tob;
Q : TQuery;
i,x, Position : Integer;
begin
//Cas de l'édition complète ou normale
if (TypeEdition='') then
   begin
   StQ:= 'SELECT PDL_DADSSEGMENT, PDL_DADSNATURE, PDL_DADSVALEUR'+
         ' FROM DADSLEXIQUE WHERE'+
         ' PDL_DADSSEGMENT>"S30" AND'+
         ' PDL_DADSSEGMENT<"S50" AND'+
         ' PDL_EXERCICEDEB <= "'+Annee+'" AND'+
         ' (PDL_EXERCICEFIN >= "'+Annee+'" OR PDL_EXERCICEFIN = "")'+
         ' ORDER BY PDL_DADSSEGMENT';
   Q:= OpenSQL (StQ, True);
   TobLexique.LoadDetailDB('LesSegments','','',Q,False);
   Ferme(Q);
   end
else
//Cas de l'édition récapitulative
If (TypeEdition='R') then
   begin
//PT39-3
   TL:= Tob.Create('LesSegments',TobLexique,-1);
   TL.AddchampSupValeur('PDL_DADSSEGMENT','S30.G01.00.001');
//FIN PT39-3
   TL:= Tob.Create('LesSegments',TobLexique,-1);
   TL.AddchampSupValeur('PDL_DADSSEGMENT','S30.G01.00.002');
   TL:= Tob.Create('LesSegments',TobLexique,-1);
   TL.AddchampSupValeur('PDL_DADSSEGMENT','S30.G01.00.003');
   TL:= Tob.Create('LesSegments',TobLexique,-1);
   TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.019');
   TL:= Tob.Create('LesSegments',TobLexique,-1);
   TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.022');
   If (Annee>='2003') then
      begin
      TL:= Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.029.001');
      TL:= Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.030.001');
      TL:= Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.032.001');
      TL:= Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.033.001');
      TL:= Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.035.001');
      TL:= Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.063.001');
      end
   else
      begin
      TL:= Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.029');
      TL:= Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.030');
      TL:= Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.032');
      TL:= Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.033');
      TL:= Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.035');
      TL:= Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.063');
      end;
   If (Annee>='2005') then
      begin
      TL:= Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.037.001');
      TL:= Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.044.001');
      TL:= Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.056.001');
      TL:= Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.057.001');
      TL := Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.058.001');
      TL:= Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.066.001');
      end
   else
      begin
      TL:= Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.037');
      TL:= Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.044');
      TL:= Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.056');
      TL:= Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.057');
      TL := Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.058');
      TL:= Tob.Create('LesSegments',TobLexique,-1);
      TL.AddchampSupValeur('PDL_DADSSEGMENT','S41.G01.00.066');
      end;
//PT47
   TL:= Tob.Create ('LesSegments', TobLexique, -1);
   TL.AddchampSupValeur ('PDL_DADSSEGMENT', 'S41.G30.35.001.001');
   TL:= Tob.Create ('LesSegments', TobLexique, -1);
   TL.AddchampSupValeur ('PDL_DADSSEGMENT', 'S41.G30.35.001.002');
   TL:= Tob.Create ('LesSegments', TobLexique, -1);
   TL.AddchampSupValeur ('PDL_DADSSEGMENT', 'S41.G30.35.002');
   TL:= Tob.Create ('LesSegments', TobLexique, -1);
   TL.AddchampSupValeur ('PDL_DADSSEGMENT', 'S41.G30.35.003.001');
   TL:= Tob.Create ('LesSegments', TobLexique, -1);
   TL.AddchampSupValeur ('PDL_DADSSEGMENT', 'S41.G30.35.004.001');
//FIN PT47
   For i := 0 to TobLexique.Detail.Count-1 do
       begin
       Segment:= TobLexique.Detail[i].GetValue('PDL_DADSSEGMENT');
       If (Segment <> 'S30.G01.00.002') and (Segment <> 'S30.G01.00.003') and
          (Segment <> 'S41.G01.00.019') then
          begin
          TobLexique.Detail[i].AddchampSupValeur('ORDRESEG','');
          TobLexique.Detail[i].AddchampSupValeur('PDL_DADSNATURE','N');
          TobLexique.Detail[i].AddchampSupValeur('TABLETTE','');
          end
       else
          begin
          TobLexique.Detail[i].AddchampSupValeur('ORDRESEG','');
          TobLexique.Detail[i].AddchampSupValeur('PDL_DADSNATURE','');
          TobLexique.Detail[i].AddchampSupValeur('TABLETTE','');
          end;
       end;
   end
else
//Cas de l'édition BTP
If (TypeEdition='B') then
   begin
   StQ:= 'SELECT PDL_DADSSEGMENT, PDL_DADSNATURE, PDL_DADSVALEUR'+
         ' FROM DADSLEXIQUE WHERE'+
         ' PDL_DADSSEGMENT>="S30" AND'+
         ' PDL_DADSSEGMENT<"S70" AND'+
         ' PDL_DADSSEGMENT NOT LIKE "S46%" AND'+
         ' PDL_EXERCICEDEB <= "'+Annee+'" AND'+
         ' (PDL_EXERCICEFIN >= "'+Annee+'" OR PDL_EXERCICEFIN = "") AND'+
         ' PDL_DADSUCCPBTP="P" ORDER BY PDL_DADSSEGMENT';
   Q:= OpenSQL (StQ, True);
   TobLexique.LoadDetailDB('LesSegments','','',Q,False);
   Ferme(Q);
   end
else
//Cas de l'édition des périodes d'inactivité
if (TypeEdition='I') then
   begin
   StQ:= 'SELECT PDL_DADSSEGMENT, PDL_DADSNATURE, PDL_DADSVALEUR'+
         ' FROM DADSLEXIQUE WHERE'+
         ' PDL_DADSSEGMENT LIKE "S46%" AND'+
         ' PDL_EXERCICEDEB <= "'+Annee+'" AND'+
         ' (PDL_EXERCICEFIN >= "'+Annee+'" OR PDL_EXERCICEFIN = "")'+
         ' ORDER BY PDL_DADSSEGMENT';
   Q:= OpenSQL (StQ, True);
   TobLexique.LoadDetailDB('LesSegments','','',Q,False);
   Ferme(Q);
   TL:= Tob.Create('LesSegments',TobLexique,-1);
   TL.AddchampSupValeur('PDL_DADSSEGMENT','S30.G01.00.002');
   TL:= Tob.Create('LesSegments',TobLexique,-1);
   TL.AddchampSupValeur('PDL_DADSSEGMENT','S30.G01.00.003');
   For i := 0 to TobLexique.Detail.Count-1 do
       begin
       Segment:= TobLexique.Detail[i].GetValue('PDL_DADSSEGMENT');
       If ((Segment='S30.G01.00.002') or (Segment='S30.G01.00.003')) then
          begin
          TobLexique.Detail[i].AddchampSupValeur('ORDRESEG','');
          TobLexique.Detail[i].AddchampSupValeur('PDL_DADSNATURE','');
          TobLexique.Detail[i].AddchampSupValeur('PDL_DADSVALEUR','');
          TobLexique.Detail[i].AddchampSupValeur('TABLETTE','');
          end;
       end;
   end
else
//Cas de l'édition des honoraires
if (TypeEdition='H') then
   begin
   StQ:= 'SELECT PDL_DADSSEGMENT, PDL_DADSNATURE, PDL_DADSVALEUR'+
         ' FROM DADSLEXIQUE WHERE'+
         ' PDL_DADSSEGMENT LIKE "S70%" AND'+
         ' PDL_EXERCICEDEB <= "'+Annee+'" AND'+
         ' (PDL_EXERCICEFIN >= "'+Annee+'" OR PDL_EXERCICEFIN = "")'+
         ' ORDER BY PDL_DADSSEGMENT';
   Q:= OpenSQL (StQ, True);
   TobLexique.LoadDetailDB('LesSegments','','',Q,False);
   Ferme(Q);
   end
else
//Cas de l'édition des traitements et salaires
if (TypeEdition='S') then
   begin
   StQ:= 'SELECT PDL_DADSSEGMENT, PDL_DADSNATURE, PDL_DADSVALEUR'+
         ' FROM DADSLEXIQUE WHERE'+
         ' (PDL_DADSSEGMENT LIKE "S30.G01.00.008%" OR'+
         ' PDL_DADSSEGMENT="S41.G01.00.005" OR'+
         ' PDL_DADSSEGMENT="S41.G01.00.021" OR'+
         ' PDL_DADSSEGMENT="S41.G01.00.022" OR'+
         ' PDL_DADSSEGMENT="S41.G01.00.035.001" OR'+
         ' PDL_DADSSEGMENT="S41.G01.00.044.001" OR'+
         ' PDL_DADSSEGMENT="S41.G01.00.063.001" OR'+
         ' PDL_DADSSEGMENT="S41.G01.00.066.001" OR'+
         ' PDL_DADSSEGMENT="S41.G01.00.067.001" OR'+
         ' PDL_DADSSEGMENT="S41.G30.35.001.002") AND'+
         ' PDL_EXERCICEDEB <= "'+Annee+'" AND'+
         ' (PDL_EXERCICEFIN >= "'+Annee+'" OR PDL_EXERCICEFIN = "")'+
         ' ORDER BY PDL_DADSSEGMENT';
   Q:= OpenSQL (StQ, True);
   TobLexique.LoadDetailDB('LesSegments','','',Q,False);
   Ferme(Q);
   end;

if ((TypeEdition<>'R') and (TypeEdition<>'S')) then
   begin
   For i := 0 to TobLexique.Detail.Count - 1 do
       begin
       Segment := TobLexique.Detail[i].GetValue('PDL_DADSSEGMENT');
       Nature := TobLexique.Detail[i].GetValue('PDL_DADSNATURE');
       TobLexique.Detail[i].AddchampSupValeur('ORDRESEG','',False);
       TobLexique.Detail[i].AddchampSupValeur('TABLETTE','',False);
       If Segment = 'S41.G01.01.001' then
          begin
//PT47
          If (Annee>='2007') then
             TobLexique.Detail[i].PutValue('ORDRESEG','111')
          else
             TobLexique.Detail[i].PutValue('ORDRESEG','101');
          TobLexique.Detail[i].PutValue('TABLETTE','');
          For x := 1 to 4 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              If (Annee>='2007') then
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(111 + (x * 2)),False)
              else
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(101 + (x * 2)),False);
              TL.AddchampSupValeur('TABLETTE','',False);
              end;
          For x := 1 to 5 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              If (Annee>='2007') then
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(109 + (x * 2))+'L',False)
              else
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(99 + (x * 2))+'L',False);
              TL.AddchampSupValeur('TABLETTE','PGINSTITUTION',False);
              end;
          end
       else
       If (Segment='S41.G01.00.013') Then
          TobLexique.Detail[i].PutValue('TABLETTE','PGCONDEMPLOI')
       else
       If (Segment='S41.G01.00.034') Then
          TobLexique.Detail[i].PutValue('TABLETTE','PGTRAVAILETRANGER')
       else
       If Segment = 'S41.G01.01.002' then
          begin
          If (Annee>='2007') then
             TobLexique.Detail[i].PutValue('ORDRESEG','112')
          else
             TobLexique.Detail[i].PutValue('ORDRESEG','102');
          TobLexique.Detail[i].PutValue('TABLETTE','');
          For x := 1 to 4 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              If (Annee>='2007') then
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(112 + (x * 2)),False)
              else
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(102 + (x * 2)),False);
//FIN PT47
              TL.AddchampSupValeur('TABLETTE','',False);
              end;
          end
       else
       If Segment = 'S41.G01.02.001' then
          begin
          TobLexique.Detail[i].PutValue('ORDRESEG','201');
          TobLexique.Detail[i].PutValue('TABLETTE','');
          For x := 1 to 3 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              TL.AddchampSupValeur('ORDRESEG',IntToStr(201 + (x * 3)),False);
              TL.AddchampSupValeur('TABLETTE','',False);
              end;
          For x := 1 to 4 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              TL.AddchampSupValeur('ORDRESEG',IntToStr(198 + (x * 3))+'L',False);
              TL.AddchampSupValeur('TABLETTE','PGBASEBRUTESPEC',False);
              end;
          end
       else
       If (Segment = 'S41.G01.02.002.001') or (Segment = 'S41.G01.02.002') then
          begin
          TobLexique.Detail[i].PutValue('ORDRESEG','202');
          TobLexique.Detail[i].PutValue('TABLETTE','');
          For x := 1 to 3 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              TL.AddchampSupValeur('ORDRESEG',IntToStr(202 + (x * 3)),False);
              TL.AddchampSupValeur('TABLETTE','',False);
              end;
          end
       else
       If Segment = 'S41.G01.03.001' then
          begin
          TobLexique.Detail[i].PutValue('ORDRESEG','301');
          TobLexique.Detail[i].PutValue('TABLETTE','');
          For x := 1 to 3 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              TL.AddchampSupValeur('ORDRESEG',IntToStr(301 + (x * 3)),False);
              TL.AddchampSupValeur('TABLETTE','',False);
              end;
          For x := 1 to 4 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              TL.AddchampSupValeur('ORDRESEG',IntToStr(298 + (x * 3))+'L',False);
              TL.AddchampSupValeur('TABLETTE','PGBASEPLAFSPEC',False);
              end;
          end
       else
       If (Segment = 'S41.G01.03.002.001') or (Segment = 'S41.G01.03.002') then
          begin
          TobLexique.Detail[i].PutValue('ORDRESEG','302');
          TobLexique.Detail[i].PutValue('TABLETTE','');
          For x := 1 to 3 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              TL.AddchampSupValeur('ORDRESEG',IntToStr(302 + (x * 3)),False);
              TL.AddchampSupValeur('TABLETTE','',False);
              end;
          end
       else
       If Segment = 'S41.G01.04.001' then
          begin
          TobLexique.Detail[i].PutValue('ORDRESEG','401');
          TobLexique.Detail[i].PutValue('TABLETTE','');
          For x := 1 to 3 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              TL.AddchampSupValeur('ORDRESEG',IntToStr(401 + (x * 6)),False);
              TL.AddchampSupValeur('TABLETTE','',False);
              end;
          For x := 1 to 4 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              TL.AddchampSupValeur('ORDRESEG',IntToStr(395 + (x * 6))+'L',False);
              TL.AddchampSupValeur('TABLETTE','PGSOMISO',False);
              end;
          end
       else
       If Segment = 'S41.G01.04.002' then
          begin
          TobLexique.Detail[i].PutValue('ORDRESEG','402');
          TobLexique.Detail[i].PutValue('TABLETTE','');
          For x := 1 to 3 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              TL.AddchampSupValeur('ORDRESEG',IntToStr(402 + (x * 6)),False);
              TL.AddchampSupValeur('TABLETTE','',False);
              end;
          end
       else
       If (Segment = 'S41.G01.04.003.001') or (Segment = 'S41.G01.04.003') then
          begin
          TobLexique.Detail[i].PutValue('ORDRESEG','403');
          TobLexique.Detail[i].PutValue('TABLETTE','');
          For x := 1 to 3 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              TL.AddchampSupValeur('ORDRESEG',IntToStr(403 + (x * 6)),False);
              TL.AddchampSupValeur('TABLETTE','',False);
              end;
          end
       else
       If ((Segment='S41.G01.04.004.001') or
          (Segment='S41.G01.04.004.001')) then
          begin
          TobLexique.Detail[i].PutValue('ORDRESEG','405');
          TobLexique.Detail[i].PutValue('TABLETTE','');
          For x := 1 to 3 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              TL.AddchampSupValeur('ORDRESEG',IntToStr(405 + (x * 6)),False);
              TL.AddchampSupValeur('TABLETTE','',False);
              end;
          end
       else
       If Segment = 'S41.G01.06.001' then
          begin
          TobLexique.Detail[i].PutValue('ORDRESEG','601');
          TobLexique.Detail[i].PutValue('TABLETTE','');
          For x := 1 to 9 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              TL.AddchampSupValeur('ORDRESEG',IntToStr(601 + (x * 5)),False);
              TL.AddchampSupValeur('TABLETTE','',False);
              end;
           For x := 1 to 10 do
               begin
               TL := Tob.Create('UneFille',TobLexique,-1);
               TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
               TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
               TL.AddchampSupValeur('ORDRESEG',IntToStr(596 + (x * 5))+'L',False);
               TL.AddchampSupValeur('TABLETTE','PGEXONERATION',False);
               end;
           end
       else
       If (Segment = 'S41.G01.06.002.001') or (Segment = 'S41.G01.06.002') then
          begin
          TobLexique.Detail[i].PutValue('ORDRESEG','602');
          TobLexique.Detail[i].PutValue('TABLETTE','');
          For x := 1 to 9 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              TL.AddchampSupValeur('ORDRESEG',IntToStr(602 + (x * 5)),False);
              TL.AddchampSupValeur('TABLETTE','',False);
              end;
          end
       else
       If (Segment = 'S41.G01.06.003.001') or (Segment = 'S41.G01.06.003') then
          begin
          TobLexique.Detail[i].PutValue('ORDRESEG','604');
          TobLexique.Detail[i].PutValue('TABLETTE','');
          For x := 1 to 9 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              TL.AddchampSupValeur('ORDRESEG',IntToStr(604 + (x * 5)),False);
              TL.AddchampSupValeur('TABLETTE','',False);
              end;
          end
       else
       If (Segment='S41.G02.00.009') Then
          TobLexique.Detail[i].PutValue('TABLETTE','PGCOLLEGEPRUD')
       else
       If (Segment='S41.G02.00.010') Then
          TobLexique.Detail[i].PutValue('TABLETTE','PGSECTIONPRUD')
       else
//PT47
       If (Segment='S41.G30.10.002.001') then
          begin
          TobLexique.Detail[i].PutValue ('ORDRESEG', '802');
          TobLexique.Detail[i].PutValue ('TABLETTE', '');
          For x := 1 to 3 do
              begin
              TL:= Tob.Create ('UneFille', TobLexique, -1);
              TL.AddchampSupValeur ('PDL_DADSSEGMENT', segment, False);
              TL.AddchampSupValeur ('PDL_DADSNATURE', Nature, False);
              TL.AddchampSupValeur ('ORDRESEG', IntToStr (802+(x*3)), False);
              TL.AddchampSupValeur ('TABLETTE', '', False);
              end;
          end
       else
       If (Segment='S41.G30.11.001.001') then
          begin
          TobLexique.Detail[i].PutValue ('ORDRESEG', '821');
          TobLexique.Detail[i].PutValue ('TABLETTE', '');
          For x := 1 to 2 do
              begin
              TL:= Tob.Create ('UneFille', TobLexique, -1);
              TL.AddchampSupValeur ('PDL_DADSSEGMENT', segment, False);
              TL.AddchampSupValeur ('PDL_DADSNATURE', Nature, False);
              TL.AddchampSupValeur ('ORDRESEG', IntToStr (821+(x*2)), False);
              TL.AddchampSupValeur ('TABLETTE', '', False);
              end;
          end
       else
       If (Segment='S41.G30.11.001.002') then
          begin
          TobLexique.Detail[i].PutValue ('ORDRESEG', '822');
          TobLexique.Detail[i].PutValue ('TABLETTE', '');
          For x := 1 to 2 do
              begin
              TL:= Tob.Create ('UneFille', TobLexique, -1);
              TL.AddchampSupValeur ('PDL_DADSSEGMENT', segment, False);
              TL.AddchampSupValeur ('PDL_DADSNATURE', Nature, False);
              TL.AddchampSupValeur ('ORDRESEG', IntToStr (822+(x*2)), False);
              TL.AddchampSupValeur ('TABLETTE', '', False);
              end;
          end
       else
       If (Segment='S41.G30.15.002.001') then
          begin
          TobLexique.Detail[i].PutValue ('ORDRESEG', '832');
          TobLexique.Detail[i].PutValue ('TABLETTE', '');
          TL:= Tob.Create ('UneFille', TobLexique, -1);
          TL.AddchampSupValeur ('PDL_DADSSEGMENT', segment, False);
          TL.AddchampSupValeur ('PDL_DADSNATURE', Nature, False);
          TL.AddchampSupValeur ('ORDRESEG', '835', False);
          TL.AddchampSupValeur ('TABLETTE', '', False);
          end
       else
       If (Segment='S41.G30.20.002.001') then
          begin
          TobLexique.Detail[i].PutValue ('ORDRESEG', '842');
          TobLexique.Detail[i].PutValue ('TABLETTE', '');
          For x := 1 to 3 do
              begin
              TL:= Tob.Create ('UneFille', TobLexique, -1);
              TL.AddchampSupValeur ('PDL_DADSSEGMENT', segment, False);
              TL.AddchampSupValeur ('PDL_DADSNATURE', Nature, False);
              TL.AddchampSupValeur ('ORDRESEG', IntToStr (842+(x*3)), False);
              TL.AddchampSupValeur ('TABLETTE', '', False);
              end;
          end
       else
       If (Segment='S41.G30.25.002.001') then
          begin
          TobLexique.Detail[i].PutValue ('ORDRESEG', '862');
          TobLexique.Detail[i].PutValue ('TABLETTE', '');
          TL:= Tob.Create ('UneFille', TobLexique, -1);
          TL.AddchampSupValeur ('PDL_DADSSEGMENT', segment, False);
          TL.AddchampSupValeur ('PDL_DADSNATURE', Nature, False);
          TL.AddchampSupValeur ('ORDRESEG', '865', False);
          TL.AddchampSupValeur ('TABLETTE', '', False);
          end
       else
       If (Segment='S41.G30.35.001.001') Then
          TobLexique.Detail[i].PutValue ('TABLETTE','PGDADSEXO')
       else
       If Segment = 'S45.G01.01.001' then
          begin
          If (Annee>='2007') then
             TobLexique.Detail[i].PutValue('ORDRESEG','922')
          else
             TobLexique.Detail[i].PutValue('ORDRESEG','802');
          TobLexique.Detail[i].PutValue('TABLETTE','PGDADSPREV');
          For x := 1 to 3 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              If (Annee>='2007') then
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(922 + (x * 12)),False)
              else
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(802 + (x * 11)),False);
              TL.AddchampSupValeur('TABLETTE','PGDADSPREV',False);
              end;
          end
       else
       If Segment = 'S45.G01.01.002' then
          begin
          If (Annee>='2007') then
             TobLexique.Detail[i].PutValue('ORDRESEG','923')
          else
             TobLexique.Detail[i].PutValue('ORDRESEG','803');
          TobLexique.Detail[i].PutValue('TABLETTE','');
          For x := 1 to 3 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              If (Annee>='2007') then
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(923 + (x * 12)),False)
              else
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(803 + (x * 11)),False);
              TL.AddchampSupValeur('TABLETTE','',False);
              end;
          end
       else
       If Segment = 'S45.G01.01.003' then
          begin
          If (Annee>='2007') then
             TobLexique.Detail[i].PutValue('ORDRESEG','924')
          else
             TobLexique.Detail[i].PutValue('ORDRESEG','804');
          TobLexique.Detail[i].PutValue('TABLETTE','');
          For x := 1 to 3 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              If (Annee>='2007') then
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(924 + (x * 12)),False)
              else
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(804 + (x * 11)),False);
              TL.AddchampSupValeur('TABLETTE','',False);
              end;
          end
       else
       If Segment = 'S45.G01.01.004' then
          begin
          TobLexique.Detail[i].PutValue('ORDRESEG','805');
          TobLexique.Detail[i].PutValue('TABLETTE','');
          For x := 1 to 3 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              TL.AddchampSupValeur('ORDRESEG',IntToStr(805 + (x * 11)),False);
              TL.AddchampSupValeur('TABLETTE','',False);
              end;
          end
//PT45
       else
       If (Segment='S45.G01.01.004.001') then
          begin
          TobLexique.Detail[i].PutValue('ORDRESEG','925');
          TobLexique.Detail[i].PutValue('TABLETTE','');
          For x := 1 to 3 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              TL.AddchampSupValeur('ORDRESEG',IntToStr(925 + (x * 12)),False);
              TL.AddchampSupValeur('TABLETTE','',False);
              end;
          end
//FIN PT45       
       else
       If Segment = 'S45.G01.01.005' then
          begin
          If (Annee>='2007') then
             TobLexique.Detail[i].PutValue('ORDRESEG','927')
          else
             TobLexique.Detail[i].PutValue('ORDRESEG','806');
          TobLexique.Detail[i].PutValue('TABLETTE','');
          For x := 1 to 3 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              If (Annee>='2007') then
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(927 + (x * 12)),False)
              else
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(806 + (x * 11)),False);
              TL.AddchampSupValeur('TABLETTE','',False);
              end;
          end
       else
       If Segment = 'S45.G01.01.006' then
          begin
          If (Annee>='2007') then
             TobLexique.Detail[i].PutValue('ORDRESEG','928')
          else
             TobLexique.Detail[i].PutValue('ORDRESEG','807');
          TobLexique.Detail[i].PutValue('TABLETTE','PGDADSBASEPREV');
          For x := 1 to 3 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              If (Annee>='2007') then
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(928 + (x * 12)),False)
              else
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(807 + (x * 11)),False);
              TL.AddchampSupValeur('TABLETTE','PGDADSBASEPREV',False);
              end;
          end
       else
       If (Segment = 'S45.G01.01.007.001') or (Segment = 'S45.G01.01.007') then
          begin
          If (Annee>='2007') then
             TobLexique.Detail[i].PutValue('ORDRESEG','929')
          else
             TobLexique.Detail[i].PutValue('ORDRESEG','808');
          TobLexique.Detail[i].PutValue('TABLETTE','');
          For x := 1 to 3 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              If (Annee>='2007') then
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(929 + (x * 12)),False)
              else
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(808 + (x * 11)),False);
              TL.AddchampSupValeur('TABLETTE','',False);
              end;
          end
       else
       If Segment = 'S45.G01.01.008' then
          begin
          If (Annee>='2007') then
             TobLexique.Detail[i].PutValue('ORDRESEG','931')
          else
             TobLexique.Detail[i].PutValue('ORDRESEG','810');
          TobLexique.Detail[i].PutValue('TABLETTE','PGDADSPREVPOP');
          For x := 1 to 3 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              If (Annee>='2007') then
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(931 + (x * 12)),False)
              else
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(810 + (x * 11)),False);
              TL.AddchampSupValeur('TABLETTE','PGDADSPREVPOP',False);
              end;
          end
       else
       If Segment = 'S45.G01.01.009' then
          begin
          If (Annee>='2007') then
             TobLexique.Detail[i].PutValue('ORDRESEG','932')
          else
             TobLexique.Detail[i].PutValue('ORDRESEG','811');
          TobLexique.Detail[i].PutValue('TABLETTE','PGSITUATIONFAMIL');
          For x := 1 to 3 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              If (Annee>='2007') then
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(932 + (x * 12)),False)
              else
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(811 + (x * 11)),False);
              TL.AddchampSupValeur('TABLETTE','PGSITUATIONFAMIL',False);
              end;
          end
       else
       If Segment = 'S45.G01.01.010' then
          begin
          If (Annee>='2007') then
             TobLexique.Detail[i].PutValue('ORDRESEG','933')
          else
             TobLexique.Detail[i].PutValue('ORDRESEG','812');
          TobLexique.Detail[i].PutValue('TABLETTE','');
          For x := 1 to 3 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              If (Annee>='2007') then
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(933 + (x * 12)),False)
              else
                 TL.AddchampSupValeur('ORDRESEG',IntToStr(812 + (x * 11)),False);
              TL.AddchampSupValeur('TABLETTE','',False);
              end;
          end
//FIN PT47
       else
       If (Segment='S66.G01.00.009') Then
          TobLexique.Detail[i].PutValue('TABLETTE','PGDADSBTPHORAIRE')
       else
       If Segment = 'S70.G01.01.001' then
          begin
          TobLexique.Detail[i].PutValue('ORDRESEG','101');
          TobLexique.Detail[i].PutValue('TABLETTE','PGDADSHONTYPEREM');
          For x := 1 to 4 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              TL.AddchampSupValeur('ORDRESEG',IntToStr(101 + (x * 3)),False);
              TL.AddchampSupValeur('TABLETTE','PGDADSHONTYPEREM',False);
              end;
          end
       else
       If Segment = 'S70.G01.01.002' then
          begin
          TobLexique.Detail[i].PutValue('ORDRESEG','102');
          For x := 1 to 5 do
              begin
              TL := Tob.Create('UneFille',TobLexique,-1);
              TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
              TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
              TL.AddchampSupValeur('ORDRESEG',IntToStr(102 + (x * 3)),False);
              TL.AddchampSupValeur('TABLETTE','',False);
              end;
          end
       else
       if (TobLexique.Detail[i].GetValue('PDL_DADSVALEUR')<>'') then
          begin
          DadsValeur:= TobLexique.Detail[i].GetValue('PDL_DADSVALEUR');
          Position:= Pos ('T', DadsValeur);
          if (Position=1) then
             begin
             System.Delete (DadsValeur, 1, 2);
             Position:= Pos (';', DadsValeur);
             DadsValeur:= Copy (DadsValeur, 1, Position-1);
             TobLexique.Detail[i].PutValue('TABLETTE',DadsValeur);
             end;
          end
       else
       If (Segment='S41.G01.00.005') Then
          begin
          if (Fichier=False) then
             TobLexique.Detail[i].PutValue('TABLETTE','TTETABLISSEMENT');
          end
       else
       If (Segment='S41.G01.00.011') Then
          begin
          if VH_Paie.PGPCS2003 then
             TobLexique.Detail[i].PutValue('TABLETTE','PGCODEPCSESE')
          else
             TobLexique.Detail[i].PutValue('TABLETTE','PGCODEEMPLOI');
          end
       else
       If (Segment='S41.G01.00.017') Then
{PT39-4
          TobLexique.Detail[i].PutValue('TABLETTE','PGCOEFFICIENT')
}
          begin
          if (Fichier=False) then
             TobLexique.Detail[i].PutValue('TABLETTE','PGCOEFFICIENT');
          end
       else
       If (Segment='S66.G01.00.017') Then
{PT39-4
          TobLexique.Detail[i].PutValue('TABLETTE','PGCOEFFICIENT')
}
          begin
          if (Fichier=False) then
             TobLexique.Detail[i].PutValue('TABLETTE','PGCOEFFICIENT');
          end
       else
       If (Segment='S66.G01.00.022') Then
          TobLexique.Detail[i].PutValue('TABLETTE','PGDADSBTPAFFILIRC')
       else
       If Not Fichier then
          begin
          If (Segment='S30.G01.00.007') Then
             TobLexique.Detail[i].PutValue('TABLETTE','YYCIVILITE')
          else
          If (Segment='S30.G01.00.012') Then
             TobLexique.Detail[i].PutValue('TABLETTE','TTPAYS')
          else
          If (Segment='S30.G01.00.013') Then
             TobLexique.Detail[i].PutValue('TABLETTE','TTPAYS')
          else
          If (Segment='S41.G01.00.010') Then
             TobLexique.Detail[i].PutValue('TABLETTE','PGLIBEMPLOI')
          else
          If (Segment='S41.G01.00.012') Then
             TobLexique.Detail[i].PutValue('TABLETTE','PGTYPECONTRAT')
          else
          If (Segment='S41.G01.00.012.001') Then
             TobLexique.Detail[i].PutValue('TABLETTE','PGTYPECONTRAT')
          else
          If (Segment='S66.G01.00.012') Then
             TobLexique.Detail[i].PutValue('TABLETTE','PGSITUATIONFAMIL');
          end;
       end;
   end;
end;

//PT39-2
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 27/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDIT_DADSPER.DateDecaleChange(Sender: TObject);
begin
AglLanceFiche ('PAY', 'DADS_DATE', '', '', '');
SetControlText ('DATEDEB', DateToStr(DebExer));
SetControlText ('DATEFIN', DateToStr(FinExer));
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 28/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDIT_DADSPER.Parametrage(Sender: TObject);
var
StExer : string;
QExer : TQuery;
begin
PGAnnee:= THAnneeRef.value;
StExer:= 'SELECT PEX_DATEDEBUT, PEX_DATEFIN, PEX_ANNEEREFER'+
         ' FROM EXERSOCIAL WHERE'+
         ' PEX_EXERCICE="'+PGAnnee+'"';
QExer:=OpenSQL(StExer,TRUE) ;
if (NOT QExer.EOF) then
   PGExercice := QExer.FindField ('PEX_ANNEEREFER').AsString
else
   begin
   PGExercice := '';
   Ferme(QExer);
   exit;
   end;

if (Arg='B') then
   begin
   if ((Car.Value>='M00') and (Car.Value <='M99')) then
      begin
      TypeD := '005';
      if (Car.Value = 'M01') then
         begin
         DebExer := StrToDate('01/01/'+PGExercice);
         FinExer := StrToDate('31/01/'+PGExercice);
         end
      else
      if (Car.Value = 'M02') then
         begin
         DebExer := StrToDate('01/02/'+PGExercice);
         FinExer := FinDeMois(StrToDate('28/02/'+PGExercice));
         end
      else
      if (Car.Value = 'M03') then
         begin
         DebExer := StrToDate('01/03/'+PGExercice);
         FinExer := StrToDate('31/03/'+PGExercice);
         end
      else
      if (Car.Value = 'M04') then
         begin
         DebExer := StrToDate('01/04/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('30/04/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M05') then
         begin
         DebExer := StrToDate('01/05/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('31/05/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M06') then
         begin
         DebExer := StrToDate('01/06/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('30/06/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M07') then
         begin
         DebExer := StrToDate('01/07/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('31/07/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M08') then
         begin
         DebExer := StrToDate('01/08/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('31/08/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M09') then
         begin
         DebExer := StrToDate('01/09/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('30/09/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M10') then
         begin
         DebExer := StrToDate('01/10/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('31/10/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M11') then
         begin
         DebExer := StrToDate('01/11/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('30/11/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M12') then
         begin
         DebExer := StrToDate('01/12/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('31/12/'+IntToStr(StrToInt(PGExercice)-1));
         end;
      end;
   if ((Car.Value>='T00') and (Car.Value <='T99')) then
      begin
      TypeD := '004';
      if (Car.Value = 'T01') then
         begin
{PT48
         DebExer := StrToDate('01/04/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('30/06/'+IntToStr(StrToInt(PGExercice)-1));
}
         DebExer := StrToDate('01/01/'+PGExercice);
         FinExer := StrToDate('31/03/'+PGExercice);
         end
      else
      if (Car.Value = 'T02') then
         begin
{
         DebExer := StrToDate('01/07/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('30/09/'+IntToStr(StrToInt(PGExercice)-1));
}
         DebExer := StrToDate('01/04/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('30/06/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'T03') then
         begin
{
         DebExer := StrToDate('01/10/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('31/12/'+IntToStr(StrToInt(PGExercice)-1));
}
         DebExer := StrToDate('01/07/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('30/09/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'T04') then
         begin
{
         DebExer := StrToDate('01/01/'+PGExercice);
         FinExer := StrToDate('31/03/'+PGExercice);
}
         DebExer := StrToDate('01/10/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('31/12/'+IntToStr(StrToInt(PGExercice)-1));
//FIN PT48
         end;
      end;
   if ((Car.Value>='S00') and (Car.Value <='S99')) then
      begin
      PGIBox ('Périodicité interdite', Ecran.Caption);
      Car.Value:= 'A00';
      end;
   if Car.Value='A00' then
      begin
      TypeD := '002';
      DebExer := StrToDate('01/04/'+IntToStr(StrToInt(PGExercice)-1));
      FinExer := StrToDate('31/03/'+PGExercice);
      end;
   end
else
   begin
   TypeD := '001';
   if NOT QExer.eof then
      begin
      DebExer := QExer.FindField ('PEX_DATEDEBUT').AsDateTime;
      FinExer := QExer.FindField ('PEX_DATEFIN').AsDateTime;
      end;
   end;
Ferme(QExer);

SetControlText ('DATEDEB', DateTostr(DebExer));
SetControlText ('DATEFIN', DateTostr(FinExer));
end;
//FIN PT39-2


Initialization
  registerclasses ( [ TOF_PGEDIT_DADSPER ] ) ;
end.

