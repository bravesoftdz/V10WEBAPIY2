{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 26/07/2001
Modifié le ... :   /  /
Description .. : Unité de gestion du multicritère DADSU
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
{
PT1   : 16/01/2002 VG V571 Dans le multi-critère, on ne faisait pas apparaître
                           les salariés qui sont sortis mais qui ont quand-même
                           une paye
PT2   : 22/01/2002 VG V571 Ajout de PGExercice
PT3-1 : 27/02/2002 VG V571 Suppression de l'affichage lorsque le calcul a planté
                           Fiche de bug N°442
PT3-2 : 27/02/2002 VG V571 Mauvais affichage du titre de l'écran mul
                           Fiche de bug N°409
PT4   : 19/03/2002 VG V571 Gestion de la DADSU BTP
PT5   : 25/04/2002 VG V571 Affichage de la même liste de salariés quelle que
                           soit l'utilisation : Calcul (standard ou BTP), Saisie
                           des périodes d'activité ou d'inactivité
PT6   : 26/04/2002 VG V571 Par défaut, année précédente jusqu'au mois de
                           septembre . A partir du mois d'octobre, année en
                           cours (pour tests)
PT7   : 31/05/2002 VG V582 Version S3
PT8   : 15/07/2002 VG V585 Rectification du Multi-critère : L'écran "sautait"
                           lorsqu'on modifiait un critère de sélection
PT9   : 03/09/2002 VG V585 Ajout d'un traitement de mise à jour des éléments
                           salarié
PT10  : 21/10/2002 VG V585 Gestion du journal des evenements
PT11-1: 29/01/2003 VG V591 Affichage d'un message si l'exercice par défaut
                           n'existe pas - FQ N°10469
PT11-2: 12/02/2003 VG V_42 FQ N°10469 bis
PT12  : 20/02/2003 VG V_42 Cas du double-click sur une liste vide
PT13  : 14/04/2003 VG V_42 Ajout des dates de début et de fin d'exercice sur le
                           multi-critère pour une meilleure compréhension .
                           FQ N°10624
PT14  : 16/09/2003 VG V_42 Adaptation cahier des charges V7R01
PT15-1: 10/10/2003 VG V_42 Modification du message indiquant une disconcordance
                           entre la DADS-U BTP et le cahier des charges
PT15-2: 10/10/2003 VG V_42 Modification de l'année par défaut pour la DADS-U BTP
PT15-3: 15/10/2003 VG V_42 DADS-U BTP : Saisie des périodes d'inactivités au
                           niveau des périodes d'activité - FQ N°10896
PT16  : 25/02/2004 VG V_50 Ajout d'un message lors du calcul - FQ N°11038
PT17  : 26/02/2004 VG V_50 Traitement de suppression d'une période - FQ N°11074
PT18-1: 31/03/2004 VG V_50 On ne pose plus la question pour la réinitialisation
                           du fichier de log. Le fichier est réinitialisé si le
                           précédent calcul datait de plus de 6 mois.
                           Le fichier est systématiquement ouvert en fin de
                           traitement
PT18-2: 31/03/2004 VG V_50 Ajout d'un bouton sur le multi-critère pour
                           visualiser le fichier de log
PT19  : 23/04/2004 VG V_50 Désactivation des filtres par défaut pour le calcul
                           DADS-U BTP - FQ N°11170
PT20-1: 12/05/2004 VG V_50 Ajout du traitement lié aux raccourcis clavier
PT20-2: 12/05/2004 VG V_50 Optimisation
PT20-3: 12/05/2004 VG V_50 On complète le message lorsque le siret n'est pas
                           valide - FQ N°11152
PT21-1: 05/07/2004 VG V_50 Calcul des périodes d'inactivité
PT21-2: 05/07/2004 VG V_50 Optimisation du traitement
PT21-3: 05/07/2004 VG V_50 Ergonomie
PT21-4: 05/07/2004 VG V_50 Gestion des prud'hommes
PT21-5: 05/07/2004 VG V_50 Adaptation cahier des charges V8R00
PT22  : 27/08/2004 VG V_50 Le message indiquant que le SIRET de la société n'est
                           pas valide devient bloquant dès le calcul
                           FQ N°11550
PT23  : 02/02/2005 VG V_60 Modification du message lors du calcul des éléments
                           de la fiche salarié - FQ N°11956
PT24  : 07/10/2005 VG V_60 Adaptation cahier des charges DADS-U V8R02
PT25  : 03/01/2006 PH V_650 FQ 12281 Multi-période DADS_U
PT26  : 10/01/2006 VG V_65 Affichage des champs de recherche de l'onglet
                           complément
PT27-1: 17/10/2006 VG V_70 Suppression du fichier de contrôle - mise en table
                           des erreurs
PT27-2: 17/10/2006 VG V_70 Gestion du décalage de paie - FQ N°12860
PT27-3: 17/10/2006 VG V_70 Suppression du statut professionnel "ZZZ"
                           FQ N°12996
PT27-4: 17/10/2006 VG V_70 Traitement DADS-U complémentaire
PT27-5: 17/10/2006 VG V_70 Adaptation cahier des charges DADS-U V8R04
PT28  : 06/11/2006 VG V_70 Correction gestion du décalage
PT29  : 07/11/2006 VG V_70 Correction traitement DADS-U complémentaire
                           FQ N°13654
PT30  : 21/11/2006 VG V_70 Permettre de déclarer des honoraires en DADS-U
                           complémentaire - FQ N°13613
PT31  : 23/11/2006 VG V_70 Modification automatique du menu - FQ N°13621
PT32  : 28/12/2006 GG V_80 Message d'erreur en cas de salarié déjà intégré dans
                           une déclaration complémentaire - FQ N°1375
PT33  : 03/01/2007 VG V_72 Impossible de calculer les inactivités de Décembre en
                           paye décalée - FQ N°13806
PT34  : 11/01/2007 VG V_72 Suppression des enregistrements ayant le champ
                           exercicedads mal alimenté - FQ N°13827
PT35  : 05/11/2007 VG V_80 Adaptation cahier des charges V8R06 
PT36  : 28/11/2007 VG V_80 Gestion du champ ET_FICTIF - FQ N°13925
PT37  : 29/11/2007 FC V_80 Gestion des habilitations
PT38  : 03/12/2007 VG V_80 Prise en compte du trimestre civil dans le cas d'une
                           déclaration trimestrielle - FQ N°13245
PT39  : 10/12/2007 VG V_80 Correction PT36
PT40  : 11/12/2007 VG V_80 Gestion du champ PDE_DADSCDC
}
unit UTofPG_MulDADS;

interface
uses UTOF,
     HCtrls,
     HEnt1,
     Hqry,
     HTB97,
     Pgoutils2,
     classes,
     sysutils,
     ed_tools,
     hmsgbox,
     hstatus,
     EntPaie,
     PgDADSOutils,
     PgDADSCommun,
     AGLInit,
     ParamSoc,
     windows,
     controls,
{$IFNDEF DADSUSEULE}
     P5Def,
{$ENDIF}
{$IFNDEF EAGLCLIENT}
     mul,
     FE_Main,
     HDB,
{$IFNDEF DBXPRESS}
     dbTables,
{$ELSE}
     uDbxDataSet,
{$ENDIF}
{$ELSE}
     emul,
     UtileAGL,
     MaineAgl,
     UTob,
{$ENDIF}
     StdCtrls;

Type
     TOF_PGMULDADS= Class (TOF)
       public
       procedure OnArgument(Arguments : String ) ; override ;
       procedure OnUpdate ; override;
       procedure OnLoad ; override ;

       private
       WW : THEdit;     // Clause XX_WHERE
       SALAR : THEdit;
       Annee, Car, Nature : THValComboBox;
       Q_Mul:THQuery ;  // Query pour changer la liste associee
       BCherche, Calculer, Delete : TToolbarButton97;
       Compl : TCheckBox;

{$IFNDEF EAGLCLIENT}
       Liste : THDBGrid;
{$ELSE}
       Liste : THGrid;
{$ENDIF}
       param : String;

       procedure ActiveWhere (Sender: TObject);
       procedure GrilleDblClick (Sender: TObject);
       procedure CalculerClick (Sender: TObject);
       procedure InitCalcul(NomFic : String);
       procedure Calcul_un();
       procedure Calcul_un_inact();
       procedure DeleteClick (Sender: TObject);
       procedure Delete_un();
       procedure SalarieExit (Sender: TObject);
       procedure CarDateChange (Sender: Tobject);
       procedure NatureChange (Sender: Tobject);
       procedure Parametrage();
       procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
       procedure DecalageChange(Sender: TObject);
       procedure DateDecaleChange(Sender: TObject);
       procedure DateDecalageChange(Sender: TObject);
       procedure ComplClick (Sender: TObject);
       procedure MAJTypeD();
     END ;

implementation

//PT8
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 12/07/2002
Modifié le ... :   /  /
Description .. : OnLoad
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADS.OnLoad;
begin
Inherited ;
ActiveWhere(Nil);
end;
//FIN PT8

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 02/10/2001
Modifié le ... :   /  /
Description .. : XX_WHERE
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADS.ActiveWhere(Sender: TObject);
var
St, StDateFin, StEtab : String;
LesEtablis,Etabl,Habilitation,LstEtab :string;
k:integer;
begin
FinDAnnee:= StrToDate ('31/12/'+Copy (DateToStr (FinExer), 7, 4));  //PT35
if (PGExercice='') then
   exit;

{Pour la suppression des périodes, présentation d'une liste basée sur la table
DADSPERIODES}
if (param='U') then
   begin
   if (WW <> NIL) then
{PT35
      WW.Text:= ' PDE_TYPE="'+TypeD+'" AND'+
                ' PDE_DATEDEBUT<="'+UsDateTime(FinExer)+'" AND'+
                ' PDE_DATEFIN>="'+UsDateTime(DebExer)+'" AND'+
                ' PDE_EXERCICEDADS="'+PGExercice+'"';
}
      begin
      if ((Nature.Value>='0100') and (Nature.Value<'0300')) then
         StDateFin:= ' PDE_DATEDEBUT<="'+UsDateTime(FinDAnnee)+'" AND'
      else
         StDateFin:= ' PDE_DATEDEBUT<="'+UsDateTime(FinExer)+'" AND';
      St:= ' PDE_TYPE="'+TypeD+'" AND'+StDateFin+
           ' PDE_EXERCICEDADS="'+PGExercice+'"';
      WW.Text:= St;
      end;
//FIN PT35
   end
else
{Pour les autres menus, présentation d'une liste basée sur la table SALARIES}
   begin
   if (WW <> NIL) then
      begin
      if ((DebExer = 0) or (FinExer = 0)) then
         WW.Text := ''
      else
         begin
{PT35
         if ((Nature.Value>='0400') and (Nature.Value<'0500')) then
            WW.Text := ' ((PSA_DATEENTREE <= "'+UsDateTime(FinExer)+'" AND'+
                       ' (PSA_DATESORTIE >= "'+UsDateTime(DebExer)+'" OR'+
                       ' PSA_DATESORTIE = "'+UsDateTime(IDate1900)+'" OR'+
                       ' PSA_DATESORTIE IS NULL)) OR'+
                       ' (EXISTS(SELECT PPU_SALARIE FROM PAIEENCOURS WHERE'+
                       ' PPU_DATEDEBUT>= "'+UsDateTime(DebExer)+'" AND'+
                       ' PPU_DATEFIN <= "'+UsDateTime(FinExer)+'" AND'+
                       ' PPU_SALARIE=SALARIES.PSA_SALARIE)) OR'+
                       ' (EXISTS(SELECT PHC_SALARIE FROM HISTOCUMSAL WHERE'+
                       ' PHC_DATEDEBUT>= "'+UsDateTime(DebExer)+'" AND'+
                       ' PHC_DATEFIN <= "'+UsDateTime(FinExer)+'" AND'+
                       ' PHC_SALARIE=SALARIES.PSA_SALARIE))) AND'+
                       ' (EXISTS(SELECT PSE_SALARIE FROM DEPORTSAL WHERE'+
                       ' PSE_BTP="X" AND'+
                       ' PSE_SALARIE=SALARIES.PSA_SALARIE))'
         else
            WW.Text := ' (PSA_DATEENTREE <= "'+UsDateTime(FinExer)+'" AND'+
                       ' (PSA_DATESORTIE >= "'+UsDateTime(DebExer)+'" OR'+
                       ' PSA_DATESORTIE <= "'+UsDateTime(IDate1900)+'" OR'+
                       ' PSA_DATESORTIE IS NULL)) OR'+
                       ' (EXISTS(SELECT PPU_SALARIE FROM PAIEENCOURS WHERE'+
                       ' PPU_DATEDEBUT>= "'+UsDateTime(DebExer)+'" AND'+
                       ' PPU_DATEFIN <= "'+UsDateTime(FinExer)+'" AND'+
                       ' PPU_SALARIE=SALARIES.PSA_SALARIE)) OR'+
                       ' (EXISTS(SELECT PHC_SALARIE FROM HISTOCUMSAL WHERE'+
                       ' PHC_DATEDEBUT>= "'+UsDateTime(DebExer)+'" AND'+
                       ' PHC_DATEFIN <= "'+UsDateTime(FinExer)+'" AND'+
                       ' PHC_SALARIE=SALARIES.PSA_SALARIE))';
}
         if ((Nature.Value>='0100') and (Nature.Value<'0300')) then
            StDateFin:= ' PSA_DATEENTREE <= "'+UsDateTime(FinDAnnee)+'" AND'
         else
            StDateFin:= ' PSA_DATEENTREE <= "'+UsDateTime(FinExer)+'" AND';
         St:= ' ('+StDateFin+
              ' (PSA_DATESORTIE >= "'+UsDateTime(DebExer)+'" OR'+
              ' PSA_DATESORTIE <= "'+UsDateTime(IDate1900)+'" OR'+
              ' PSA_DATESORTIE IS NULL)) OR'+
              ' (EXISTS(SELECT PPU_SALARIE FROM PAIEENCOURS WHERE'+
              ' PPU_DATEDEBUT>= "'+UsDateTime(DebExer)+'" AND'+
              ' PPU_DATEFIN <= "'+UsDateTime(FinExer)+'" AND'+
              ' PPU_SALARIE=SALARIES.PSA_SALARIE)) OR'+
              ' (EXISTS(SELECT PHC_SALARIE FROM HISTOCUMSAL WHERE'+
              ' PHC_DATEDEBUT>= "'+UsDateTime(DebExer)+'" AND'+
              ' PHC_DATEFIN <= "'+UsDateTime(FinExer)+'" AND'+
              ' PHC_SALARIE=SALARIES.PSA_SALARIE))';

        //DEB PT37
        LstEtab := '';
        if (Assigned(MonHabilitation)) and (MonHabilitation.LeSQL<>'') then
        begin
          LesEtablis := MonHabilitation.LesEtab;
          Etabl := ReadTokenSt(LesEtablis);
          Habilitation := '';
          k := 0;
          while Etabl <> '' do
          begin
            k := k + 1;
            if Etabl <> '' then
            begin
              if k > 1 then Habilitation := Habilitation + ',';
              Habilitation := Habilitation + '"' + Etabl + '"';
            end;
            Etabl := ReadTokenSt(LesEtablis);
          end;
          if k > 0 then
           LstEtab := ' AND ET_ETABLISSEMENT IN (' + Habilitation + MonHabilitation.SqlPop + ')';
        end;
        //FIN PT37

        //PT36
         StEtab:= 'AND PSA_ETABLISSEMENT IN (SELECT ET_ETABLISSEMENT'+
                  ' FROM ETABLISS WHERE'+
                  ' ET_FICTIF<>"X" ' + LstEtab + ')'; //PT37
         St:= ' ('+St+')'+StEtab;
//FIN PT36

         if ((Nature.Value>='0400') and (Nature.Value<'0500')) then
            WW.Text := ' ('+St+') AND'+
                       ' (EXISTS(SELECT PSE_SALARIE FROM DEPORTSAL WHERE'+
                       ' PSE_BTP="X" AND'+
                       ' PSE_SALARIE=SALARIES.PSA_SALARIE))'
         else
            WW.Text := St;
//FIN PT35
         end;
   end;

{A la fin des calculs, on positionne par défaut le bouton calculer à disable
jusqu'à l'application des critères}
   if ((param = 'C') OR (param = 'E') OR (param = 'B') OR (param = 'J')) then
      begin
      if Calculer <> NIL then
         Calculer.Enabled := False;
      end;
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 02/10/2001
Modifié le ... :   /  /
Description .. : OnArgument
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADS.OnArgument(Arguments: String);
var
AccesGrp, arg, AnneeE, AnneePrec, ComboExer, EtabSelect, MoisE : String;
StDelete, StMenu, StPlus : string;
TT : TFMul;
JourJ : TDateTime;
AnneeA, Jour, MoisM : Word;
Decal : THValComboBox;
Num: Integer;
ModifDateDecale : TToolbarButton97;
DDecalage : THEdit;
Exist : boolean;
QMenu : TQuery;
begin
inherited;
arg:= Arguments;
param:= Trim (ReadTokenPipe (arg, ';'));
// DEB PT25
{PT27-2
if ((param = 'J') OR (param ='I') OR (param ='U')) then
}
{PT33
if (param = 'J') then
   begin
   SetControlVisible ('DECALAGE', FALSE);
   SetControlVisible ('TDECALAGE', FALSE);
   end;
}   
// FIN PT25

TT:= TFMul(Ecran);

TFMul(Ecran).OnKeyDown:= FormKeyDown;

TFMul(Ecran).SetDBListe('PGMULSALDADS');

if (param = 'C') then
//PT31
   begin
   TFMul(Ecran).Caption:= 'Calcul de la DADS-U';
   StMenu:= 'SELECT MN_VERSIONDEV'+
            ' FROM MENU WHERE'+
            ' MN_1=42 AND'+
            ' MN_2=6 AND'+
            ' MN_3=2 AND'+
            ' MN_4=0';
   Exist:= ExisteSQL (StMenu);
   if (Exist=TRUE) then
      begin
      StMenu:= 'SELECT MN_ACCESGRP'+
               ' FROM MENU WHERE'+
               ' MN_1=42 AND'+
               ' MN_2=6 AND'+
               ' MN_3=1 AND'+
               ' MN_4=0';
      QMenu:= OpenSQL (StMenu, TRUE);
      if (NOT QMenu.EOF) then
         begin
         AccesGrp:= QMenu.FindField ('MN_ACCESGRP').AsString;
         StMenu:= 'UPDATE MENU SET'+
                  ' MN_ACCESGRP="'+AccesGrp+'",'+
                  ' MN_VERSIONDEV="-" WHERE'+
                  ' MN_1=42 AND'+
                  ' MN_2=6 AND'+
                  ' MN_3=2 AND'+
                  ' MN_4=0';
         ExecuteSQL (StMenu);
         end;
      Ferme (QMenu);

      StMenu:= 'SELECT MN_ACCESGRP'+
               ' FROM MENU WHERE'+
               ' MN_1=41 AND'+
               ' MN_2=8 AND'+
               ' MN_3=2 AND'+
               ' MN_4=1';
      QMenu:= OpenSQL (StMenu, TRUE);
      if (NOT QMenu.EOF) then
         begin
         AccesGrp:= QMenu.FindField ('MN_ACCESGRP').AsString;
         StMenu:= 'UPDATE MENU SET'+
                  ' MN_ACCESGRP="'+AccesGrp+'",'+
                  ' MN_VERSIONDEV="-" WHERE'+
                  ' MN_1=41 AND'+
                  ' MN_2=8 AND'+
                  ' MN_3=2 AND'+
                  ' MN_4=5';
         ExecuteSQL (StMenu);
         end;
      Ferme (QMenu);
      end;
   end
//FIN PT31
else
if (param = 'B') then
   begin
   TFMul(Ecran).Caption:= 'Calcul de la DADS-U pour les caisses de congés du BTP';
   TFMul(Ecran).FiltreDisabled:= True;
   end
else
if (param = 'E') then
   TFMul(Ecran).Caption:= 'Recalcul des éléments de la fiche salarié pour la DADS-U'
else
if (param = 'J') then
   TFMul(Ecran).Caption:= 'Calcul des périodes d''inactivité DADS-U'
else
if (param = 'S') then
   begin
   TFMul(Ecran).Caption:= 'Saisie des périodes d''activité DADS-U';
   SetControlVisible ('BOuvrir', True);
   end
else
if (param = 'I') then
   begin
   TFMul(Ecran).Caption:= 'Saisie des périodes d''inactivité DADS-U';
   SetControlVisible ('BOuvrir', True);
   end
else
if (param = 'U') then
   begin
   TFMul(Ecran).Caption:= 'Suppression des périodes';
   TFMul(Ecran).SetDBListe('PGDADSPERIODE');
   end;

if (TT <> nil) then
   UpdateCaption(TT);

{Calculs}
if ((param = 'C') or (param = 'B') or (param = 'E') or (param = 'J')) then
   begin
   Calculer:= TToolbarButton97 (GetControl ('BCALCULER'));
   if (Calculer <> nil) then
      begin
      Calculer.Visible:= True;
      Calculer.Enabled:= False;
      Calculer.OnClick:= CalculerClick;
      end;
   end;

{Saisies}
if ((param = 'S') or (param = 'I')) then
   begin
   TFMul (Ecran).bSelectAll.Visible:= False;

{$IFNDEF EAGLCLIENT}
   Liste:= THDBGrid (GetControl ('FListe'));
{$ELSE}
   Liste:= THGrid (GetControl ('FListe'));
{$ENDIF}
   if (Liste <> nil) then
      begin
{$IFNDEF EAGLCLIENT}
      Liste.MultiSelection:= False;
{$ENDIF}
      Liste.OnDblClick:= GrilleDblClick;
      end;
   end;

//Cas où périodes absentes, MAJ fiche salarié
if ((param = 'S') or (param = 'C') or (param = 'B') or (param = 'E') or
   (param = 'J')) then
   begin
   StDelete:= 'UPDATE SALARIES SET'+
              ' PSA_DADSDATE="'+UsDateTime (IDate1900)+'" WHERE'+
              ' PSA_DADSDATE <>"'+UsDateTime (IDate1900)+'" AND'+
              ' PSA_SALARIE NOT IN(SELECT PDE_SALARIE FROM DADSPERIODES)';
   ExecuteSQL (StDelete);
   end;

//PT27-3
StDelete:= 'UPDATE SALARIES SET PSA_DADSPROF="01" WHERE PSA_DADSPROF="ZZZ"';
ExecuteSQL (StDelete);
//FIN PT27-3

if (param = 'U') then
   begin
   Delete:= TToolbarButton97 (GetControl ('BDELETE'));
   if (Delete <> nil) then
      begin
      Delete.Visible:= True;
      Delete.Enabled:= True;
      Delete.OnClick:= DeleteClick;
      end;
   MajExercice;        //PT34
   end;

{PT27-1
Visualiser := TToolbarButton97(GetControl('BVISUALISER'));
if Visualiser <> nil then
   Visualiser.OnClick := VisuClick;
}

Q_Mul:= THQuery (Ecran.FindComponent ('Q'));

SALAR:= THEdit (GetControl ('PSA_SALARIE'));
SetControlText ('PSA_DADSFRACTION', '1');

Annee:= THValComboBox (GetControl ('ANNEE'));
Nature:= THValComboBox (GetControl ('CNATURE'));
Car:= THValComboBox (GetControl ('CCAR'));
WW:= THEdit (GetControl ('XX_WHERE'));

if SALAR <> nil then
   SALAR.OnExit:= SalarieExit;

JourJ:= Date;
DecodeDate (JourJ, AnneeA, MoisM, Jour);
if (param = 'B') then
   AnneePrec:= IntToStr (AnneeA)
else
   begin
   if MoisM > 9 then
      AnneePrec:= IntToStr (AnneeA)
   else
      AnneePrec:= IntToStr (AnneeA-1);
   end;

if RendExerSocialPrec (MoisE, AnneeE, ComboExer, DebExer, FinExer, AnneePrec) = TRUE then
   begin
   if Annee <> nil then
      begin
      Annee.value:= ComboExer;
      PGExercice := AnneeE;
      end;
   end
else
   begin
   PGIBox ('L''exercice '+AnneePrec+' n''existe pas', TFMul (Ecran).Caption);
   end;

if Annee <> nil then
   begin
   PGAnnee:= Annee.value;
   Annee.OnChange:= CarDateChange;
   end;

if ((param = 'B') and (PGExercice <> '')) then
   begin
   DebExer:= StrToDate ('01/04/'+IntToStr (StrToInt (PGExercice)-1));
   FinExer:= StrToDate ('31/03/'+PGExercice);
   end;

SetControlText ('L_DDU', DateToStr (DebExer));
SetControlText ('L_DAU', DateToStr (FinExer));

SetControlText ('PSA_DADSDATE', DateToStr (IDate1900));
// DEB PT25
{PT27-2
SetControlText('DD1', DateToStr(DebExer));
SetControlText('DD2', DateToStr(FinExer));
if VH_Paie.PGDecalage then
   SetControlText ('DECALAGE', '03')
else
   SetControlText ('DECALAGE', '01');
LaDate := THEdit (GetControl ('DD1'));
if LaDate <> NIL then
   LaDate.OnExit := DD1Exit;
LaDate := THEdit (GetControl ('DD2'));
if LaDate <> NIL then
   LaDate.OnExit := DD2Exit;
}
{PT33
if (param = 'J') then
   SetControlText ('DECALAGE', '01')
else
   begin
}
if (VH_Paie.PGDecalage) then
   SetControlText ('DECALAGE', '03')
else
   SetControlText ('DECALAGE', '01');
{PT33
   end;
}
//FIN PT27-2
Decal:= THValComboBox (GetControl ('DECALAGE'));
if Decal <> nil then
   Decal.OnChange:= DecalageChange;
// FIN PT25

if Nature <> nil then
   begin
   if (param = 'B') then
      begin
      Nature.OnChange:= NatureChange;
      Nature.Value:= '0451';
      Nature.Enabled:= FALSE;
      end
   else
      begin
      Nature.OnChange:= NatureChange;
      Nature.Value:= '0151';
      end;
   end;

if (param = 'B') then
   SetControlEnabled ('LNATURE', False)
else
   SetControlEnabled ('LNATURE', True);

if Car <> nil then
   begin
   if (param = 'B') then
      Car.Enabled:= True
   else
      begin
      Car.Enabled:= False;
      Car.Visible:= False;
      end;
   Car.OnChange:= CarDateChange;
   Car.Value:= 'A00';
   end;

if (param = 'B') then
   SetControlVisible ('LCAR', True)
else
   SetControlVisible ('LCAR', False);

BCherche:= TToolbarButton97 (GetControl ('BCherche'));
ActiveWhere (nil);

//PT26
{$IFNDEF DADSUSEULE}
for Num := 1 to VH_Paie.PGNbreStatOrg do
    begin
    if Num > 4 then
       Break;
    VisibiliteChampSalarie (IntToStr (Num),
                            GetControl ('PSA_TRAVAILN'+IntToStr (Num)),
                            GetControl ('TPSA_TRAVAILN'+IntToStr (Num)));
    end;
VisibiliteStat (GetControl ('PSA_CODESTAT'), GetControl ('TPSA_CODESTAT'));
{$ENDIF}
//FIN PT26

//PT27-2
// Gestion du bouton de modification des dates
ModifDateDecale:= TToolbarButton97 (GetControl ('BMODIFDATE'));
if ModifDateDecale<>nil then
   ModifDateDecale.OnClick:= DateDecaleChange;

DDecalage:= THEDIT (GetControl ('DATEDECALAGE'));
if DDecalage <> nil then
{PT28
   Decalage.OnChange:= DateDecalageChange;
}
   DDecalage.OnExit:= DateDecalageChange;
//FIN PT28
//FIN PT27-2
//PT29
Compl:= TCheckBox (GetControl ('CHCOMPL'));
if (Compl <> nil) then
   Compl.OnClick:= ComplClick;
//FIN PT29

//PT39
EtabSelect:= 'SELECT ET_ETABLISSEMENT'+
             ' FROM ETABLISS WHERE'+
             ' ET_FICTIF IS NULL';
if (ExisteSQL (EtabSelect)) then
   begin
   StPlus:= 'UPDATE ETABLISS SET ET_FICTIF="-" WHERE ET_FICTIF IS NULL';
   ExecuteSQL (StPlus);
   end;
//FIN PT39

//PT36
StPlus:= ' WHERE ET_FICTIF<>"X"';
SetControlProperty ('PSA_ETABLISSEMENT', 'Plus', StPlus);
//FIN PT36
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 19/10/2001
Modifié le ... :   /  /
Description .. : OnUpdate
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADS.OnUpdate;
begin
inherited;
{PT40
Parametrage;
}
{PT4
if (param = 'C') then
}
{PT9
if ((param = 'C') OR (param = 'B')) then
}
{PT21-1
if ((param = 'C') OR (param = 'B') OR (param = 'E')) then
}
if ((param = 'C') or (param = 'B') or (param = 'E') or (param = 'J')) then
   begin
   if Calculer <> nil then
      Calculer.Enabled := True;
   end;

//PT11-2
if (PGAnnee = '') then
   begin
   if Calculer <> nil then
      Calculer.Enabled := False;
   end;
//FIN PT11-2
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 02/10/2001
Modifié le ... :   /  /
Description .. : Double-click sur la grille
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADS.GrilleDblClick(Sender: TObject);
var
Salarie, StExiste, StSal, TypeDe : String;
reponse : integer;
Existe : boolean;
QCdc : TQuery;
begin
if ((Q_Mul <> nil) and (Q_Mul.RecordCount = 0)) then
   exit;

{$IFDEF EAGLCLIENT}
TFMul (Ecran).Q.TQ.Seek (TFMul (Ecran).FListe.Row-1) ;
{$ENDIF}

//PT27-4
Salarie:= Q_Mul.FindField ('PSA_SALARIE').AsString;
if (Copy (TypeD, 1, 1)='2') then
   begin
   StExiste:= 'SELECT PDS_SALARIE'+
              ' FROM DADSDETAIL WHERE'+
              ' PDS_SALARIE="'+Salarie+'" AND'+
              ' PDS_TYPE="0'+Copy (TypeD, 2, 2)+'" AND'+
              ' PDS_EXERCICEDADS = "'+PGExercice+'"';
   Existe:= ExisteSQL (StExiste);
   end
else
   begin
   StExiste:= 'SELECT PDS_SALARIE'+
              ' FROM DADSDETAIL WHERE'+
              ' PDS_SALARIE="'+Salarie+'" AND'+
              ' PDS_TYPE="2'+Copy (TypeD, 2, 2)+'" AND'+
              ' PDS_EXERCICEDADS = "'+PGExercice+'"';
   Existe:= ExisteSQL (StExiste);
   end;

if (Existe=False) then
   begin
//FIN PT27-4
   if (param = 'S') then
      begin
      StSal:= 'SELECT PDE_SALARIE'+
              ' FROM DADSPERIODES WHERE'+
              ' PDE_SALARIE="'+Salarie+'" AND'+
              ' PDE_TYPE="'+TypeD+'" AND'+
              ' PDE_ORDRE > 0 AND'+
              ' PDE_EXERCICEDADS = "'+PGExercice+'"';
{$IFNDEF EAGLCLIENT}
      TheMulQ:= THQuery (Ecran.FindComponent ('Q'));
{$ELSE}
      TheMulQ:= TOB (Ecran.FindComponent ('Q'));
{$ENDIF}
      if PGAnnee <> '' then
         begin
         reponse:= mrYes;
{PT27-4 + PT27-5
         if (TypeD<>'001') then
            begin
            if (PGExercice <> '2006') then
               reponse:=PGIAsk('Attention ! Le cahier des charges utilisé#13#10'+
                               'est à la norme DADS-U 2005.#13#10'+
                               'Voulez-vous continuer ?', TFMul(Ecran).Caption);
            end
         else
            begin
            if (PGExercice <> '2005') then
               reponse:=PGIAsk('Attention ! Le cahier des charges utilisé#13#10'+
                               'est à la norme DADS-U 2005.#13#10'+
                               'Voulez-vous continuer ?', TFMul(Ecran).Caption);
            end;
}
{PT35
         if ((TypeD<>'001') and (TypeD<>'201')) then
            begin
            if (PGExercice <> '2007') then
               reponse:= PGIAsk ('Attention ! Le cahier des charges#13#10'+
                                 'utilisé est à la norme DADS-U 2006.#13#10'+
                                 'Voulez-vous continuer ?',
                                 TFMul (Ecran).Caption);
            end
         else
            begin
            if (PGExercice <> '2006') then
               reponse:= PGIAsk ('Attention ! Le cahier des charges#13#10'+
                                 'utilisé est à la norme DADS-U 2006.#13#10'+
                                 'Voulez-vous continuer ?',
                                 TFMul (Ecran).Caption);
            end;
}
         if ((TypeD<>'001') and (TypeD<>'201')) then
            begin
            if (PGExercice <> '2008') then
               reponse:= PGIAsk ('Attention ! Le cahier des charges#13#10'+
                                 'utilisé est à la norme DADS-U 2007.#13#10'+
                                 'Voulez-vous continuer ?',
                                 TFMul (Ecran).Caption);
            end
         else
            begin
            if (PGExercice <> '2007') then
               reponse:= PGIAsk ('Attention ! Le cahier des charges#13#10'+
                                 'utilisé est à la norme DADS-U 2007.#13#10'+
                                 'Voulez-vous continuer ?',
                                 TFMul (Ecran).Caption);
            end;
//FIN PT35
//FIN PT27-4 + PT27-5
         if reponse = mrYes then
            begin
            if ((ExisteSQL (StSal)=FALSE) and (PGAnnee <> '')) then
               AGLLanceFiche ('PAY','DADSU',  '',Salarie+';'+TypeD+';1',
                              'CREATION;'+Salarie+';'+TypeD+';'+GetControlText ('PSA_DADSFRACTION'))
            else
{PT40
               AGLLanceFiche ('PAY','DADSU',  '',Salarie+';'+TypeD+';1',
                              'MODIFICATION;'+Salarie+';'+TypeD+';'+GetControlText ('PSA_DADSFRACTION'));
}
               begin
               StSal:= 'SELECT PDE_DADSCDC'+
                       ' FROM DADSPERIODES WHERE'+
                       ' PDE_SALARIE="'+Salarie+'" AND'+
                       ' PDE_TYPE="'+TypeD+'" AND'+
                       ' PDE_ORDRE=1 AND'+
                       ' PDE_EXERCICEDADS="'+PGExercice+'"';
               QCdc:= OpenSQL (StSal, True);
               if (QCdc.FindField ('PDE_DADSCDC').AsString<>CDC) then
                  PGIBox ('Saisie impossible ! la période n''a pas été#13#10'+
                          ' calculée avec la version '+CDC+' du cahier#13#10'+
                          ' des charges.', TFMul (Ecran).Caption)
               else
                  AGLLanceFiche ('PAY','DADSU',  '',Salarie+';'+TypeD+';1',
                                 'MODIFICATION;'+Salarie+';'+TypeD+';'+GetControlText ('PSA_DADSFRACTION'));
               Ferme (QCdc);
               end;
//FIN PT40
            end;
         end
      else
         PGIBox ('L''année n''est pas valide', TFMul (Ecran).Caption);
      end
   else
      begin
      Salarie:= Q_Mul.FindField ('PSA_SALARIE').AsString;
      StSal:= 'SELECT PDE_SALARIE'+
              ' FROM DADSPERIODES WHERE'+
              ' PDE_SALARIE="'+Salarie+'" AND'+
              ' PDE_TYPE="'+TypeD+'" AND'+
              ' PDE_ORDRE < 0 AND'+
              ' PDE_EXERCICEDADS = "'+PGExercice+'"';

{$IFNDEF EAGLCLIENT}
      TheMulQ:= THQuery (Ecran.FindComponent ('Q'));
{$ELSE}
      TheMulQ:= TOB (Ecran.FindComponent ('Q'));
{$ENDIF}
      if (PGAnnee <> '') then
         begin
         if ((ExisteSQL (StSal)=FALSE) and (PGAnnee <> '')) then
            begin
            StSal:= 'SELECT PDS_SALARIE'+
                    ' FROM DADSDETAIL WHERE'+
                    ' PDS_SALARIE="'+Salarie+'" AND'+
                    ' PDS_TYPE="'+TypeD+'" AND'+
                    ' PDS_ORDRE = 0 AND'+
                    ' PDS_EXERCICEDADS = "'+PGExercice+'"';
            if (ExisteSQL (StSal)=TRUE) then
               AGLLanceFiche ('PAY','DADS_INACTIVITE',  '',Salarie+';'+TypeD,
                              'CREATION;'+Salarie+';'+TypeD)
            else
               PGIBox ('Traitement impossible : Le salarié '+Salarie+'#13#10'+
                       ' n''a jamais été calculé', TFMul (Ecran).Caption);
            end
         else
{PT40
            AGLLanceFiche ('PAY','DADS_INACTIVITE',  '',Salarie+';'+TypeD,
                           'MODIFICATION;'+Salarie+';'+TypeD);
}
            begin
            StSal:= 'SELECT PDE_DADSCDC'+
                    ' FROM DADSPERIODES WHERE'+
                    ' PDE_SALARIE="'+Salarie+'" AND'+
                    ' PDE_TYPE="'+TypeD+'" AND'+
                    ' PDE_ORDRE=-1 AND'+
                    ' PDE_EXERCICEDADS="'+PGExercice+'"';
            QCdc:= OpenSQL (StSal, True);
            if (QCdc.FindField ('PDE_DADSCDC').AsString<>CDC) then
               PGIBox ('Saisie impossible ! la période n''a pas été#13#10'+
                       ' calculée avec la version '+CDC+' du cahier#13#10'+
                       ' des charges.', TFMul (Ecran).Caption)
            else
               AGLLanceFiche ('PAY','DADS_INACTIVITE',  '',Salarie+';'+TypeD,
                              'MODIFICATION;'+Salarie+';'+TypeD);
            Ferme (QCdc);
            end;
//FIN PT40
         end
      else
         PGIBox ('L''année n''est pas valide', TFMul (Ecran).Caption);
      end;
//PT27-4
   end
else
   begin
   if (Copy (TypeD, 1, 1)='2') then
      TypeDe:= 'normale'
   else
      TypeDe:= 'complémentaire';
   PGIBox ('Traitement impossible : Le salarié '+Salarie+' a été#13#10'+
           'intégré dans une déclaration '+TypeDe, TFMul (Ecran).Caption);
   end;
//FIN PT27-4
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 08/06/2001
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "Calculer"
Suite ........ : ou "Valider" (pour l'instant)
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PGMULDADS.CalculerClick(Sender: TObject);
var
BufDest, NomFic : String;
i, reponse : integer;
Maintenant : TDateTime;
begin
{$IFNDEF EAGLCLIENT}
Liste:= THDBGrid(GetControl('FListe'));
{$ELSE}
Liste:= THGrid(GetControl('FListe'));
{$ENDIF}

reponse:= mrYes;
if ((VH_Paie.PGPCS2003 = FALSE) and (PGExercice = '2003')) then
   reponse :=PGIAsk('Vous ne devez pas calculer la DADS-U 2003 sans#13#10'+
                    'avoir réaffecté les codes PCS.#13#10'+
                    'Pour cela, choisissez le menu#13#10'+
                    '"Paramètres\Calcul de la Paie\Gestion des codes PCS"#13#10'+
                    'Voulez-vous continuer ?', TFMul(Ecran).Caption);
if (reponse <> mrYes) then
   exit;

if param = 'E' then
   reponse:= PGIAsk ('Attention, ce traitement va remettre à jour les#13#10'+
                     'données issues de la fiche salarié.#13#10'+
                     'Voulez-vous continuer ?', TFMul(Ecran).Caption)
else
   reponse:= PGIAsk ('Cette commande supprimera les modifications#13#10'+
                     'manuelles effectuées en saisie des périodes pour#13#10'+
                     'l''exercice.#13#10'+
                     'Voulez-vous continuer ?', TFMul(Ecran).Caption);
if (reponse <> mrYes) then
   exit;

if Liste <> nil then
   begin
   if (Liste.NbSelected=0) and (not Liste.AllSelected) then
      begin
      MessageAlerte('Aucun élément sélectionné');
      exit;
      end;

{PT27-5
   if (param = 'B') then
      begin
      if (PGExercice <> '2006') then
         reponse:= PGIAsk('Attention ! Le cahier des charges utilisé#13#10'+
                          'est à la norme DADS-U 2005.#13#10'+
                          'Voulez-vous continuer ?', TFMul(Ecran).Caption);
      end
   else
      begin
      if (PGExercice <> '2005') then
        reponse:= PGIAsk('Attention ! Le cahier des charges utilisé#13#10'+
                         'est à la norme DADS-U 2005.#13#10'+
                         'Voulez-vous continuer ?', TFMul(Ecran).Caption);
      end;
}
{PT35
   if (param = 'B') then
      begin
      if (PGExercice <> '2007') then
         reponse:= PGIAsk('Attention ! Le cahier des charges utilisé#13#10'+
                          'est à la norme DADS-U 2006.#13#10'+
                          'Voulez-vous continuer ?', TFMul(Ecran).Caption);
      end
   else
      begin
      if (PGExercice <> '2006') then
        reponse:= PGIAsk('Attention ! Le cahier des charges utilisé#13#10'+
                         'est à la norme DADS-U 2006.#13#10'+
                         'Voulez-vous continuer ?', TFMul(Ecran).Caption);
      end;
}
   if (param = 'B') then
      begin
      if (PGExercice <> '2008') then
         reponse:= PGIAsk('Attention ! Le cahier des charges utilisé#13#10'+
                          'est à la norme DADS-U 2007.#13#10'+
                          'Voulez-vous continuer ?', TFMul(Ecran).Caption);
      end
   else
      begin
      if (PGExercice <> '2007') then
        reponse:= PGIAsk('Attention ! Le cahier des charges utilisé#13#10'+
                         'est à la norme DADS-U 2007.#13#10'+
                         'Voulez-vous continuer ?', TFMul(Ecran).Caption);
      end;
//FIN PT35
//FIN PT27-5
   if (reponse <> mrYes) then
      exit;

   ForceNumerique(GetParamSoc('SO_SIRET'), BufDest);
   if ControlSiret(BufDest)=False then
      begin
      PGIBox('Le SIRET de la société n''est pas valide.#13#10'+
             'Vous devez le vérifier en y accédant par le module#13#10'+
             '"Paramètres/menu Comptabilité/commande Paramètres comptables/Coordonnées".#13#10'+
             'Si vous travaillez en environnement multi-dossiers,#13#10'+
             'vous pouvez y accéder par le Bureau PGI/Annuaire.',
             TFMul(Ecran).Caption);
      if (Liste.AllSelected=TRUE) then
         begin
         Liste.AllSelected:= False;
         TFMul(Ecran).bSelectAll.Down:= Liste.AllSelected;
         end
      else
         Liste.ClearSelected;
      exit;
      end;

{$IFDEF EAGLCLIENT}
   NomFic:= VH_Paie.PgCheminEagl+'\'+BufDest+'_DADSU_PGI.log';
{$ELSE}
   NomFic:= V_PGI.DatPath+'\'+BufDest+'_DADSU_PGI.log';
{$ENDIF}
   Trace:= TStringList.Create;
   if (Liste.AllSelected=TRUE) then
      begin
      InitMoveProgressForm (NIL,'Calcul en cours', 'Veuillez patienter SVP ...',
                            TFmul(Ecran).Q.RecordCount,FALSE,TRUE);
      InitMove(TFmul(Ecran).Q.RecordCount, '');
      InitCalcul(NomFic);
      ChargeTOBCommun;
{$IFDEF EAGLCLIENT}
      if (TFMul(Ecran).bSelectAll.Down) then
         TFMul(Ecran).Fetchlestous;
{$ENDIF}
      TFmul(Ecran).Q.First;
      while not TFmul(Ecran).Q.EOF do
            begin
            if (param='J') then
               Calcul_un_inact
            else
               Calcul_un;
            TFmul(Ecran).Q.Next;
            end;

      LibereTOBCommun;
      Liste.AllSelected:= False;
      TFMul(Ecran).bSelectAll.Down:= Liste.AllSelected;
      end
   else
      begin
      InitMoveProgressForm (NIL,'Calcul en cours', 'Veuillez patienter SVP ...',
                            Liste.NbSelected,FALSE,TRUE);
      InitMove(Liste.NbSelected, '');
      InitCalcul(NomFic);
      ChargeTOBCommun;

      for i:=0 to Liste.NbSelected-1 do
          begin
          Liste.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
          TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row-1) ;
{$ENDIF}
          if (param='J') then
             Calcul_un_inact
          else
             Calcul_un;
          end;

      LibereTOBCommun;
      Liste.ClearSelected;
      end;

   FiniMove;
   FiniMoveProgressForm;
   PGIBox ('Traitement terminé. Vous devez maintenant procéder au#13#10'+
           'contrôle DADS-U.', TFMul (Ecran).Caption);
   Maintenant:= Now;
{PT27-1
   Writeln(FRapport, 'Calcul de la DADS-U terminé : ' + DateTimeToStr(Maintenant));
}
   if Trace <> nil then
      Trace.Add ('Calcul de la DADS-U terminé à '+TimeToStr (Maintenant));
{PT27-1
   CloseFile(FRapport);
   ShellExecute(0, PCHAR('open'), PChar('WordPad'), PChar(NomFic), nil, SW_RESTORE);
}
{$IFNDEF DADSUSEULE}
   CreeJnalEvt('001', '040', 'OK', nil, nil, Trace);
{$ENDIF}   
   FreeAndNil (Trace);
   end;
if Calculer <> nil then
   Calculer.Enabled := False;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/05/2004
Modifié le ... :   /  /
Description .. : Initialisation du calcul de la DADSU
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADS.InitCalcul(NomFic: string);
var
Maintenant : TDateTime;
begin
{PT27-1
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

   if (PlusMois(Maintenant, -6)>DateCalcul) then
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
   ReWrite(FRapport);
   Writeln(FRapport, 'Attention, Le dernier calcul se trouve en fin du fichier');
   end;

Writeln(FRapport, '_____________________________________');
Writeln(FRapport, TFMul(Ecran).Caption);
Maintenant := Now;
Writeln(FRapport, 'Début de calcul : '+DateTimeToStr(Maintenant));
Writeln(FRapport, '');
Writeln(FRapport, 'Paramètres : Nature '+
        RechDom('PGDADSNATURE', Nature.Value, False));
if (Car.Visible = True) then
   Writeln(FRapport, '             Caractéristiques '+
           RechDom('PGCARDADS', Car.Value, False));
Writeln(FRapport, '             '+
        RechDom('PGCODESECTION', GetControlText('PSA_DADSFRACTION'), False));
Writeln(FRapport, '             Période du '+GetControlText('L_DDU')+
        ' au '+GetControlText('L_DAU'));
}
Maintenant:= Now;
//FIN PT27-1
if Trace <> nil then
   Trace.Add ('Début du calcul à '+TimeToStr (Maintenant));
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/05/2004
Modifié le ... :   /  /
Description .. : Calcul d'un élément
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADS.Calcul_un();
var
St, StExiste : String;
TypeDe : String;  //PT32
Existe : boolean;
Maintenant : TDateTime;
begin
St:= TFmul (Ecran).Q.FindField('PSA_SALARIE').asstring;
if St <> '' then
   begin
   if (isnumeric (St) and (VH_PAIE.PgTypeNumSal = 'NUM')) then
      St:= ColleZeroDevant (StrToInt (St), 10);
   try
      begintrans;
      if Trace <> nil then
         Trace.Add ('Traitement du salarié '+St);
//PT27-4
      if (Copy (TypeD, 1, 1)='2') then
         begin
         StExiste:= 'SELECT PDS_SALARIE'+
                    ' FROM DADSDETAIL WHERE'+
                    ' PDS_SALARIE="'+St+'" AND'+
                    ' PDS_TYPE="0'+Copy (TypeD, 2, 2)+'" AND'+
                    ' PDS_EXERCICEDADS = "'+PGExercice+'"';
         Existe:= ExisteSQL (StExiste);
         end
      else
         begin
         StExiste:= 'SELECT PDS_SALARIE'+
                    ' FROM DADSDETAIL WHERE'+
                    ' PDS_SALARIE="'+St+'" AND'+
                    ' PDS_TYPE="2'+Copy (TypeD, 2, 2)+'" AND'+
                    ' PDS_EXERCICEDADS = "'+PGExercice+'"';
         Existe:= ExisteSQL (StExiste);
         end;
      if (Existe=False) then
         begin
         Existe:= TRUE;
//FIN PT27-4
         if (param <> 'E') then
            begin
{PT30
            DeletePeriode (St, TypeD, -100);
            DeleteDetail (St, TypeD, -100);
}
            DeletePeriode (St, -100);
            DeleteDetail (St, -100);
//FIN PT30
            end
         else
            begin
            StExiste:= 'SELECT PDS_SALARIE'+
                       ' FROM DADSDETAIL WHERE'+
                       ' PDS_SALARIE="'+St+'" AND'+
                       ' PDS_TYPE="'+TypeD+'" AND'+
                       ' PDS_ORDRE=0 AND'+
                       ' PDS_EXERCICEDADS = "'+PGExercice+'"';
            Existe:= ExisteSQL (StExiste);
            end;
         if (Existe = TRUE) then
            begin
{PT30
            DeleteDetail (St, TypeD, 0);
}
            DeleteDetail (St, 0);
            ChargeTOBSal (St);
//PT27-1
            if (param = 'E') then
               DeleteErreur (St, 'S3')
            else
               DeleteErreur (St, 'S');
            DeleteErreur ('', 'SKO');
//FIN PT27-1
{PT30
            CalculElemSal (St, TypeD);
}
            CalculElemSal (St);
            if (param <> 'E') then
{PT30
               CalculSal (St, TypeD);
}
               CalculSal (St);
            EcrireErreurKO;	//PT27-1
            LibereTOB;
            if (param <> 'E') then
               ExecuteSQL ('UPDATE SALARIES SET PSA_DADSDATE='+
                           '"'+UsDateTime (Date)+'" WHERE'+
                           ' PSA_SALARIE="'+St+'"');
            end
         else
            PGIBox ('Traitement impossible : Le salarié '+St+' n''a#13#10'+
                    'jamais été calculé', TFMul (Ecran).Caption);
{PT27-1
            Writeln (FRapport, 'Calcul de la DADS-U impossible : Le salarié '+St+
                     ' n''a jamais été calculé');
}
         end
      else
{PT32
         PGIBox ('Traitement impossible : Le salarié '+St+' a été#13#10'+
                 'intégré dans une déclaration normale', TFMul (Ecran).Caption);
}
         begin
           if (Copy (TypeD, 1, 1)='2') then
              TypeDe:= 'normale'
           else
              TypeDe:= 'complémentaire';
           PGIBox ('Traitement impossible : Le salarié '+St+' a été#13#10'+
                   'intégré dans une déclaration '+TypeDe, TFMul (Ecran).Caption);
         end;
//Fin PT32
      CommitTrans;
   except
      Rollback;
      Maintenant:= Now;
{PT27-1
      Writeln (FRapport, 'Salarié '+St+' : Calcul de la DADS-U annulé : '+
               DateTimeToStr(Maintenant));
}
      if Trace <> nil then
         Trace.Add ('Salarié '+St+' : Calcul de la DADS-U annulé à '+
                    TimeToStr (Maintenant));
      end;
   end;
MoveCur (False);
MoveCurProgressForm (St);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/06/2004
Modifié le ... :   /  /
Description .. : Calcul d'un élément d'inactivité
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADS.Calcul_un_inact();
var
St, StExiste : String;
TypeDe : String;  //PT32
Existe : boolean;
Maintenant : TDateTime;
begin
St:= TFmul (Ecran).Q.FindField ('PSA_SALARIE').asstring;
if St <> '' then
   begin
   if (isnumeric(St) and (VH_PAIE.PgTypeNumSal = 'NUM')) then
      St:= ColleZeroDevant (StrToInt (St), 10);
   try
      begintrans;
      if Trace <> nil then
         Trace.Add ('Traitement du salarié '+St);

//PT27-4
      if (Copy (TypeD, 1, 1)='2') then
         begin
         StExiste:= 'SELECT PDS_SALARIE'+
                    ' FROM DADSDETAIL WHERE'+
                    ' PDS_SALARIE="'+St+'" AND'+
                    ' PDS_TYPE="0'+Copy (TypeD, 2, 2)+'" AND'+
                    ' PDS_EXERCICEDADS = "'+PGExercice+'"';
         Existe:= ExisteSQL (StExiste);
         end
      else
         begin
         StExiste:= 'SELECT PDS_SALARIE'+
                    ' FROM DADSDETAIL WHERE'+
                    ' PDS_SALARIE="'+St+'" AND'+
                    ' PDS_TYPE="2'+Copy (TypeD, 2, 2)+'" AND'+
                    ' PDS_EXERCICEDADS = "'+PGExercice+'"';
         Existe:= ExisteSQL (StExiste);
         end;
      if (Existe=False) then
         begin
//FIN PT27-4
{PT30
         DeletePeriode (St, TypeD, -200);
         DeleteDetail (St, TypeD, -200);
}
         DeletePeriode (St, -200);
         DeleteDetail (St, -200);
//FIN PT30
         StExiste:= 'SELECT PDS_SALARIE'+
                    ' FROM DADSDETAIL WHERE'+
                    ' PDS_SALARIE="'+St+'" AND'+
                    ' PDS_TYPE="'+TypeD+'" AND'+
                    ' PDS_ORDRE=0 AND'+
                    ' PDS_EXERCICEDADS = "'+PGExercice+'"';
         Existe:= ExisteSQL (StExiste);
         if (Existe = TRUE) then
            begin
            ChargeTOBInact (St);
            DeleteErreur ('', 'SKO');	//PT27-1
{PT30
            RemplitInact (St, TypeD);
}
            RemplitInact (St);
            EcrireErreurKO;	//PT27-1
            LibereTOBInact;
            end
         else
{PT27-1
            Writeln(FRapport, 'Calcul de la DADS-U impossible : Le salarié '+St+' n''a jamais été calculé');
}
            PGIBox ('Traitement impossible : Le salarié '+St+' n''a#13#10'+
                    'jamais été calculé', TFMul (Ecran).Caption);
//FIN PT27-1
         end
      else
{PT32
         PGIBox ('Traitement impossible : Le salarié '+St+' a été#13#10'+
                 'intégré dans une déclaration normale', TFMul (Ecran).Caption);
}
         begin
           if (Copy (TypeD, 1, 1)='2') then
              TypeDe:= 'normale'
           else
              TypeDe:= 'complémentaire';
           PGIBox ('Traitement impossible : Le salarié '+St+' a été#13#10'+
                   'intégré dans une déclaration '+TypeDe, TFMul (Ecran).Caption);
         end;
//Fin PT32
      CommitTrans;
   except
      Rollback;
      Maintenant:= Now;
{PT27-1
      Writeln(FRapport, 'Salarié '+St+' : Calcul de la DADS-U annulé : '+DateTimeToStr(Maintenant));
}
      if Trace <> nil then
         Trace.Add ('Salarié '+St+' : Calcul de la DADS-U annulé à '+
                    TimeToStr (Maintenant));
      end;
   end;
MoveCur (False);
MoveCurProgressForm (St);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 25/02/2004
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "Delete"
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PGMULDADS.DeleteClick(Sender: TObject);
Var
i, reponse : integer;
Maintenant : TDateTime;
begin
{$IFNDEF EAGLCLIENT}
Liste := THDBGrid(GetControl('FListe'));
{$ELSE}
Liste := THGrid(GetControl('FListe'));
{$ENDIF}

reponse:= PGIAsk ('Cette commande supprimera les périodes sélectionnées.#13#10'+
                  'Voulez-vous continuer ?', TFMul(Ecran).Caption);
if (reponse <> mrYes) then
   exit;

if Liste <> nil then
   begin
   if (Liste.NbSelected = 0) and (not Liste.AllSelected) then
      begin
      MessageAlerte('Aucun élément sélectionné');
      exit;
      end;

   Trace := TStringList.Create;
   if (Liste.AllSelected = TRUE) then
      begin
      InitMoveProgressForm (nil, 'Suppression en cours',
                            'Veuillez patienter SVP ...',
                            TFmul(Ecran).Q.RecordCount, FALSE, TRUE);
      InitMove(TFmul(Ecran).Q.RecordCount, '');

      Maintenant:= Now;
      if Trace <> nil then
         Trace.Add ('Début de suppression à '+TimeToStr(Maintenant));
{$IFDEF EAGLCLIENT}
      if (TFMul(Ecran).bSelectAll.Down) then
         TFMul(Ecran).Fetchlestous;
{$ENDIF}
      TFmul(Ecran).Q.First;
      while not TFmul(Ecran).Q.EOF do
            begin
            Delete_un;
            TFmul(Ecran).Q.Next;
            end;

      Liste.AllSelected:= False;
      TFMul(Ecran).bSelectAll.Down := Liste.AllSelected;
      end
   else
      begin
      InitMoveProgressForm (NIL,'Suppression en cours',
                            'Veuillez patienter SVP ...',
                            Liste.NbSelected, FALSE, TRUE);
      InitMove(Liste.NbSelected, '');

      Maintenant:= Now;
      if Trace <> nil then
         Trace.Add ('Début de suppression à '+TimeToStr (Maintenant));

      for i := 0 to Liste.NbSelected - 1 do
          begin
          Liste.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
          TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row-1) ;
{$ENDIF}
          Delete_un;
          end;

      Liste.ClearSelected;
      end;

   FiniMove;
   FiniMoveProgressForm;
   PGIBox ('Traitement terminé', TFMul (Ecran).Caption);
   Maintenant:= Now;
   if Trace <> nil then
      Trace.Add ('Suppression terminée à '+TimeToStr (Maintenant));

{$IFNDEF DADSUSEULE}
   CreeJnalEvt('001', '043', 'OK', nil, nil, Trace);
{$ENDIF}    
   FreeAndNil(Trace);
   end;

if BCherche <> nil then
   BCherche.click;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/05/2004
Modifié le ... :   /  /
Description .. : Suppression d'un seul salarié
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADS.Delete_un();
var
St, StExiste : String;
Existe : boolean;
Maintenant : TDateTime;
begin
St:= TFmul(Ecran).Q.FindField('PDE_SALARIE').asstring;
NumPer:= TFmul(Ecran).Q.FindField('PDE_ORDRE').asinteger;
if (St <> '') then
   begin
   if (isnumeric(St) and (VH_PAIE.PgTypeNumSal = 'NUM')) then
      St:= ColleZeroDevant(StrToInt(St), 10);
   try
      begintrans;
      if Trace <> nil then
         Trace.Add ('Traitement du salarié '+St+' Periode : '+IntToStr (NumPer));
{PT30
      DuplicPrudh(St, TypeD, 'D', NumPer);
      DeletePeriode(St, TypeD, NumPer);
      DeleteDetail(St, TypeD, NumPer);
}
      DuplicPrudh(St, 'D', NumPer);
      DeletePeriode(St, NumPer);
      DeleteDetail(St, NumPer);
//FIN PT30
      DeleteErreur (St, 'S4', NumPer);	//PT27-1
      StExiste:= 'SELECT PDS_SALARIE'+
                 ' FROM DADSDETAIL WHERE'+
                 ' PDS_SALARIE="'+St+'" AND'+
                 ' PDS_TYPE="'+TypeD+'" AND'+
                 ' PDS_ORDRE>0 AND'+
                 ' PDS_EXERCICEDADS = "'+PGExercice+'"';
      Existe := ExisteSQL(StExiste);
      if (Existe = False) then
         begin
         if Trace <> nil then
            Trace.Add ('Traitement du salarié '+St);
{PT30
         DeleteDetail(St, TypeD, 0);
         DeletePeriode(St, TypeD, -200);
         DeleteDetail(St, TypeD, -200);
}
         DeleteDetail(St, 0);
         DeletePeriode(St, -200);
         DeleteDetail(St, -200);
//FIN PT30
         DeleteErreur (St, 'S3', 0);	//PT27-1
         StExiste:= 'UPDATE SALARIES SET'+
                    ' PSA_DADSDATE="'+UsDateTime(IDate1900)+'" WHERE'+
                    ' PSA_SALARIE = "'+St+'"';
         ExecuteSQL(StExiste);
         end;
      CommitTrans;
    except
      Rollback;
      Maintenant:= Now;
      if Trace <> nil then
         Trace.Add ('Salarié '+St+' : Suppression annulée à '+TimeToStr (Maintenant));
      end;
   end;
MoveCur(False);
MoveCurProgressForm (St); // PORTAGECWAS
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 31/03/2004
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "Visualiser"
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
(*PT27-1
procedure TOF_PGMULDADS.VisuClick(Sender: TObject);
var
NomFic, St : String;
begin
ForceNumerique(GetParamSoc('SO_SIRET'), st);
if ControlSiret(st)=False then
   PGIAsk('Le SIRET de la société n''est pas valide.#13#10'+
          'Vous devez le vérifier en y accédant par le module#13#10'+
          'Paramètres/menu Comptabilité/commande Paramètres comptables/Coordonnées.#13#10'+
          'Si vous travaillez en environnement multi-dossiers,#13#10'+
          'vous pouvez y accéder par le Bureau PGI/Annuaire',
          'Visualisation du contrôle du calcul');
{$IFDEF EAGLCLIENT}
NomFic:= VH_Paie.PgCheminEagl+'\'+st+'_DADSU_PGI.log';
{$ELSE}
NomFic:= V_PGI.DatPath+'\'+st+'_DADSU_PGI.log';
{$ENDIF}
ShellExecute( 0, PCHAR('open'),PChar('WordPad'), PChar(NomFic),Nil,SW_RESTORE);
end;
*)

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/10/2001
Modifié le ... :   /  /
Description .. : Sortie de la zone de saisie du code salarié
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADS.SalarieExit(Sender: TObject);
begin

if (isnumeric(SALAR.Text) AND (VH_PAIE.PgTypeNumSal='NUM')) then
   SALAR.Text:=ColleZeroDevant(StrToInt(SALAR.Text),10);
ActiveWhere (Sender);

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 02/10/2001
Modifié le ... :   /  /
Description .. : Modification de la date
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}

procedure TOF_PGMULDADS.CarDateChange(Sender: TObject);
begin
Parametrage;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/10/2001
Modifié le ... :   /  /
Description .. : Gestion de modification de la nature
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADS.NatureChange(Sender: TObject);
begin
if (Car <> nil) then
   begin
   if ((Nature.Value>='0400') and (Nature.Value<'0500')) then
      begin
      Car.Enabled := True;
      Car.Visible := True;
      SetControlEnabled('LCAR', True);
      SetControlVisible('LCAR', True);
      end
   else
      begin
      Car.Enabled := False;
      Car.Visible := False;
      Car.Value := 'A00';
      SetControlEnabled('LCAR', False);
      SetControlVisible('LCAR', False);
      end;
   end;

if ((Nature.Value>='0400') and (Nature.Value<'0500')) then
   begin
   if (VH_Paie.PGBTP = False) then
      begin
      PGIBox ('Vous ne gérez pas le module BTP', 'DADS-U standard');
      Nature.Value := '0151';
      end
   else
      if (param = 'C') then
         begin
         PGIBox ('Choisissez le menu Calcul DADS-U congés BTP',
                 'DADS-U standard');
         Nature.Value := '0151';
         end;
//PT15-3
{PT21-5
      else
         if (param = 'I') then
            begin
            PGIBox('Pour une DADS-U congés BTP, la saisie des périodes#13#10'+
                   'd''inactivité s''effectue dans l''onglet BTP de la#13#10'+
                   'période d''activité', 'DADS-U standard');
            Nature.Value := '0151';
            end;
}            
   end;
Parametrage;
ActiveWhere (Sender);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/02/2002
Modifié le ... :   /  /
Description .. : Gestion de modification de la nature
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADS.Parametrage();
var
Car1, StExer : string;
QExer : TQuery;
NbMois : integer;
begin
PGAnnee:= Annee.value;
StExer:= 'SELECT PEX_DATEDEBUT, PEX_DATEFIN, PEX_ANNEEREFER'+
         ' FROM EXERSOCIAL WHERE'+
         ' PEX_EXERCICE="'+PGAnnee+'"';
QExer:= OpenSQL (StExer,TRUE) ;
if (NOT QExer.EOF) then
   PGExercice:= QExer.FindField ('PEX_ANNEEREFER').AsString
else
   begin
   PGExercice:= '';
   Ferme (QExer);
   exit;
   end;

if ((Nature.Value>='0400') and (Nature.Value<'0500')) then
   begin
   if ((Car.Value>='M00') and (Car.Value <='M99')) then
      begin
      TypeD:= '005';
      if (Car.Value = 'M01') then
         begin
         DebExer:= StrToDate ('01/01/'+PGExercice);
         FinExer:= StrToDate ('31/01/'+PGExercice);
         end
      else
      if (Car.Value = 'M02') then
         begin
         DebExer:= StrToDate ('01/02/'+PGExercice);
         FinExer:= FinDeMois (StrToDate ('28/02/'+PGExercice));
         end
      else
      if (Car.Value = 'M03') then
         begin
         DebExer:= StrToDate ('01/03/'+PGExercice);
         FinExer:= StrToDate ('31/03/'+PGExercice);
         end
      else
      if (Car.Value = 'M04') then
         begin
         DebExer:= StrToDate ('01/04/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/04/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M05') then
         begin
         DebExer:= StrToDate ('01/05/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('31/05/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M06') then
         begin
         DebExer:= StrToDate ('01/06/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/06/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M07') then
         begin
         DebExer:= StrToDate ('01/07/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('31/07/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M08') then
         begin
         DebExer:= StrToDate ('01/08/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('31/08/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M09') then
         begin
         DebExer:= StrToDate ('01/09/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/09/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M10') then
         begin
         DebExer:= StrToDate ('01/10/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('31/10/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M11') then
         begin
         DebExer:= StrToDate ('01/11/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/11/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M12') then
         begin
         DebExer:= StrToDate ('01/12/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('31/12/'+IntToStr (StrToInt (PGExercice)-1));
         end;
      end;
   if ((Car.Value>='T00') and (Car.Value <='T99')) then
      begin
      TypeD:= '004';
      if (Car.Value = 'T01') then
         begin
{PT38
         DebExer:= StrToDate ('01/04/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/06/'+IntToStr (StrToInt (PGExercice)-1));
}
         DebExer:= StrToDate ('01/01/'+PGExercice);
         FinExer:= StrToDate ('31/03/'+PGExercice);
         end
      else
      if (Car.Value = 'T02') then
         begin
{
         DebExer:= StrToDate ('01/07/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/09/'+IntToStr (StrToInt (PGExercice)-1));
}
         DebExer:= StrToDate ('01/04/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/06/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'T03') then
         begin
{
         DebExer:= StrToDate ('01/10/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('31/12/'+IntToStr (StrToInt (PGExercice)-1));
}
         DebExer:= StrToDate ('01/07/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/09/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'T04') then
         begin
{
         DebExer:= StrToDate ('01/01/'+PGExercice);
         FinExer:= StrToDate ('31/03/'+PGExercice);
}
         DebExer:= StrToDate ('01/10/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('31/12/'+IntToStr (StrToInt (PGExercice)-1));
//FIN PT38         
         end;
      end;
   if ((Car.Value>='S00') and (Car.Value <='S99')) then
      begin
{PT27-5
      TypeD := '003';
      if (Car.Value = 'S01') then
         begin
         DebExer := StrToDate('01/04/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('30/09/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'S02') then
         begin
         DebExer := StrToDate('01/10/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('31/03/'+PGExercice);
         end;
}
      PGIBox ('Périodicité interdite', TFMul (Ecran).Caption);
      Car.Value:= 'A00';
//FIN PT27-5
      end;
   if Car.Value='A00' then
      begin
      TypeD:= '002';
      DebExer:= StrToDate ('01/04/'+IntToStr (StrToInt (PGExercice)-1));
      FinExer:= StrToDate ('31/03/'+PGExercice);
      end;
   end
else
   begin
   TypeD:= '001';
   if NOT QExer.eof then
      begin
      DebExer:= QExer.FindField ('PEX_DATEDEBUT').AsDateTime;
      FinExer:= QExer.FindField ('PEX_DATEFIN').AsDateTime;
      end;
   end;
Ferme (QExer);

//PT27-4
{PT29
if (GetControlText ('CHCOMPL')='X') then
   TypeD:= '2'+Copy (TypeD, 2, 2);
}
MajTypeD;
//FIN PT29
//FIN PT27-4

//PT27-2
Decalage:= GetControlText ('DECALAGE');
Car1:= Copy (Car.Value, 1, 1);
if (Car1='A') then
   NbMois:= 12
else
if (Car1='S') then
   NbMois:= 6
else
if (Car1='T') then
   NbMois:= 3
else
if (Car1='M') then
   NbMois:= 1
else
   NbMois:= 12;

if (Decalage = '02') then  //Adoption du décalage
   begin
   if (Car1 <> 'M') then
      begin
      FinExer:= PlusDate (DebExer, NbMois-1, 'M');
      FinExer:= PlusDate (FinExer, -1, 'J');
      SetControlText ('L_DATEDECALAGE', 'Date d''adoption');
      SetControlVisible ('L_DATEDECALAGE', True);
      SetControlVisible ('DATEDECALAGE', True);
      SetControlVisible ('BMODIFDATE', True);
      end
   else
      begin
      PGIBox ('Vous devez effectuer une DADS-U Néant', 'Adoption du décalage');
      SetControlEnabled ('BCALCULER', False);
      end;
   end
else
if (Decalage = '04') then             //Suppression du décalage
   begin
   FinExer:= PlusDate (DebExer, NbMois+1, 'M');
   FinExer:= PlusDate (FinExer, -1, 'J');
   SetControlText ('L_DATEDECALAGE', 'Date de suppression');
   SetControlVisible ('L_DATEDECALAGE', True);
   SetControlVisible ('DATEDECALAGE', True);
   SetControlVisible ('BMODIFDATE', True);
   end
else
   begin
   SetControlVisible ('L_DATEDECALAGE', False);
   SetControlVisible ('DATEDECALAGE', False);
   if (((Decalage = '01') and (VH_Paie.PGDecalage=False)) or
      ((Decalage = '03') and (VH_Paie.PGDecalage=True))) then
      SetControlVisible ('BMODIFDATE', False)
   else
      begin
      if (Decalage = '01') then
         begin
         DebExer:= PlusDate (DebExer, 1, 'M');
         FinExer:= PlusDate (FinExer, 1, 'M');
         end
      else
      if (Decalage = '03') then
         begin
         DebExer:= PlusDate (DebExer, -1, 'M');
         FinExer:= PlusDate (FinExer, -1, 'M');
         end;
      SetControlVisible ('BMODIFDATE', True);
      end;
   end;
SetControlText ('DATEDECALAGE', DateToStr (FinExer));
//FIN PT27-2

SetControlText ('L_DDU', DateToStr (DebExer));
SetControlText ('L_DAU', DateToStr (FinExer));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 04/05/2004
Modifié le ... :   /  /
Description .. : Complément des raccourcis claviers
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADS.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
TFMul(Ecran).FormKeyDown(Sender, Key, Shift);
case Key of
     VK_F6: if ((GetControlVisible ('BCALCULER')) and
               (GetControlEnabled ('BCALCULER'))) then
               Calculer.Click; //Calcul des éléments
{PT27-1
     ord('I'): if ((GetControlVisible('BVISUALISER')) and
                  (GetControlEnabled('BVISUALISER')) and (ssCtrl in Shift)) then
                  Visualiser.Click; //Visualisation des éléments
}
     VK_Delete: if ((GetControlVisible ('BDELETE')) and
                   (GetControlEnabled ('BDELETE'))) then
                   Delete.Click; //Suppression des éléments
     end;
end;
// DEB PT25
{$IFDEF RYVIA}
procedure TOF_PGMULDADS.DD1Exit (Sender: TObject);
Var
DD : TDateTime;
begin
DD:= StrToDate (GetControlText ('DD1'));
SetControlText ('L_DDU', DateToStr (DD));
DebExer:= DD;
end;

procedure TOF_PGMULDADS.DD2Exit (Sender: TObject);
Var
DD : TDateTime;
begin
DD:= StrToDate (GetControlText ('DD2'));
SetControlText ('L_DAU', DateToStr (DD));
FinExer:= DD;
PGIInfo ('Attention, vous devez modifier le décalage pour indiquer le#13#10'+
         'type de décalage à prendre en compte dans le calcul.',
         Ecran.Caption);
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 24/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGMULDADS.DecalageChange (Sender: TObject);
begin
{PT27-2
  st := GetControlText ('DECALAGE');
  if st = '03' then PGIINFO ('Attention, vous devez modifier les dates des périodes #13#10 pour'+
           ' prendre en compte le décalage dans le calcul.', Ecran.Caption)
           else PGIINFO ('Attention, vérifiez les dates des périodes #13#10 pour '+
           ' effectuer les calculs en fonction dudécalage.', Ecran.Caption);
  SetFocusControl ('DD1');
}
Parametrage;
Decalage:= GetControlText ('DECALAGE');
//FIN PT27-2
ActiveWhere (Sender);     //PT28
end;
// FIN PT25

//PT27-2
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 24/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGMULDADS.DateDecaleChange (Sender: TObject);
begin
AglLanceFiche ('PAY', 'DADS_DATE', '', '', '');
SetControlText ('L_DDU', DateToStr (DebExer));
SetControlText ('L_DAU', DateToStr (FinExer));
SetControlText ('DATEDECALAGE', DateToStr (FinExer));
ActiveWhere (Sender);   //PT28
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/04/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGMULDADS.DateDecalageChange (Sender: TObject);
begin
DateDecalage:= StrToDate (GetControlText ('DATEDECALAGE'));
ActiveWhere (Sender);   //PT28
end;
//FIN PT27-2

//PT29
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 07/11/2006
Modifié le ... :   /  /
Description .. : Clic sur case à cocher "Déclaration complémentaire"
Mots clefs ... :
*****************************************************************}
procedure TOF_PGMULDADS.ComplClick (Sender: TObject);
begin
MajTypeD;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 07/11/2006
Modifié le ... :   /  /
Description .. : Mise à jour du type de déclaration
Mots clefs ... :
*****************************************************************}
procedure TOF_PGMULDADS.MajTypeD ();
begin
if (GetControlText ('CHCOMPL')='X') then
   TypeD:= '2'+Copy (TypeD, 2, 2)
else
   TypeD:= '0'+Copy (TypeD, 2, 2);
end;
//FIN PT29


initialization
  registerclasses([TOF_PGMULDADS]);
end.

