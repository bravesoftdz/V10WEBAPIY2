{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 18/10/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : PG_DADSHONOR ()
Mots clefs ... : TOF;PG_DADSHONOR
*****************************************************************
PT1   : 11/01/2002 JL V571 ajout de la destruction de la tob TDetailSal, et
                           suppression de la TOB TDetail car inutilisée
PT2   : 26/03/2002 VG V571 La date de clôture ne doit être indiquée que si
                           différente de 31/12
PT3   : 05/04/2002 JL V571 Modification du contrôle SIRET : message bloquant si
                           moins de 14 caractères (pour éviter la saisie d'un
                           SIREN autorisée avant).
PT4   : 17/04/2002 JL V571 Modification chargement combo : PDS_DATEDEBUT
                           reprise au lieu de la valeur de la date de cloture
                           (car le champ n'existe plus si valeur 31/12)
PT5-1 : 15/07/2002 VG V585 Adaptation cahier des charges V6R02
PT5-2 : 15/07/2002 VG V585 En création par le bouton situé sur le mul, le
                           navigateur était actif
PT5-3 : 15/07/2002 VG V585 Date de clôture plus alimentée par requête mais par
                           PGAnnee
PT6   : 02/09/2002 VG V585 On ne permet plus de modifier l'exercice lorsqu'on
                           est dans la saisie des honoraires
PT7   : 05/09/2002 VG V585 Si on avait touché à la combo bis ou ter, un
                           enregistrement S70.G01.00.004.004 était créé à blanc
PT8   : 21/10/2002 VG V585 Gestion du journal des evenements
PT9   : 31/01/2003 VG V591 Arrondi des montants à l'euro le plus proche
PT10  : 13/03/2003 VG V_42 Gestion du journal des evenements
PT11  : 14/04/2003 VG V_42 Modification de la tablette PGADRESSECOMPL
PT12  : 16/09/2003 VG V_42 Adaptation cahier des charges V7R01
PT13  : 12/02/2004 VG V_50 Sur les honoraires, le cahier des charges interdit
                           d'avoir Raison sociale et Nom/Prénom renseignés
                           simultanément - FQ N°11107
PT14  : 08/03/2004 VG V_50 Le SIRET du bénéficiaire est obligatoire pour une
                           personne morale - FQ N°11108
PT15  : 05/07/2004 VG V_50 Adaptation cahier des charges V8R00
PT16  : 12/10/2004 VG V_50 Il n'existait pas de validation automatique en cas de
                           navigation - FQ N°11676
PT17  : 17/11/2004 VG V_60 Si navigation et validation des modifications, les
                           boutons "suivant" et "dernier" devenaient grisés
                           FQ N°11744
PT18  : 06/01/2005 VG V_60 La touche F10 ne fonctionnait pas - FQ N°11779
PT19  : 10/01/2005 VG V_60 Erreur de format : JJMAAAA au lieu de JJMMAAAA si
                           mois de clôture antérieur à octobre - FQ N°11913
PT20  : 12/01/2005 VG V_60 Le SIRET du bénéficiaire est obligatoire pour une
                           entreprise ayant un établissement situé sur le
                           territoire national - FQ N°11921
PT21  : 24/01/2005 VG V_60 Le code postal est numérique sur 5 positions
                           uniquement si pays est "FRA" ou "" - FQ N°11921
PT22  : 07/10/2005 VG V_60 Adaptation cahier des charges DADS-U V8R02
PT23  : 24/11/2005 VG V_65 Initialisation au 31/12 de la date de fin
                           d'exercice comptable - FQ N°12681
PT24  : 02/12/2005 VG V_65 L'enregistrement du "code modalité de prise en charge
                           des honoraires" n'enregistrait que les 2 premiers
                           codes lorsque la valeur était "F;R;P;" - FQ N°12718
PT25  : 08/12/2005 VG V_65 Lors de la validation contrôle de validité de la date
                           de fin d'exercice comptable - FQ N°12681
PT26  : 04/01/2006 VG V_65 Lors de la validation contrôle du "code modalité de
                           prise en charge des honoraires" - FQ N°12719
PT27-1: 12/10/2006 VG V_70 Modification du champ PDS_SALARIE pour les honoraires
PT27-2: 12/10/2006 VG V_70 Utilisation d'un type pour la cle DADS-U
PT27-3: 12/10/2006 VG V_70 Suppression du fichier de contrôle - mise en table
                           des erreurs
PT28  : 23/10/2006 VG V_70 Le message n'était pas toujours affiché - FQ N°13612
PT29  : 21/11/2006 VG V_70 Permettre de déclarer des honoraires en DADS-U
                           complémentaire - FQ N°13613
PT30-1: 17/08/2007 VG V_80 Adaptation cahier des charges V8R05 - FQ N°14424
PT30-2: 01/10/2007 VG V_80 Valeur 0 non acceptée pour
                           "Montant TVA droits d'auteurs" - FQ N°12853
PT31  : 28/11/2007 VG V_80 Gestion du champ ET_FICTIF - FQ N°13925
PT32  : 04/12/2007 NA V_80 FQ N°14449 Interdire 2 types identiques de
                           remunération
PT33  : 10/12/2007 VG V_80 Pas de montant à zéro - FQ N°13869
}
Unit PG_DADSHONOR_TOF ;

Interface
uses     UTOF,
{$IFNDEF EAGLCLIENT}
         DBCtrls,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
         UtileAGL,
{$ENDIF}
         Hctrls,
         HEnt1,
         HMsgBox,
         StdCtrls,
         Classes,
         sysutils,
         UTob,
         HTB97,
         Vierge,
         PgDADSCommun,
         PgDADSOutils,
         EntPaie,
         Controls,
{$IFNDEF DADSUSEULE}
         P5Def,
{$ENDIF}
         PgOutils2,
         windows,
         ParamSoc;

Type
  TOF_PG_DADSHONOR = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (stArgument : String ) ; override ;
    procedure OnClose                  ; override ;
    private
    QSal:TQuery;

    Autres, Voiture, Logement, Nourriture, NTIC : TCheckBox;

    TypeRem1, TypeRem2, TypeRem3, TypeRem4, TypeRem5 : THValComboBox;

    Arg, State, Sal, NbreTot, ExerSocial:String;

    SalPrem, SalPrec, SalSuiv, SalDern, BDaccord : TToolBarButton97;

    choixreq : Boolean;

    NumOrdre : Integer;

    procedure SalPremClick (Sender: TObject);
    procedure SalPrecClick (Sender: TObject);
    procedure SalSuivClick (Sender: TObject);
    procedure SalDernClick (Sender: TObject);
    procedure ChargeZones();
    procedure SauveZones();
    procedure Validation(Sender:TObject);
    function UpdateTable() : boolean;
    Function BougeSal(Button: TNavigateBtn) : boolean ;
    procedure MetABlanc();
    procedure New(Sender: TObject);
    procedure Del(Sender: TObject);
    procedure AccesMontant(Sender: TObject);
    procedure MAJQuery();
    function Controles() : boolean;
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  end;

Implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 18/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PG_DADSHONOR
*****************************************************************}
procedure TOF_PG_DADSHONOR.OnLoad;
var QCaption: TQuery;
begin
  inherited;
ChargeZones;
QCaption:=OpenSQL('SELECT PEX_ANNEEREFER'+
                  ' FROM EXERSOCIAL WHERE'+
                  ' PEX_EXERCICE="'+GetControlText('CEXERSOC')+'"',True);
If not QCaption.eof then
   begin
   if VH_Paie.PGTenueEuro=True Then
      Ecran.Caption:= 'Saisie des honoraires en Euro de l''année '+
                      QCaption.FindField('PEX_ANNEEREFER').AsString
   Else
       Ecran.Caption:= 'Saisie des honoraires en Francs de l''année '+
                       QCaption.FindField('PEX_ANNEEREFER').AsString;
   end
Else
   Ecran.Caption :='Saisie des honoraires';

Ecran.Caption := TraduireMemoire(Ecran.Caption); // Traduction
Ferme(QCaption);
UpdateCaption(Ecran);
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 18/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PG_DADSHONOR
*****************************************************************}
procedure TOF_PG_DADSHONOR.OnArgument (stArgument : String ) ;
var
StPer, StPlus : String;
begin
Inherited ;
choixreq:=False;
Sal:='--H';
BDaccord:= TToolBarButton97(GetControl('BDACCORD'));
If BDaccord<>NIL Then
   BDaccord.OnClick:=Validation;
State :=(Trim(ReadTokenSt(stArgument)));
Arg :=trim(ReadTokenPipe(stArgument,';'));
ExerSocial :=trim(ReadTokenPipe(stArgument,';'));

TFVierge(Ecran).OnKeyDown:= FormKeyDown;

Nourriture := TCheckBox (GetControl ('THNOURRITURE'));
Logement := TCheckBox (GetControl ('THLOGEMENT'));
Voiture := TCheckBox (GetControl ('THVOITURE'));
Autres := TCheckBox (GetControl ('THAUTRES'));
NTIC := TCheckBox (GetControl ('THNTIC'));
TypeRem1 := THValComboBox (GetControl ('CTYPEREM1'));
TypeRem2 := THValComboBox (GetControl ('CTYPEREM2'));
TypeRem3 := THValComboBox (GetControl ('CTYPEREM3'));
TypeRem4 := THValComboBox (GetControl ('CTYPEREM4'));
TypeRem5 := THValComboBox (GetControl ('CTYPEREM5'));

if TypeRem1 <> NIL then
   TypeRem1.OnChange:=AccesMontant;
if TypeRem2 <> NIL then
   TypeRem2.OnChange:=AccesMontant;
if TypeRem3 <> NIL then
   TypeRem3.OnChange:=AccesMontant;
if TypeRem4 <> NIL then
   TypeRem4.OnChange:=AccesMontant;
if TypeRem5 <> NIL then
   TypeRem5.OnChange:=AccesMontant;

if State = 'CREATION' then
   begin
   TFVierge(Ecran).Binsert.Enabled := False;
   TFVierge(Ecran).BDelete.Enabled := False;
   end;
TFVierge(Ecran).Binsert.OnClick := New;
TFVierge(Ecran).BDelete.OnClick := Del;

// Gestion du navigateur
SalPrem := TToolbarButton97(GetControl('BSALPREM'));
SalPrec := TToolbarButton97(GetControl('BSALPREC'));
SalSuiv := TToolbarButton97(GetControl('BSALSUIV'));
SalDern := TToolbarButton97(GetControl('BSALDERN'));

{$IFNDEF EAGLCLIENT}
if SalPrem <> NIL then
   begin
   if State <> 'CREATION' then
      SalPrem.Enabled := True;
   SalPrem.Visible := True;
   SalPrem.OnClick := SalPremClick;
   end;

if SalPrec <> NIL then
   begin
   if State <> 'CREATION' then
      SalPrec.Enabled := True;
   SalPrec.Visible := True;
   SalPrec.OnClick := SalPrecClick;
   end;

if SalSuiv <> NIL then
   begin
   if State <> 'CREATION' then
      SalSuiv.Enabled := True;
   SalSuiv.Visible := True;
   SalSuiv.OnClick := SalSuivClick;
   end;

if SalDern <> NIL then
   begin
   if State <> 'CREATION' then
      SalDern.Enabled := True;
   SalDern.Visible := True;
   SalDern.OnClick := SalDernClick;
   end;
{$ELSE}
if SalPrem <> NIL then
   SalPrem.Visible := False;

if SalPrec <> NIL then
   SalPrec.Visible := False;

if SalSuiv <> NIL then
   SalSuiv.Visible := False;

if SalDern <> NIL then
   SalDern.Visible := False;
{$ENDIF}

// recuperation de la query du multicritere
StPer:= 'SELECT *'+
        ' FROM DADSPERIODES'+
        ' WHERE PDE_SALARIE LIKE "'+Sal+'%" AND'+
        ' PDE_DATEDEBUT="'+UsDateTime(DebExer)+'"'+
        ' ORDER BY PDE_TYPE,PDE_ORDRE,PDE_DATEFIN';
QSal:=OpenSql(StPer, True);
IF State='MODIFICATION' Then
   begin
   if not QSal.eof then
      begin
      While QSal.FindField('PDE_ORDRE').AsInteger<>StrToInt(Arg) do
            QSal.Next;
      end;
   end;

//PT31
StPlus:= ' WHERE ET_FICTIF<>"X"';
SetControlProperty ('CETAB', 'Plus', StPlus);
//FIN PT31
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 18/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PG_DADSHONOR
*****************************************************************}
procedure TOF_PG_DADSHONOR.ChargeZones();
var
StDADSD : String;
TDetailSal, TDetailD:Tob;
QRechDADSD : TQuery;
annee, jour, mois : word;
begin
MetABlanc;
FreeAndNil (TDetailSal);

If State='CREATION' Then
   SetControlText('CEXERSOC', PGAnnee);

if (State = 'MODIFICATION') then // chargement uniquement en mode modification
   begin
   StDADSD:= 'SELECT PDS_SALARIE, PDS_TYPE, PDS_ORDRE, PDS_DATEDEBUT,'+
             ' PDS_DATEFIN, PDS_ORDRESEG, PDS_SEGMENT, PDS_DONNEE,'+
             ' PDS_DONNEEAFFICH'+
             ' FROM DADSDETAIL WHERE'+
             ' PDS_SALARIE LIKE "'+Sal+'%" AND'+
             ' PDS_TYPE="'+TypeD+'" AND'+
             ' PDS_ORDRE='+Arg+' AND'+
             ' PDS_DATEDEBUT="'+UsDateTime(DebExer)+'"'+
             ' ORDER BY PDS_ORDRESEG,PDS_SEGMENT,PDS_DATEFIN';
   QRechDADSD:=OpenSql(StDADSD,TRUE);
   TDetailSal := TOB.Create('Les détails honoraire', NIL, -1);
   TDetailSal.LoadDetailDB('DADSDETAIL','','',QRechDADSD,False);
   Ferme(QRechDADSD);

   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.001'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('ERAISONSOC', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.002.001'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('ENOM', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.002.002'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('EPRENOM', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.003.001'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('ESIRET', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.004.001'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('EADRCOMP', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.004.003'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('EADRNUM', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.004.004'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('CADRBIS', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.004.006'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('EADRNOMVOIE', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.004.007'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('EADRINSEE', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.004.009'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('EADRNOMCOM', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.004.010'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('EADRCP', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.004.012'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('EADRBURDISTRIB', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.004.013'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('CADRPAYS', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.005'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('EPROFESSION', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.006'], TRUE);
   if ((Nourriture <> NIL) and (TDetailD <> NIL)) then
      Nourriture.Checked := TDetailD.GetValue('PDS_DONNEEAFFICH')='N';

   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.007'], TRUE);
   if ((Logement <> NIL) and (TDetailD <> NIL)) then
      Logement.Checked := TDetailD.GetValue('PDS_DONNEEAFFICH')='L';

   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.008'], TRUE);
   if ((Voiture <> NIL) and (TDetailD <> NIL)) then
      Voiture.Checked := TDetailD.GetValue('PDS_DONNEEAFFICH')='V';

   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.009'], TRUE);
   if ((Autres <> NIL) and (TDetailD <> NIL)) then
      Autres.Checked := TDetailD.GetValue('PDS_DONNEEAFFICH')='A';

   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.010'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('MCCODEINDEMNITES', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.011'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('CCODETAUXREDUIT', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.013'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('EMONTANTTVADA', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.014'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('CETAB', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   DecodeDate (FinExer, annee, mois, jour);
   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.015'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('EEXERCOMPTA', TDetailD.GetValue('PDS_DONNEEAFFICH'))
   else
      SetControlText('EEXERCOMPTA', DateToStr(EncodeDate(annee, 12, 31)));

   TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S70.G01.00.016'], TRUE);
   if ((NTIC <> NIL) and (TDetailD <> NIL)) then
      NTIC.Checked := TDetailD.GetValue('PDS_DONNEEAFFICH')='T';
   SetControlText('CEXERSOC', PGAnnee);

   TDetailD := TDetailSal.FindFirst(['PDS_ORDRESEG'], ['101'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('CTYPEREM1', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_ORDRESEG'], ['102'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('EMONTREM1', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_ORDRESEG'], ['104'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('CTYPEREM2', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_ORDRESEG'], ['105'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('EMONTREM2', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_ORDRESEG'], ['107'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('CTYPEREM3', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_ORDRESEG'], ['108'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('EMONTREM3', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_ORDRESEG'], ['110'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('CTYPEREM4', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_ORDRESEG'], ['111'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('EMONTREM4', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_ORDRESEG'], ['113'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('CTYPEREM5', TDetailD.GetValue('PDS_DONNEEAFFICH'));

   TDetailD := TDetailSal.FindFirst(['PDS_ORDRESEG'], ['114'], TRUE);
   if (TDetailD <> NIL) then
      SetControlText('EMONTREM5', TDetailD.GetValue('PDS_DONNEEAFFICH'));
   end;
FreeAndNil(TDetailSal);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 18/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PG_DADSHONOR
*****************************************************************}
procedure TOF_PG_DADSHONOR.SauveZones();
var
Q, QEtab : TQuery;
BufDest, BufOrig, CodeIso, Libelle, NIC, NICe, SDateCloture : String;
SDateCloture4, SIRENe : String;
annee, jour, mois : word;
CleDADS : TCleDADS;
begin
Trace:= TStringList.Create;

if (State='CREATION') then
   begin
   NumOrdre:= 1;
   Q:= OpenSQL ('SELECT MAX(PDE_ORDRE) AS NUMMAX'+
                ' FROM DADSPERIODES WHERE'+
                ' PDE_SALARIE LIKE "'+Sal+'%"', True);
   if (not Q.eof) then
      try
      NumOrdre:= (Q.FindField ('NUMMAX').AsInteger)+1;
      except
            on E: EConvertError do
               NumOrdre:= 1;
      end;
   Ferme(Q);
   end;

if (State='MODIFICATION') then
   begin
   ExecuteSQL ('DELETE FROM DADSDETAIL WHERE'+
               ' PDS_ORDRE='+Arg+' AND PDS_SALARIE LIKE "'+Sal+'%"');
   ExecuteSQL ('DELETE FROM DADSPERIODES WHERE'+
               ' PDE_ORDRE='+Arg+' AND PDE_SALARIE LIKE "'+Sal+'%"');
   NumOrdre:= StrToInt (Arg);
   end;
CleDADS.Salarie:= Sal+Arg;
CleDADS.TypeD:= TypeD;
CleDADS.Num:= NumOrdre;
CleDADS.DateDeb:= DebExer;
CleDADS.DateFin:= FinExer;
CleDADS.Exercice:= PGExercice;
CreeEntete (CleDADS, '', '', '01');

if (GetControlText ('ERAISONSOC')<>'') then
   CreeDetail (CleDADS, 1, 'S70.G01.00.001', GetControlText ('ERAISONSOC'),
               GetControlText ('ERAISONSOC'))
else
   begin
   if (GetControlText ('ENOM')<>'') then
      CreeDetail (CleDADS, 2, 'S70.G01.00.002.001', GetControlText ('ENOM'),
                  GetControlText ('ENOM'));
   if (GetControlText ('EPRENOM')<>'') then
      CreeDetail (CleDADS, 3, 'S70.G01.00.002.002', GetControlText ('EPRENOM'),
                  GetControlText ('EPRENOM'));
   end;

if (GetControlText ('ESIRET')<>'') then
   begin
   BufOrig:= GetControlText ('ESIRET');
   ForceNumerique (BufOrig, BufDest);
   SIRENe:= Copy (BufDest, 1, 9);
   CreeDetail (CleDADS, 4, 'S70.G01.00.003.001', SIRENe,
               GetControlText ('ESIRET'));
   NICe:= Copy (BufDest, 10, 5);
   CreeDetail (CleDADS, 5, 'S70.G01.00.003.002', NICe, GetControlText ('ESIRET'));
   end;

if (GetControlText ('EADRCOMP')<>'') then
   CreeDetail (CleDADS, 6, 'S70.G01.00.004.001', GetControlText ('EADRCOMP'),
               GetControlText ('EADRCOMP'));

if (GetControlText ('EADRNUM')<>'') then
   CreeDetail (CleDADS, 7, 'S70.G01.00.004.003', GetControlText ('EADRNUM'),
               GetControlText ('EADRNUM'));
if ((GetControlText ('CADRBIS')<>'') and
   (GetControlText ('CADRBIS')<>'           ')) then
   CreeDetail (CleDADS, 8, 'S70.G01.00.004.004',
               RechDom ('PGADRESSECOMPL', GetControlText ('CADRBIS'), True),
               GetControlText ('CADRBIS'));

if (GetControlText ('EADRNOMVOIE')<>'') then
   CreeDetail (CleDADS, 9, 'S70.G01.00.004.006', GetControlText ('EADRNOMVOIE'),
               GetControlText ('EADRNOMVOIE'));

if (GetControlText ('EADRINSEE')<>'') then
   CreeDetail (CleDADS, 10, 'S70.G01.00.004.007', GetControlText ('EADRINSEE'),
               GetControlText ('EADRINSEE'));

if (GetControlText ('EADRNOMCOM')<>'') then
   CreeDetail (CleDADS, 11, 'S70.G01.00.004.009', GetControlText ('EADRNOMCOM'),
               GetControlText ('EADRNOMCOM'));

CreeDetail (CleDADS, 12, 'S70.G01.00.004.010', GetControlText ('EADRCP'),
            GetControlText ('EADRCP'));

CreeDetail (CleDADS, 13, 'S70.G01.00.004.012',
            PGUpperCase (GetControlText ('EADRBURDISTRIB')),
            GetControlText ('EADRBURDISTRIB'));

{PT30-1
if ((GetControlText ('CADRPAYS')<>'FRA') and
   (GetControlText ('CADRPAYS')<>'')) then
}
if ((GetControlText ('CADRPAYS')<>'FRA') and
   (GetControlText ('CADRPAYS')<>'GUF') and
   (GetControlText ('CADRPAYS')<>'GLP') and
   (GetControlText ('CADRPAYS')<>'MCO') and
   (GetControlText ('CADRPAYS')<>'MTQ') and
   (GetControlText ('CADRPAYS')<>'NCL') and
   (GetControlText ('CADRPAYS')<>'PYF') and
   (GetControlText ('CADRPAYS')<>'SPM') and
   (GetControlText ('CADRPAYS')<>'REU') and
   (GetControlText ('CADRPAYS')<>'ATF') and
   (GetControlText ('CADRPAYS')<>'WLF') and
   (GetControlText ('CADRPAYS')<>'MYT') and
   (GetControlText ('CADRPAYS')<>'')) then
//FIN PT30-1
   begin
   PaysISOLib (GetControlText ('CADRPAYS'), CodeIso, Libelle);
   CreeDetail (CleDADS, 14, 'S70.G01.00.004.013', CodeIso,
               GetControlText ('CADRPAYS'));
   CreeDetail (CleDADS, 15, 'S70.G01.00.004.014', Libelle, Libelle);
   end;

CreeDetail (CleDADS, 16, 'S70.G01.00.005', GetControlText ('EPROFESSION'),
            GetControlText ('EPROFESSION'));

if (Nourriture<>nil) and (Nourriture.Checked=True) then
   CreeDetail (CleDADS, 17, 'S70.G01.00.006', 'N', 'N');

if (Logement<>nil) and (Logement.Checked=True) then
   CreeDetail (CleDADS, 18, 'S70.G01.00.007', 'L', 'L');

if (Voiture<>nil) and (Voiture.Checked=True) then
   CreeDetail (CleDADS, 19, 'S70.G01.00.008', 'V', 'V');

if (Autres<>nil) and (Autres.Checked=True) then
   CreeDetail (CleDADS, 20, 'S70.G01.00.009', 'A', 'A');

if (GetControlText ('MCCODEINDEMNITES')<>'') then
   begin
   BufOrig:= GetControlText ('MCCODEINDEMNITES');
   if (THMultiValCombobox (GetControl ('MCCODEINDEMNITES')).Tous) then
      BufDest:= 'FRP'
   else
      BufDest:= StringReplace (BufOrig, ';', '', [rfReplaceAll]);
   CreeDetail (CleDADS, 21, 'S70.G01.00.010', BufDest,
               GetControlText ('MCCODEINDEMNITES'));
   end;

if (GetControlText ('CCODETAUXREDUIT')<>'') then
   CreeDetail (CleDADS, 22, 'S70.G01.00.011', GetControlText ('CCODETAUXREDUIT'),
               GetControlText ('CCODETAUXREDUIT'));

{PT30-2
if (GetControlText ('EMONTANTTVADA')<>'') then
}
if ((GetControlText ('EMONTANTTVADA')<>'') and
   (GetControlText ('EMONTANTTVADA')<>'0')) then
//FIN PT30-2
   CreeDetail (CleDADS, 23, 'S70.G01.00.013',
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EMONTANTTVADA')), 0)),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EMONTANTTVADA')), 0)));

QEtab:= OpenSQL ('SELECT ET_SIRET'+
                 ' FROM ETABLISS WHERE'+
                 ' ET_ETABLISSEMENT="'+GetControlText ('CETAB')+'"', True);
BufOrig:= '';
if not QEtab.eof then
   BufOrig:= QEtab.FindField ('ET_SIRET').AsString;
Ferme (QEtab);
ForceNumerique (BufOrig, BufDest);
NIC:= Copy (BufDest, 10, 5);
CreeDetail (CleDADS, 25, 'S70.G01.00.014', NIC, GetControlText ('CETAB'));

DecodeDate (StrToDate (GetControlText ('EEXERCOMPTA')), annee, mois, jour);

SDateCloture:= ColleZeroDevant (jour, 2)+ColleZeroDevant(mois, 2)+
               IntToStr (annee);
SDateCloture4:= Copy (SDateCloture, 1, 4);
if (SDateCloture4<>'3112') then
   CreeDetail (CleDADS, 26, 'S70.G01.00.015', SDateCloture,
               GetControlText ('EEXERCOMPTA'));

if (NTIC<>nil) and (NTIC.Checked=True) then
   CreeDetail (CleDADS, 27, 'S70.G01.00.016', 'T', 'T');

CreeDetail (CleDADS, 101, 'S70.G01.01.001', GetControlText ('CTYPEREM1'),
            GetControlText ('CTYPEREM1'));

{PT33
if (GetControlText ('EMONTREM1')<>'') then
}
if ((GetControlText ('EMONTREM1')<>'') and
   (GetControlText ('EMONTREM1')<>'0')) then
//FIN PT33
   CreeDetail (CleDADS, 102, 'S70.G01.01.002',
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EMONTREM1')), 0)),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EMONTREM1')), 0)));

if (GetControlText ('CTYPEREM2')<>'') then
   begin
{PT33
   if (GetControlText ('EMONTREM2')='') then
}
   if ((GetControlText ('EMONTREM2')='') or
      (GetControlText ('EMONTREM2')='0')) then
//FIN PT33
      PGIBox ('Attention, le type de montant n°2 ne sera pas sauvegardé,'+
              ' car aucun montant n''a été affecté', 'Honoraires')
   else
      CreeDetail (CleDADS, 104, 'S70.G01.01.001', GetControlText ('CTYPEREM2'),
                  GetControlText ('CTYPEREM2'));
   end;

{PT33
if (GetControlText ('EMONTREM2')<>'') then
}
if ((GetControlText ('EMONTREM2')<>'') and
   (GetControlText ('EMONTREM2')<>'0')) then
//FIN PT33
   CreeDetail (CleDADS, 105, 'S70.G01.01.002',
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EMONTREM2')), 0)),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EMONTREM2')), 0)));

if (GetControlText('CTYPEREM3')<>'') then
   begin
{PT33
   if (GetControlText ('EMONTREM3')='') then
}
   if ((GetControlText ('EMONTREM3')='') or
      (GetControlText ('EMONTREM3')='0')) then
//FIN PT33
      PGIBox ('Attention, le type de montant n°3 ne sera pas sauvegardé,'+
              ' car aucun montant n''a été affecté', 'Honoraires')
   else
      CreeDetail (CleDADS, 107, 'S70.G01.01.001', GetControlText ('CTYPEREM3'),
                  GetControlText ('CTYPEREM3'));
   end;

{PT33
if (GetControlText ('EMONTREM3')<>'') then
}
if ((GetControlText ('EMONTREM3')<>'') and
   (GetControlText ('EMONTREM3')<>'0')) then
//FIN PT33
   CreeDetail (CleDADS, 108, 'S70.G01.01.002',
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EMONTREM3')), 0)),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EMONTREM3')), 0)));

if (GetControlText ('CTYPEREM4')<>'') then
   begin
{PT33
   if (GetControlText ('EMONTREM4')='') then
}
   if ((GetControlText ('EMONTREM4')='') or
      (GetControlText ('EMONTREM4')='0')) then
//FIN PT33
      PGIBox ('Attention, le type de montant n°4 ne sera pas sauvegardé,'+
              ' car aucun montant n''a été affecté', 'Honoraires')
   else
      CreeDetail (CleDADS, 110, 'S70.G01.01.001', GetControlText ('CTYPEREM4'),
                  GetControlText ('CTYPEREM4'));
   end;

{PT33
if (GetControlText ('EMONTREM4')<>'') then
}
if ((GetControlText ('EMONTREM4')<>'') and
   (GetControlText ('EMONTREM4')<>'0')) then
//FIN PT33
   CreeDetail (CleDADS, 111, 'S70.G01.01.002',
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EMONTREM4')), 0)),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EMONTREM4')), 0)));

if (GetControlText ('CTYPEREM5')<>'') then
   begin
{PT33
   if (GetControlText ('EMONTREM5')='') then
}
   if ((GetControlText ('EMONTREM5')='') or
      (GetControlText ('EMONTREM5')='0')) then
//FIN PT33
      PGIBox ('Attention, le type de montant n°5 ne sera pas sauvegardé,'+
              ' car aucun montant n''a été affecté', 'Honoraires')
   else
      CreeDetail (CleDADS, 113, 'S70.G01.01.001', GetControlText ('CTYPEREM5'),
                  GetControlText ('CTYPEREM5'));
   end;

{PT33
if (GetControlText ('EMONTREM5')<>'') then
}
if ((GetControlText ('EMONTREM5')<>'') and
   (GetControlText ('EMONTREM5')<>'0')) then
//FIN PT33
   CreeDetail (CleDADS, 114, 'S70.G01.01.002',
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EMONTREM5')), 0)),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EMONTREM5')), 0)));

{$IFNDEF DADSUSEULE}
CreeJnalEvt ('001', '044', 'OK', nil, nil, Trace);
{$ENDIF}
if (Trace<>nil) then
   begin
   Trace.Free;
   Trace:= nil;
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 18/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PG_DADSHONOR
*****************************************************************}
procedure TOF_PG_DADSHONOR.Validation(Sender:TObject);
var controleOK:Boolean;
begin
controleOK:=Controles;
If controleOK=FALSE Then
   Exit;
UpdateTable;
TFVierge(Ecran).Binsert.Enabled := True;
TFVierge(Ecran).BDelete.Enabled := True;
SalPrem.Enabled := True;
SalPrec.Enabled := True;
SalSuiv.Enabled := True;
SalDern.Enabled := True;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 18/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PG_DADSHONOR
*****************************************************************}
function TOF_PG_DADSHONOR.UpdateTable(): boolean;
var
Rep : integer;
begin
result := FALSE;
Rep:=PGIAsk ('Voulez vous sauvegarder votre saisie ?', 'Saisie complémentaire DADS-U') ;
if Rep=mrNo then
   exit
else
   result := TRUE;

try
   begintrans;
   ChargeTOBDADS;
   DeleteErreur ('', 'SKO');
   SauveZones;
   EcrireErreurKO;
   LibereTOBDADS;
   CommitTrans;
   Arg:=IntToStr(NumOrdre);
   MAJQuery;
Except
   result := FALSE;
   Rollback;
   PGIBox ('Une erreur est survenue lors de la mise à jour de la base', 'DADS-U');
   END;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 18/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PG_DADSHONOR
*****************************************************************}
procedure TOF_PG_DADSHONOR.OnClose ;
begin
  Inherited ;
If QSal<>NIL Then
   Ferme(QSal);
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 18/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PG_DADSHONOR
*****************************************************************}
procedure TOF_PG_DADSHONOR.SalPremClick(Sender: TObject);
begin
BougeSal(nbFirst) ;
SalPrem.Enabled := FALSE;
SalPrec.Enabled := FALSE;
SalSuiv.Enabled := TRUE;
SalDern.Enabled := TRUE;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 18/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PG_DADSHONOR
*****************************************************************}
procedure TOF_PG_DADSHONOR.SalPrecClick(Sender: TObject);
begin
BougeSal(nbPrior) ;
if Qsal.BOF then
   begin
   SalPrem.Enabled := FALSE;
   SalPrec.Enabled := FALSE;
   end;
SalSuiv.Enabled := TRUE;
SalDern.Enabled := TRUE;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 18/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PG_DADSHONOR
*****************************************************************}
procedure TOF_PG_DADSHONOR.SalSuivClick(Sender: TObject);
begin
BougeSal(nbNext) ;
SalPrem.Enabled := TRUE;
SalPrec.Enabled := TRUE;
if Qsal.EOF then
   begin
   SalSuiv.Enabled := FALSE;
   SalDern.Enabled := FALSE;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 18/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PG_DADSHONOR
*****************************************************************}
procedure TOF_PG_DADSHONOR.SalDernClick(Sender: TObject);
begin
BougeSal(nbLast) ;
SalPrem.Enabled := TRUE;
SalPrec.Enabled := TRUE;
SalSuiv.Enabled := FALSE;
SalDern.Enabled := FALSE;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 18/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PG_DADSHONOR
*****************************************************************}
Function TOF_PG_DADSHONOR.BougeSal(Button: TNavigateBtn) : boolean ;
begin
UpdateTable;
result:=TRUE ;

if Button=nbDelete then
   begin
   if Qsal.EOF = FALSE then
      begin
      Qsal.Next;
      if Qsal.EOF = TRUE then
         begin
         Qsal.prior ;
         if Qsal.BOF then
            Close;
         end
      end
   else
      begin
      if Qsal.BOF = FALSE then
         begin
         Qsal.prior;
         if Qsal.BOF = TRUE then
            Close;
         end;
      end;
   end;

if Qsal <> NIL then
   begin
   Case Button of
        nblast : Qsal.Last;
        nbfirst : Qsal.First;
        nbnext : Qsal.Next;
        nbprior : Qsal.prior;
        end;
   end;
Arg:=IntToStr(Qsal.FindField('PDE_ORDRE').AsInteger);
ChargeZones;
END ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 18/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PG_DADSHONOR
*****************************************************************}
procedure TOF_PG_DADSHONOR.MetABlanc();
var
annee, jour, mois : word;
begin
SetControlText('ESIRET', '');
SetControlText('ERAISONSOC', '');
SetControlText('ENOM', '');
SetControlText('EPRENOM', '');
SetControlText('EADRCOMP', '');
SetControlText('EADRNUM', '');
SetControlText('CADRBIS', '');
SetControlText('EADRNOMVOIE', '');
SetControlText('EADRINSEE', '');
SetControlText('EADRNOMCOM', '');
SetControlText('EADRCP', '');
SetControlText('EADRBURDISTRIB', '');
SetControlText('EPROFESSION', '');
SetControlText('EMONTANTTVADA', '');
DecodeDate (FinExer, annee, mois, jour);
SetControlText('EEXERCOMPTA', DateToStr(EncodeDate(annee, 12, 31)));
SetControlText('EMONTREM1', '');
SetControlText('EMONTREM2', '');
SetControlText('EMONTREM3', '');
SetControlText('EMONTREM4', '');
SetControlText('EMONTREM5', '');
SetControlText('CEXERSOC', '');
SetControlText('CADRPAYS', '');
SetControlText('CETAB', '');
SetControlText('MCCODEINDEMNITES', '');
SetControlText('CTYPEREM1', '');
SetControlText('CTYPEREM2', '');
SetControlText('CTYPEREM3', '');
SetControlText('CTYPEREM4', '');
SetControlText('CTYPEREM5', '');
SetControlText('CCODETAUXREDUIT', '');
if Nourriture <> NIL then
   Nourriture.Checked := False;
if Logement <> NIL then
   Logement.Checked := False;
if Voiture <> NIL then
   Voiture.Checked := False;
if Autres <> NIL then
   Autres.Checked := False;
if NTIC <> NIL then
   NTIC.Checked := False;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 18/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PG_DADSHONOR
*****************************************************************}
procedure TOF_PG_DADSHONOR.Del(Sender: TObject);
begin
try
   begintrans;
   DeletePeriode(Sal, StrToInt(Arg));
   DeleteDetail(Sal, StrToInt(Arg));
   Trace := TStringList.Create;
   if Trace <> nil then
      Trace.Add (Sal+' : Suppression de l''honoraire '+Arg);
{$IFNDEF DADSUSEULE}
   CreeJnalEvt('001', '043', 'OK', NIL, NIL, Trace);
{$ENDIF}   
   if Trace <> nil then
      begin
      Trace.Free;
      Trace := nil;
      end;
   CommitTrans;
   Qsal.prior;
   Arg:=IntToStr(QSal.FindField('PDE_ORDRE').AsInteger);
   MAJQuery;
   SalPrem.Enabled := TRUE;
   SalPrec.Enabled := TRUE;
   SalSuiv.Enabled := TRUE;
   SalDern.Enabled := TRUE;
Except
   Rollback;
   PGIBox ('Une erreur est survenue lors de la mise à jour de la base', 'DADS-U');
   END;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 18/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PG_DADSHONOR
*****************************************************************}
procedure TOF_PG_DADSHONOR.New(Sender: TObject);
begin
MetABlanc;
State := 'CREATION';
ChargeZones;
// Gestion du navigateur
if SalPrem <> NIL then
   SalPrem.Enabled := False;

if SalPrec <> NIL then
   SalPrec.Enabled := False;

if SalSuiv <> NIL then
   SalSuiv.Enabled := False;

if SalDern <> NIL then
   SalDern.Enabled := False;

TFVierge(Ecran).Binsert.Enabled := False;
TFVierge(Ecran).BDelete.Enabled := False;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 18/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PG_DADSHONOR
*****************************************************************}
procedure TOF_PG_DADSHONOR.AccesMontant(Sender: TObject);
begin
If Sender=NIL Then Exit;
If Sender=TypeRem1 Then
   begin
     If GetControlText('CTYPEREM1')='' Then
        begin
        SetControlEnabled('EMONTREM1', False);
        SetControlText('EMONTREM1', '');
        SetControlEnabled('CTYPEREM2', False);
        SetControlText('CTYPEREM2', '');
        end
     Else
        begin
        SetControlEnabled('EMONTREM1', True);
        SetControlEnabled('CTYPEREM2', True);
         end;
   end;
If Sender=TypeRem2 Then
   begin
     If GetControlText('CTYPEREM2')='' Then
        begin
        SetControlEnabled('EMONTREM2', False);
        SetControlText('EMONTREM2', '');
        SetControlEnabled('CTYPEREM3', False);
        SetControlText('CTYPEREM3', '');
        end
     Else
        begin
        SetControlEnabled('EMONTREM2', True);
        SetControlEnabled('CTYPEREM3', True);
        end;
   end;
If Sender=TypeRem3 Then
   begin
     If GetControlText('CTYPEREM3')='' Then
        begin
        SetControlEnabled('EMONTREM3', False);
        SetControlText('EMONTREM3', '');
        SetControlEnabled('CTYPEREM4', False);
        SetControlText('CTYPEREM4', '');
        end
     Else
        begin
        SetControlEnabled('EMONTREM3', True);
        SetControlEnabled('CTYPEREM4', True);
        end;
   end;
If Sender=TypeRem4 Then
   begin
     If GetControlText('CTYPEREM4')='' Then
        begin
        SetControlEnabled('EMONTREM4', False);
        SetControlText('EMONTREM4', '');
        SetControlEnabled('CTYPEREM5', False);
        SetControlText('CTYPEREM5', '');
        end
     Else
        begin
        SetControlEnabled('EMONTREM4', True);
        SetControlEnabled('CTYPEREM5', True);
        end;
   end;
If Sender=TypeRem5 Then
   begin
     If GetControlText('CTYPEREM5')='' Then
        begin
        SetControlEnabled('EMONTREM5', False);
        SetControlText('EMONTREM5', '');
        end
     Else
        SetControlEnabled('EMONTREM5', True);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 18/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PG_DADSHONOR
*****************************************************************}
procedure TOF_PG_DADSHONOR.MAJQuery();
var
Req, StPer : string;
QCount : TQuery;
begin
ChoixReq:=True;
Ferme(QSal);
StPer:= 'SELECT COUNT(PDE_SALARIE) AS NBRE'+
        ' FROM DADSPERIODES WHERE'+
        ' PDE_SALARIE LIKE "'+Sal+'%"'+
        ' AND PDE_DATEDEBUT="'+UsDateTime(DebExer)+'"';
QCount:=OpenSql(StPer, True);
NbreTot:='0';
if not QCount.eof then
   NbreTot := IntToStr(QCount.FindField('NBRE').AsInteger);
Ferme(QCount);

StPer:= 'SELECT *'+
        ' FROM DADSPERIODES'+
        ' WHERE PDE_SALARIE LIKE "'+Sal+'%"'+
        ' AND PDE_DATEDEBUT="'+UsDateTime(DebExer)+'"'+
        ' ORDER BY PDE_TYPE,PDE_ORDRE,PDE_DATEFIN';
QSal:=OpenSql(StPer, True);
if not QSal.eof then
   begin
//   QSal.First;  // PortageCWAS
   IF State='MODIFICATION' Then
      begin
      Req:= 'SELECT PDE_ORDRE,PDE_SALARIE'+
            ' FROM DADSPERIODES WHERE'+
            ' PDE_SALARIE LIKE "'+Sal+'%" AND'+
            ' PDE_ORDRE='+Arg+'';
      If ExisteSQL(Req) Then
         begin
         While (IntToStr(QSal.FindField('PDE_ORDRE').AsInteger)<>Arg) do
               QSal.Next;
         Arg:=IntToStr(QSal.FindField('PDE_ORDRE').AsInteger);
         ChargeZones;
         end
      Else
         begin
         Arg:=IntToStr(QSal.FindField('PDE_ORDRE').AsInteger);
         ChargeZones;
         end;
      end;
   If State='CREATION' Then
      begin
      While (IntToStr(QSal.FindField('PDE_ORDRE').AsInteger)<>Arg) do
            QSal.Next;
      State:='MODIFICATION';
      Arg:=IntToStr(QSal.FindField('PDE_ORDRE').AsInteger);
      ChargeZones;
      end;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 18/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PG_DADSHONOR
*****************************************************************}
function TOF_PG_DADSHONOR.Controles() : boolean;
var
BufDest, BufOrig, Mess, Mess2, Mess3, SIREN, StEtab : String;
ControleChamp, REMOK :Boolean;  // pt32
buffer : string;
ErreurDADSU : TControle;
QRechEtab : TQuery;
i, j : integer ; // pt32
begin
Result:= True;
ControleChamp:= True;
Mess:= '';      //Vous devez renseigner
Mess2:= '';     //Contrôle conformité
Mess3:= '';     //Incohérence
ErreurDADSU.Salarie:= Sal+Arg;
ErreurDADSU.TypeD:= TypeD;
ErreurDADSU.Num:= NumOrdre;
ErreurDADSU.DateDeb:= DebExer;
ErreurDADSU.DateFin:= FinExer;
ErreurDADSU.Exercice:= PGExercice;

// Champs obligatoires :
If ((GetControlText ('ERAISONSOC')='') and (GetControlText ('ENOM')='')) Then
   begin
   result:= FALSE;
   Mess:= Mess+ '#13#10 - la raison sociale du bénéficiaire est obligatoire #13#10'+
                'pour une personne morale.#13#10'+
                'Le nom du bénéficiaire est obligatoire #13#10'+
                'pour une personne physique.#13#10';
   SetFocusControl ('ENOM');
   ErreurDADSU.Segment:= 'S70.G01.00.001';
   ErreurDADSU.Explication:= 'La raison sociale du bénéficiaire est obligatoire'+
                             ' pour une personne morale. Le nom du bénéficiaire'+
                             ' est obligatoire pour une personne physique';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur (ErreurDADSU);
   end;

If (GetControlText ('ERAISONSOC')<>'') then
   begin
   if (GetControlText ('ENOM')<>'') Then
      begin
      Mess3:= Mess3+ 'La raison sociale du bénéficiaire et le nom du#13#10'+
                     'bénéficiaire ne peuvent pas être renseignés #13#10'+
                     'simultanément. La raison sociale étant renseignée,#13#10'+
                     'le nom et le prénom sont effacés.';
      SetControlText ('ENOM', '');
      SetControlText ('EPRENOM', '');
      ErreurDADSU.Segment:= 'S70.G01.00.002.001';
      ErreurDADSU.Explication:= 'Le nom du bénéficiaire a été effacé';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur (ErreurDADSU);
      end;

   if ((GetControlText ('ESIRET')='') and ((GetControlText ('CADRPAYS')='') or
      (GetControlText ('CADRPAYS')='FRA'))) then
      begin
      Result:= False;
      Mess:= Mess+ '#13#10 - le siret du bénéficiaire';
      SetFocusControl ('ESIRET');
      ErreurDADSU.Segment:= 'S70.G01.00.003.001';
      ErreurDADSU.Explication:= 'Le siret du bénéficiaire n''est pas renseigné';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur (ErreurDADSU);
      end;
   end;

If ((GetControlText ('ENOM')<>'') and (GetControlText ('EPRENOM')='')) Then
   begin
   result:= FALSE;
   Mess:= Mess+ '#13#10 - le nom et le prénom du bénéficiaire sont obligatoires #13#10'+
                'pour une personne physique.#13#10';
   SetFocusControl ('EPRENOM');
   ErreurDADSU.Segment:= 'S70.G01.00.002.002';
   ErreurDADSU.Explication:= 'Le prénom du bénéficiaire n''est pas renseigné';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur (ErreurDADSU);
   end;

If (GetControlText ('EADRCP')='') Then
   begin
   Result:= False;
   Mess:= Mess+ '#13#10 - le code postal';
   SetFocusControl ('EADRCP');
   ErreurDADSU.Segment:= 'S70.G01.00.004.010';
   ErreurDADSU.Explication:= 'Le code postal n''est pas renseigné';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur (ErreurDADSU);
   end;

If (GetControlText ('EPROFESSION')='') Then
   begin
   Result:= False;
   Mess:= Mess+ '#13#10 - la profession du bénéficiaire';
   SetFocusControl ('EPROFESSION');
   ErreurDADSU.Segment:= 'S70.G01.00.005';
   ErreurDADSU.Explication:= 'La profession du bénéficiaire n''est pas renseignée';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur (ErreurDADSU);
   end;

if ((((Nourriture<>nil) and (Nourriture.Checked=True)) or
   ((Logement<>nil) and (Logement.Checked=True)) or
   ((Voiture<>nil) and (Voiture.Checked=True)) or
   ((Autres<>nil) and (Autres.Checked=True)) or
   ((NTIC<>nil) and (NTIC.Checked=True))) and
   (GetControlText ('CTYPEREM1')<>'10') and
   (GetControlText ('CTYPEREM2')<>'10') and
   (GetControlText ('CTYPEREM3')<>'10') and
   (GetControlText ('CTYPEREM4')<>'10') and
   (GetControlText ('CTYPEREM5')<>'10')) then
   begin
   Result:= False;
   Mess:= Mess+ '#13#10 - le type de rémunération "Avantages en nature"';
   SetFocusControl ('CTYPEREM1');
   ErreurDADSU.Segment:= 'S70.G01.01.001';
   ErreurDADSU.Explication:= 'Le type de rémunération "Avantages en nature"'+
                             ' n''est pas renseigné';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur (ErreurDADSU);
   end;

if ((Nourriture<>nil) and (Nourriture.Checked=False) and
   (Logement<>nil) and (Logement.Checked=False) and
   (Voiture<>nil) and (Voiture.Checked=False) and
   (Autres<>nil) and (Autres.Checked=False) and
   (NTIC<>nil) and (NTIC.Checked=False) and
   ((GetControlText ('CTYPEREM1')='10') or
   (GetControlText ('CTYPEREM2')='10') or
   (GetControlText ('CTYPEREM3')='10') or
   (GetControlText ('CTYPEREM4')='10') or
   (GetControlText ('CTYPEREM5')='10'))) then
   begin
   Result:= False;
   Mess:= Mess+ '#13#10 - Une coche "Avantages en nature"';
   SetFocusControl ('THNOURRITURE');
   ErreurDADSU.Segment:= 'S70.G01.00.006';
   ErreurDADSU.Explication:= 'Aucune coche "Avantages en nature" n''est'+
                             ' renseignée';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur (ErreurDADSU);
   end;

  //Deb pt32 : controle que 2 types identiques de rem ne peuvent pas être saisis
   REMOK := true;
   for i := 1 to 4 do
   begin
     if  getcontroltext('CTYPEREM' + inttostr(i)) <> '' then
     begin
      for j := i+1 to 5 do
      begin
        if getcontroltext('CTYPEREM' + inttostr(i)) = getcontroltext('CTYPEREM' + Inttostr(j)) then
        begin
          REMOK := false;
          Result:= False;
          ControleChamp:= False;
          Mess2:= Mess2 + 'On peut pas avoir 2 types de rémunération identiques.';
          SetFocusControl ('CTYPEREM' + inttostr(i));
          ErreurDADSU.Segment:= 'S70.G01.01.001';
          ErreurDADSU.Explication:= 'On peut pas avoir 2 types de rémunération identiques';
          ErreurDADSU.CtrlBloquant:= True;
          EcrireErreur (ErreurDADSU);
        end;
        if not REMOK then break;
       end;
      end;
    if not REMOK then break;
   end;
  // fin PT32

If (GetControlText ('CETAB')='') Then
   begin
   Result:= False;
   Mess:= Mess+ '#13#10 - l''établissement';
   SetFocusControl('CETAB');
   ErreurDADSU.Segment:= 'S70.G01.00.014';
   ErreurDADSU.Explication:= 'L''établissement n''est pas renseigné';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur (ErreurDADSU);
   end;

If (GetControlText ('CTYPEREM1')='') Then
   begin
   Result:= False;
   Mess:= Mess+ '#13#10 - le premier code rémunération';
   SetFocusControl ('CTYPEREM1');
   ErreurDADSU.Segment:= 'S70.G01.01.001';
   ErreurDADSU.Explication:= 'Le premier code rémunération n''est pas renseigné';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur (ErreurDADSU);
   end;

{PT33
If (GetControlText ('EMONTREM1')='') Then
}
If ((GetControlText ('EMONTREM1')='') or
   (GetControlText ('EMONTREM1')='0')) Then
//FIN PT33
   begin
   Result:= False;
   Mess:= Mess+ '#13#10 - le premier montant';
   SetFocusControl ('EMONTREM1');
   ErreurDADSU.Segment:= 'S70.G01.01.002';
   ErreurDADSU.Explication:= 'Le premier montant n''est pas renseigné';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur (ErreurDADSU);
   end;

If (GetControlText ('EADRBURDISTRIB')='') Then
   begin
   Result:= False;
   Mess:= Mess+ '#13#10 - le bureau distributeur';
   ErreurDADSU.Segment:= 'S70.G01.00.004.012';
   ErreurDADSU.Explication:= 'Le bureau distributeur n''est pas renseigné';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur (ErreurDADSU);
   end;

if (Mess3 <> '') then
   PGIBox (Mess3,'Incohérence');

if (not(IsValidDate (GetControlText ('EEXERCOMPTA')))) then
   begin
   result:= False;
   Mess3:= 'La date de clôture de l''exercice comptable n''est pas valide';
   PGIBox (Mess3,'Contrôle conformité');
   ErreurDADSU.Segment:= 'S70.G01.00.015';
   ErreurDADSU.Explication:= 'La date de clôture de l''exercice comptable'+
                             ' n''est pas valide';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur (ErreurDADSU);
   end;

if (result = FALSE)  and (mess <> '') then                      // pt32
   PGIBox ('Vous devez renseigner :'+Mess,'Informations obligatoires');

// Formats des champs
If (GetControlText ('EADRCP')<>'') Then
   begin
   If ((Not (IsNumeric (GetControlText ('EADRCP')))) and
      ((GetControlText ('CADRPAYS')='FRA') or
      (GetControlText ('CADRPAYS')=''))) then
      begin
      result:= False;
      ControleChamp:= False;
      mess2:= mess2+ '#13#10 Le code postal doit être numérique';
      ErreurDADSU.Segment:= 'S70.G01.00.004.010';
      ErreurDADSU.Explication:= 'Le code postal doit être numérique';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur (ErreurDADSU);
      end
   Else
      begin
      if ((GetControlText ('CADRPAYS')='FRA') or
         (GetControlText ('CADRPAYS')='')) then
         begin
         If StrToInt (GetControlText ('EADRCP'))=0 Then
            begin
            result:= False;
            ControleChamp:= False;
            mess2:= mess2+ '#13#10 Le code postal ne peut pas avoir la valeur 0';
            ErreurDADSU.Segment:= 'S70.G01.00.004.010';
            ErreurDADSU.Explication:= 'Le code postal ne peut pas avoir la'+
                                      ' valeur 0';
            ErreurDADSU.CtrlBloquant:= True;
            EcrireErreur (ErreurDADSU);
            end
         Else
            begin
            If (length (GetControlText ('EADRCP'))<>5) Then
               begin
               result:= False;
               ControleChamp:= False;
               mess2:= mess2+ '#13#10 Le code postal doit comporter 5 caractères';
               ErreurDADSU.Segment:= 'S70.G01.00.004.010';
               ErreurDADSU.Explication:= 'Le code postal doit comporter 5'+
                                         ' caractères';
               ErreurDADSU.CtrlBloquant:= True;
               EcrireErreur (ErreurDADSU);
               end;
            end;
         end;
      end;
   end;

If (GetControlText ('ESIRET')<>'')  then
   begin
   If Length (GetControlText ('ESIRET'))<>14 Then
      begin
      mess2:= mess2+ '#13#10 Le SIRET est incomplet (14 caractères obligatoires)';
      result:= False;
      ControleChamp:= False;
      ErreurDADSU.Segment:= 'S70.G01.00.003.001';
      ErreurDADSU.Explication:= 'Le SIRET est incomplet (14 caractères'+
                                ' obligatoires)';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur (ErreurDADSU);
      end
   Else
      begin
      If ControlSiret (GetControlText ('ESIRET'))=False Then
         begin
         result:= False;
         ControleChamp:= False;
         mess2:= mess2+ '#13#10 Le SIRET n''est pas valide';
         ErreurDADSU.Segment:= 'S70.G01.00.003.001';
         ErreurDADSU.Explication:= 'Le SIRET n''est pas valide';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur (ErreurDADSU);
         end;
      end;
   end;

StEtab := 'SELECT ET_SIRET'+
          ' FROM ETABLISS WHERE'+
          ' ET_ETABLISSEMENT = "'+GetControlText ('CETAB')+'"';
QRechEtab:=OpenSQL (StEtab,TRUE);
if (not QRechEtab.EOF) then
   BufOrig:= QRechEtab.FindField ('ET_SIRET').Asstring;
Ferme (QRechEtab);
ForceNumerique (BufOrig, BufDest);
if ControlSiret (BufDest)=False then
   begin
   Result:=False;
   ControleChamp:= False;
   Mess2:= Mess2+ '#13#10 Le SIRET de l''établissement n''est pas valide';
   SetFocusControl ('CETAB');
   ErreurDADSU.Segment:= 'S70.G01.00.014';
   ErreurDADSU.Explication:= 'Le SIRET de l''établissement n''est pas valide';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur (ErreurDADSU);
   end;
ForceNumerique (GetParamSoc ('SO_SIRET'), SIREN);
if (Copy (BufDest, 1, 9)<> Copy (SIREN, 1, 9)) then
   begin
   Result:=False;
   ControleChamp:= False;
   Mess2:= Mess2+ '#13#10 Le SIREN de l''établissement n''est pas valide';
   SetFocusControl ('CETAB');
   ErreurDADSU.Segment:= 'S70.G01.00.014';
   ErreurDADSU.Explication:= 'Le SIREN de l''établissement '+
                             GetControlText ('CETAB')+' ne correspond pas au'+
                             ' SIREN de l''entreprise';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur (ErreurDADSU);
   end;

If (length(GetControlText ('EADRINSEE'))<>5) and
   (GetControlText ('EADRINSEE')<>'') Then
   begin
   result:= False;
   ControleChamp:= False;
   mess2:= mess2+ '#13#10 Le code INSEE doit comporter 5 caractères ou être nul';
   ErreurDADSU.Segment:= 'S70.G01.00.004.007';
   ErreurDADSU.Explication:= 'Le code INSEE doit comporter 5 caractères ou'+
                             ' être vide';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur (ErreurDADSU);
   end;

buffer:= GetControlText ('MCCODEINDEMNITES');
buffer:= StringReplace (buffer, ';', '', [rfReplaceAll]);
If ((buffer<>'FPR') and (buffer<>'FR') and (buffer<>'FP') and (buffer<>'PR') and
   (buffer<>'F') and (buffer<>'R') and (buffer<>'P') and (buffer<>'')) Then
   begin
   result:= False;
   ControleChamp:= False;
   mess2:= mess2+ '#13#10 Le code modalité de prise en charge des indemnités'+
                  ' n''est pas valide';
   ErreurDADSU.Segment:= 'S70.G01.00.010';
   ErreurDADSU.Explication:= 'Le code modalité de prise en charge des'+
                             ' indemnités n''est pas valide';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur (ErreurDADSU);
   end
else
   begin
   If (GetControlText ('MCCODEINDEMNITES')<>'') and
      (GetControlText ('CTYPEREM1')<>'09') and
      (GetControlText ('CTYPEREM2')<>'09') and
      (GetControlText ('CTYPEREM3')<>'09') and
      (GetControlText ('CTYPEREM4')<>'09') and
      (GetControlText ('CTYPEREM5')<>'09') then
      begin
      result:= False;
      ControleChamp:= False;
      Mess2:= Mess2+ '#13#10 Le type de rémunération "Indemnités et'+
                     ' remboursements" n''est pas renseigné';
      ErreurDADSU.Segment:= 'S70.G01.00.010';
      ErreurDADSU.Explication:= 'Le type de rémunération "Indemnités et'+
                                ' remboursements" n''est pas renseigné';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur (ErreurDADSU);
      end;

   if ((GetControlText ('MCCODEINDEMNITES')= '') and
      ((GetControlText ('CTYPEREM1')='09') or
      (GetControlText ('CTYPEREM2')='09') or
      (GetControlText ('CTYPEREM3')='09') or
      (GetControlText ('CTYPEREM4')='09') or
      (GetControlText ('CTYPEREM5')='09'))) then
      begin
      Result:= False;
      ControleChamp:= False;
      Mess2:= Mess2+ '#13#10 Le code modalités de prise en charges des'+
                     ' indemnités n''est pas renseigné';
      SetFocusControl ('MCCODEINDEMNITES');
      ErreurDADSU.Segment:= 'S70.G01.00.010';
      ErreurDADSU.Explication:= 'Le "code modalités de prise en charges des'+
                                ' indemnités" n''est pas renseigné';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur (ErreurDADSU);
      end;
   end;

If ControleChamp=False Then
   PGIBox (Mess2,'Contrôle conformité');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 06/01/2005
Modifié le ... :   /  /
Description .. : Complément des raccourcis claviers
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSHONOR.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
case Key of
     VK_F10: if ((GetControlVisible('BDACCORD')) and
               (GetControlEnabled('BDACCORD'))) then
               BDaccord.Click; //Validation
     end;
end;

Initialization
registerclasses ([TOF_PG_DADSHONOR]);
end.
