{***********UNITE*************************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 02/04/2007
Modifié le ... : 23/05/2007
Description .. : Source TOF de la FICHE : AMREMPLACEMENT ()
Suite ........ : BTY FQ 20256 Mettre le type de dérogatoire pour l'immo fille
Suite ........ : MBO 08/06/2007 - stockage de X ds I_STRING1 pour identifier un composant remplaçant
Suite ........ : MBO 18/06/2007 - stockage du montant ht dans i_valeurAchat
Suite ........ : MBO 10/07/2007 - bloquer l'opération si dépréciation sur l'exercice
Suite ........ : MBO 12/07/2007 - controle que le code immo est bien alimenté sur le onexit
Suite ........ : MBO 29/10/2007 - fq 21754 - appel à plan.traitesortie avec 2 nveaux paramètres pour calcul cession
                                  appel à planInfo.calcul avec 2 nveaux paramètres
                                  attention : en remplacement on calcule toujours la cession y compris le jour de cession
Mots clefs ... : TOF;AMREMPLACEMENT
*****************************************************************}
Unit AMREMPLACEMENT_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     maineagl,
{$ENDIF}
     ParamSoc,
     ImEnt,
     ImPlan,
     ImPlanInfo,
     forms,
     Outils,
     sysutils,
     ComCtrls,
     Dialogs,
     HCtrls,
     HEnt1,
     HTB97,
     HMsgBox,
     HRichOle,
     UTOM,
     UTOB,
     UTOF,
     ImOuPlan,
     Math,
     UtobDebug,
     Vierge;

Type
  TOF_AMREMPLACEMENT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure OnAfterFormShowEcran;
  private
    TobImmoMere    : TOB;
    TobLogMere     : TOB;
    PlanMere       : TPlanAmort;

    TobImmoFille    : TOB;
    TobLogFille     : TOB;
    PlanFille       : TPlanAmort;

    fblocnote: THRichEditOLE;
    DateDebEcoMere : tDateTime;
    QuantiteMere : double;
    TPnulle : boolean;
    CodeImmoMere : string;
    PlanActifMere : integer;

    procedure OnExitMontantHT(Sender: TObject);
    procedure OnExitTVA1(Sender: TObject);
    procedure OnExitTVA2(Sender: TObject);
    procedure DureeDejaAmortie;
    procedure OnClickPlafondDeductibilite(Sender: TObject);
    procedure OnexitDuree(Sender: TObject);
    procedure OnExitMethode(Sender: TObject);
    procedure OnbFermeClick(Sender : TObject);
    procedure OnExitDateOp(Sender : TObject);
    procedure CalculVNC;
    Function ControleMontantHT : integer;
    Function ControleTVA1 : integer;
    Function ControleTVA2 : integer;
    Function ControleCodeImmo : integer;
    procedure EnregistreRemplacement;
  end ;

function AMLanceFiche_FicheRemplacement(Stcode: string): Integer ;

Implementation



const
 HM2: array[0..19] of string =
 {0} ('%d mois - %d jour(s)',
 {1}  'La date de remplacement doit appartenir à l''exercice en cours.',
 {2}  'La date de remplacement doit être supérieure ou égale à la plus haute date d''opération déjà enregistrée.',
 {3}  'La base d''amortissement économique doit être inférieure ou égale au montant TTC.',
 {4}  'Le code immobilisation existe déjà.',
 {5}  'Vous devez choisir une méthode d''amortissement.',
 {6}  'La date de remplacement doit être supérieure à la date de début d''amortissement économique de l''immobilisation remplacée.',
 {7} 'Voulez-vous enregistrer cette opération ?',
 {8} 'L''enregistrement de l''opération a échoué.',
 {9} 'Vous devez renseigner une quantité inférieure ou égale à la quantité d''origine et non nulle.',
 {10}'Le montant de la TVA récupérée ne peut en aucun cas être supérieur au montant de la TVA récupérable.',
 {11}'Le montant HT doit être supérieur à 0.',
 {12}'Veuillez saisir un montant positif ou nul.',
 {13}'En mode dégressif, la durée d''amortissement doit être supérieure ou égale à 36 mois.',
 {14}'La durée d''amortissement dépasse 180 mois.'
      + #10#13'Voulez-vous modifier cette valeur ?',
 {15}'Vous devez renseigner une quantité inférieure ou égale à la quantité de l''immobilisation remplacée et non nulle.',
 {16}'Le plafond de déductibilité fiscale doit être inférieur ou égal à la base d''amortissement économique.',
 {17}'%d mois',
 {18}'Opération non disponible car cette immobilisation a fait l''objet d''une dépréciation d''actif',
 {19}'Le code immobilisation est obligatoire.'
      );

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 02/04/2007
Modifié le ... :   /  /
Description .. : Lancement de la fiche sous certaines conditions
Mots clefs ... :
*****************************************************************}
function AMLanceFiche_FicheRemplacement(Stcode: string): Integer;
begin
  result := 0;
   // Ctl présence d'une dépréciation dans la vie de l'immo
   // opération de dépréciation sur N
   if ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+
                 Stcode+'" AND (I_REVISIONECO<>0)') then
      PGIBox(HM2[18])
   else
   begin
      //opération de dépréciation sur exo antérieurs cloturés (montant maxi reprise deprec)
      if ExisteSQL ('SELECT IL_IMMO FROM IMMOLOG WHERE IL_IMMO="'+
         Stcode+'" AND IL_TYPEOP="CLO" AND IL_REVISIONECO<>0 AND IL_DATEOP = "'+
         USDATETIME((VHImmo^.EnCours.deb)-1)+'" ') then

         PGIBox(HM2[18])
      else
      begin
         // antérieurs dépréciation saisis
         if ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+
                       Stcode+'" AND (I_REPRISEDEP<>0)') then
            PGIBox(HM2[18])
         else
            result:=StrToInt(AGLLanceFiche('AM', 'AMREMPLACEMENT', '', '', Stcode));
      end;
   end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 02/04/2007
Modifié le ... :   /  /
Description .. : On argument : init de la fenêtre avec alimentation
Suite ........ : des controles
Mots clefs ... :
*****************************************************************}
procedure TOF_AMREMPLACEMENT.OnArgument (S : String ) ;
var
  QPlan : TQuery;

begin
  Inherited ;

  { Evènements }
  TCheckBox(GetControl('RPL_CKPLAF')).OnClick := OnClickPlafondDeductibilite;
  THEdit(GetControl('RPL_METHODE')).OnExit := OnExitMethode;

  THEdit(GetControl('RPL_HT')).OnExit := OnExitMontantHT;
  THEdit(GetControl('RPL_TVA1')).OnExit := OnExitTVA1;
  THEdit(GetControl('RPL_TVA2')).OnExit := OnExitTVA2;

  THEdit(GetControl('RPL_DUREE')).OnExit := OnExitDuree;

  THEdit(GetControl('RPL_DATEOP')).OnExit := OnExitDateOp;

  TToolbarButton97(GetControl('BFerme')).Onclick := OnbFermeClick;

  TFVIERGE(ECRAN).retour := IntToStr(MrNo);
  TFVierge(Ecran).OnAfterFormShow := OnAfterFormShowEcran;

  // Assignation du bloc note
  fblocnote := THRichEditOLE(GetControl('RPL_BN'));

  // Lecture de l'immo en cours
  QPlan:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+S+'"', FALSE) ;

  DateDebEcoMere := StrToDate(QPlan.FindField('I_DATEDEBECO').AsString);
  QuantiteMere   := QPlan.FindField('I_QUANTITE').AsInteger;
  CodeImmoMere   := QPlan.FindField('I_IMMO').AsString;
  PlanActifMere := QPlan.FindField('I_PLANACTIF').AsInteger;

  SetControlCaption('RPL_CODE_OLD',QPlan.FindField('I_IMMO').AsString);
  SetControlCaption('RPL_DESI',QPlan.FindField('I_LIBELLE').AsString);
  SetControlCaption('RPL_DATEACHAT',QPlan.FindField('I_DATEPIECEA').AsString);
  SetControlCaption('RPL_VALACHAT',FloatToStr(QPlan.FindField('I_MONTANTHT').AsFloat));
  SetControlCaption('RPL_DEBECO', QPlan.FindField('I_DATEDEBECO').AsString);
  SetControlCaption('RPL_CPTE',QPlan.FindField('I_COMPTEIMMO').AsString);
  SetControlCaption('RPL_QTE_OLD',IntToStr(QPlan.FindField('I_QUANTITE').AsInteger));
  SetControlCaption('RPL_METH_OLD',RechDom('TIMETHODEIMMO',QPlan.FindField('I_METHODEECO').AsString,False));
  SetControlCaption('RPL_OLD_TAUX',FormatFloat('#,##0.00',QPlan.FindField('I_TAUXECO').AsFloat));
  SetControlCaption('RPL_DUREE_OLD', Format(HM2[17], [QPlan.FindField('I_DUREEECO').AsInteger]));

  SetControlText('RPL_DUREEAMORTIE', '');
  //calcul durée déjà amortie
  DureeDejaAmortie;

  SetControlText('RPL_VRC', '');

  THValComboBox(GetControl('RPL_METHODE')).Items.Delete(THValComboBox(GetControl('RPL_METHODE')).Items.IndexOf('Variable'));
  THValComboBox(GetControl('RPL_METHODE')).Items.Delete(THValComboBox(GetControl('RPL_METHODE')).Items.IndexOf('Non amortissable'));

  SetControlEnabled('RPL_DATEOP',True);
  SetControlEnabled('RPL_CODEIMMO',True);
  SetControlEnabled('RPL_QUANTITE',True);
  SetControlEnabled('RPL_METHODE',True);
  SetControlEnabled('RPL_DUREE',true);
  SetControlEnabled('RPL_TAUX',False);
  SetControlEnabled('RPL_TVA1',true);
  SetControlEnabled('RPL_TVA2',true);
  SetControlEnabled('RPL_TTC',False);
  SetControlEnabled('RPL_BASE',true);
  SetControlEnabled('RPL_PLAF',False);

  SetControlText('RPL_METHODE','LIN');

  //methodeMere := QPlan.FindField('I_METHODEECO').AsString;
  //montanthtMere := QPlan.FindField('I_MONTANTHT').AsFloat;
  quantiteMere := QPlan.FindField('I_QUANTITE').AsInteger;

  SetControlText('RPL_QTE', '1');
  SetControlText('RPL_CODE', NouveauCodeImmo);

  // pas de base TP sur radicaux 20, 26 et 27
  TPnulle := false;

  if (COPY(GetControlText('RPL_CPTE'),1, 2) = '20') or
     (COPY(GetControlText('RPL_CPTE'),1, 2) = '26') or
     (COPY(GetControlText('RPL_CPTE'),1, 2) = '27') then
  begin
     SetControlText('RPL_TP','0');
     SetControlEnabled('RPL_TP', false);
     TPnulle := true;
  end;

  SetFocusControl('RPL_CPDESI');
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 03/04/2007
Modifié le ... :   /  /
Description .. : Pour masque
Mots clefs ... :
*****************************************************************}
procedure TOF_AMREMPLACEMENT.OnAfterFormShowEcran;
begin
  SetControlProperty('RPL_VALACHAT', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('RPL_HT', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('RPL_TVA1', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('RPL_TVA2', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('RPL_TTC', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('RPL_PLAF', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('RPL_TP', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('RPL_BASE', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));

end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 02/04/2007
Modifié le ... :   /  /
Description .. : Calcul de la durée déjà amortie entre date début
Suite ........ : d'amortissement éco et fin exercice n-1
Mots clefs ... :
*****************************************************************}
procedure TOF_AMREMPLACEMENT.DureeDejaAmortie;
var
  PremMois, PremAnnee, NbMois: Word;
  nJour: integer;
  date_immo: TdateTime;

begin
  //si méthode est dégressif : différence entre date d'achat et date fin ex.précédent
  //sinon                    : différence entre date de mise en service et date fin ex.précédent
  date_immo := StrToDate(GetControlText('RPL_DEBECO')) ;

  if (date_immo >= VHImmo^.Encours.Deb) or (date_immo = iDate1900) then
  begin
    SetControlText('RPL_DUREEAMORTIE', '');
    SetControlText('RPL_LIDUREEAMORTIE', '');
    exit;
  end;

  NOMBREMOIS(date_immo, VHImmo^.Encours.Deb, PremMois, PremAnnee,NbMois);
  nJour := NombreJour360(date_immo,VHImmo^.Encours.Deb-1);

  if nJour < 0 then nJour := 0;

  SetControlCaption('RPL_DUREEAMORTIE', Format(HM2[0], [NbMois - 1, nJour]));
end;



{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 02/04/2007
Modifié le ... :   /  /
Description .. : Procedure appelée en modif montant HT ou TVA
Suite ........ : récupérable ou TVA récupérée pour alimentation du
Suite ........ : montant TTC et de la base TP et la base ECO
Mots clefs ... :
*****************************************************************}
Procedure TOF_AMREMPLACEMENT.OnExitMontantHT(Sender: TObject);
var verifMontant : integer;
    montantTTC : double;
begin
      VerifMontant := ControleMontantHT;
      if VerifMontant <> 0 then
      begin
         ecran.ModalResult := mrNone;
         SetFocusCOntrol('RPL_HT');
      end else
      begin
         MontantTTC := Valeur(GetControlText('RPL_HT')) + Valeur(GetControlText('RPL_TVA1')) -
                       Valeur(GetControlText('RPL_TVA2'));

         SetControlText('RPL_TTC', StrFMontant(MontantTTC,15,V_PGI.OkDecV,'',True));

         if not(TPnulle) then SetControlText('RPL_TP', StrFMontant(MontantTTC,15,V_PGI.OkDecV,'',True));

         SetControlText('RPL_BASE', StrFMontant(MontantTTC,15,V_PGI.OkDecV,'',True));
      end;
end;

//------------------------------------------------------------------------

Procedure TOF_AMREMPLACEMENT.OnExitTVA1(Sender: TObject);
var verifMontant : integer;
    montantTTC : double;
begin
      VerifMontant := ControleTVA1;
      if VerifMontant <> 0 then
      begin
         ecran.ModalResult := mrNone;
         SetFocusControl('RPL_TVA1');
      end else
      begin
         MontantTTC := Valeur(GetControlText('RPL_HT')) + Valeur(GetControlText('RPL_TVA1')) -
                       Valeur(GetControlText('RPL_TVA2'));

         SetControlText('RPL_TTC', StrFMontant(MontantTTC,15,V_PGI.OkDecV,'',True));

         if not(TPnulle) then SetControlText('RPL_TP', StrFMontant(MontantTTC,15,V_PGI.OkDecV,'',True));

         SetControlText('RPL_BASE', StrFMontant(MontantTTC,15,V_PGI.OkDecV,'',True));
      end;
end;

//------------------------------------------------------------------------

Procedure TOF_AMREMPLACEMENT.OnExitTVA2(Sender: TObject);
var verifMontant : integer;
    montantTTC : double;
begin
      VerifMontant := ControleTVA2;
      if VerifMontant <> 0 then
      begin
         ecran.ModalResult := mrNone;
         SetFocusControl('RPL_TVA2');
      end else
      begin
         MontantTTC := Valeur(GetControlText('RPL_HT')) + Valeur(GetControlText('RPL_TVA1')) -
                       Valeur(GetControlText('RPL_TVA2'));

         SetControlText('RPL_TTC', StrFMontant(MontantTTC,15,V_PGI.OkDecV,'',True));

         if not(TPnulle) then SetControlText('RPL_TP', StrFMontant(MontantTTC,15,V_PGI.OkDecV,'',True));

         SetControlText('RPL_BASE', StrFMontant(MontantTTC,15,V_PGI.OkDecV,'',True));
      end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 12/07/2007
Modifié le ... :   /  /
Description .. : vérifie que le code est différent de ''
Suite ........ : bourre avec des zéros à gauche
Suite ........ : vérifie qu'il n'existe pas déja
Mots clefs ... :
*****************************************************************}
Function TOF_AMREMPLACEMENT.ControleCodeImmo : integer;
begin
  result:=0;

  //Controle code immo alimenté
  if (GetControlText('RPL_CODE') = '') or (GetControlText('RPL_CODE') = '0000000000') then
  begin
    pgibox(HM2[19]);
    result:=1;
  end else
  begin
     //Contrôle Code immobilisation inexistant
     if ExisteSQL('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+GetControlText('RPL_CODE')+'"') then
     begin
       pgibox(HM2[4]);
       result:=1;
     end;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 05/04/2007
Modifié le ... :   /  /
Description .. : Controle montant HT , tva récupérable, tva récupérée
Mots clefs ... :
*****************************************************************}
Function TOF_AMREMPLACEMENT.ControleMontantHT : integer;
//Var
  //MontantTTC : double;

begin
    if Valeur(GetControlText('RPL_HT'))<= 0.00 then
    begin
       pgibox(HM2[11]);         // montantHT > 0
       result:= 1;
    end else
       result := 0;
end;

//------------------------------------------------------------------------

Function TOF_AMREMPLACEMENT.ControleTVA1 : integer;
begin
   if (Valeur(GetControlText('RPL_TVA1'))< 0.00) then
   begin
      pgibox(HM2[12]);         // montantTVA positif non nul
      Result := 1;
   end else
      result := 0;
end;

//------------------------------------------------------------------------

Function TOF_AMREMPLACEMENT.ControleTVA2 : integer;
begin
  if Valeur(GetControlText('RPL_HT'))<= 0.00 then
  begin
     pgibox(HM2[11]);         // montantHT > 0
     result:= 1;
  end else
     result := 0;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 02/04/2007
Modifié le ... :   /  /
Description .. : on exit du control methode : affichage du taux
Mots clefs ... :
*****************************************************************}
procedure TOF_AMREMPLACEMENT.OnExitDuree(Sender: TObject);
var
  taux: double;
begin
  if GetControlText('RPL_METHODE') = '' then
  begin
    pgibox(HM2[5]);
    ecran.ModalResult := mrNone;
    SetFocusControl('RPL_METHODE');
    exit;
  end;

  if ((GetControlText('RPL_METHODE') <> '')
     and (valeur(GetControlText('RPL_DATEOP'))<> 0)
     and (GetControlText('RPL_DUREE')<> '')) then
  begin
    taux := GetTaux(GetControlText('RPL_METHODE')
         , StrToDate(GetControlText('RPL_DATEOP'))
         , StrToDate(GetControlText('RPL_DATEOP'))
         , StrToInt(GetControlText('RPL_DUREE')));
  end
  else taux := 0;

  SetControlText('RPL_TAUX',FormatFloat('#,##0.00',taux));
end;


{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 03/04/2007
Modifié le ... :   /  /    
Description .. : Changement de méthode = recalcul du taux
Mots clefs ... :
*****************************************************************}
procedure TOF_AMREMPLACEMENT.OnExitMethode(Sender: TObject);
var
  taux: double;
begin
  if GetControlText('RPL_METHODE') = '' then
  begin
    pgibox(HM2[5]);
    ecran.ModalResult := mrNone;
    SetFocusControl('RPL_METHODE');
    exit;
  end;

  //Contrôle de la durée mini et maxi pour la méthode dégressive
  if (GetControlText('RPL_METHODE') = 'DEG') then
  begin
     if Valeur(GetControlText('RPL_DUREE')) < 36 then
        SetControlText('RPL_DUREE', '36');
  end;

  if ((GetControlText('RPL_METHODE') <> '')
     and (valeur(GetControlText('RPL_DATEOP'))<> 0)
     and (Valeur(GetControlText('RPL_DUREE'))<>0)) then
  begin
    taux := GetTaux(GetControlText('RPL_METHODE')
         , StrToDate(GetControlText('RPL_DATEOP'))
         , StrToDate(GetControlText('RPL_DATEOP'))
         , StrToInt(GetControlText('RPL_DUREE')));
  end
  else taux := 0;

  SetControlText('RPL_TAUX',FormatFloat('#,##0.00',taux));
end;



{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 02/04/2007
Modifié le ... :   /  /
Description .. : le choix plafond de déductibilité ouvre la zone montant du
Suite ........ : plafond
Mots clefs ... :
*****************************************************************}
procedure TOF_AMREMPLACEMENT.OnClickPlafondDeductibilite(Sender: TObject);
begin
  if (GetCheckBoxState('RPL_CKPLAF') = cbChecked) then
     SetControlEnabled('RPL_PLAF', true)
  else
  begin
     SetControlText('RPL_PLAF', '0');
     SetControlEnabled('RPL_PLAF', false);
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 02/04/2007
Modifié le ... :   /  /
Description .. : la date de remplacement doit être valide car elle devient la
Suite ........ : date de début d'amortissement économique
Mots clefs ... :
*****************************************************************}
Procedure TOF_AMREMPLACEMENT.OnExitDateOp(Sender: TObject);

var
   date_saisie : TdateTime;

begin
  if valeur(GetControlText('RPL_DATEOP'))= 0 then
  begin
    pgibox(HM2[1]);
    ecran.ModalResult := mrNone;
    SetFocusControl('RPL_DATEOP');
    exit;
  end
  else
  begin
    // la date de remplacement doit être supérieure ou égale à la date deb éco de l'immo remplacée
    date_saisie := StrToDate(GetControlText('RPL_DATEOP'));
    if (date_saisie <= DateDebEcoMere) then
    begin
      pgibox(HM2[6]);
      ecran.ModalResult := mrNone;
      SetFocusControl('RPL_DATEOP');
      exit;
    end
    // la date doit appartenir à l'exercice en cours
    else if (date_saisie >  VHImmo^.Encours.Fin) OR ( date_saisie < VHImmo^.Encours.Deb) then
    begin
      pgibox(HM2[1]);
      ecran.ModalResult := mrNone;
      SetFocusControl('RPL_DATEOP');
      exit;
    end
    //la date de remplacement doit être >= date dernière opération
    else if ExisteSQL('SELECT IL_DATEOP FROM IMMOLOG WHERE IL_IMMO="'+GetControlText('RPL_CODE_OLD')+'" AND IL_DATEOP>"'+USDATETIME(date_saisie)+'"') then
    begin
      pgibox(HM2[2]);
      ecran.ModalResult := mrNone;
      SetFocusControl('RPL_DATEOP');
      exit;
    end else
      // la date est ok
      //calcul de la valeur nette comptable à date du remplacement
      CalculVnc;

  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 02/04/2007
Modifié le ... :   /  /
Description .. : Calcul de la Vnc à date de remplacement
Mots clefs ... :
*****************************************************************}
Procedure TOF_AMREMPLACEMENT.CalculVNC;
var
  PlanInfo  : TPlanInfo;
  Plan : TplanAmort;
  QPlan : TQuery;
  DtOpe : TDateTime;
  VncEco : double;
  VRC : double;
  Year,Month,Day,DayFin : Word;
  VNCDebExo : double;
begin

  DtOpe := StrToDate(GetControlText('RPL_DATEOP'));
  DecodeDate(dtOpe,Year,Month,Day);
  DecodeDate(FinDeMois(dtOpe),Year,Month,DayFin);
  if Day = DayFin then Day:=0;


  // Lecture de l'immo en cours
  QPlan:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+CodeImmoMere+'"', FALSE) ;
  Plan := TPlanAmort.Create(true);
  try
     Plan.Charge(Qplan);
     Plan.Recupere(CodeImmoMere,IntToStr(PlanActifMere));

     PlanInfo:=TPlanInfo.Create(CodeImmoMere) ;
     try
       PlanInfo.Plan.copie(Plan);
       PlanInfo.Calcul(dtOpe, True, False, '');     //mbo - 29/10/07 - fq 21754

       VNCDebExo:=Plan.GetVNCAvecMethode(Plan.AmortEco, VHImmo^.Encours.Deb);
       VncEco := VNCDebExo - PlanInfo.DotationEco;

       VRC:= Plan.GetValResiduelle(CodeImmoMere,Plan.Amorteco, VHImmo^.Encours.Deb, false);
       VRC:= VRC - PlanInfo.DotationEco;

       SetControlText('RPL_VRC', StrFMontant(VRC,15,V_PGI.OkDecV,'',True));
       SetControlText('RPL_VNC', StrFMontant(VncEco,15,V_PGI.OkDecV,'',True));
     finally
        PlanInfo.free;
     end;
  finally
    Plan.Free;
    Ferme(Qplan);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 03/04/2007
Modifié le ... :   /  /
Description .. : Fermeture de l'écran de saisie par la croix
Mots clefs ... :
*****************************************************************}
procedure TOF_AMREMPLACEMENT.OnbFermeClick(Sender: TObject);

begin

    // Action Echap ou Fermer par la croix
    case PGIAskCancel(HM2[7],Ecran.Caption) of
      mrYes:
        // contrôle validation et sortie
        TToolbarButton97(GetControl('BValider')).Click;
      mrNo:
        // Sortir
        Ecran.ModalResult := mrNo;
      else
        // Rester dans l'écran
        Ecran.ModalResult := mrNone;
    end;

end;

procedure TOF_AMREMPLACEMENT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AMREMPLACEMENT.OnDelete ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 02/04/2007
Modifié le ... :   /  /
Description .. : l'immo d'origine va être cédée pour la différence entre
Suite ........ : caractéristiques d'origine et caractéristiques du composant
Mots clefs ... :
*****************************************************************}
procedure TOF_AMREMPLACEMENT.OnUpdate ;
var  VerifMontant : integer;
     Taux : double;
begin
  Inherited ;
  ecran.ModalResult := mrYes;

  //Contrôle date de remplacement dans l'exercice et >= à la dernière opération
  //fait dans un onexit

  if ControleCodeImmo = 1 then
  begin
     ecran.ModalResult:=mrNone;
     SetFocusControl('RPL_CODE');
     exit;
  end;

  VerifMontant := ControleMontantHT;
  if VerifMontant <> 0 then
  begin
      ecran.ModalResult := mrNone;
      SetFocusControl('RPL_HT');
      exit;
   end else
   begin
      if ControleTVA1 <>0 then
      begin
         ecran.ModalResult := mrNone;
         SetFocusControl('RPL_TVA1');
         exit;
      end else
      begin
         if controleTVA2 <> 0 then
         begin
            ecran.ModalResult := mrNone;
            SetFocusControl('RPL_TVA2');
            exit;
         end;
      end;
  end;

  // Controle de la base éco <> 0 et inférieure ou égale au montant TTC
  if Valeur(GetControlText('RPL_BASE')) >  Valeur(GetControlText('RPL_TTC')) then
  begin
    pgibox(HM2[3]);
    ecran.ModalResult := mrNone;
    SetFocusControl('RPL_BASE');
    exit;
  end;

  // Controle de la méthode
  if (GetcontrolText('RPL_METHODE') = '') then
  begin
    //pgibox(HM2[5]);
    ecran.ModalResult := mrNone;
    SetFocusControl('RPL_METHODE');
    exit;
  end;

  if (GetControlText('RPL_METHODE') = 'DEG') then
  begin
     if Valeur(GetControlText('RPL_DUREE')) < 36 then
     begin
        pgibox(HM2[13]);
        ecran.ModalResult := mrNone;
        SetFocusControl('RPL_METHODE');
        exit;
     end else
     begin
      if Valeur(GetControlText('RPL_DUREE')) >= 180 then
        begin
           if PGIAsk(HM2[14]) = mrYes then
           begin
              SetFocusControl('RPL_DUREE');
              //result:= -1;
              Ecran.ModalResult := mrNone;
              exit;
           end;

        end;

     end;
   end;

  // on réaffiche le taux au cas où on ne serait pas passé zone à zone
  taux := GetTaux(GetControlText('RPL_METHODE')
         , StrToDate(GetControlText('RPL_DATEOP'))
         , StrToDate(GetControlText('RPL_DATEOP'))
         , StrToInt(GetControlText('RPL_DUREE')));
  SetControlText('RPL_TAUX',FormatFloat('#,##0.00',taux));

  //Contrôle Quantité < quantité origine, si > 1
  if (Valeur(GetControlText('RPL_QTE')) > QuantiteMere) OR
    (Valeur(GetControlText('RPL_QTE')) < 1) then
  begin
    pgibox(HM2[15]);
    ecran.ModalResult := mrNone;
    SetFocusControl('RPL_QTE');
    exit;
  end;

  //contrôle plafond de déductibilité
  if (Valeur(GetControlText('RPL_PLAF')) > Valeur(GetControlText('RPL_BASE'))) then
  begin
    pgibox(HM2[16]);
    ecran.ModalResult := mrNone;
    SetFocusControl('RPL_PLAF');
    exit;
  end;

  EnregistreRemplacement;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 03/04/2007
Modifié le ... :   /  /
Description .. : Opération de sortie de l'immo d'origine et création du composant
Mots clefs ... :
*****************************************************************}
Procedure TOF_AmRemplacement.EnregistreRemplacement;
var
  CumulAntEco, CumulAntFisc : double;
  // 06.06.07 CumulAntDR, CumulAntFEC: double;
  VRFisc, DotCesEco, DotCesFisc : double;
  fDateOp : tDateTime;
  // FinAmortFiscMere : tDateTime;
  QPlan : tQuery;
  Q : tQuery;
  PlanInfo  : TPlanInfo;
  PremMois, PremAnnee, NbMois: Word;
  NewPlanSeq, MoisRestants : integer;
  GestFisc : boolean;
  
begin
  NewPlanSeq := 1;
  fDateOp := StrToDate(GetControlText('RPL_DATEOP'));

  BEGINTRANS ;
  try

    Q := OpenSQL ('SELECT * FROM IMMO WHERE I_IMMO="'+CodeImmoMere+'"',True);
    try
      TobImmoMere := TOB.Create ('IMMO', nil, -1);
      TobImmoMere.SelectDB ('',Q);

      PlanMere := TPlanAmort.Create(True);
      PlanMere.Charge(Q);
      PlanMere.Recupere(CodeImmoMere,TobImmoMere.GetValue('I_PLANACTIF'));

      //PlanMere.CalculDateFinAmortissement(PlanMere.AmortFisc);
      //FinAmortFiscMere := PlanMere.AmortFisc.DateFinAmort;

      PlanInfo:=TPlanInfo.Create(CodeImmoMere) ;
      try
        PlanInfo.Plan.copie(PlanMere);
        PlanInfo.Calcul(fdateOp,true, false, '');  // mbo - 29/10/07 - fq 21754

        //VRFisc := PlanInfo.VRFisc;   // valeur résiduelle fiscale
        DotCesEco := PlanInfo.DotationEco;
        DotCesFisc := PlanInfo.DotationFisc;
        CumulAntEco := PlanInfo.GetCumulAntEco(fDateOp);
        CumulAntFisc := PlanInfo.GetCumulAntFisc(fDateOp);
        VRFisc := PlanInfo.BaseFisc - (CumulAntFisc + DotCesFisc);   // valeur résiduelle fiscale
        if VRFisc < 0 then VRFisc := 0;
      finally
        PlanInfo.free;
      end;


      //06.06.07 PlanMere.GetCumulsDotExerciceDR(VHImmo^.Encours.Deb,CumulAntDR,CumulAntFEC, False, true, False);

      PlanMere.Calcul(Q,idate1900);
      PlanMere.SetTypeOpe('CES');
      PlanMere.TraiteSortie (fDateOp,'', 0,'OUI');  //fq 21754 - mbo 02/10/2007 param false on calcule dot° y compris jour de cession

      CumulAntEco := CumulAntEco + DotCesEco;
      CumulAntFisc := CumulAntFisc + DotCesFisc;

      //06.06.07 CumulAntFec := CumulAntFec + DotCesEco - DotCesFisc;

    finally
      Ferme (Q);
    end;

    TobImmoFille := TOB.Create ('IMMO', nil, -1);

    // mise à jour des éléments de la fille
    TobImmoFille.PutValue('I_IMMO', GetControlText('RPL_CODE'));
    TobImmoFille.PutValue('I_LIBELLE', GetControlText('RPL_CPDESI'));
    TobImmoFille.PutValue('I_REFINTERNEA', GetControlText('RPL_CPFAC')); // Facture
    TobImmoFille.PutValue('I_TIERSA', GetControlText('RPL_CPFOUR'));     // Fournisseur
    TobImmoFille.PutValue('I_QUOTEPART', TobImmoMere.GetValue('I_QUOTEPART'));
    TobImmoFille.PutValue('I_COMPTEIMMO', TobImmoMere.GetValue('I_COMPTEIMMO'));
    TobImmoFille.PutValue('I_COMPTEAMORT', TobImmoMere.GetValue('I_COMPTEAMORT'));
    TobImmoFille.PutValue('I_COMPTEDOTATION', TobImmoMere.GetValue('I_COMPTEDOTATION'));
    TobImmoFille.PutValue('I_COMPTEDEROG', TobImmoMere.GetValue('I_COMPTEDEROG'));
    TobImmoFille.PutValue('I_REPRISEDEROG', TobImmoMere.GetValue('I_REPRISEDEROG'));
    TobImmoFille.PutValue('I_PROVISDEROG', TobImmoMere.GetValue('I_PROVISDEROG'));
    TobImmoFille.PutValue('I_DOTATIONEXC', TobImmoMere.GetValue('I_DOTATIONEXC'));
    TobImmoFille.PutValue('I_VACEDEE', TobImmoMere.GetValue('I_VACEDEE'));
    TobImmoFille.PutValue('I_AMORTCEDE', TobImmoMere.GetValue('I_AMORTCEDE'));
    TobImmoFille.PutValue('I_REPEXPLOIT', TobImmoMere.GetValue('I_REPEXPLOIT'));
    TobImmoFille.PutValue('I_REPEXCEP', TobImmoMere.GetValue('I_REPEXCEP'));
    TobImmoFille.PutValue('I_VAOACEDEE', TobImmoMere.GetValue('I_VAOACEDEE'));
    TobImmoFille.PutValue('I_COMPTEREF', TobImmoMere.GetValue('I_COMPTEREF'));
    TobImmoFille.PutValue('I_SOCIETE', TobImmoMere.GetValue('I_SOCIETE'));
    TobImmoFille.PutValue('I_ETABLISSEMENT', TobImmoMere.GetValue('I_ETABLISSEMENT'));
    TobImmoFille.PutValue('I_ABREGE',GetControlText('RPL_CODE'));
    TobImmoFille.PutValue('I_NATUREIMMO', TobImmoMere.GetValue('I_NATUREIMMO'));
    TobImmoFille.PutValue('I_QUALIFIMMO', TobImmoMere.GetValue('I_QUALIFIMMO'));
    TobImmoFille.PutValue('I_CODEPOSTAL', TobImmoMere.GetValue('I_CODEPOSTAL'));
    TobImmoFille.PutValue('I_VILLE', TobImmoMere.GetValue('I_VILLE'));
    TobImmoFille.PutValue('I_PAYS', TobImmoMere.GetValue('I_PAYS'));
    TobImmoFille.PutValue('I_VENTILABLE', TobImmoMere.GetValue('I_VENTILABLE'));
    TobImmoFille.PutValue('I_VENTILABLE1', TobImmoMere.GetValue('I_VENTILABLE1'));
    TobImmoFille.PutValue('I_VENTILABLE2', TobImmoMere.GetValue('I_VENTILABLE2'));
    TobImmoFille.PutValue('I_VENTILABLE3', TobImmoMere.GetValue('I_VENTILABLE3'));
    TobImmoFille.PutValue('I_VENTILABLE4', TobImmoMere.GetValue('I_VENTILABLE4'));
    TobImmoFille.PutValue('I_VENTILABLE5', TobImmoMere.GetValue('I_VENTILABLE5'));
    TobImmoFille.PutValue('I_LIEUGEO', TobImmoMere.GetValue('I_LIEUGEO'));
    TobImmoFille.PutValue('I_NATUREBIEN', TobImmoMere.GetValue('I_NATUREBIEN'));
    TobImmoFille.PutValue('I_DEVISE', TobImmoMere.GetValue('I_DEVISE'));
    TobImmoFille.PutValue('I_GROUPEIMMO', TobImmoMere.GetValue('I_GROUPEIMMO'));
    TobImmoFille.PutValue('I_DOCGUID', TobImmoMere.GetValue('I_DOCGUID'));
    TobImmoFille.PutValue('I_NONDED', TobImmoMere.GetValue('I_NONDED'));
    TobImmoFille.PutValue('I_COMPTEDEP', TobImmoMere.GetValue('I_COMPTEDEP'));
    TobImmoFille.PutValue('I_ETAT', 'OUV');
    TobImmoFille.PutValue('I_DATEDEBECO', fdateOp+1);
    TobImmoFille.PutValue('I_DATEDEBFIS', fdateOp+1);
    TobImmoFille.PutValue('I_DATEPIECEA', fdateOp+1);
    TobImmoFille.PutValue('I_DATEAMORT', fdateOp+1);

    TobImmoFille.PutValue('I_MONTANTHT', Valeur(GetControlText('RPL_HT')));
    TobImmoFille.PutValue('I_VALEURACHAT', Valeur(GetControlText('RPL_HT'))); //mbo 18.06.07

    TobImmoFille.PutValue('I_MONTANTBASEAMORT', Valeur(GetControlText('RPL_TTC')));
    TobImmoFille.PutValue('I_BLOCNOTE', RichToString (fblocnote));
    TobImmoFille.PutValue('I_TVARECUPERABLE', Valeur(GetControlText('RPL_TVA1')));
    TobImmoFille.PutValue('I_TVARECUPEREE', Valeur(GetControlText('RPL_TVA2')));
    TobImmoFille.PutValue('I_TVARECUPERABLE', Valeur(GetControlText('RPL_TVA1')));
    TobImmoFille.PutValue('I_BASETAXEPRO', Valeur(GetControlText('RPL_TP')));

    TobImmoFille.PutValue('I_BASEECO', Valeur(GetControlText('RPL_BASE')));
    TobImmoFille.PutValue('I_BASEAMORDEBEXO', Valeur(GetControlText('RPL_BASE')));
    TobImmoFille.PutValue('I_BASEAMORFINEXO', Valeur(GetControlText('RPL_BASE')));
    TobImmoFille.PutValue('I_METHODEECO', GetControlText('RPL_METHODE'));
    TobImmoFille.PutValue('I_DUREEECO', Valeur(GetControlText('RPL_DUREE')));
    TobImmoFille.PutValue('I_TAUXECO', Valeur(GetControlText('RPL_TAUX')));

    // on force le fiscal en linéaire avec calcul du plan futur sur la vnf
    //TobImmoFille.PutValue('I_METHODEFISC', TobImmoMere.GetValue('I_METHODEFISC'));
    TobImmoFille.PutValue('I_METHODEFISC','LIN');
    TobImmoFille.PutValue('I_BASEFISC', VRFisc);    // valeur résiduelle fiscale
    TobImmoFille.PutValue('I_FUTURVNFISC', '***');  // on force calcul sur durée restante


    NOMBREMOIS(TobImmoMere.GetValue('I_DATEDEBFIS'), fDateOp, PremMois, PremAnnee, NbMois);
    MoisRestants := (TobImmoMere.GetValue('I_DUREEFISC')) - NbMois;

    TobImmoFille.PutValue('I_DUREEFISC', MaxIntValue([0, MoisRestants]));

    //TobImmoFille.PutValue ('I_TAUXFISC', GetTaux(TobImmoMere.GetValue('I_METHODEFISC')
    //                                     , fDateOp
    //                                     , fDateOp
    //                                     , NbMois));

    TobImmoFille.PutValue ('I_TAUXFISC', TobImmoMere.GetValue('I_TAUXFISC'));
    TobImmoFille.PutValue('I_DATEDERNMVTFISC',TobImmoMere.GetValue('I_DATEDERNMVTFISC'));
    TobImmoFille.PutValue('I_QUANTITE', Valeur(GetControlText('RPL_QTE')));
    TobImmoFille.PutValue('I_DATECREATION', Date);
    TobImmoFille.PutValue('I_DATEMODIF', Date);
    TobImmoFille.PutValue('I_PLANACTIF', NewPlanSeq);
    TobImmoFille.PutValue('I_REINTEGRATION', Valeur(GetControlText('RPL_PLAF')));
    TobImmoFille.PutValue('I_REMPLACE', CodeImmoMere);
    TobImmoFille.PutValue('I_TYPER','IRM');
    TobImmoFille.PutValue('I_VERSION', V_PGI.NumVersion);
    TobImmoFille.PutValue('I_BASEAMORDEBEXO',Valeur(GetControlText('RPL_BASE')));
    TobImmoFille.PutValue('I_BASEAMORFINEXO',Valeur(GetControlText('RPL_BASE')));

    // on stocke la notion de composant remplaçant qui sera reconduite de clôture en clôture
    // car i_remplace est réinitialisé par la cloture et on ne sait plus que l'immo est un
    // composant (on a besoin de le savoir dans la fiche immo pour ne pas faire de controles en
    // validation

    TobImmoFille.PutValue('I_STRING1', 'X');
    TobImmoFille.PutValue('I_TYPEDEROGLIA', TypeDerogatoire(TobImmoFille, nil) ); // FQ 20256

    // mise à jour immolog de la mère
    TobLogMere := TOB.Create ('IMMOLOG',nil,-1);
    TobLogMere.PutValue('IL_IMMO',CodeImmoMere);

    TobLogMere.PutValue('IL_DOTCESSECO',Arrondi((DotCesEco),V_PGI.OKDecV));
    TobLogMere.PutValue('IL_DOTCESSFIS',Arrondi((DotCesFisc),V_PGI.OKDecV));

    TobLogMere.PutValue('IL_CUMANTCESECO',Arrondi((CumulAntEco),V_PGI.OKDecV));
    TobLogMere.PutValue('IL_CUMANTCESFIS',Arrondi((CumulAntFisc),V_PGI.OKDecV));

    TobLogMere.PutValue('IL_LIBELLE','Composant remplacé par le composant n° '+ GetControlText('RPL_CODE'));
    TobLogMere.PutValue('IL_TYPEMODIF',AffecteCommentaireOperation('CES'));
    TobLogMere.PutValue('IL_DATEOP',fDateOp);
    TobLogMere.PutValue('IL_TYPEOP','CES');
    TobLogMere.PutValue('IL_MOTIFCES','999');
    TobLogMere.PutValue('IL_CALCCESSION','NOR');
    TobLogMere.PutValue('IL_VOCEDEE',TobImmoMere.GetValue('I_MONTANTHT'));
    TobLogMere.PutValue('IL_QTECEDEE',TobImmoMere.GetValue('I_QUANTITE'));
    TobLogMere.PutValue('IL_MONTANTCES',0);
    TobLogMere.PutValue('IL_TVAAREVERSER',0);
    TobLogMere.PutValue('IL_TVARECUPEREE',TobImmoMere.GetValue('I_TVARECUPEREE'));
    TobLogMere.PutValue('IL_TVARECUPERABLE',TobImmoMere.GetValue('I_TVARECUPERABLE'));
    TobLogMere.PutValue('IL_REPRISEECO',TobImmoMere.GetValue('I_REPRISEECO'));
    TobLogMere.PutValue('IL_REPRISEFISC',TobImmoMere.GetValue('I_REPRISEFISCAL'));

    TobLogMere.PutValue('IL_PVALUE', 0);
    TobLogMere.PutValue('IL_VNC', TobImmoMere.GetValue('I_VNC'));
    TobLogMere.PutValue('IL_PLANACTIFAV',TobImmoMere.GetValue('I_PLANACTIF'));
    TobLogMere.PutValue('IL_PLANACTIFAP',TobImmoMere.GetValue('I_PLANACTIF')+1);
    TobLogMere.PutValue('IL_ORDRE',TrouveNumeroOrdreLogSuivant(CodeImmoMere));
    TobLogMere.PutValue('IL_ORDRESERIE',-1);

    TobLogMere.PutValue('IL_MONTANTEXC',0);
    TobLogMere.PutValue('IL_TYPEEXC','');
    TobLogMere.PutValue('IL_BASEECOAVMB',TobImmoMere.GetValue('I_BASEECO'));
    TobLogMere.PutValue('IL_BASEFISCAVMB',TobImmoMere.GetValue('I_BASEFISC'));
    TobLogMere.PutValue('IL_BASETAXEPRO',TobImmoMere.GetValue('I_BASETAXEPRO'));
    TobLogMere.PutValue('IL_CODEECLAT',GetControlText('RPL_CODE'));

    { Log mère}
    TobLogMere.InsertDB(nil);

    {duplication des ventil de la mère sur la fille}
    if TobImmoMere.GetValue('I_VENTILABLE')='X' then ImDupliqueVentil(CodeImmoMere,GetControlText('RPL_CODE')) ;

    { Mise à jour de la fiche de la mère }
    TobImmoMere.PutValue('I_OPERATION','X');
    TobImmoMere.PutValue('I_OPECESSION','X');
    TobImmoMere.PutValue('I_MONTANTHT',0);
    TobImmoMere.PutValue('I_VNC',Arrondi(Valeur(GetControlText('RPL_VNC')), V_PGI.OkDecV));

    TobImmoMere.PutValue('I_BASEECO',0);
    TobImmoMere.PutValue('I_BASEFISC',0);
    TobImmoMere.PutValue('I_QTCEDE',TobImmoMere.GetValue('I_QUANTITE'));
    TobImmoMere.PutValue('I_QUANTITE',0);
    TobImmoMere.PutValue('I_BASEAMORFINEXO',0);
    TobImmoMere.PutValue('I_PLANACTIF',TobImmoMere.GetValue('I_PLANACTIF')+1);
    TobImmoMere.PutValue('I_REPCEDECO',TobImmoMere.GetValue('I_REPRISEECO'));
    TobImmoMere.PutValue('I_REPCEDFISC',TobImmoMere.GetValue('I_REPRISEFISCAL'));
    TobImmoMere.PutValue('I_REPRISEECO',0);
    TobImmoMere.PutValue('I_REPRISEFISCAL',0);
    TobImmoMere.PutValue('I_DATEDERMVTECO', fDateOp);
    TobImmoMere.PutValue('I_REPRISEDEPCEDEE',TobImmoMere.GetValue('I_REPRISEDEP'));
    TobImmoMere.PutValue('I_REPRISEDEP',0);
    TobImmoMere.PutValue('I_DATECESSION', fDateOp);
    TobImmoMere.PutValue('I_REGLECESSION','NOR');
    TobImmoMere.PutValue('I_DATEDERNMVTFISC',fDateOp);
    TobImmoMere.PutValue('I_REPRISEFDRCEDEE',TobImmoMere.GetValue('I_REPRISEDR'));
    TobImmoMere.PutValue('I_REPRISEFECCEDEE',TobImmoMere.GetValue('I_REPRISEFEC'));
    TobImmoMere.PutValue('I_REPRISEFEC',0);
    TobImmoMere.PutValue('I_REPRISEDR',0);


    //valeurs spécifiques au remplacement
    TobImmoMere.PutValue('I_REMPLACEE', GetControlText('RPL_CODE'));
    TobImmoMere.PutValue('I_TYPER', 'VNC');
    TobImmoMere.PutValue('I_OPEREMPL', 'X');
    TobImmoMere.PutValue('I_OPERATION', 'X');


    TobImmoMere.UpdateDateModif;
    {mise à jour de la fiche immo}
    TobImmoMere.UpdateDB;


    PlanMere.NumSeq := PlanActifMere ;
    PlanMere.Sauve;
    PlanMere.Free;
    Ferme(Qplan);

    // on va charger l'immo pour calcul des annuités et création ds immoamor
    //QPlan:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+GetControlText('RPL_CODE')+'"', FALSE) ;
    //QPlan.Edit;

    PlanFille := TPlanAmort.Create(True);
    Planfille.ChargeTOb(TobImmoFille);
    PlanFille.NumSeq := 0;  // calcul fait + 1
    PlanFille.SetTypeOpe('CRE');
    PlanFille.CalculTob(TobImmoFille,iDate1900);

    // ajout 06.06.07
    if PlanFille.AmortEco.TableauDot[0] >= PlanFille.AmortFisc.TableauDot[0] then
       GestFisc := true
    else
       GestFisc := false;

    PlanFille.NumSeq := 0; // sauve fait + 1
    PlanFille.Sauve;
    PlanFille.CalculDateFinAmortissement(PlanFille.AmortEco);
    TobImmoFille.PutValue('I_DATEDERMVTECO',PlanFille.AmortEco.DateFinAmort);

    // ajout 06.06.07
    // si la mère avait la gestion fiscale on la reconduit sur la fille
    // autre cas : la mère n'avait pas la gestion fiscale mais le nouveau plan éco
    // saisi sur la fille fait que la fille prend la gestion fiscale

    If GestFisc then
       TobImmoFille.PutValue('I_NONDED', 'X');

    TobImmoFille.InsertDB(nil);

    PlanFille.Free;
    TobImmoFille.Free;

    TobImmoMere.Free;

    TobLogFille := TOB.Create ('IMMOLOG',nil,-1);
    try
      TobLogFille.PutValue('IL_IMMO',GetControlText('RPL_CODE'));
      TobLogFille.PutValue('IL_LIBELLE', RechDom('TIOPEAMOR', 'ACQ', FALSE) + ' ' +DateToStr(fDateOp+1));
      TobLogFille.PutValue('IL_TYPEMODIF',AffecteCommentaireOperation('ACQ'));
      TobLogFille.PutValue('IL_TVARECUPERABLE', Valeur(GetControlText('RPL_TVA1')));
      TobLogFille.PutValue('IL_TVARECUPEREE', Valeur(GetControlText('RPL_TVA1')));


      TobLogFille.PutValue('IL_DATEOP', fDateOp+1);
      TobLogFille.PutValue('IL_METHODEECO', GetControlText('RPL_METHODE'));

      TobLogFille.PutValue('IL_DUREEECO', GetControlText('RPL_DUREE'));

      TobLogFille.PutValue('IL_TYPEOP', 'ACQ');
      TobLogFille.PutValue('IL_ORDRE', 1);
      TobLogFille.PutValue('IL_OrdreSerie', 1);
      TobLogFille.PutValue('IL_PLANACTIFAV', 0);
      TobLogFille.PutValue('IL_PLANACTIFAP', 1);
      TobLogFille.PutValue('IL_BLOCNOTE', RichToString (fBlocnote));

      TobLogFille.InsertDB(nil);
    finally
       TobLogFille.Free;
    end;

    COMMITTRANS ;

    VHImmo^.ChargeOBImmo := True;
    ImMarquerPublifi (True);

 except
   PGIBox(HM2[8]);
   ROLLBACK ;
 end ;

  TFVIERGE(ECRAN).retour := IntToStr(MrYes);

end;


//====================================================================
procedure TOF_AMREMPLACEMENT.OnLoad ;
begin
  Inherited ;
end ;


procedure TOF_AMREMPLACEMENT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AMREMPLACEMENT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AMREMPLACEMENT.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_AMREMPLACEMENT ] ) ;
end.


