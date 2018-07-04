{***********UNITE*************************************************
Auteur  ...... : TGA
Créé le ...... : 08/09/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AMDEPRECIATION ()
Mots clefs ... : TOF;AMDEPRECIATION
Suite..........: FQ 17259 BTY 01/06 Nelle colonne dans IMMO
Suite..........: MBO - 13/02/2006 - Impact de la présence d'un antérieur dépréciation
                                    dans le calcul de la reprise maximale
Suite..........: MBO - 13/02/2006 - FQ 17476 - Stockage de la dépréciation sur le plan fiscal
Suite..........: BTY - 05/06 - FQ 18119 Positionner les indicateurs de modif
Suite..........: MBO - 02/06/2006 - FQ 18295 - message d'erreur si immo cédée
Suite..........: MBO - 14/06/2006 - FQ 18386 - pas de dépréciation si datedebeco > fin exo en cours
Suite..........: BTY - 06/06 FQ 17500 Prise en compte de l'antérieur dépréciation saisi
Suite..........: MVG - 12/07/2006 Correction conseils de compilation
Suite..........: MBO - 28/03/2007 - FQ 17512 - pas de dépréciation si gestion fiscale
Suite..........: MBO - 16/04/2007 - Ajout test + message si remplacement de composant
*****************************************************************}
Unit AMDEPRECIATION_TOF ;

interface

Uses StdCtrls,
     Controls,
     Forms,
     Classes,
     Windows,
     Menus,
     Graphics,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     MaineAGL,
{$ENDIF}
     ImEnt,
     ImPlan,
     ImOuPlan,
     Chanplan,
     Outils,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     uTOB,
     ParamSoc,
     HRichOle,
     LookUp,
     HTB97,
     AGLInit,
     UTOF,
     Vierge;

Type

  TOF_AMDEPRECIATION = Class (TOF)

    procedure OnArgument (S : String ) ; override ;
    procedure OnAfterFormShowEcran     ;
    procedure OnUpdate ; override ;

  private
    { Variables }
    fCode_immo : String;
    fPlan:TPlanAmort;
    fInitVNC : double;
    fCumulEco: double;
    fDotationExe: double;
    fdatetest: string;
    fPlanActifAv: integer;
    fPlanActifAp: integer;
    fMntDeprec : double;
    fDP_BLOCNOTE: THRichEditOLE;

    procedure OnExitDP_MONTANTOPE(Sender:TObject);
    procedure OnClickRbAjout(Sender:TObject);
    procedure OnClickRbReprise(Sender:TObject);
    procedure OnbFermeClick(Sender: TObject);
    Function  ControleDuMontant :Boolean;
    Function  ControleDeLaDate  :Boolean;
    procedure AffichageDesMontants;

  end ;

Function AMLanceFiche_DepreciationActif(stCode: string):String;

Implementation

const
  // fq 18386 - ajout du message si datebeco > fin exo en cours
  HM: array[0..12] of string =
  ('Il n''est pas possible de pratiquer plusieurs dépréciations/reprises dans l''exercice.',
    'La saisie de la date est obligatoire.',
    'La date doit appartenir à l''exercice en cours.',
    'La date doit être supérieure ou égale à la plus haute date d''opération déjà enregistrée.',
    'Le montant doit être inférieur ou égal à ',
    'Voulez-vous enregistrer cette opération ?',
    'L''enregistrement de la dépréciation à échoué.',
    'L''immobilisation est sortie.',
    'Après l''enregistrement de la dépréciation, toute modification ' +
    'rétrospective du plan d''amortissement sera interdite. Voulez-vous continuer ?',
    'Il n''est pas possible de pratiquer une dépréciation ' +
    'si l''immobilisation n''est pas amortie sur l''exercice en cours.' ,
    'Montant négatif  ou nul',
    'Il n''est pas possible de pratiquer une dépréciation ' +
    'si l''immobilisation est amortie avec une gestion fiscale.',
    'L''immobilisation a fait l''objet d''un remplacement de composant.');



{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 04/11/2005
Modifié le ... :   /  /
Description .. :  Activation de la fenêtre Dépréciation
Mots clefs ... :
*****************************************************************}
Function AMLanceFiche_DepreciationActif(Stcode: string):String;
  begin

  // ajout mbo pour retour dans la fiche immo - fq 18295 02.06.06
  result := '0';

   // Ctl présence immo avec opération dépréciation
   if ExisteSQL ('SELECT IL_IMMO FROM IMMOLOG WHERE IL_IMMO="'+
    Stcode+'" AND IL_TYPEOP="DPR" AND IL_DATEOP >="'+
    USDATETIME(VHImmo^.EnCours.Deb)+'" AND IL_DATEOP <="'+
    USDATETIME(VHImmo^.EnCours.Fin)+'" ') then
      PGIBox(HM[0])
   // Ctl présence immo sortie mais pas remplacement de composant (ajout mbo)
   Else if ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+
    Stcode+'" AND (I_QUANTITE=0) AND (I_OPEREMPL="-")') then
      PGIBox(HM[7])
   // mbo ajout : Ctl présence immo sortie mais remplacement de composant
   Else if ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+
    Stcode+'" AND (I_QUANTITE=0) AND (I_OPEREMPL="X")') then
      PGIBox(HM[12])
   // fq 18386 blocage si amortissement non commencé
   // BTY > et non pas >=
   else if ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+
    Stcode+'" AND I_DATEDEBECO >"'+
    USDATETIME(VHImmo^.EnCours.Fin)+'" ') then
      PGIBox(HM[9])
   // mbo fq 17512 chantier fiscal
   Else if ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+
    Stcode+'" AND I_NONDED="X"') then
      PGIBox(HM[11])
   Else
      result:= AGLLanceFiche('AM', 'AMDEPRECIATION','' , '', Stcode);
 end;


{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 04/11/2005
Modifié le ... :   /  /
Description .. :  Préparation de la fenêtre Dépréciation
Mots clefs ... :
*****************************************************************}
procedure TOF_AMDEPRECIATION.OnArgument (S: String ) ;
var CumulFisc: double;
    QPlan : TQuery;
 begin
  Inherited ;

    // Initialisation
    fInitVNC :=0.00;
    CumulFisc:=0.00;
    fCumulEco:=0.00;

    ThEdit(GetControl ('DP_MONTANTOPE')).OnExit := OnExitDP_MONTANTOPE;
    TradioButton (GetControl ('DP_RbAjout')).OnClick := OnClickRbAjout;
    TradioButton (GetControl ('DP_RbReprise')).OnClick := OnClickRbReprise;
    TToolbarButton97(GetControl('BFerme')).Onclick := OnbFermeClick;
    TFVIERGE(ECRAN).retour := IntToStr(MrNo);
    TFVIERGE(ECRAN).OnAfterFormShow := OnAfterFormShowEcran;

    // Assignation du bloc note
    fDP_BLOCNOTE := THRichEditOLE(GetControl('DP_BLOCNOTE'));

    // Lecture de l'immo en cours
    fPlan:=TPlanAmort.Create(true) ;
    fCode_immo := S;
    QPlan:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+fCode_immo+'"', FALSE) ;
    fPlan.Charge(QPlan);
    fPlan.Recupere(fCode_immo,QPlan.FindField('I_PLANACTIF').AsString);

    SetControlCaption('DP_CODEIMMO',fCode_immo);
    SetControlCaption('DP_DESIGNATION',fPlan.libelleImmo);
    SetControlCaption('DP_DATEACHAT', DateToStr(fPlan.dateachat));
    SetControlCaption('DP_ValeurHT',StrFMontant(fPlan.ValeurHT,15,V_PGI.OkDecV,'',True));

    // Affichage initial Ajout par défaut
    SetControlChecked('DP_RBAjout', True);

    // Calcul de la VNC de fin d'exercice
    fPlan.GetCumulsDotExercice(VHImmo^.Encours.Fin,fCumulEco,CumulFisc,false,false,true);

    // suite modif calcul mbo 27.10.05 cumuldepre est déjà dans fcumuleco
    // BTY FQ 17500 impact antérieur dépréciation, le soustraire aussi
    //fInitVNC := Arrondi(fPlan.BaseDebutExoEco-fPlan.AmortEco.Reprise-fCumulEco, V_PGI.OkDecV );
    fInitVNC := Arrondi(fPlan.BaseDebutExoEco-fPlan.AmortEco.Reprise-fCumulEco-fPlan.AmortEco.RepriseDep, V_PGI.OkDecV );
    SetControlText('DP_VNC', StrFMontant(fInitVNC,15,V_PGI.OkDecV,'',True));

    // Plafond = Vnc
    SetControlText('DP_PLAFOND', StrFMontant(fInitVNC,15,V_PGI.OkDecV,'',True));

    fDotationExe := fPlan.GetDotationExercice(VHImmo^.EnCours.Fin,fPlan.AmortEco,false);
    SetControlCaption('DP_DOTAVTOPE', StrFMontant(fDotationExe,15,V_PGI.OkDecV,'',True));
    SetControlCaption('DP_DOTAPROPE', StrFMontant(fDotationExe,15,V_PGI.OkDecV,'',True));
end ;


{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 04/11/2005
Modifié le ... :   /  /
Description .. :  Mise à Jour des masques apres affichage
Mots clefs ... :
*****************************************************************}
procedure TOF_AMDEPRECIATION.OnAfterFormShowEcran;
begin
  Inherited ;
  SetControlProperty('DP_VNC','DisplayFormat',StrfMask(V_PGI.OkDecV,'', True));;
  SetControlProperty('DP_PLAFOND','DisplayFormat',StrfMask(V_PGI.OkDecV,'', True));;
  SetControlProperty('DP_MONTANTOPE','DisplayFormat',StrfMask(V_PGI.OkDecV,'', True));;
  SetControlProperty('DP_DEPRECIATION','DisplayFormat',StrfMask(V_PGI.OkDecV,'', True));;

  // sauvegarde de la date initiale
  fdatetest := GetControlText('DP_DATEOPE');

end;


{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 04/11/2005
Modifié le ... :   /  /
Description .. : gestion du Click sur ajout
Mots clefs ... :
*****************************************************************}
procedure TOF_AMDEPRECIATION.OnClickRbAjout (Sender:TObject);
begin
   // Réinitialisation des zones
   SetControlText('DP_VNC', StrFMontant(fInitVNC,15,V_PGI.OkDecV,'',True));
   SetControlText('DP_PLAFOND', StrFMontant(fInitVNC,15,V_PGI.OkDecV,'',True));
   SetControlText('DP_MONTANTOPE',StrFMontant(0.00,15,V_PGI.OkDecV,'',True));;
   SetControlCaption('DP_DOTAVTOPE', StrFMontant(fDotationExe,15,V_PGI.OkDecV,'',True));
   SetControlCaption('DP_DOTAPROPE', StrFMontant(fDotationExe,15,V_PGI.OkDecV,'',True));
   SetControlText('DP_DEPRECIATION', StrFMontant(0.00,15,V_PGI.OkDecV,'',True));
   SetFocusControl('DP_MONTANTOPE');
end;


{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 04/11/2005
Modifié le ... :   /  /
Description .. : gestion du  Click sur reprise
Suite..........: mbo - 02.06 - impact de l'antérieur dépréciation sur calcul de la reprise maximale
Mots clefs ... :
*****************************************************************}
procedure TOF_AMDEPRECIATION.OnClickRbReprise(Sender:TObject);
Var montant:Double;
begin

  // mbo 4.11.05 le plafond est stocké dans immolog par la cloture
  montant := fPlan.CumulDepre;
  // impact antérieur dépréciation mbo 02.06
  if (montant = 0) and (fplan.AmortEco.RepriseDep <> 0) then
    // il faut recalculer le montant maximum de la reprise
    montant := fPlan.CalculRepriseDepreciation(fplan.AmortEco, VHImmo^.EnCours.Deb,
                                               VHImmo^.EnCours.Fin, false);
 // fin 02.06

  SetControlText('DP_PLAFOND', StrFMontant(montant,15,V_PGI.OkDecV,'',True));

  // Réinitialisation des zones
  SetControlText('DP_VNC', StrFMontant(fInitVNC,15,V_PGI.OkDecV,'',True));
  SetControlText('DP_MONTANTOPE',StrFMontant(0.00,15,V_PGI.OkDecV,'',True));
  SetControlCaption('DP_DOTAVTOPE', StrFMontant(fDotationExe,15,V_PGI.OkDecV,'',True));
  SetControlCaption('DP_DOTAPROPE', StrFMontant(fDotationExe,15,V_PGI.OkDecV,'',True));

  SetControlText('DP_DEPRECIATION', StrFMontant(0.00,15,V_PGI.OkDecV,'',True));
  SetFocusControl('DP_MONTANTOPE');

end;


{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 04/11/2005
Modifié le ... :   /  /
Description .. : Contrôle de la date d'opération
Mots clefs ... :
*****************************************************************}
Function TOF_AMDEPRECIATION.ControleDeLaDate :Boolean;
var date_saisie: TdateTime;
    dd :string;
begin

  result := False;

  dd:= GetControlText('DP_DATEOPE');
  if dd = fdatetest then
    Begin
      result := True;
      PGIBox(HM[1]);
      SetFocusControl('DP_DATEOPE');
    End
  Else
    Begin
      date_saisie := StrToDate(GetControlText('DP_DATEOPE'));
      // la date doit appartenir à l'exercice en cours
      if (date_saisie >  VHImmo^.Encours.Fin) OR ( date_saisie < VHImmo^.Encours.Deb) then
        begin
          PGIBox(HM[2]);
          result := True;
          SetFocusControl('DP_DATEOPE');
        end
        // date de dépréciation doit être > date dernière opération
      else if ExisteSQL('SELECT IL_DATEOP FROM IMMOLOG WHERE IL_IMMO="'+fCode_immo+'" AND IL_DATEOP>"'+USDATETIME(date_saisie)+'"') then
        begin
          PGIBox(HM[3]);
          result := True;
          SetFocusControl('DP_DATEOPE');
        end;
    end;
end;



{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 04/11/2005
Modifié le ... :   /  /
Description .. : Contrôle du montant de l'opération
Mots clefs ... :
*****************************************************************}
Function TOF_AMDEPRECIATION.ControleDuMontant :Boolean;
var montant:Double;
    plafond:double;
begin

//   plafond:=0.00;  MVG 12/07/2006
//   montant:=0.00;  MVG 12/07/2006
   result := False;

   // Récupération du montant de l'opération
   montant := Valeur(GetControlText('DP_MONTANTOPE'));

   // Test montant <= au plafond
   plafond := Valeur(GetControlText('DP_PLAFOND'));

   IF (montant > plafond) then
   Begin
      PGIBox(HM[4]+ FloatToStr(plafond));
      SetFocusControl('DP_MONTANTOPE');
      result:=True;
      exit;
   end;
   IF (montant <=0) then
   Begin
      PGIBox(HM[10]);
      SetFocusControl('DP_MONTANTOPE');
      result:=True;
      exit;
   end;


   // Mise à jour de la VNC ,de la dépréciation, de la dotation après
   AffichageDesMontants;

end;


{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 04/11/2005
Modifié le ... :   /  /
Description .. : Affichage des montants de l'opération
Mots clefs ... :
*****************************************************************}
Procedure TOF_AMDEPRECIATION.AffichageDesMontants;
var montant:Double;
    vnc:Double;
    dot:Double;
begin

   // Récupération du montant de l'opération
   montant := Valeur(GetControlText('DP_MONTANTOPE'));

   // Mise à jour de la VNC ,de la dépréciation, de la dotation après
   IF TRadioButton(GetControl('DP_RBAjout')).Checked = True then
    Begin
      vnc := fInitvnc - montant;
      dot := fDotationExe + montant;
      SetControlText('DP_DEPRECIATION', StrFMontant(montant,15,V_PGI.OkDecV,'',True));
    End
   Else
    Begin
      vnc := fInitvnc + montant;
      dot := fDotationExe - montant;
      SetControlText('DP_DEPRECIATION', StrFMontant((montant*-1),15,V_PGI.OkDecV,'',True));
    End;
   SetControlText('DP_VNC', StrFMontant(vnc,15,V_PGI.OkDecV,'',True));
   SetControlCaption('DP_DOTAPROPE', StrFMontant(Dot,15,V_PGI.OkDecV,'',True));

end;


{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 04/11/2005
Modifié le ... :   /  /
Description .. : Contrôle en cours de saisie du montant de l'opération
Mots clefs ... :
*****************************************************************}
procedure TOF_AMDEPRECIATION.OnExitDP_MONTANTOPE(Sender:TObject);
begin
   AffichageDesMontants;
end;



{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 04/11/2005
Modifié le ... :   /  /
Description .. : Sortie par la X sans validation
Mots clefs ... :
*****************************************************************}
procedure TOF_AMDEPRECIATION.OnbFermeClick(Sender: TObject);
Var montant : double;
    dd :string;
    question : TmodalResult;
begin

  dd := GetControlText('DP_DATEOPE');
  montant:= valeur(GetControlText('DP_MONTANTOPE'));

  IF (montant=0) AND (dd=fdatetest) THEN
    Begin
      // Pas de saisie on sort
      Ecran.ModalResult := mrNo;
    end
  Else
    Begin
     // Action Echap ou Fermer par la croix
     Question := PGIAskCancel(HM[5],Ecran.Caption);
     If Question = mrYes then
        // contrôle validation
        OnUpdate
     else If Question = mrNo then
        // Sortir
        Ecran.ModalResult := mrNo
     else
        // Cancel
        Ecran.ModalResult := mrNone;
    end

end;




{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 04/11/2005
Modifié le ... :   /  /
Description .. : Sortie par validation de la Dépréciation
Mots clefs ... :
*****************************************************************}
procedure TOF_AMDEPRECIATION.OnUpdate;

var Ordre : integer;
    TLog : TOB;
    lib :String;
    erreur : Boolean;
    QPlan : TQuery;
    PlanNew : TPlanAmort;
    question : TmodalResult;

begin
  Inherited ;

   erreur := ControleDeLaDate;
   IF erreur=True then
     Begin
       ecran.ModalResult :=mrNone;
       SetFocusControl('DP_DATEOPE');
       exit;
     end;

   erreur := ControleDuMontant;
   IF erreur=True then
     Begin
       ecran.ModalResult :=mrNone;
       SetFocusControl('DP_MONTANTOPE');
       exit;
     end;

   // Enregistrement de la dépréciation
   IF Fplan.GestDeprec = false Then
     Begin
      Question := PGIAsk(HM[8],Ecran.Caption);
      If Question = mrNo then
        Begin
          // Sortir
          Ecran.ModalResult := mrNo;
          exit;
        End
     end;

   // Récupération du montant de la dépréciation
   IF TRadioButton(GetControl('DP_RBAjout')).Checked = True then
    Begin
     lib :='Dépréciation';
     fMntDeprec := Valeur(GetControlText('DP_MONTANTOPE'))
    End
   Else
    Begin
     lib :='Reprise de dépréciation';
     fMntDeprec := Valeur(GetControlText('DP_MONTANTOPE'))* -1;
    End;


   // Mise à jour de la Dépréciation dans  les tables IMMO,IMMOAMOR et IMMOLOG
   BEGINTRANS ;
   try

      // Mise à jour de IMMO
      QPlan:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+fCode_immo+'"', FALSE) ;

      // Maj de la fiche immo
      QPlan.Edit;
      QPlan.FindField('I_OPERATION').AsString:= 'X' ;
      // BTY 01/06 Nelle zone de IMMO
      QPlan.FindField('I_OPEDEPREC').AsString:= 'X' ;
      QPlan.FindField('I_OPECHANGEPLAN').AsString:= 'X' ;
      QPlan.FindField('I_REVISIONECO').AsFloat := fMntDeprec ;

      // fq 17476 - Mbo
      QPlan.FindField('I_REVISIONFISCALE').AsFloat := fMntDeprec ;

      QPlan.Post;

      // Maj du plan d'amortissement IMMOAMOR
      QPlan.Edit;
      PlanNew:=TPlanAmort.Create(true) ;
      try
         PlanNew.Charge(QPlan);
         PlanNew.Recupere(fCode_immo,QPlan.FindField('I_PLANACTIF').AsString);
         fPlanActifAv := PlanNew.NumSeq;
         PlanNew.SetTypeOpe('DPR');
         PlanNew.Calcul(Qplan, idate1900);
         PlanNew.Sauve;
         // Attribuer à l'immo le n° du nouveau plan d'amortissement
         Qplan.FindField('I_PLANACTIF').AsInteger := PlanNew.NumSeq;
         QPlan.FindField('I_DATEDERMVTECO').AsDateTime := PlanNew.GetDateFinAmortEx(PlanNew.AmortEco);
         QPlan.FindField('I_DATEDERNMVTFISC').AsDateTime := PlanNew.GetDateFinAmortEx(PlanNew.AmortFisc);
         Qplan.Post;
         fPlanActifAp := PlanNew.NumSeq;
      finally
         PlanNew.Free;
         Ferme(Qplan);
      end;

      // Mise à jour de IMMOLOG
      Ordre := TrouveNumeroOrdreLogSuivant(fPlan.CodeImmo);

      TLog := TOB.Create ('IMMOLOG',nil,-1);
      try
         TLog.PutValue('IL_IMMO',fCode_immo);
         TLog.PutValue('IL_LIBELLE',lib);
         TLog.PutValue('IL_TYPEMODIF',AffecteCommentaireOperation('DPR'));
         TLog.PutValue('IL_DATEOP', StrToDate(GetControlText('DP_DATEOPE')));
         TLog.PutValue('IL_TYPEOP', 'DPR');
         TLog.PutValue('IL_ORDRE', Ordre);
         TLog.PutValue('IL_OrdreSerie', -1);
         TLog.PutValue('IL_TYPEDOT', 'DPR');
         TLog.PutValue('IL_MONTANTDOT', fMntDeprec);
         TLog.PutValue('IL_PLANACTIFAV', fPlanActifAv);
         TLog.PutValue('IL_PLANACTIFAP', fPlanActifAp);
         TLog.PutValue('IL_BLOCNOTE', RichToString (fDP_Blocnote));
         TLog.InsertDB(nil);
      finally
         TLog.Free;
      end;

      COMMITTRANS ;

      // 05/06 FQ 18119 Positionner les indicateurs de modif
      VHImmo^.ChargeOBImmo := True;
      ImMarquerPublifi (True);

   except
      PGIBox(HM[6]);
      ROLLBACK ;
   end ;

   TFVIERGE(ECRAN).retour := IntToStr(MrYes);
end;




Initialization
  registerclasses ( [ TOF_AMDEPRECIATION ] ) ;
end.









