{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 06/11/2006
Modifié le ... : 05/12/2006
Description .. : Source TOF de la FICHE : AMREDUCSBV ()
Suite ........ : FQ 19273 Durée échelonnement mal affichée
Suite ........ : FQ 19274 Subvention doit être < ancien montant-amortissements déjà dégagés
Suite ........ : FQ 19281
Mots clefs ... : TOF;AMREDUCSBV
*****************************************************************}
Unit AMREDUCSBV_TOF ;

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
     MaineAGL,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     ImEnt,
     ImPlan,
     ImPlanMeth,
     HMsgBox,
     HTB97,
     outils,
     HRichOle,
     Vierge,
     UTOF ;

Type
  TOF_AMREDUCSBV = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
      fCode:string;
      fPlan:TPlanAmort;
      fANTOLD : double;
      fDateDeb : TDateTime;
      fReduction : double;
      fReprise : double;
      fBlocNote : THRichEditOle;
      procedure OnbFermeClick(Sender: TObject);
      procedure OnExitReduction(Sender: TObject);
      procedure OnExitReprise(Sender: TObject);
      function  ControleDateOpe : boolean;
      function  ControleReduction : boolean;
      function  ControleReprise : boolean;
      function  ControleZones : boolean;
      //procedure RecalculRepriseTheorique; //XVI Conseil Compile...
      procedure EnregistrerOpe ;
  end ;


function AMLanceFiche_AMReducSubvention(stParam: string) : integer;


Implementation

const
  HMD: array[1..8] of string =
   {1} ('Voulez-vous enregistrer cette opération ?',
   {2}  'La date de l''opération est incorrecte.',
   {3}  'La date d''opération doit être une date de l''exercice en cours',
   {4}  'Une opération a déjà été effectuée à une date postérieure.',
   {5}  'Le montant de la réduction ne peut pas dépasser la subvention, une fois retirés les amortissements déjà pratiqués.',
   {6}  'Vous n''avez saisi aucun montant.',
   {7}  'L''enregistrement de la réduction de la subvention d''investissement a échoué.',
   {8}  'La reprise déjà pratiquée ne peut pas excéder celle déjà enregistrée.'
    );

function AMLanceFiche_AMReducSubvention(stParam: string) : integer;
begin
 result:= StrToInt(AGLLanceFiche('AM', 'AMREDUCSBV','' , '', stParam));
end;


procedure TOF_AMREDUCSBV.OnExitReduction(Sender: TObject);
var Reduc : double;
begin
   Reduc := Valeur (GetControlText('Reduc'));
   if  (Reduc <> fReduction)
   and (Reduc <> 0)
   and (Reduc <= Valeur (GetControlText('SBVOLD'))) then
       begin
       fReduction := Reduc;
       // FQ 19281 RecalculRepriseTheorique;
       SetControlCaption('SBV', StrFMontant((fPlan.MntSBV-fReduction), 15, V_PGI.OkDecV, '', True));
       end;
end;

procedure TOF_AMREDUCSBV.OnExitReprise(Sender: TObject);
begin
   if (Valeur (GetControlText('Reprise')) <> fReprise) then
      begin
      fReprise := Valeur (GetControlText('Reprise'));
      end;
end;

function TOF_AMREDUCSBV.ControleDateOpe : boolean;
var Q : TQuery;
    stop : boolean;
begin
  Result := True;
  // La date est-elle valide ?
  if (not IsValidDate (GetControlText('DateOpe'))) then
  begin
    PGIBox (HMD[2]);
    SetFocusControl('DateOpe');
    Result := false; Exit;
  end;
  // La date est-elle un date de l'exercice en cours ?
  if (StrToDate(GetControlText('DateOpe')) < VHImmo^.Encours.Deb)
  or (StrToDate(GetControlText('DateOpe')) > VHImmo^.Encours.Fin) then
  begin
    PGIBox (HMD[3]);
    SetFocusControl('DateOpe');
    Result := false;  Exit;
  end;
  // La date est-elle postérieure à la date de dernière opération ?
  stop := False;
  Q := OpenSQL ('SELECT IL_DATEOP FROM IMMOLOG WHERE IL_IMMO="'
                 + fCode +'" ORDER BY IL_IMMO,Il_DATEOP DESC', True);
  if not Q.Eof then
     if Q.FindField ('IL_DATEOP').AsDateTime > StrToDate(GetControlText('DateOpe')) then
        stop:= True;
     Ferme (Q);
  if stop then
    begin
    PGIBox (HMD[4]);
    SetFocusControl('DateOpe');
    Result := false;
    end;
end;

function TOF_AMREDUCSBV.ControleReduction : boolean;
begin
   result := True;
   if Valeur (GetControlText('Reduc')) < 0 then
         begin
         PGIInfo('Montant négatif !','Attention');
         SetFocusControl ('Reduc');
         Result := False; Exit;
         end;

   if Valeur (GetControlText('Reduc')) = 0 then
         begin
         PGIBox (HMD[6]);
         SetFocusControl ('Reduc');
         Result := False; Exit;
         end;
   // FQ 19274
   //if Valeur (GetControlText('Reduc')) > Arrondi(Valeur(GetControlText('SBVOLD')
   if Valeur (GetControlText('Reduc')) >
      Arrondi(Valeur(GetControlText('SBVOLD'))-Valeur(GetControlText('RepriseOLD')), V_PGI.OkDecV) then
         begin
         PGIBox (HMD[5]);
         SetFocusControl ('Reduc');
         Result := False; Exit;
         end;
end;

function TOF_AMREDUCSBV.ControleReprise : boolean;
var MtRep :double;
begin

   Result := True;
   exit; // FQ 19281
   MtRep :=  Valeur (GetControlText('Reprise'));
   if MtRep < 0 then
         begin
         SetActiveTabSheet('SheetNEW');
         SetFocusControl ('Reprise');
         PGIInfo('Montant négatif !','Attention');
         Exit;
         end;
   if MtRep > Valeur (GetControlText('RepriseOLD')) then
         begin
         SetActiveTabSheet('SheetNEW');
         SetFocusControl ('Reprise');
         PGIBox(HMD[8]);
         Exit;
         end;
   Result := True;
end;

function TOF_AMREDUCSBV.ControleZones : boolean;
begin
  result := False;

  if ControleDateOpe then
     begin
     if ControleReduction  then
        Result := ControleReprise;
     end;
end;

(*procedure TOF_AMREDUCSBV.RecalculRepriseTheorique;
var TPlan : TPlanAmort;
    Q : TQuery;
    PresencePRI : boolean;
begin

  // Màj champs
  SetControlCaption('SBV', StrFMontant((fPlan.MntSBV-fReduction), 15, V_PGI.OkDecV, '', True));

  TPlan:= TPlanAmort.Create(true) ;
  Q:= OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="' + fCode + '"', False) ;

  try
    TPlan.Charge(Q);

    PresencePRI :=  (Q.FindField('I_SBVPRI').AsFloat <> 0);

    // Alimenter AmortSBV avec les éléments nécessaires au calcul
    TPlan.AmortSBV.Base := Valeur(GetControlText('SBV'));
   // TPlan.SBVDateDeb := fDateDeb;
   // TPlan.SBVDuree := fPlan.SBVDuree;
    TPlan.AmortSBV.Creation := True;
    if PresencePRI then
       TPlan.AmortPRI.Creation := True;

    // Calcul
    TPlan.CalculDateFinAmortissement(TPlan.AmortSBV);
    TPlan.CalculReprises;
    fReprise := TPlan.AmortSBV.Reprise;
    SetControlText('Reprise', StrFMontant(fReprise, 15, V_PGI.OkDecV, '', True));

  finally
    Ferme(Q);
  end;
  TPlan.Free;
end; *) //XVI Conseil Compile...


procedure TOF_AMREDUCSBV.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AMREDUCSBV.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AMREDUCSBV.OnUpdate ;
begin
  Inherited ;
  // Ne pas sortir a priori
  Ecran.ModalResult := mrNone;

  if ControleZones then
     EnregistrerOpe;
end ;

procedure TOF_AMREDUCSBV.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AMREDUCSBV.OnArgument (S : String ) ;
var QPlan : TQuery;
    RepriseOLD : double;
begin
  Inherited ;
  TFVIERGE(ECRAN).retour := IntToStr(MrNo);

  // Formatage des champs
  SetControlProperty('Reduc', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('Reprise', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  // FQ 19281
  SetControlVisible('Reprise', False);
  SetControlVisible('LiReprise', False);
  fBlocNote := THRichEditOle (GetControl('BN'));

  // Evènements
  TToolbarButton97(GetControl('BFerme')).Onclick := OnbFermeClick;
  THEdit(GetControl('Reduc')).OnExit    := OnExitReduction;
  THEdit(GetControl('Reprise')).OnExit   := OnExitReprise;

  // Initialisation des champs
  fCode := S;
  fReprise := 0;
  fReduction := 0;
  fPlan:= TPlanAmort.Create(true) ;
  QPlan:= OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="' + fCode + '"', FALSE) ;
  try
    fPlan.Charge(QPlan);
    fPlan.Recupere(fCode,QPlan.FindField('I_PLANACTIF').AsString);

    // Récup date début calcul
    if fPlan.Fiscal = True then
       fDateDeb := fPlan.DateDebFis
    else
       if fPlan.AmortEco.Methode = 'NAM' then
         fDateDeb := fPlan.DateMiseEnService
       else
         fDateDeb := fPlan.DateDebEco;

    // Récup durée subvention
    if fPlan.AmortEco.Methode = 'NAM' then
      begin
      SetControlCaption('LiDureeOLD', 'Durée d''inaliénabilité');
      SetControlCaption('DureeOLD', IntToStr(fPlan.SBVDuree) + ' mois')
      end
    else
      begin
      SetControlCaption('LiDureeOLD', 'Durée d''échelonnement');
      SetControlCaption('DureeOLD', IntToStr(fPlan.SBVDuree) + ' mois') // FQ 19273 'EchelonOld'
      end;

    // Récup antérieurs saisis + ceux des exos précédents
    fPlan.GetCumulSBV (VHImmo^.Encours.Deb-1, RepriseOLD, False, True, False);
    // Stocker antérieurs saisis
    fANTOLD := fPlan.AmortSbv.Reprise;

    // Champs connus
    SetControlCaption('PCode', fPlan.Codeimmo);
    SetControlCaption('PDesi', fPlan.LibelleImmo);
    SetControlCaption('PAchat', DateToStr(fPlan.DateAchat));
    SetControlCaption('PValeurHT', StrFMontant(fPlan.ValeurHT, 15, V_PGI.OkDecV, '', True));
    SetControlCaption('SBVOLD', StrFMontant(fPlan.MntSBV, 15, V_PGI.OkDecV, '', True));
    SetControlCaption('RepriseOLD', StrFMontant(RepriseOLD, 15, V_PGI.OkDecV, '', True));
    SetControlCaption('SBV', StrFMontant((fPlan.MntSBV-fReduction), 15, V_PGI.OkDecV, '', True));

  finally
    Ferme(Qplan);
  end;
end ;

procedure TOF_AMREDUCSBV.OnClose ;
begin
  Inherited ;
  fPlan.Free;
end ;

procedure TOF_AMREDUCSBV.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AMREDUCSBV.OnCancel () ;
begin
  Inherited ;
end ;

// Action Echap ou Fermer par la croix
procedure TOF_AMREDUCSBV.OnbFermeClick(Sender: TObject);
begin

  if (Valeur(GetControlText('Reduc')) <> fReduction) or
     (Valeur(GetControlText('Reprise')) <> fReprise)   then

    case PGIAskCancel(HMD[1], Ecran.Caption) of
    mrYes:
      begin
      if ControleZones then
           // OK enregistrer et sortir
           EnregistrerOpe
      else Ecran.ModalResult := mrNone; // Rester dans l'écran
      end;
    mrNo:
      // Sortir
      Ecran.ModalResult := mrNo;
    else
      // Rester dans l'écran
      Ecran.ModalResult := mrNone;
    end
  else
    // pas de modifs => sortir
    Ecran.ModalResult := mrNo;
end;

procedure TOF_AMREDUCSBV.EnregistrerOpe ;
var Ordre, PlanActifAv, PlanActifAp : integer;
    TLog : TOB;
    Lib :String;
    QPlan : TQuery;
    PlanNew : TPlanAmort;
    PremReprise : string;
begin
  // Valider la sortie
  Ecran.ModalResult := mrYes;
  TFVierge(Ecran).retour := IntToStr (mrYes);

    BEGINTRANS ;
  try
      QPlan:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="' + fCode + '"', FALSE) ;

      // Maj de la fiche immo
      QPlan.Edit;

      // ajout mbo 13.11.06 pour stockage dans immolog
      PremReprise := QPlan.FindField('I_DPIEC').AsString;

      QPlan.FindField('I_OPERATION').AsString:= 'X' ;
      QPlan.FindField('I_SBVMT').AsFloat:= Valeur(GetControlText('SBV'));
      // MVG 06/12/2006 FQ 19281
      QPlan.FIndField('I_CORRECTIONVR').AsFloat:= Valeur(GetControlText('REPRISEOLD'));
      QPlan.FindField('I_DPIEC').AsString:= '-';
      QPlan.Post;

      // Maj du plan d'amortissement IMMOAMOR
      QPlan.Edit;
      PlanNew:=TPlanAmort.Create(true) ;
      try
         PlanNew.Charge(QPlan);
         PlanNew.Recupere(fCode, QPlan.FindField('I_PLANACTIF').AsString);
         // RAZ TableauDot qui contient les anciennes valeurs
         PlanNew.ResetTableauDot(PlanNew.AmortSBV, PlanNew.GetDateDebutAmort (PlanNew.AmortSBV) );

         PlanActifAv := PlanNew.NumSeq;
         PlanNew.SetTypeOpe('RSB');
         PlanNew.Calcul(Qplan, idate1900);
         PlanNew.Sauve;

         // Attribuer à l'immo le n° du nouveau plan d'amortissement
         Qplan.FindField('I_PLANACTIF').AsInteger := PlanNew.NumSeq;

         Qplan.Post;
         PlanActifAp := PlanNew.NumSeq;
      finally
         PlanNew.Free;
         Ferme(Qplan);
      end;


      // Maj de IMMOLOG
      Lib := 'Réduction d''une subvention d''investissement';
      Ordre := TrouveNumeroOrdreLogSuivant(fCode);

      TLog := TOB.Create ('IMMOLOG',nil,-1);
      try
         TLog.PutValue('IL_IMMO', fCode);
         TLog.PutValue('IL_LIBELLE', Lib);
         TLog.PutValue('IL_TYPEMODIF', AffecteCommentaireOperation('RSB'));

         TLog.PutValue('IL_DATEOP', StrToDate(GetControlText('DateOpe')));

         // ancienne subvention
         TLog.PutValue('IL_MONTANTEXC', GetControlText('SBVOLD'));
         // ancienne reprise
         TLog.PutValue('IL_MONTANTAVMB', Arrondi (fANTOLD, V_PGI.OkDecV));

         TLog.PutValue('IL_TYPEOP', 'RSB');
         TLog.PutValue('IL_ORDRE', Ordre);
         TLog.PutValue('IL_OrdreSerie', -1);
         TLog.PutValue('IL_PLANACTIFAV', PlanActifAv);
         TLog.PutValue('IL_PLANACTIFAP', PlanActifAp);

         // ajout mbo 13.11.06 pour stocker le fait que l'on est en reprise s/ 1ère dotation
         TLog.PutValue('IL_CODECB', PremReprise);

         TLog.PutValue('IL_BLOCNOTE', RichToString (fBlocnote));
         TLog.InsertDB(nil);
      finally
         TLog.Free;
      end;

      COMMITTRANS ;

      // Positionner les indicateurs de modif
      VHImmo^.ChargeOBImmo := True;
      ImMarquerPublifi (True);

  except
    PGIBox(HMd[7]);
    ROLLBACK ;
  end ;


end;

Initialization
  registerclasses ( [ TOF_AMREDUCSBV ] ) ;
end.
