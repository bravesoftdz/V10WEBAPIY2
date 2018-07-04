{***********UNITE*************************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 09/10/2006
Modifié le ... : 20/10/2006
Description .. : Source TOF de la FICHE : AMREDUCPRIME ()
Suite ........ : Ajout message 8
Mots clefs ... : TOF;AMREDUCPRIME
*****************************************************************}
Unit AMREDUCPRIME_TOF ;

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
     HMsgBox,
     HTB97,
     ImEnt,
     ImPlan,
     ImOuPlan,
     HRichOle,
     outils,
     Vierge,
     UTOF ;

Type
  TOF_AMREDUCPRIME = Class (TOF)
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
      fDateDeb : TDateTime;
      fPrimeOLD : double;
      fRepriseOLD : double;
      fAntOLD : double;
      fReduction : double;
      fAntNEW : double;
      fBlocNote : THRichEditOle;
      procedure AfficheNouvellesValeurs;
      procedure OnbFermeClick(Sender: TObject);
      procedure OnExitReduction(Sender: TObject);
      procedure OnExitAntNEW(Sender: TObject);
      function  ControleDateOpe : boolean;
      function  ControleReduction : boolean;
      function  ControleAntNEW : boolean;
      function  ControleZones : boolean;
      procedure RecalculSuramortissementTheorique;
      procedure EnregistrerOpe ;
  end ;

function AMLanceFiche_AMReducPrime(stParam: string) : integer;

Implementation

const
  HMD: array[1..10] of string =
   {1} ('Voulez-vous enregistrer cette opération ?',
   {2}  'La date de l''opération est incorrecte.',
   {3}  'La date d''opération doit être une date de l''exercice en cours',
   {4}  'Une opération a déjà été effectuée à une date postérieure.',
   {5}  'Le montant de la réduction ne peut pas être supérieur à la prime.',
   {6}  'Vous n''avez saisi aucun montant.',
   {7}  'L''enregistrement de la réduction de la prime d''équipement a échoué.',
   {8}  'Le suramortissement déjà pratiqué ne doit pas être supérieur à la nouvelle base.',
   {9}  'Le suramortissement déjà pratiqué ne peut pas excéder celui déjà enregistré ',
  {10}  'Après l''enregistrement de la réduction, il sera nécessaire ' + #10#13 +
        ' de réintégrer manuellement le montant de '
       );

function AMLanceFiche_AMReducPrime(stParam: string) : integer;
begin
 result:= StrToInt(AGLLanceFiche('AM', 'AMREDUCPRIME','' , '', stParam));
end;

procedure TOF_AMREDUCPRIME.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AMREDUCPRIME.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AMREDUCPRIME.OnUpdate ;
begin
  Inherited ;
  // Ne pas sortir a priori
  Ecran.ModalResult := mrNone;

  if ControleZones then
     EnregistrerOpe;
end ;

procedure TOF_AMREDUCPRIME.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AMREDUCPRIME.OnArgument (S : String ) ;
var QPlan : TQuery;
begin
  Inherited ;
  TFVIERGE(ECRAN).retour := IntToStr(MrNo);

  // Formatage des champs
  SetControlProperty('PReduc', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('PAntNEW', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  fBlocNote := THRichEditOle (GetControl('PBN'));

  // Evènements
  TToolbarButton97(GetControl('BFerme')).Onclick := OnbFermeClick;
  THEdit(GetControl('PReduc')).OnExit    := OnExitReduction;
  THEdit(GetControl('PAntNEW')).OnExit   := OnExitAntNEW;

  // Initialisation des champs
  fPlan:= TPlanAmort.Create(true) ;
  fCode := S;
  QPlan:= OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="' + fCode + '"', FALSE) ;
  try
    fPlan.Charge(QPlan);
    fPlan.Recupere(fCode,QPlan.FindField('I_PLANACTIF').AsString);

    // Récup date début du calcul de suramortissement
    if fPlan.Fiscal = true then
       fDateDeb := fPlan.DateDebFis
    else
       if fPlan.AmortEco.Methode = 'NAM' then
          fDateDeb := fPlan.DateMiseEnService
       else
          fDateDeb := fPlan.DateDebEco;

    // Récup antérieurs saisis + ceux des exos précédents
    fPlan.GetCumulPRI(VHImmo^.Encours.Deb-1, fAntOLD, False, True, False);
    fPrimeOLD := fPlan.MntPrime;
    // Récup antérieurs saisis
    fRepriseOLD := fPlan.AmortPri.Reprise;

    SetControlCaption('PCode',fPlan.Codeimmo);
    SetControlCaption('PDesi',fPlan.LibelleImmo);
    SetControlCaption('PAchat', DateToStr(fPlan.DateAchat));
    SetControlCaption('PValeurHT',StrFMontant(fPlan.ValeurHT, 15, V_PGI.OkDecV, '', True));
    SetControlCaption('PPrimeOLD',StrFMontant(fPrimeOLD, 15, V_PGI.OkDecV, '', True));
    SetControlCaption('PBaseOLD',StrFMontant((fPrimeOLD/2), 15, V_PGI.OkDecV, '', True));
    SetControlCaption('PAntOLD',StrFMontant(fAntOLD, 15, V_PGI.OkDecV, '', True));
    if (QPlan.FindField('I_METHODEECO').AsString = 'NAM') then
      SetControlCaption('PDureeOLD', IntToStr(fPlan.PRIDuree) +' mois')
    else
      begin
      SetControlVisible('PDureeOLD', False);
      SetControlVisible('PLiDuree', False);
      end;

  finally                                 
     Ferme(Qplan);
  end;

  AfficheNouvellesValeurs;
  SetFocusControl('PDateOpe');
end ;

procedure TOF_AMREDUCPRIME.OnClose ;
begin
  Inherited ;
  fPlan.Free;
end ;

procedure TOF_AMREDUCPRIME.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AMREDUCPRIME.OnCancel () ;
begin
  Inherited ;
end ;

// Action Echap ou Fermer par la croix
procedure TOF_AMREDUCPRIME.OnbFermeClick(Sender: TObject);
begin

  if (Valeur(GetControlText('PReduc')) <> fReduction) or
     (Valeur(GetControlText('PAntNEW')) <> fAntNEW)   then

    case PGIAskCancel(HMD[1], Ecran.Caption) of
    mrYes:
      begin
      if ControleZones then
           // OK enregistrer et sortir
           EnregistrerOpe
      else Ecran.ModalResult := mrNone; // Rester dans l'écran
      end;
    mrNo:
      begin
      // Sortir
      Ecran.ModalResult := mrNo;
      end;
    else
      // Rester dans l'écran
      Ecran.ModalResult := mrNone;
    end

  else
  // pas de modifs => sortir
  Ecran.ModalResult := mrNo;

end;

procedure TOF_AMREDUCPRIME.OnExitReduction(Sender: TObject);
var Reduc : double;
begin
   {if ControleReduction then fReducErr :=False
   else
      begin
      fReducErr :=True;
      SetFocusControl('PReduc');
      end; }

   Reduc := Valeur (GetControlText('PReduc'));
   if  (Reduc <> fReduction)
   and (Reduc <> 0)
   and (Reduc <= Valeur (GetControlText('PPrimeOLD'))) then
       begin
       fReduction := Reduc;
       RecalculSuramortissementTheorique;
       AfficheNouvellesValeurs;
       end;
end;

procedure TOF_AMREDUCPRIME.OnExitAntNEW(Sender: TObject);
var AntNEW : double;
begin
   {if ControleAntNEW then fAntNEWErr :=False
   else
      begin
      fAntNEWErr :=True;
      SetFocusControl('PAntNEW');
      end; }

   AntNEW := Valeur (GetControlText('PAntNEW'));
   if (AntNEW <> fAntNEW) then
      begin
      fAntNEW := AntNEW;
      SetControlCaption('PAntReintegrer',
         StrFMontant((fAntOLD-fAntNEW), 15, V_PGI.OkDecV, '', True));
      end;
end;


function TOF_AMREDUCPRIME.ControleDateOpe : boolean;
var Q : TQuery;
    stop : boolean;
begin
  Result := True;
  // La date est-elle valide ?
  if (not IsValidDate (GetControlText('PDateOpe'))) then
  begin
    PGIBox (HMD[2]);
    SetFocusControl('PDateOpe');
    Result := false;  Exit;
  end;
  // La date est-elle un date de l'exercice en cours ?
  if (StrToDate(GetControlText('PDateOpe')) < VHImmo^.Encours.Deb)
  or (StrToDate(GetControlText('PDateOpe')) > VHImmo^.Encours.Fin) then
  begin
    PGIBox (HMD[3]);
    SetFocusControl('PDateOpe');
    Result := false;  Exit;
  end;
  // La date est-elle postérieure à la date de dernière opération ?
  stop := False;
  Q := OpenSQL ('SELECT IL_DATEOP FROM IMMOLOG WHERE IL_IMMO="'
                 + fCode +'" ORDER BY IL_IMMO,Il_DATEOP DESC', True);
  if not Q.Eof then
     if Q.FindField ('IL_DATEOP').AsDateTime > StrToDate(GetControlText('PDateOpe')) then
        stop:= True;
     Ferme (Q);
  if stop then
    begin
    PGIBox (HMD[4]);
    SetFocusControl('PDateOpe');
    Result := false;
    end;
end;

function TOF_AMREDUCPRIME.ControleReduction : boolean;
begin
   result := True;
   if Valeur (GetControlText('PReduc')) < 0 then
         begin
         PGIInfo('Montant négatif !','Attention');
         SetFocusControl ('PReduc');
         Result := False; Exit;
         end;

   if Valeur (GetControlText('PReduc')) = 0 then
         begin
         PGIBox (HMD[6]);
         SetFocusControl ('PReduc');
         Result := False; Exit;
         end;

   if Valeur (GetControlText('PReduc')) > Valeur (GetControlText('PPrimeOLD')) then
         begin
         PGIBox (HMD[5]);
         SetFocusControl ('PReduc');
         Result := False; Exit;
         end;

{   if (fReduction <> Valeur (GetControlText('PReduc'))) then
     // Proposer antérieur prime théorique
     RecalculSuramortissementTheorique;
   // Réduction OK
   fReduction := Valeur (GetControlText('PReduc'));
   // Réafficher pavé Nouvelles Valeurs
   AfficheNouvellesValeurs; }
end;

function TOF_AMREDUCPRIME.ControleAntNEW : boolean;
var MtAntNew : double;
begin
   result := False;
   MtAntNew :=  Valeur (GetControlText('PAntNEW'));
   if MtAntNew < 0 then
         begin
         SetActiveTabSheet('PSheetNEW');
         SetFocusControl ('PAntNEW');
         PGIInfo('Montant négatif !','Attention');
         Exit;
         end;
   if MtAntNew > Valeur (GetControlText('PBaseNEW')) then
         begin
         SetActiveTabSheet('PSheetNEW');
         SetFocusControl ('PAntNEW');
         PGIBox(HMD[8]);
         Exit;
         end;
   if MtAntNew > fAntOLD then
         begin
         SetActiveTabSheet('PSheetNEW');
         SetFocusControl ('PAntNEW');
         PGIBox(HMD[9] + StrFMontant(fAntOLD,15,V_PGI.OkDecV,'',True) + '.');
         Exit;
         end;

   Result := True;
   {// Suramortissement OK => réafficher pavé Nouvelles Valeurs
   fAntNEW := MtAntNew;
   AfficheNouvellesValeurs; }
end;

procedure TOF_AMREDUCPRIME.RecalculSuramortissementTheorique;
var TPlan : TPlanAmort;
    Q : TQuery;
    PresenceSBV : boolean;
begin

  // Màj champs
  SetControlCaption('PPrimeNEW',
    StrFMontant((fPrimeOLD-fReduction), 15, V_PGI.OkDecV, '', True));
  SetControlCaption('PBaseNEW',
    StrFMontant(((fPrimeOLD-fReduction)/2), 15, V_PGI.OkDecV, '', True));


  TPlan:= TPlanAmort.Create(true) ;
  Q:= OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="' + fCode + '"', False) ;

  try
    TPlan.Charge(Q);

    PresenceSBV := (Q.FindField('I_SBVMT').AsFloat <> 0);

    TPlan.AmortPri.Base := Valeur(GetControlText('PBaseNEW'));
    TPlan.PRIDateDeb := fDateDeb;
    TPlan.AmortPRI.Creation := True;

    if PresenceSBV then TPlan.AmortSBV.Creation := true;  // ajout mbo pour pb sur immo déjà cloturée

    TPlan.CalculDateFinAmortissement(TPlan.AmortPri);
    TPlan.CalculReprises;
    SetControlText('PAntNew', StrFMontant(TPlan.AmortPri.Reprise, 15, V_PGI.OkDecV, '', True));
    fAntNEW :=  Valeur (GetControlText('PAntNEW'));

  finally
    Ferme(Q);
  end;
  TPlan.Free;
end;


procedure TOF_AMREDUCPRIME.AfficheNouvellesValeurs;
begin
  SetControlCaption('PPrimeNEW',
     StrFMontant((fPrimeOLD-fReduction), 15, V_PGI.OkDecV, '', True));
  SetControlCaption('PBaseNEW',
     StrFMontant(((fPrimeOLD-fReduction)/2), 15, V_PGI.OkDecV, '', True));
  SetControlCaption('PAntReintegrer',
     StrFMontant((fAntOLD-fAntNEW), 15, V_PGI.OkDecV, '', True));
end;


function TOF_AMREDUCPRIME.ControleZones : boolean;
begin
  result := False;

  // Aucun montant saisi
 { if (Valeur(GetControlText('PReduc')) = 0) then
     begin
     PGIBox (HMD[6]);
     Ecran.FocusControl (THEdit(GetControl('PReduc')));
     exit;
     end;  }

  if ControleDateOpe then
     begin
     if ControleReduction  then
        begin
        Result := ControleAntNEW;
        end;
     end;
end;

procedure TOF_AMREDUCPRIME.EnregistrerOpe ;
var Ordre, PlanActifAv, PlanActifAp : integer;
    TLog : TOB;
    Lib :String;
    QPlan : TQuery;
    PlanNew : TPlanAmort;
begin
  // Valider la sortie
  Ecran.ModalResult := mrYes;
  TFVierge(Ecran).retour := IntToStr (mrYes);

  // Informer l'utilisateur de la réintégration manuelle
  if Valeur (GetControlText('PAntReintegrer')) <> 0 then
     PGIBox(HMD[10] + StrFMontant(Valeur(GetControlText('PAntReintegrer')),15,V_PGI.OkDecV,'',True) + '.');

  BEGINTRANS ;
  try
      QPlan:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="' + fCode + '"', FALSE) ;

      // Maj de la fiche immo
      QPlan.Edit;
      QPlan.FindField('I_OPERATION').AsString:= 'X' ;
      QPlan.FindField('I_SBVPRI').AsFloat:= Valeur(GetControlText('PPrimeNEW'));
      QPlan.FIndField('I_REPRISEUO').AsFloat:= Valeur(GetControlText('PAntNEW'));
      QPlan.Post;

      // Maj du plan d'amortissement IMMOAMOR
      QPlan.Edit;
      PlanNew:=TPlanAmort.Create(true) ;
      try
         PlanNew.Charge(QPlan);
         PlanNew.Recupere(fCode, QPlan.FindField('I_PLANACTIF').AsString);
         // RAZ TableauDot qui contient les anciennes valeurs
         PlanNew.ResetTableauDot(PlanNew.AmortPri, PlanNew.GetDateDebutAmort (PlanNew.AmortPri) );

         PlanActifAv := PlanNew.NumSeq;
         PlanNew.SetTypeOpe('RPR');
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
      Lib := 'Réduction d''une prime d''équipement';
      Ordre := TrouveNumeroOrdreLogSuivant(fCode);

      TLog := TOB.Create ('IMMOLOG',nil,-1);
      try
         TLog.PutValue('IL_IMMO', fCode);
         TLog.PutValue('IL_LIBELLE', Lib);
         TLog.PutValue('IL_TYPEMODIF', AffecteCommentaireOperation('RPR'));

         TLog.PutValue('IL_DATEOP', StrToDate(GetControlText('PDateOpe')));

         // ancienne prime
         TLog.PutValue('IL_MONTANTEXC', GetControlText('PPrimeOLD'));
         // ancienne base
         TLog.PutValue('IL_BASETAXEPRO',
                        Arrondi((Valeur(GetControlText('PPrimeOLD'))/2),V_PGI.OkDecV)
                        );
         // ancien cumul antérieurs hors ceux des exos précédents
         TLog.PutValue('IL_MONTANTAVMB', Arrondi (fRepriseOLD, V_PGI.OkDecV));

         TLog.PutValue('IL_TYPEOP', 'RPR');
         TLog.PutValue('IL_ORDRE', Ordre);
         TLog.PutValue('IL_OrdreSerie', -1);
         TLog.PutValue('IL_PLANACTIFAV', PlanActifAv);
         TLog.PutValue('IL_PLANACTIFAP', PlanActifAp);
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
  registerclasses ( [ TOF_AMREDUCPRIME ] ) ;
end.
