{-------------------------------------------------------------------------------------
    Version   | Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
               19/03/03  GCO  Création de l'unité : Tom de la Table EEXBQ (Référence / Session de pointage)
 8.01.001.004  05/02/07  JP   Uniformisation de l'utilisation de la Table EEXBQ dans le pointage et le
                              rapprochement : Maintenant, découplage de EEXBQ et EEXBQLIG, EEXBQ servant
                              uniquement au session de pointage et EEXBQLIG est un pendant à écriture
                              pour les mouvements bancaires. Lorsque l'on verse le contenu de CETEBAC dans
                              EEXBQLIG, on rattache les mvts bancaires à EEXBQ par CEL_NUMRELEVE = EE_NUMRELEVE,
                              mais maintenant la référence et la date de pointage peuvent être différents.
                              Par ailleurs, en pointage manuel, il est maintenant possible de créer des enregistrements
                              dans EEXBQLIG afin d'équilibrer un état de rappro en cas de désaccord avec la banque
 8.01.001.012  24/04/07  JP   FQ TRESO 10441 : Gestion du Is Null pour Oracle sur CET_IMO
 8.01.001.012  24/04/07  JP   FQ 20040 : Gestion du format en fonction de la Devise
 8.01.001.012  26/04/07  JP   FQ TRESO 10445 : Activation des zones soldes en pointage manuel
 8.01.001.013  04/05/07  JP   FQ 20240 : en eAGL, en retour du lookup sur le compte le champ ee_general n'est pas à jour
 8.00.001.020  19/06/07  JP   FQ TRESO 10479 : récupération de la clef primaire complète si l'on a que le compte et le numéro
 8.00.001.020  19/06/07  JP   FQ TRESO 10477 : désactivation toujours les zones de solde pour les relevés automatiques
 8.00.001.022  28/06/07  JP   FQ 20879 : effets de bord des suite de la FQ 10477
 8.00.001.025  18/07/07  JP   FQ 21119 : Raccourci F6 pour afficher la liste des relevés
 8.00.001.025  19/07/07  JP   FQ 20603 : Si à la validation le focus est sur la référence ou sur le général et que l'on
                              a forcé une référence, celle-est écrasée
 8.10.001.010  19/09/07  JP   FQ 21361 : Réinitialisation des montants, si on vient de changer de relevé
 8.10.001.013  08/10/07  JP   FQ TRESO 10528 : en CWAS, il y avait des problèmes de focus et d'exécution des OnExit quand
                              on utilisait les racourcis claviers ou sur les clicks de bitbuttons
 8.10.002.002  30/10/07  JP   FQ 21763 : Les pièces de la saisie automatique ne sont pas correctement pointées
 8.10.004.001  08/11/07  JP   FQ 21788 : Pas de contrôle sur les soldes à zéro lors du choix d'un relevé
08.10.005.001  14/11/07  JP   Gestion des comptes pointables qui ne sont pas bancaires  
--------------------------------------------------------------------------------------}
unit uTOMEEXBQ;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  Maineagl, eFiche,
  {$ELSE}
  Fe_main, db, Hdb, Fiche,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  uTOB, SysUtils, ComCtrls, HCtrls, HEnt1, UTOM, Htb97,
  ULibEcriture, ZEteBac,   ZEtebacGuide, ZEtebacVersEcr;

type
  TOM_EEXBQ = class(TOM)
    procedure OnNewRecord              ; override;
    procedure OnDeleteRecord           ; override;
    procedure OnUpdateRecord           ; override;
    procedure OnAfterUpdateRecord      ; override;
    procedure OnLoadRecord             ; override;
    procedure OnChangeField(F : TField); override;
    procedure OnArgument   (S : string); override;
    procedure OnClose                  ; override;
  private
    cTitreCreation        : string;
    FDtDernDatePointage   : TDateTime;
    FBoOkDernStatutReleve : Boolean;
    FStDevise             : string;
    FNbDecimales          : Integer; {FQ 20040}
    FNoDossier            : string;
    FBoConsultation       : Boolean;
    FBoCreationEnreg      : Boolean;
    FBoPtgeNonBQ          : Boolean; {14/11/07 : Si pointage sur TRECRITURE et que l'on est en compta}
    FBoRefNonBq           : Boolean; {14/11/07 : Si on choisit un compte non bancaire et que l'on est en compta}
    FGeneral              : string;
    FDeviseConvertie      : string; {Pour gérer le cohérence CodeIso et CodeDevise}
    OldValued             : TDateTime; {19/07/07 : FQ 20603}
    OldValueG             : string;    {19/07/07 : FQ 20603}
    CanInitReference      : Boolean;   {19/07/07 : FQ 20603}
    FLoading              : Boolean;   {24/07/07}
    FEcranKeyDown         : TKeyEvent;
    FStCompteOuJournalEnCreation : string;
    FMessCompta           : TMessageCompta ;        // affichage des messages

    procedure ChargeDernReferencePointage;

    procedure MajApresUpdate;
    procedure ControleDuRIB;
    procedure ControleCompte;
    procedure ControleJournal;
    procedure ControleDatePointage;
    procedure ControleRefPointage;
    procedure InitControles;
    procedure InitEvenements;
    procedure InitAffichage(Arg : string);
    procedure ChangeGeneral;
    procedure MajTitre(DateOk : Boolean);
    procedure AfficheErreur;
    procedure GereAffichageMnt(GriseToutOk : Boolean = False);
    procedure InitRefPointage;
    {24/04/07 : FQ 20040 : Gestion du format des zones montants en fonction de la devise}
    procedure MajFormatMnt;
    {26/04/07 : FQ TRESO 10445 : Activation des zones soldes en pointage manuel}
    procedure DeviseChange;
    {14/11/07 : Le pointage manuel concerne les natures MAN, NBQ et vide}
    function  EstPointageManuel : Boolean;
    procedure OnError ( sender : TObject ; Error : TRecError);
  protected
    FExisteRef  : Boolean;
    FNumReleve  : Integer;
    TobEexBqLig : TOB;

    procedure GereSoldes;
    procedure RecupereLeReleve;
    function  SaisieAutomatique(var tRel : TOB) : Boolean;
    function  TrSaisieAuto     (var tRel : TOB) : Boolean;
    procedure PrepareStructure (var FInfo : TInfoEcriture; var FZEtebac : TZEtebac;
                                var FZGuide : TZEtebacGuide; var FZEcr : TZTresoVersEcr);
    procedure PrepareEcrRel    (TRel : TOB; FInfo : TInfoEcriture; FZEtebac : TZEtebac; FZGuide : TZEtebacGuide);
    procedure SetReferences(FZEtebac : TZEtebac; tLigneRel : TOB);
  public
    {$IFDEF EAGLCLIENT}
    EE_GENERAL         : THEdit;
    EE_DATEPOINTAGE    : THEdit;
    EE_REFPOINTAGE     : THEdit;
    EE_NEWSOLDEDEBEURO : THEdit;
    EE_NEWSOLDECREEURO : THEdit;
    EE_NEWSOLDEDEB     : THEdit;
    EE_NEWSOLDECRE     : THEdit;
    {$ELSE}
    EE_GENERAL         : THDBEdit;
    EE_DATEPOINTAGE    : THDBEdit;
    EE_REFPOINTAGE     : THDBEdit;
    EE_NEWSOLDEDEBEURO : THDBEdit;
    EE_NEWSOLDECREEURO : THDBEdit;
    EE_NEWSOLDEDEB     : THDBEdit;
    EE_NEWSOLDECRE     : THDBEdit;
    {$ENDIF}

    BInsert : TToolBarButton97;

    procedure OnElipsisClickEE_General(Sender : TObject);
    procedure OnKeyDownMontant        (Sender : TObject; var Key: Word; Shift: TShiftState);
    procedure ReleveOnClick           (Sender : TObject);
    procedure ReleveOnExit            (Sender : TObject);
    procedure EcranOnKeyDown          (Sender : TObject; var Key : Word; Shift : TShiftState); {18/07/07 : FQ 21119}
    procedure ReleveOnEnter           (Sender : TObject); {19/07/07 : FQ 20603}
    procedure CompteOnEnter           (Sender : TObject); {19/07/07 : FQ 20603}
    procedure CompteOnExit            (Sender : TObject); {19/07/07 : FQ 20603}
  end;

function CPLanceFiche_PointageFic(vStLequel : string = ''; vStParam : string = '') : string;

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  {$IFDEF VER150} Variants, {$ENDIF}
  {$IFDEF TRESO}ULibPieceCompta, UProcEcriture, {$ENDIF TRESO}
  HMsgBox, ParamSoc, LookUp, uLibWindows, Windows, Ent1, Constantes, Commun,
  UCstEtebac, UtilPgi, CpteSav {RecalculTotPointeNew1}, UprocGen, UObjGen,
  uTobDebug, UlibPointage,
  CPPOINTAGERAP_TOF;


{---------------------------------------------------------------------------------------}
function CPLanceFiche_PointageFic(vStLequel : string = ''; vStParam : string = '') : string;
{---------------------------------------------------------------------------------------}
var
  Tmp : string;
  Gen : string;
  DtP : string;
  Ref : string;
  Num : string;
  Q   : TQuery;
begin
  {19/06/07 : FQ TRESO 10476 : si on est en modification, consultation, il faut récupérer la clef complète}
  if Pos('CREATION', vStParam) = 0 then begin
    Tmp := vStLequel;
    Gen := ReadTokenSt(Tmp);
    DtP := ReadTokenSt(Tmp);
    Ref := ReadTokenSt(Tmp);
    Num := ReadTokenSt(Tmp);

    if (Gen = '') or (Num = '') then begin
      PgiError(TraduireMemoire('La session est inaccessible'));
      Exit;
    end;

    if Ref = '' then begin
      Q := OpenSql('SELECT EE_DATEPOINTAGE, EE_REFPOINTAGE FROM EEXBQ WHERE EE_GENERAL = "' + Gen +
                   '" AND  EE_NUMERO = ' + Num, True);
      if not Q.EOF then begin
        Ref := Q.FindField('EE_REFPOINTAGE').AsString;
        DtP := Q.FindField('EE_DATEPOINTAGE').AsString;
        vStLequel := Gen + ';' + DtP + ';' + Ref + ';' + Num + ';';
        Ferme(Q);
      end
      else begin
        Ferme(Q);
        PgiError(TraduireMemoire('La session est inaccessible'));
        Exit;
      end;
    end;
  end;

  AGLLanceFiche('CP', 'CPPOINTAGEFIC', '', vStLequel, vStParam);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
var
  lSt : string;
begin
  inherited;
  if ctxTreso in V_PGI.PGIContexte then Ecran.HelpContext := 50000126
                                   else Ecran.HelpContext := 7604100 ;
  FLoading := True;
  {Récupération des contrôles}
  InitControles;
  {Branchement des évènements}
  InitEvenements;
  lSt := S;
  {Suppression de ACTION=}
  ReadTokenSt(lSt);
  {Présentation de la fiche}
  InitAffichage(lSt);
  {19/07/07 : FQ 20603 : Initialisation des variables pour l'initialisation de la référence de pointage}
  CanInitReference := True;
  OldValueG := '';
  OldValueD := iDate1900;
  FMessCompta := TMessageCompta.Create('Pointage');
end;


{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.OnNewRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  SetField('EE_GENERAL', FStCompteOuJournalEnCreation);
  SetField('EE_DATEPOINTAGE', iDate1900);
  SetField('EE_DATEOLDSOLDE', iDate1900);
  SetField('EE_NUMERO', 1);
  SetField('EE_NUMRELEVE', 1);
  SetField('EE_DEVISE', V_PGI.DevisePivot);
  {JP 28/06/07 : FQ 20879 : Par défaut on se met en pointage manuel, sinon on reste en pointage
                 automatique jusqu'à la validation et donc on ne peut saisir les soldes
   JP 14/11/07 : Gestion du pointage sur les comptes non bancaires (NQB)}
  if FBoPtgeNonBQ then SetField('EE_ORIGINERELEVE', CODEPOINTAGENBQ)
                  else SetField('EE_ORIGINERELEVE', CODEPOINTAGEMAN);
  SetField('EE_DATEINTEGRE', Date);
  SetField('EE_REFPOINTAGE', '');
  SetField('EE_AVANCEMENT', '-');
  {14/11/07 : Quant on est en pointage sur Tréso, il faut distinguer, sur la base (Pool Tréso),
              les relevés visible en compta de ceux visibles en Tréso}
  if EstPointageSurTreso then SetField('EE_STATUTRELEVE', SESSIONTRESO)
                         else SetField('EE_STATUTRELEVE', 'NON');
  SetFocusControl('EE_GENERAL');
  CanInitReference := True;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.OnLoadRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  FBoConsultation := CEstPointageEnConsultationSurDossier;
  if FBoConsultation then begin
    SetControlEnabled('BINSERT', False);
    SetControlEnabled('BDELETE', False);
    SetControlEnabled('BVALIDER', False);
  end;
  SetControlEnabled('BRELEVE', (DS.State = DsInsert));
  CanInitReference := (DS.State in [dsEdit, dsInsert]);
  FLoading := False;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.OnDeleteRecord;
{---------------------------------------------------------------------------------------}
var
  aCpte : string;
  aDatP : TDateTime;
begin
  aCpte := VarToStr(GetField('EE_GENERAL'));
  aDatP := VarToDateTime(GetField('EE_DATEPOINTAGE'));

  if (aCpte = '') or (aDatP <= iDate1900) then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Impossible de récupérer la date ou le compte');
  end

  else if CanDeleteRefPtge(aCpte, aDatP) then begin
    if CDepointeEcriture(aCpte, VarToStr(GetField('EE_REFPOINTAGE')), VarToStr(GetField('EE_NUMERO')), aDatP) then
      RecalculTotPointeNew1(aCpte);
  end

  else begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Il y a des sessions de pointages postérieures');
  end;
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.OnUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  LastError    := 0;
  LastErrorMsg := '';

  {08/10/07 : FQ TRESO 10528 : Pour forcer l'exit de EE_DATEPOINTAGE}
  if Ecran.ActiveControl = EE_DATEPOINTAGE then
    SetFocusControl('EE_REFPOINTAGE');

  if GetControlEnabled('EE_GENERAL') then begin
    if not VH^.PointageJal then ControleCompte
                           else ControleJournal;
  end;
  if LastError = -1 then Exit;

  if GetControlEnabled('EE_DATEPOINTAGE') then
    ControleDatePointage;
  if LastError = -1 then Exit;

  if GetControlEnabled('EE_REFPOINTAGE') then
    ControleRefPointage;
  if LastError = -1 then Exit;

  if GetControlEnabled('EE_GENERAL') and (not FBoOkDernStatutReleve) then begin
    if GetParamSocSecur('SO_CPPOINTAGESECU', False) then begin
      LastErrorMsg := TraduireMemoire('La session de pointage en cours de réalisation n''est pas cohérente');
      LastError    := -1;
    end
    else begin
      if PgiAsk(TraduireMemoire('La session de pointage en cours de réalisation n''est pas cohérente.') + #13#10 +
                TraduireMemoire('Voulez-vous créer la référence de pointage ?'), Ecran.Caption) = MrNo then
        LastError := -1;
    end;
  end;

  {Si solde debit et credit saisi > erreur}
  if (LastError = 0) and
     ((ValeurI(GetControlText('EE_NEWSOLDEDEB')) <> 0) and (ValeurI(GetControlText('EE_NEWSOLDECRE')) <> 0)) or
     ((ValeurI(GetControlText('EE_NEWSOLDEDEBEURO')) <> 0) and  (ValeurI(GetControlText('EE_NEWSOLDECREEURO')) <> 0)) then begin
    LastErrorMsg := TraduireMemoire('Impossible de créer une session de pointage avec un solde débiteur et créditeur');
    LastError    := - 1;
  end;

  inherited;
  if LastError = 0 then begin
    {Récupération des enregistrements de CETEBAC pour les déverser dans EEXBQLIG
     et saisie de Trésorerie automatique
     19/06/07 : Ajout du test sur dsInsert, car sinon en modification toutes les sessions passaient en manuel}
    if DS.State = DsInsert then RecupereLeReleve;

    FBoCreationEnreg := (DS.State = DsInsert);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.OnAfterUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if LastError = 0 then begin
    {Insertion dans EEXBQLIG, saisie de trésorerie et purge de CETEBAC}
    MajApresUpdate;

    Ecran.Close;
  end
  else
    FBoCreationEnreg := False;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.OnChangeField(F : TField);
{---------------------------------------------------------------------------------------}
begin
  inherited;
       if (F.FieldName = 'EE_GENERAL'     ) then ChangeGeneral
  else if (F.FieldName = 'EE_REFPOINTAGE' ) then MajTitre(False)
  else if (F.FieldName = 'EE_DATEPOINTAGE') then MajTitre(True)
  else if (F.FieldName = 'EE_DEVISE'      ) then DeviseChange; {FQ TRESO 10445}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.OnClose;
{---------------------------------------------------------------------------------------}
begin
  FMessCompta.Free ;
  if Assigned(TobEexBqLig) then FreeAndNil(TobEexBqLig);
  if FBoCreationEnreg then
    CPLanceFiche_PointageRappro('ACTION=MODIFICATION;' + GetField('EE_GENERAL') + ';' + DateToStr(GetField('EE_DATEPOINTAGE')) + ';' +
                          GetField('EE_REFPOINTAGE') + ';' +  IntToStr(GetField('EE_NUMERO')) + ';' +
                                  VarToStr(GetField('EE_ORIGINERELEVE')) + ';');

  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.OnElipsisClickEE_General(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  LookUpGenePtge(Sender);

  if not VH^.PointageJal then
    {On vérifie que le compte existe et contrôle du RIB}
    ControleCompte
  else
    {On vérifie que le Journal existe et contrôle du RIB}
    ControleJournal;

  {Réinitialisation de toutes les zones}
  {24/04/07 : FQ 20040}
  s := '.';
  s := PadR(s, '0', FNbDecimales + 1);
  if s = '.' then s := '0'
             else s := '0' + s;

  SetControlText('EE_REFPOINTAGE', '');
  SetControlText('EE_NEWSOLDEDEBEURO', s);
  SetControlText('EE_NEWSOLDECREEURO', s);
  SetControlText('EE_NEWSOLDEDEB', s);
  SetControlText('EE_NEWSOLDECRE', s);

  {Eventuelle gestion des erreurs}
  if LastError <> 0 then
    AfficheErreur;
end;

{Alt+F2 : fait appel à la calculatrice
{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.OnKeyDownMontant(Sender : TObject; var Key : Word; Shift : TShiftState);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if (Key = VK_F2) and (Shift = [ssAlt]) then
  begin
    if not (DS.State in [DSEdit..DsInsert]) then
      DS.Edit;
  end;
end;

{24/04/07 : FQ 20040 : Gestion du format des zones montants en fonction de la devise}
{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.MajFormatMnt;
{---------------------------------------------------------------------------------------}
begin
  SetControlProperty('EE_NEWSOLDEDEBEURO', 'DisplayFormat', StrfMask(FNbDecimales, '', True));
  SetControlProperty('EE_NEWSOLDECREEURO', 'DisplayFormat', StrfMask(FNbDecimales, '', True));
  SetControlProperty('EE_NEWSOLDEDEB'    , 'DisplayFormat', StrfMask(FNbDecimales, '', True));
  SetControlProperty('EE_NEWSOLDECRE'    , 'DisplayFormat', StrfMask(FNbDecimales, '', True));
  Ecran.Refresh;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.ChangeGeneral;
{---------------------------------------------------------------------------------------}
var
  SQL    : string;
  Valeur : string;
begin
  {Vérification de l'existence du Compte / Journal et contrôle du RIB}
  if (DS.State in [dsInsert, dsEdit]) then begin
    {On ne teste pas ici le général à vide si on est en insertion, notamment pour
     éviter un message d'erreur en ouverture de fiche}
    if (DS.State = dsInsert) and
      (Trim(VarToStr(GetField('EE_GENERAL'))) = '') then Exit;

    if VH^.PointageJal then ControleJournal
                       else ControleCompte;
  end;

  if LastError <> 0 then
    AfficheErreur
  else begin
    {Mise à jour du titre de l'écran}
    MajTitre(True);

    Valeur := Trim(VarToStr(GetField('EE_GENERAL')));
    if Valeur = '' then Exit;

    {Le choix du relevé n'est possible que si on est en insertion ?}
    if (GetControlText('EE_GENERAL') <> '') and (DS.State = DsInsert) then begin
      SQL := 'SELECT CET_DATEOPERATION FROM CETEBAC, BANQUECP WHERE CET_ETABBQ = BQ_ETABBQ AND ';
      SQL := SQL + 'CET_GUICHET = BQ_GUICHET AND CET_NUMEROCOMPTE = BQ_NUMEROCOMPTE AND CET_TYPELIGNE = "07" ';
      {24/04/07 : FQ TRESO 10441 : Gestion du Is NULL pour Oracle}
      if EstPointageSurTreso then
        SQL := SQL + 'AND (CET_IMO <> "X" OR CET_IMO IS NULL) AND BQ_CODE = "' + FGeneral + '"'
      else
        SQL := SQL + 'AND (CET_IMO <> "X" OR CET_IMO IS NULL) AND BQ_GENERAL = "' + FGeneral + '"';
      SQL := SQL + ' ORDER BY BQ_NODOSSIER, BQ_GENERAL, CET_DATEOPERATION DESC';
      FExisteRef := ExisteSQL(SQL);
    end
    else
      FExisteRef := False;
    {FExisteRef : Sert a dire si on a choisi une référence de pointage par click sur BRELEVE}
    SetControlEnabled('BRELEVE', FExisteRef);
  end;

  {Si on est en mode rapprochement, on ne peut saisir les soldes qui sont issus du relevé
   19/06/07 : FQ 10477 : il suffit que l'une de ces conditions soit remplie pour que l'on désactive les solde
  GereAffichageMnt((FStDevise = '') or FBoConsultation or (GetField('EE_ORIGINERELEVE') = ORIGINERELEVE) or (GetField('EE_ORIGINERELEVE') = CODENEWPOINTAGE));}
  GereAffichageMnt((FStDevise = '') or FBoConsultation or not EstPointageManuel);

  {On repasse FExisteRef à False et FNumReleve = -1. Il seront réinitialiser lors du clik sur les relevés}
  FExisteRef := False;
  FNumReleve := -1;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.DeviseChange;
{---------------------------------------------------------------------------------------}
begin
  FStDevise := GetField('EE_DEVISE');

  {24/04/07 : FQ 20040 : gestion du nombre de décimales}
  FNbDecimales := CalcDecimaleDevise(FStDevise);
  MajFormatMnt;

  {26/04/07 : FQ TRESO 10445 : on laisse actif en pointage manuel les zones des soldes}
  GereAffichageMnt((FStDevise = '') or not EstPointageManuel);
                   {(GetField('EE_ORIGINERELEVE') = ORIGINERELEVE) or
                   (GetField('EE_ORIGINERELEVE') = CODENEWPOINTAGE));}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.AfficheErreur;
{---------------------------------------------------------------------------------------}
begin
  PGIError(LastErrorMsg, Ecran.Caption);
  FExisteRef   := False;
  LastErrorMsg := '';
  LastError    := 0;
  FStDevise    := V_PGI.DevisePivot;
  GereAffichageMnt(True);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.GereAffichageMnt(GriseToutOk : Boolean = False);
{---------------------------------------------------------------------------------------}
begin
  {Même en pointage manuel, on n'active pas les soldes si la session est équilibré}
  GriseToutOk := GriseToutOk or (VarToStr(GetField('EE_AVANCEMENT')) = 'X');

  if GriseToutOk then begin
    SetControlEnabled('EE_NEWSOLDEDEBEURO', False);
    SetControlEnabled('EE_NEWSOLDECREEURO', False);
    SetControlEnabled('EE_NEWSOLDEDEB', False);
    SetControlEnabled('EE_NEWSOLDECRE', False);
  end
  else begin
    if FStDevise = V_PGI.DevisePivot then begin
      SetControlEnabled('EE_NEWSOLDEDEBEURO', not FExisteRef);
      SetControlEnabled('EE_NEWSOLDECREEURO', not FExisteRef);
      SetControlEnabled('EE_NEWSOLDEDEB', False);
      SetControlEnabled('EE_NEWSOLDECRE', False);
    end
    else begin
      SetControlEnabled('EE_NEWSOLDEDEBEURO', False);
      SetControlEnabled('EE_NEWSOLDECREEURO', False);
      SetControlEnabled('EE_NEWSOLDEDEB', not FExisteRef);
      SetControlEnabled('EE_NEWSOLDECRE', not FExisteRef);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.MajTitre(DateOk : Boolean);
{---------------------------------------------------------------------------------------}
var
  Valeur : string;
  Titre  : string;
begin
  Valeur := Trim(VarToStr(GetField('EE_GENERAL')));
  if Valeur = '' then begin
    Ecran.Caption := cTitreCreation + ' : ';
    UpdateCaption(Ecran);
    Exit;
  end;

  if not VH^.PointageJal then Titre := TraduireMemoire('Compte') + ' ' + Valeur
                         else Titre := TraduireMemoire('Journal') + ' ' + Valeur;
  if VarToDateTime(GetField('EE_DATEPOINTAGE')) > iDate1900 then
    Titre := Titre + ' - Pointage au ' + VarToStr(GetField('EE_DATEPOINTAGE'));
  if Trim(VarToStr(GetField('EE_DATEPOINTAGE'))) <> '' then
    Titre := Titre + ' (' + VarToStr(GetField('EE_REFPOINTAGE')) + ')';

  Ecran.Caption := Titre;
  UpdateCaption(Ecran);
  if DateOk and (DS.State in [dsInsert, dsEdit]) then 
     SetField('EE_DATESOLDE', GetField('EE_DATEPOINTAGE'));
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.InitControles;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF EAGLCLIENT}
  EE_GENERAL         := THEdit(GetControl('EE_GENERAL'        , True));
  EE_DATEPOINTAGE    := THEdit(GetControl('EE_DATEPOINTAGE'   , True));
  EE_REFPOINTAGE     := THEdit(GetControl('EE_REFPOINTAGE'    , True));
  EE_NEWSOLDEDEBEURO := THEdit(GetControl('EE_NEWSOLDEDEBEURO', True));
  EE_NEWSOLDECREEURO := THEdit(GetControl('EE_NEWSOLDECREEURO', True));
  EE_NEWSOLDEDEB     := THEdit(GetControl('EE_NEWSOLDEDEB'    , True));
  EE_NEWSOLDECRE     := THEdit(GetControl('EE_NEWSOLDEDEB'    , True));
  {$ELSE}
  EE_GENERAL         := THDBEdit(GetControl('EE_GENERAL'        , True));
  EE_DATEPOINTAGE    := THDBEdit(GetControl('EE_DATEPOINTAGE'   , True));
  EE_REFPOINTAGE     := THDBEdit(GetControl('EE_REFPOINTAGE'    , True));
  EE_NEWSOLDEDEBEURO := THDBEdit(GetControl('EE_NEWSOLDEDEBEURO', True));
  EE_NEWSOLDECREEURO := THDBEdit(GetControl('EE_NEWSOLDECREEURO', True));
  EE_NEWSOLDEDEB     := THDBEdit(GetControl('EE_NEWSOLDEDEB'    , True));
  EE_NEWSOLDECRE     := THDBEdit(GetControl('EE_NEWSOLDEDEB'    , True));
  {$ENDIF}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.InitEvenements;
{---------------------------------------------------------------------------------------}
begin
  (GetControl('BRELEVE') as TToolBarButton97).OnClick := ReleveOnClick;
  EE_DATEPOINTAGE.OnExit := ReleveOnExit;
  EE_GENERAL.OnElipsisClick := OnElipsisClickEE_GENERAL;
  {19/07/07 : FQ 20603 : Pour tester une éventuelle modification de valeur du compte et de la référence}
  EE_DATEPOINTAGE.OnEnter := ReleveOnEnter;
  EE_GENERAL     .OnEnter := CompteOnEnter;
  EE_GENERAL     .OnExit  := CompteOnExit;

  EE_NEWSOLDEDEBEURO.OnKeyDown := OnKeyDownMontant;
  EE_NEWSOLDECREEURO.OnKeyDown := OnKeyDownMontant;
  EE_NEWSOLDEDEB.OnKeyDown     := OnKeyDownMontant;
  EE_NEWSOLDECRE.OnKeyDown     := OnKeyDownMontant;
  
  FEcranKeyDown := TFFiche(Ecran).FormKeyDown;
  Ecran.OnKeyDown := EcranOnKeyDown; {18/07/07 : FQ 21119}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.InitAffichage(Arg : string);
{---------------------------------------------------------------------------------------}
begin
  cTitreCreation := TraduireMemoire('Création d''une référence de pointage');
  FNoDossier := V_PGI.NoDossier;

  {Compte ou Journal à renseigner en création}
  FStCompteOuJournalEnCreation := ReadTokenSt(Arg);
  FBoCreationEnreg := False;
  {JP 14/11/07 : Gestion du pointage sur les comptes non bancaires (NQB)}
  FBoPtgeNonBQ := ReadTokenSt(Arg) = CODEPOINTAGENBQ;
  FBoRefNonBq      := False;

  EE_GENERAL.MaxLength := IIF(not VH^.PointageJal, VH^.CPta[fbGene].Lg, 3);

  if not VH^.PointageJal then begin
    SetControlProperty('TEE_GENERAL', 'CAPTION', TraduireMemoire('Compte général'));
    SetControlProperty('EE_GENERAL', 'DATATYPE', 'TZGBANQUE');
    SetControlProperty('EE_GENERAL', 'MAXLENGTH', VH^.CPta[fbGene].Lg);
  end
  else begin
    SetControlProperty('TEE_GENERAL', 'CAPTION', TraduireMemoire('Journal'));
    SetControlProperty('EE_GENERAL', 'DATATYPE', 'TTJALBANQUE');
    SetControlProperty('EE_GENERAL', 'MAXLENGTH', 3);
  end;

  GereAffichageMnt(True);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.ChargeDernReferencePointage;
{---------------------------------------------------------------------------------------}
var
  lQuery : TQuery;
begin
  FDtDernDatePointage := iDate1900;
  FBoOkDernStatutReleve := True;
   //EE_DATEOLDSOLDE
  lQuery := OpenSql('SELECT EE_DATEPOINTAGE, EE_AVANCEMENT, EE_DATESOLDE FROM EEXBQ WHERE ' +
                    'EE_GENERAL = "' + VarToStr(GetField('EE_GENERAL')) + '" ORDER BY EE_DATEPOINTAGE DESC', True);
  try
    try
      if not lQuery.Eof then begin
        if IsValidDate(lQuery.Findfield('EE_DATEPOINTAGE').AsString) then begin
          FDtDernDatePointage := lQuery.FindField('EE_DATEPOINTAGE').AsDateTime;

          if VarToDateTime(GetField('EE_DATESOLDE')) = iDate1900 then begin
            {Mise à jour de la date su solde précédent}
            if IsValidDate(lQuery.Findfield('EE_DATESOLDE').AsString) then
              SetField('EE_DATESOLDE', lQuery.Findfield('EE_DATESOLDE').AsDateTime)
            else
              SetField('EE_DATESOLDE', FDtDernDatePointage);
          end;
        end;

        FBoOkDernStatutReleve := (lQuery.FindField('EE_AVANCEMENT').AsString <> '-');
      end;
    except
      on E : Exception do PgiError(TraduireMemoire('Erreur SQL : ') + E.Message, TraduireMemoire('Récupération de la dernière référence de pointage'));
    end;
  finally
    Ferme(lQuery);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.ControleCompte;
{---------------------------------------------------------------------------------------}
var
  lSql : TQuery;
  G    : string;
begin
  FBoRefNonBq := False;
  
  {Pointage sur TRECRITURE}
  if EstPointageSurTreso then begin
    FGeneral := VarToStr(GetField('EE_GENERAL'));
    {04/05/07 : FQ 20240 : En CWas, à la sortie du Lookup, le GetField n'est pas renseigné !}
    if FGeneral = '' then FGeneral := GetControlText('EE_GENERAL');

    lSql := OpenSql('SELECT BQ_CODE FROM BANQUECP WHERE BQ_CODE LIKE "' + FGeneral + '%"', True);
    try
      if not lSql.EOF then begin
        if FGeneral <> lSql.FindField('BQ_CODE').AsString then
          SetField('EE_GENERAL', lSql.FindField('BQ_CODE').AsString);
        ControleDuRIB;
      end
      else begin
        LastErrorMsg := TraduireMemoire('Le compte n''existe pas');
        LastError    := - 1;
      end;
    finally
      Ferme(lSql);
    end;
  end

  {Pointage sur ECRITURE}
  else begin
    {04/05/07 : FQ 20240 : En CWas, à la sortie du Lookup, le GetField n'est pas renseigné !}
    G := GetField('EE_GENERAL');
    if G = '' then G := GetControlText('EE_GENERAL');

    lSql := OpenSql('SELECT G_GENERAL, G_NATUREGENE FROM GENERAUX WHERE G_GENERAL = "' + BourreEtLess(G, fbGene) + '" AND G_POINTABLE = "X"', True);
    try
      try
        if not lSql.Eof then begin
          if Length(GetField('EE_GENERAL')) < VH^.Cpta[fbGene].Lg then
            SetField('EE_GENERAL', BourreEtLess(G, fbGene));
          FGeneral := G;

          {Contrôle du RIB pour les comptes de banque}
          if lSql.FindField('G_NATUREGENE').AsString = 'BQE' then begin
            {14/11/07 : Gestion du pointage sur les comptes non bancaires lorsque l'on est en pointage sur TRECRITURE}
            if EstPointageCache then begin
              SetField('EE_GENERAL', '');
              LastErrorMsg := TraduireMemoire('Vous n''avez pas le droit de pointer sur des comptes bancaires en Comptabilité.');
              LastError    := - 1;
              FGeneral     := '';
              EE_GENERAL.SetFocus;
            end
            else
              ControleDuRIB
          end
          else begin
            if DS.State in [dsEdit, dsInsert] then begin
              {14/11/07 : Le pointage sur journal ne peut se faire que sur des comptes bancaires}
              if VH^.PointageJal then begin
                SetField('EE_GENERAL', '');
                LastErrorMsg := TraduireMemoire('Le pointage sur journal ne peut se faire que sur des comptes bancaires.');
                LastError    := - 1;
                FGeneral     := '';
                EE_GENERAL.SetFocus;
              end
              else begin
                SetField('EE_DEVISE', V_PGI.DevisePivot);
                if DS.State = dsInsert then
                  {JP 14/11/07 : Gestion du pointage sur les comptes non bancaires (NQB)}
                  SetField('EE_ORIGINERELEVE', CODEPOINTAGENBQ);
                FBoRefNonBq := True;
              end;
            end;

            FStDevise := V_PGI.DevisePivot;
            {24/04/07 : FQ 20040 : gestion du nombre de décimales}
            FNbDecimales := V_PGI.OkDecV;
            MajFormatMnt;
          end;
        end
        else begin
          LastErrorMsg := TraduireMemoire('Le compte n''existe pas');
          LastError    := - 1;
          FGeneral     := '';
          EE_GENERAL.SetFocus;
        end;

      except
        on E : Exception do PgiError(TraduireMemoire('Erreur SQL : ') + E.Message, TraduireMemoire('Fonction OnExitEE_General'));
      end;
    finally
      Ferme(lSql);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.ControleJournal;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  J : string;
begin
  {04/05/07 : FQ 20240 : En CWas, à la sortie du Lookup, le GetField n'est pas renseigné !}
  J := GetField('EE_GENERAL');
  if J = '' then J := GetControlText('EE_GENERAL');

  Q := OpenSQL('SELECT J_CONTREPARTIE FROM JOURNAL LEFT JOIN GENERAUX ON J_CONTREPARTIE = G_GENERAL ' +
      'WHERE J_JOURNAL = "' + J + '" AND J_NATUREJAL = "BQE" AND G_POINTABLE = "X"', True);

  {Teste l'existence du journal}
  if Q.EOF then begin
    LastErrorMsg := TraduireMemoire('Le journal sélectionné n''est pas valide pour le pointage.');
    SetFocusControl('EE_GENERAL');
    FGeneral := '';
    LastError := -1;
  end

  {Contrôle du compte de contrepartie}
  else begin
    FGeneral := Q.FindField('J_CONTREPARTIE').AsString;
    ControleDuRIB;
  end;
end;

{Controle le RIB associé à un compte de BANQUE ou au compte de contrepartie d'un JOURNAL
{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.ControleDuRIB;
{---------------------------------------------------------------------------------------}
var
  lQuery : TQuery;
  lStSql : string;
begin
  LastError    := 0;
  LastErrorMsg := '';

  {Pointage sur TRECRITURE}
  if EstPointageSurTreso then begin
    lStSql := 'SELECT BQ_DEVISE, D_CODEISO, BQ_CODEIBAN, BQ_NODOSSIER FROM BANQUECP ' +
              'LEFT JOIN DEVISE ON D_DEVISE = BQ_DEVISE ' +
              'WHERE BQ_CODE = "' + FGeneral + '"';
    lQuery := OpenSql(lStSql, True);
    try
      if not lQuery.EOF then begin
        FNoDossier       := lQuery.FindField('BQ_NODOSSIER').AsString;
        FStDevise        := lQuery.FindField('BQ_DEVISE').AsString;
        FDeviseConvertie := lQuery.FindField('D_CODEISO').AsString;
        if FstDevise = '' then begin
          LastError    := -1;
          LastErrorMsg := TraduireMemoire('Impossible de récupérer la devise');
        end
        else if FDeviseConvertie = '' then begin
          {14/06/07 : gestion du cas PCL sur la devise PIVOT}
          LastErrorMsg := TesteCodeIsoDevise(FDeviseConvertie, FStDevise);
          if LastErrorMsg <> '' then LastError := -1;
        end
        else if DS.State in [dsEdit, dsInsert] then begin
          SetField('EE_RIB', lQuery.FindField('BQ_CODEIBAN').AsString);
          SetField('EE_DEVISE', FstDevise);
        end;
      end
      else begin
        LastError := -1;
        LastErrorMsg := TraduireMemoire('Le compte est introuvable');
      end;
    finally
      Ferme(lQuery);
    end;
  end

  {POintage sur ECRITURE ou sur Journal}
  else begin
    FNoDossier := V_PGI.NoDossier;
    {Recherche de l'enregistrement dans BANQUECP}
    if VH^.PointageJal then
      lStSql := 'SELECT BQ_CODEIBAN, D_CODEISO, BQ_DEVISE FROM BANQUECP ' +
                'LEFT JOIN DEVISE ON D_DEVISE = BQ_DEVISE ' +
                'WHERE BQ_GENERAL = "' + FGeneral + '" AND BQ_NODOSSIER = "' + FNoDossier + '" '
    else
      if EstPointageSurTreso then
        lStSql := 'SELECT BQ_CODEIBAN, D_CODEISO, BQ_DEVISE FROM BANQUECP ' +
                  'LEFT JOIN DEVISE ON D_DEVISE = BQ_DEVISE ' +
                  'WHERE BQ_CODE = "' + VarToStr(GetField('EE_GENERAL')) + '"'
      else
        lStSql := 'SELECT BQ_CODEIBAN, D_CODEISO, BQ_DEVISE FROM BANQUECP ' +
                  'LEFT JOIN DEVISE ON D_DEVISE = BQ_DEVISE ' +
                  'WHERE BQ_GENERAL = "' + VarToStr(GetField('EE_GENERAL')) +
                  '" AND BQ_NODOSSIER = "' + FNoDossier + '"';

    lQuery := nil;
    try
      try
        lQuery := OpenSql(lStSql, True);
        if not lQuery.Eof then begin
          FStDevise := lQuery.FindField('BQ_DEVISE').AsString;
          FDeviseConvertie := lQuery.FindField('D_CODEISO').AsString;
          if FstDevise = '' then begin
            LastError    := -1;
            LastErrorMsg := TraduireMemoire('Impossible de récupérer la devise');
          end
          else if FDeviseConvertie = '' then begin
            {14/06/07 : gestion du cas PCL sur la devise PIVOT}
            LastErrorMsg := TesteCodeIsoDevise(FDeviseConvertie, FStDevise);
            if LastErrorMsg <> '' then LastError := -1;
          end
          else if DS.State in [dsEdit, dsInsert] then begin
            SetField('EE_RIB', lQuery.FindField('BQ_CODEIBAN').AsString);
            SetField('EE_DEVISE', FStDevise);
          end;
        end
        else begin
          FStDevise := ''; // Devise indéterminée puisque le compte n'est pas bon
          if VH^.PointageJal then
            PgiInfo(TraduireMemoire('Le compte de contrepartie du journal sélectionné n''a pas de rib associé'), Ecran.Caption)
          else
          begin
            LastErrorMsg := TraduireMemoire('Le compte sélectionné n''a pas de rib associé');
            SetFocusControl('EE_GENERAL');
            LastError := -1;
          end;
        end;

      except
        on E : Exception do begin
          LastErrorMsg := TraduireMemoire('Erreur SQL : ') + #13#13 + E.Message;
          LastError := -1;
        end;
      end;

    finally
      Ferme(lQuery);
    end;
  end;
  {24/04/07 : FQ 20040 : gestion du nombre de décimales}
  FNbDecimales := CalcDecimaleDevise(FStDevise);
  MajFormatMnt;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.ControleDatePointage;
{---------------------------------------------------------------------------------------}
begin
  {Chargement de la dernière référence de pointage}
  ChargeDernReferencePointage;

  if GetField('EE_DATEPOINTAGE') <= FDtDernDatePointage then begin
    SetFocusControl('EE_DATEPOINTAGE');
    LastErrorMsg := TraduireMemoire('La date de pointage doit être supérieure au ' + DateToStr(FDtDernDatePointage));
    LastError := -1;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.ControleRefPointage;
{---------------------------------------------------------------------------------------}
begin
  if Trim(GetField('EE_REFPOINTAGE')) = '' then begin
    LastErrorMsg := TraduireMemoire('Vous devez renseigner la référence de pointage.');
    SetFocusControl('EE_REFPOINTAGE');
    LastError := -1;
  end
  else begin
    if ExisteSql('SELECT EE_GENERAL FROM EEXBQ WHERE EE_GENERAL = "' + VarToStr(GetField('EE_GENERAL')) +
                 '" AND EE_REFPOINTAGE = "' + GetField('EE_REFPOINTAGE') + '"') then begin
      LastErrorMsg := TraduireMemoire('Le libellé de la référence de pointage existe déjà.');
      SetFocusControl('EE_REFPOINTAGE');
      LastError := -1;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.ReleveOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  SQL  : string;
  SQL1 : string;
  Gen  : string;
begin
  FNumReleve := -1;

  {08/10/07 : FQ TRESO 10528 : Pour forcer le change du général}
  if Ecran.ActiveControl = EE_GENERAL then
    SetFocusControl('EE_DATEPOINTAGE');

  if (DS.State = DsInsert) then begin

    {08/10/07 : FQ TRESO 10528 : en CWAS, GetControlText et nettement plus à jour que GetField}
    Gen := GetControlText('EE_GENERAL');
    {Théoriquement cela ne devrait pas arriver}
    if Gen = '' then begin
      LastError    := -1;
      if VH^.PointageJal then LastErrorMsg := TraduireMemoire('Veuillez choisir un journal')
                         else LastErrorMsg := TraduireMemoire('Veuillez choisir un compte');
    end;

    SQL := 'SELECT CET_DATEOPERATION, BQ_LIBELLE FROM CETEBAC, BANQUECP ';
    if VH^.PointageJal then SQL := SQL + ', JOURNAL ';
    SQL := SQL + 'WHERE CET_ETABBQ = BQ_ETABBQ AND CET_GUICHET = BQ_GUICHET AND CET_NUMEROCOMPTE = BQ_NUMEROCOMPTE ';
    {24/04/07 : FQ TRESO 10441 : Gestion du Is NULL pour Oracle}
    SQL := SQL + 'AND CET_TYPELIGNE = "07" AND (CET_IMO <> "X" OR CET_IMO IS NULL) AND ';

    if VH^.PointageJal then
      SQL := SQL + 'J_CONTREPARTIE = BQ_GENERAL AND J_NATUREJAL = "BQE" AND J_JOURNAL = "' + Gen + '"'
    else begin
      if EstPointageSurTreso then
        SQL := SQL + 'BQ_CODE = "' + Gen + '"'
      else
        SQL := SQL + 'BQ_GENERAL = "' + Gen + '"';
    end;
    
    {JP 25/01/08 : FQ 10547 : Mise en place de la confidentialité sur le choix des relevés pour gérer le cas
                   où une base du regroupement Multi-sociétés est fusionnée avec une autre base => on se
                   retrouve avec 2 enregistrements de BanqueCP ayant le même RIB : Cf. SIC avec la fusion de
                   ASINFOR dans la base Cegid}
    if CtxTreso in V_PGI.PgiContexte then
      SQL := SQL + FiltreBanqueCp(tcp_Bancaire, tcb_Bancaire, GetTrFiltreDossiers);

    SQL1 := SQL + ' ORDER BY BQ_NODOSSIER, BQ_GENERAL, CET_DATEOPERATION DESC';
    FExisteRef := LookupList(EE_DATEPOINTAGE, TraduireMemoire('Relevés en attente de pointage'), 'CETEBAC', 'CET_DATEOPERATION', 'BQ_LIBELLE', '', '', True, 0, SQL);

    {On avertit s'il y a des relevés plus anciens en attente}
    if IsValidDate(EE_DATEPOINTAGE.Text) then
      if ExisteSQL(SQL + ' AND CET_DATEOPERATION < "' + UsDateTime(StrToDate(GetControlText('EE_DATEPOINTAGE'))) + '"') then
        PGIInfo(TraduireMemoire('Il existe des relevés plus anciens ; si vous validez votre choix, vous ne pourrez plus les intégrer.'), Ecran.Caption);

    if FExisteRef then
      SetField('EE_ORIGINERELEVE', CODENEWPOINTAGE);
      
    {Création d'une référence de pointage par défaut}
    InitRefPointage;
  end
  else
    SetControlEnabled('BRELEVE', False);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.ReleveOnExit(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if FLoading then Exit;
  if FExisteRef then GereSoldes
  else if Ds.State in [dsEdit, dsInsert] then begin
    {19/07/07 : FQ 20603 : on ne lance plus systématiquement l'initialisation, on s'assure
               que le compte et la date ont changé.
     08/10/07 : FQ TRESO 10528 : en CWAS, GetControlText et nettement plus à jour que GetField}
    CanInitReference := CanInitReference and
                       ((OldValueG <> GetControlText('EE_GENERAL')) or (OldValueD <> VarToDateTime(GetControlText('EE_DATEPOINTAGE'))));
    InitRefPointage;
  end;
end;


{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.CompteOnEnter(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if FLoading then Exit;
  OldValueG := GetControlText('EE_GENERAL');
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.CompteOnExit(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if FLoading then Exit;
  {19/0/07 : FQ 20603 : on ne lance plus systématiquement l'initialisation, on s'assure
             que le compte et la date ont changé
   08/10/07 : FQ TRESO 10528 : en CWAS, GetControlText et nettement plus à jour que GetField}
  CanInitReference := CanInitReference and
                     ((OldValueG <> GetControlText('EE_GENERAL')) or (OldValueD <> VarToDateTime(GetControlText('EE_DATEPOINTAGE'))));
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.ReleveOnEnter(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if FLoading then Exit;
  OldValueD := VarToDateTime(GetControlText('EE_DATEPOINTAGE'));
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.GereSoldes;
{---------------------------------------------------------------------------------------}
var
  SQL : string;
  Q   : TQuery;
begin
  SQL := GetControlText('EE_DATEPOINTAGE');
  if IsValidDate(SQL) then begin
    {Récupération du numéro de relevé et du solde à la fin de la session de pointage}
    SQL := 'SELECT CET_CREDIT, CET_CREDITDEV, CET_DEBIT, CET_DEBITDEV, CET_DEVISE, CET_NUMRELEVE FROM CETEBAC, BANQUECP ';
    SQL := SQL + 'WHERE CET_ETABBQ = BQ_ETABBQ AND CET_GUICHET = BQ_GUICHET AND CET_NUMEROCOMPTE = BQ_NUMEROCOMPTE ';
    {24/04/07 : FQ TRESO 10441 : Gestion du Is NULL pour Oracle}
    if EstPointageSurTreso then
      SQL := SQL + 'AND CET_TYPELIGNE = "07" AND (CET_IMO <> "X" OR CET_IMO IS NULL) AND BQ_CODE = "' + FGeneral + '" AND CET_DATEOPERATION = "' +
                   UsDateTime(VarToDateTime(GetField('EE_DATEPOINTAGE'))) + '"'
    else
      SQL := SQL + 'AND CET_TYPELIGNE = "07" AND (CET_IMO <> "X" OR CET_IMO IS NULL) AND BQ_GENERAL = "' + FGeneral + '" AND CET_DATEOPERATION = "' +
                   UsDateTime(VarToDateTime(GetField('EE_DATEPOINTAGE'))) + '"';
    SQL := SQL + ' ORDER BY BQ_NODOSSIER, BQ_GENERAL, CET_DATEOPERATION DESC';

    Q := OpenSQL(SQL, True);
    try
      if not Q.EOF then begin
        {On mémorise le numéro de relevé pour remplir EEXBQ et EEXBQLIG avec le contenu de CETEBAC à la validation}
        FNumReleve := Q.FindField('CET_NUMRELEVE').AsInteger;
        {Si un des soldes est renseigné
         JP 08/11/07 : FQ 21788 : Pas de contrôle sur les soldes à zéro lors du choix d'un relevé
        if Q.FindField('CET_DEBIT'    ).AsFloat +
           Q.FindField('CET_CREDIT'   ).AsFloat +
           Q.FindField('CET_DEBITDEV' ).AsFloat +
           Q.FindField('CET_CREDITDEV').AsFloat > 0 then begin}

          {On ne permet pas la saisie du solde ? à Voir}
          GereAffichageMnt(True);

          {Si la devise du compte est celle du relevé, il n'y a pas d'opération de change à faire}
          if FDeviseConvertie = Q.FindField('CET_DEVISE').AsString then begin
            {19/09/07 : FQ 21361 : Réinitialisation des montants, si on vient de changer de relevé}
            SetField('EE_NEWSOLDEDEBEURO', 0);
            SetField('EE_NEWSOLDECREEURO', 0);
            SetField('EE_NEWSOLDEDEB'    , 0);
            SetField('EE_NEWSOLDECRE'    , 0);

                 if Q.FindField('CET_DEBIT'    ).AsFloat > 0 then SetField('EE_NEWSOLDEDEBEURO', Q.FindField('CET_DEBIT'    ).AsFloat)
            else if Q.FindField('CET_CREDIT'   ).AsFloat > 0 then SetField('EE_NEWSOLDECREEURO', Q.FindField('CET_CREDIT'   ).AsFloat)
            else if Q.FindField('CET_DEBITDEV' ).AsFloat > 0 then SetField('EE_NEWSOLDEDEB'    , Q.FindField('CET_DEBITDEV' ).AsFloat)
            else if Q.FindField('CET_CREDITDEV').AsFloat > 0 then SetField('EE_NEWSOLDECRE'    , Q.FindField('CET_CREDITDEV').AsFloat);
          end
          else begin
            {Est-ce que cela peut arriver d'avoir un relevé dans une devise différente du compte ?
             Alors gestion de la Chancellerie ?}
            //V_PGI.DevisePivot
          end;
        {end
        else begin
          {On ne récupère pas le relevé ?? à voir
           JP 08/11/07 : FQ 21788 : si on le récupère
          FNumReleve := -1;
          PGIError(TraduireMemoire('Impossible de récupérer le solde du relevé'), Ecran.Caption);
        end;}
      end
      else begin
        FNumReleve := -1;
        PGIError(TraduireMemoire('Impossible de récupérer le numéro du relevé'), Ecran.Caption);
      end;
    finally
      Ferme(Q);
    end;
  end
  else
    PGIError(TraduireMemoire('La date de pointage n''est pas correcte'), Ecran.Caption);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.RecupereLeReleve;
{---------------------------------------------------------------------------------------}
var
  Rib : TRecRib;
  D07 : TDateTime;
  D01 : TDateTime;
  Mnt : Double;
  s : string;
  Q : TQuery;
  F : TOB;
  I : Integer;
  N : Integer;
begin
  {Récupération du numéro de relevé}
  N := 0;
  s := 'SELECT MAX(EE_NUMERO) FROM ' + GetTableDossier('', 'EEXBQ') + ' WHERE EE_GENERAL = "' + VarToStr(GetField('EE_GENERAL')) + '"';
  Q := OpenSQL(S, True) ;
  if not Q.EOF then N := Q.Fields[0].AsInteger;
  Ferme(Q) ;
  Inc(N);

  if (FNumReleve = -1) or not FExisteRef then begin
    {On est en pointage manuel
     JP 14/11/07 : Gestion du pointage sur les comptes non bancaires (NQB)}
    if FBoPtgeNonBQ  or FBoRefNonBq then SetField('EE_ORIGINERELEVE', CODEPOINTAGENBQ)
                                    else SetField('EE_ORIGINERELEVE', CODEPOINTAGEMAN);

    SetField('EE_NUMERO'       , N);
    SetField('EE_NUMRELEVE'    , N);
    Exit;
  end;

  I := 0;
  {Récupération du relevé pour EEX_BQ et EEXBQLIG}
  Q := OpenSQL('SELECT * FROM CETEBAC WHERE CET_NUMRELEVE = ' + IntToStr(FNumReleve) + ' ORDER BY CET_TYPELIGNE', True);
  TobEexBqLig := TOB.Create('*RELEVE', nil, -1);
  try
    {Éventuels contrôles qui n'ont pas été exécutés lors de l'intégration des relevés}
    if not Q.EOF then begin
      {Les contrôles n'ont pas été effectués lors de l'intégration dans CETEBAC}
      if Q.FindField('CET_CONTROLOK').AsString <> 'X' then begin
        if Q.FindField('CET_TYPELIGNE').AsString = '01' then begin
          D01 := Q.FindField('CET_DATEOPERATION').AsDateTime;
          Rib.Etab    := Q.FindField('CET_ETABBQ').AsString;
          Rib.Guichet := Q.FindField('CET_GUICHET').AsString;
          Rib.Compte  := Q.FindField('CET_NUMEROCOMPTE').AsString;

          Mnt := Q.FindField('CET_CREDIT').AsFloat + Q.FindField('CET_CREDITDEV').AsFloat -
                 Q.FindField('CET_DEBIT' ).AsFloat - Q.FindField('CET_DEBITDEV' ).AsFloat;

          if Q.Locate('CET_TYPELIGNE', '07', []) then begin
            D07 := Q.FindField('CET_DATEOPERATION').AsDateTime;
            S := HasCoherenceCEtebacEexbq(D01, D07, Rib, Mnt);
            {Il y a une incohérence}
            if (Trim(S) <> '') and (Trim(S) <> SANSSESSIONPTGE) then begin
              LastErrorMsg := S;
              LastError := 1;
              Exit;
            end;
          end
          {Pas d'enregistrement 07}
          else begin
            {Cela ne devrait jamais arrivé, car ce type de contrôles a déjà été effectués. Mais on n'est jamais sûr de rien}
            LastErrorMsg := TraduireMemoire('Le relevé n''est pas correct (Pas d''enregistrement 07).'#13'Il ne peut être intégré dans la table des mouvements bancaires.');
            LastError := 1;
            Exit;
          end;
        end
        {Pas d'enregistrement 01}
        else begin
          {Cela ne devrait jamais arrivé, car ce type de contrôles a déjà été effectués. Mais on n'est jamais sûr de rien}
          LastErrorMsg := TraduireMemoire('Le relevé n''est pas correct (Pas d''enregistrement 01).'#13'Il ne peut être intégré dans la table des mouvements bancaires.');
          LastError := 1;
          Exit;
        end;
      end;
    end;

    Q.First;
    while not Q.EOF do begin
      if Q.FindField('CET_TYPELIGNE').AsString = '01' then begin
        SetField('EE_DATEOLDSOLDE', Q.FindField('CET_DATEOPERATION').AsDateTime);
        SetField('EE_NUMRELEVE'   , N);
        SetField('EE_NUMERO'      , N);
        SetField('EE_DEVISE'      , FStDevise);
        {Si la devise du compte est celle du relevé, il n'y a pas d'opération de change à faire}
        if FDeviseConvertie = Q.FindField('CET_DEVISE').AsString then begin
          if Q.FindField('CET_CREDITDEV').AsFloat > 0 then SetField('EE_OLDSOLDECRE', Q.FindField('CET_CREDITDEV').AsFloat)
                                                      else SetField('EE_OLDSOLDECRE', 0);
          if Q.FindField('CET_CREDIT').AsFloat > 0 then SetField('EE_OLDSOLDECREEURO', Q.FindField('CET_CREDIT').AsFloat)
                                                   else SetField('EE_OLDSOLDECREEURO', 0);
          if Q.FindField('CET_DEBITDEV').AsFloat > 0 then SetField('EE_OLDSOLDEDEB', Q.FindField('CET_DEBITDEV').AsFloat)
                                                     else SetField('EE_OLDSOLDEDEB', 0);
          if Q.FindField('CET_DEBIT').AsFloat > 0 then SetField('EE_OLDSOLDEDEBEURO', Q.FindField('CET_DEBIT').AsFloat)
                                                  else SetField('EE_OLDSOLDEDEBEURO', 0);

        end
        else begin
          {Est-ce possible ?}
          //SetField('EE_DEVISE', FStDevise);
        end;

        SetField('EE_RIB', Q.FindField('CET_ETABBQ').AsString + Q.FindField('CET_GUICHET').AsString + Q.FindField('CET_NUMEROCOMPTE').AsString);
        SetField('EE_ORIGINERELEVE', CODENEWPOINTAGE);
        SetField('EE_STATUTRELEVE', 'NON');
      end

      {Remplissage des lignes de EEXBQLIG}
      else if Q.FindField('CET_TYPELIGNE').AsString = '04' then begin
        Inc(I);
        F := TOB.Create('EEXBQLIG', TobEexBqLig, -1);
        F.PutValue('CEL_GENERAL'      , VarToStr(GetField('EE_GENERAL')));
        F.PutValue('CEL_RIB'          , GetField('EE_RIB'));
        F.PutValue('CEL_REFPOINTAGE'  , '');
        F.PutValue('CEL_DATEPOINTAGE' , iDate1900);
        F.PutValue('CEL_NUMRELEVE'    , N);
        F.PutValue('CEL_NUMLIGNE'     , I);
        F.PutValue('CEL_DISPONIBLE'   , -1);
        F.PutValue('CEL_LIBELLE'      , Q.FindField('CET_LIBELLE').AsString);
        F.PutValue('CEL_LIBELLE1'     , Trim(Q.FindField('CET_LIBELLE1').AsString));
        F.PutValue('CEL_LIBELLE2'     , Trim(Q.FindField('CET_LIBELLE2').AsString));
        F.PutValue('CEL_LIBELLE3'     , Trim(Q.FindField('CET_LIBELLE3').AsString));
        F.PutValue('CEL_CODEAFB'      , Q.FindField('CET_CODEAFB').AsString);
        F.PutValue('CEL_DATEOPERATION', Q.FindField('CET_DATEOPERATION').AsDateTime);
        F.PutValue('CEL_DATEVALEUR'   , Q.FindField('CET_DATEVALEUR').AsDateTime);
        F.PutValue('CEL_REFORIGINE'   , Q.FindField('CET_REFPIECE').AsString);
        F.PutValue('CEL_REFPIECE'     , Q.FindField('CET_NUMEROPIECE').AsString);
       {24/04/07 : FQ TRESO 10441 : Gestion du Is NULL pour Oracle}
        if Q.FindField('CET_IMO').AsString = '' then
          F.PutValue('CEL_IMO'        , '-')
        else
          F.PutValue('CEL_IMO'        , Q.FindField('CET_IMO').AsString);

        F.PutValue('CEL_EXONERE'      , Q.FindField('CET_EXONERE').AsString);
        F.PutValue('CEL_DEVISE'       , FstDevise);
        F.PutValue('CEL_VALIDE'       , '-');

        {Si la devise du compte est celle du relevé, il n'y a pas d'opération de change à faire}
        if FDeviseConvertie = Q.FindField('CET_DEVISE').AsString then begin
          if Q.FindField('CET_CREDITDEV').AsFloat > 0 then F.SetDouble('CEL_CREDITDEV', Q.FindField('CET_CREDITDEV').AsFloat)
                                                      else F.SetDouble('CEL_CREDITDEV', 0);
          if Q.FindField('CET_CREDIT').AsFloat > 0 then F.SetDouble('CEL_CREDITEURO', Q.FindField('CET_CREDIT').AsFloat)
                                                   else F.SetDouble('CEL_CREDITEURO', 0);
          if Q.FindField('CET_DEBITDEV').AsFloat > 0 then F.SetDouble('CEL_DEBITDEV', Q.FindField('CET_DEBITDEV').AsFloat)
                                                     else F.SetDouble('CEL_DEBITDEV', 0);
          if Q.FindField('CET_DEBIT').AsFloat > 0 then F.SetDouble('CEL_DEBITEURO', Q.FindField('CET_DEBIT').AsFloat)
                                                  else F.SetDouble('CEL_DEBITEURO', 0);
        end
        else begin
          {Est-ce possible ?}
        end;
      end

      else if Q.FindField('CET_TYPELIGNE').AsString = '07' then begin
        SetField('EE_NBMVT', I);
        SetField('EE_DATESOLDE', Q.FindField('CET_DATEOPERATION').AsDateTime);
      end;

      Q.Next;
    end;
  finally
    Ferme(Q);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.MajApresUpdate;
{---------------------------------------------------------------------------------------}
var
  Ok : Boolean;
begin
  if not Assigned(TobEexBqLig) or EstPointageManuel then Exit;
  {(VarToStr(GetField('EE_ORIGINERELEVE')) = CODEPOINTAGEMAN) then Exit;}

  try
    try
      Ok := True;
      {Saisie automatique sur le relevé et insertion dans les tables}
      if not (CtxTreso in V_PGI.PGIContexte) and not VH^.PointageJal then Ok := SaisieAutomatique(TobEexBqLig)
      else if CtxTreso in V_PGI.PGIContexte then Ok := TrSaisieAuto(TobEexBqLig);

      if not Ok then begin
        {Cela pourrait être utile pour annuler tout ??? sans conviction !}
        RollBack;
        Exit;
      end;

      {Ecriture des lignes de EEXBQLIG}
      TobEexBqLig.InsertDB(nil, True);
      {Suppression des enregistrements de CETEBAC}
      ExecuteSQL('DELETE FROM CETEBAC WHERE CET_NUMRELEVE = ' + IntToStr(FNumReleve));
    except
      on E : Exception do begin
        PGIError(TraduireMemoire('Erreur lors de la génération des écritures, avec le message :') + #13#13 + E.Message);
      end;
    end;
  finally
    FreeAndNil(TobEexBqLig);
  end;
end;

{---------------------------------------------------------------------------------------}
{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 20/02/2008
Modifié le ... :   /  /    
Description .. : - LG - 20/02/2008 - FB 22468 - on teste la valeur de retour 
Suite ........ : de genereEcriture pour stopper en cas d'erreur
Mots clefs ... : 
*****************************************************************}
function TOM_EEXBQ.SaisieAutomatique(var tRel : TOB) : Boolean;
{---------------------------------------------------------------------------------------}
var
  FInfoEcr : TInfoEcriture;
  FZEtebac : TZEtebac;
  FZEtebacGuide : TZEtebacGuide;
  FZTresoEcr : TZTresoVersEcr;
  n : Integer;
  p : Integer; {JP 30/10/07 : FQ 21763}
  F : TOB;
  B : TOB;
begin
  Result := True;

  {Si la saisie automatique est débranchée, on sort ...}
  if not GetParamSocSecur('SO_CPSAISIEAUTOPOINTAGE', False) then Exit;

  try
    try
      PrepareStructure(FInfoEcr, FZEtebac, FZEtebacGuide, FZTresoEcr);
      PrepareEcrRel(tRel, FInfoEcr, FZEtebac, FZEtebacGuide);
      if V_PGI.SAV and V_PGI.Superviseur then TobDebug(FZEtebac.TOBReleve);
      FZTresoEcr.Transfert(FZEtebac.TOBReleve);
      n := -1;
      if not FZTresoEcr.GenereEcriture(n) then
       begin
        PGIError('Impossible de générer les écritures !') ; 
        exit ;
       end ;
    except
      on E : Exception do
       begin
        PGIError(E.Message) ;
        Result := False;
        Exit;
      end;
    end;

    if V_PGI.SAV and V_PGI.Superviseur then TobDebug(FZTresoEcr.TOBEcriture);
    {Mise à jour des champs de pointage des pièces générées}
    if FZTresoEcr.TOBEcriture.Detail.Count > 0 then  begin
      {JP 30/10/07 : FQ 21763 : suite à la FQ 21342, on génère autant d pièce que d'instance correspondant à
                     un Guide et non plus une pièce : il faut donc Boucler sur les pièces
      B := FZTresoEcr.TOBEcriture.Detail[0];}
      for p := 0 to FZTresoEcr.TOBEcriture.Detail.Count - 1 do begin
        B := FZTresoEcr.TOBEcriture.Detail[p];
        for n := 0 to B.Detail.Count - 1 do begin
          F := B.Detail[n];
          if F.GetString('E_ECHE') <> 'X' then Continue;
          if F.GetString('E_NATURETRESO') <> '' then begin
            F.SetDateTime('E_DATEPOINTAGE', VarToDateTime(GetField('EE_DATEPOINTAGE')));
            F.SetString('E_REFPOINTAGE', VarToStr(GetField('EE_REFPOINTAGE')));
          end;
        end;
      end;
      n := -1;
      Result := FZTresoEcr.Save(n, False) = oeOk;
    end;

  finally
    if Assigned(FInfoEcr) then FreeAndNil(FInfoEcr);
    if Assigned(FZEtebac) then FreeAndNil(FZEtebac);
    if Assigned(FZEtebacGuide) then FreeAndNil(FZEtebacGuide);
    if Assigned(FZTResoEcr) then FreeAndNil(FZTresoEcr);
  end;
end;

{---------------------------------------------------------------------------------------}
{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 20/02/2008
Modifié le ... :   /  /    
Description .. : - LG - 22469 - parametragedes guides tous etablissement
Mots clefs ... : 
*****************************************************************}
procedure TOM_EEXBQ.PrepareStructure(var FInfo : TInfoEcriture; var FZEtebac : TZEtebac;
                                     var FZGuide : TZEtebacGuide; var FZEcr : TZTresoVersEcr);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Q := OpenSQL('SELECT J_CONTREPARTIE, J_JOURNAL, CET_NUMRELEVE ' +
               'BQ_GENERAL, J_MODESAISIE FROM CETEBAC, BANQUECP, JOURNAL ' +
               'WHERE BQ_GENERAL = J_CONTREPARTIE AND BQ_NODOSSIER = "' + FNoDossier + '" AND ' +
               'BQ_ETABBQ = CET_ETABBQ AND BQ_GUICHET = CET_GUICHET AND BQ_NUMEROCOMPTE = CET_NUMEROCOMPTE ' +
               'AND CET_TYPELIGNE = "01" AND CET_NUMRELEVE = ' + IntToStr(FNumReleve), True);
  if Q.EOF then begin
    PGIError(TraduireMemoire('Impossible de récupérer le journal pour la génération automatique d''écritures.'));
    Ferme(Q);
    Exit;
  end;

  FInfo := TInfoEcriture.Create ;
  FInfo.Devise.load([V_PGI.DevisePivot]);
  FZEtebac := TZEtebac.Create(FInfo);
  FZEtebac.InitializeObject;
  FZGuide := TZEtebacGuide.CreateTTR(FInfo, [TColDebit, TColCredit, TColLibelle]);
  try
    FZGuide.StCompteTVA            := FZEtebac.StCompteTVA;
    FZEtebac.TypeContexte          := TModeAuto;
    FZEtebac.StCompteContrepartie  := Q.FindField('J_CONTREPARTIE').AsString;
    FZEtebac.StJournalContrepartie := Q.FindField('J_JOURNAL').AsString;
    FZEtebac.Initialize;
    FZGuide.StJournalContrepartie  := FZEtebac.StJournalContrepartie;
    FZGuide.ZCompte                := FInfo.Compte;
    FZGuide.StLeNomTable           := 'CRELBQE';
    FZGuide.StEtablissement        := '' ; //VH^.EtablisDefaut ;
    FZGuide.StDevise               := FZEtebac.Devise.Code;
    FZGuide.StCompteContrepartie   := FZEtebac.StCompteContrepartie;
    FZGuide.SilentMode             := true ;
    {chargement de l'ensemble des guides}
    FZGuide.Load;

   {Création de l'objet de transfert des ecritures ( défini dans ZTresoVersEcr.pas )}
   FZEcr                       := TZTresoVersEcr.Create(FInfo);
   FZEcr.StJournalContrepartie := FZEtebac.StJournalContrepartie;
   FZEcr.StCompteContrepartie  := FZEtebac.StCompteContrepartie ;
   FZEcr.StModeSaisie          := FZEtebac.StModeSaisie;
   FZEcr.StCompteurNormal      := FZEtebac.StCompteurNormal;
   FZEcr.PTOBLettrage          := nil;
   FZEtebac.OnError            := OnError ;
   FInfo.OnError               := OnError ;
   FZEcr.OnError               := OnError ;

  finally
    Ferme(Q);
  end ;

end;

{---------------------------------------------------------------------------------------}
{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 19/09/2007
Modifié le ... :   /  /    
Description .. : - LG - 19/09/2007 - FB 21342 - nuemrotation des pieces
Mots clefs ... : 
*****************************************************************}
procedure TOM_EEXBQ.PrepareEcrRel(TRel : TOB; FInfo : TInfoEcriture; FZEtebac : TZEtebac; FZGuide : TZEtebacGuide);
{---------------------------------------------------------------------------------------}
var
 lTOB                       : TOB;
 lStChampsDateComptable     : string;
 lStChampsDateValeur        : string;
 lStTraitementImport        : string ;
 MyBool                     : Boolean;
 n, k                       : Integer;
begin
  {Chargement des info sur le journal}
  if FInfo.Journal.Load([FZEtebac.StJournalContrepartie ]) = - 1 then Exit;
  lStChampsDateComptable := FInfo.Journal.GetValue('J_CHOIXDATE');
  lStTraitementImport    := FInfo.Journal.GetValue('J_TRESOIMPORT');

  if lStChampsDateComptable = '' then
    lStChampsDateComptable := 'DATEOPERATION';

  if lStChampsDateComptable = 'DATEOPERATION' then lStChampsDateValeur := 'DATEVALEUR'
                                              else lStChampsDateValeur := 'DATEOPERATION';

  k := 1;
  for n := 0 to TRel.Detail.Count - 1 do begin

    FZEtebac.AddLigneReleve;
    FZEtebac.StIMPORT_R        := 'X';
    FZEtebac.StGENERALBQE_R    := FZEtebac.StCompteContrepartie; {compte de banque}
    FZEtebac.StJOURNAL_R       := FZEtebac.StJournalContrepartie;
    FZEtebac.DtDATECOMPTABLE_R := TRel.Detail[n].GetDateTime('CEL_' + lStChampsDateComptable);
    FZEtebac.DtDATEVALEUR_R    := TRel.Detail[n].GetDateTime('CEL_' + lStChampsDateValeur);
    SetReferences(FZEtebac, TRel.Detail[n]);
    {Gestion des creditdev et debitdev ????}
    if TRel.Detail[n].GetString('CEL_DEVISE') = V_PGI.DevisePivot then
      CSetMontants( FZEtebac.TOBLigneReleve, TRel.Detail[n].GetDouble('CEL_DEBITEURO'), TRel.Detail[n].GetDouble('CEL_CREDITEURO'), FZEtebac.Devise, True)
    else
      CSetMontants( FZEtebac.TOBLigneReleve, TRel.Detail[n].GetDouble('CEL_DEBITDEV'), TRel.Detail[n].GetDouble('CEL_CREDITDEV'), FZEtebac.Devise, True);

    FZEtebac.StNATUREPIECE_R := 'OD' ;
    FZEtebac.StLIBELLE_R     := TRel.Detail[n].GetString('CEL_LIBELLE');
    FZEtebac.StCODEAFB_R     := TRel.Detail[n].GetString('CEL_CODEAFB');
    FZEtebac.StREFPIECE_R    := TRel.Detail[n].GetString('CEL_REFPIECE');
    FZEtebac.StDEVISE_R      := TRel.Detail[n].GetString('CEL_DEVISE');
    {Gestion des libellés enrichis dans la recherche du guide}
    FZEtebac.StLibEnrichi1_R := TRel.Detail[n].GetString('CEL_LIBELLE1');
    FZEtebac.StLibEnrichi2_R := TRel.Detail[n].GetString('CEL_LIBELLE2');
    FZEtebac.StLibEnrichi3_R := TRel.Detail[n].GetString('CEL_LIBELLE3');

    FZEtebac.StETABLISSEMENT_R := VH^.EtablisDefaut ;
    FZEtebac.InCOMPTEUR_R      := FZEtebac.ProchainCompteur;  //FNumReleve;

    lTOB   := TOB.Create( '' , nil , -1 );
    MyBool := FZGuide.RechercheGuideEnBase( FZEtebac.TOBLigneReleve , lTOB , nil ) ;

    if MyBool then begin
      {On affecte les enregistrement contenus dans lTOB à la ligne de relevé courante}
      FZEtebac.CreerImputation(lTOB);
      {On supprime les lignes vides}
      FZEtebac.CalculSoldeImput;

      if FZEtebac.IsValideImputation then
        FZEtebac.StETAT_R := cImputCorrectNonValide
      else
        FZEtebac.StETAT_R := cImputIncorrecte;

      {JP 30/10/07 : FQ 21763 : Si 'limputation n'est pas correcte, l'écriture ne sera pas créée,
                     donc il faut pointer l'imputation}  
      if FZEtebac.StETAT_R = cImputCorrectNonValide then begin
        TRel.Detail[n].SetDateTime('CEL_DATEPOINTAGE', VarToDateTime(GetField('EE_DATEPOINTAGE')));
        TRel.Detail[n].SetString('CEL_REFPOINTAGE', VarToStr(GetField('EE_REFPOINTAGE')));
        TRel.Detail[n].SetString('CEL_CODEPOINTAGE', IntToStr(k));
        FZEtebac.StCodePointage_I := IntToStr(k);
        FZEtebac.StCodeAfb_i      := TRel.Detail[n].GetString('CEL_CODEAFB');
        Inc(k);
      end;
    end;

    lTOB.Free ;
  end; {for n := 0 to }
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.SetReferences(FZEtebac: TZEtebac; tLigneRel : TOB);
{---------------------------------------------------------------------------------------}
begin
  if Trim(tLigneRel.GetString('CEL_LIBELLE1')) <> '' then begin
    FZEtebac.StREFINTERNE_R := tLigneRel.GetString('CEL_LIBELLE1');
    if Trim(tLigneRel.GetString('CEL_LIBELLE2')) <> '' then
      FZEtebac.StREFEXTERNE_R := tLigneRel.GetString('CEL_LIBELLE2')
    else
      FZEtebac.StREFEXTERNE_R := tLigneRel.GetString('CEL_REFPIECE');
  end
  else
    FZEtebac.StREFINTERNE_R := tLigneRel.GetString('CEL_REFPIECE');
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.InitRefPointage;
{---------------------------------------------------------------------------------------}
var
  a, m, j : Word;
  s : string;
begin
  if IsValidDate(EE_DATEPOINTAGE.Text) and CanInitReference then begin
    DecodeDate(StrToDate(EE_DATEPOINTAGE.Text), a, m, j);
    if a < 2000 then a := a - 1900
                else a := a - 2000;

    s := Copy(EE_GENERAL.Text, 1, 8) + '-' + PadL(IntToStr(j), '0', 2) + '/' +
                                             PadL(IntToStr(m), '0', 2) + '/' +
                                             PadL(IntToStr(a), '0', 2);
    SetField('EE_REFPOINTAGE', s);
  end;

  {Pour griser / dégriser les zones des montants
   19/06/07 : FQ TRESO 10477 : on désactive toujours les zones de solde pour les relevés automatiques
  GereAffichageMnt((GetField('EE_ORIGINERELEVE') = ORIGINERELEVE) or (GetField('EE_ORIGINERELEVE') = CODENEWPOINTAGE));}
  GereAffichageMnt(not EstPointageManuel);

  {19/07/07 : FQ 20603 : Par défaut, on peut initialiser la référence de pointage}
  CanInitReference := True;
  OldValueG := VarToStr(GetField('EE_GENERAL'));
  OldValueD := VarToDateTime(GetField('EE_DATEPOINTAGE'));
end;

{Saisie automatique pour la Trésorerie
{---------------------------------------------------------------------------------------}
function TOM_EEXBQ.TrSaisieAuto(var tRel : TOB): Boolean;
{---------------------------------------------------------------------------------------}
var
  Obj  : TobjSaisieAuto;
  Gene : string;
  {$IFDEF TRESO}
  TobCompta : TobPieceCompta;
  {$ENDIF TRESO}
begin
  Result := True;
  Gene := FGeneral;
  if not EstPointageSurTreso then
    Gene := GetBqCodeFromGene(Gene, V_PGI.NoDossier);

  Obj := TobjSaisieAuto.Create(tRel, Gene, VarToStr(GetField('EE_REFPOINTAGE')), FStDevise, VarToDateTime(GetField('EE_DATEPOINTAGE')));
  try
    try
      Obj.GenereEcritures;
      {$IFDEF TRESO}
      if not GetParamSocSecur('SO_TRINTEGAUTO', False) then Exit;
      if not AutoriseFonction(dac_Integration) then Exit;
      if Obj.TobEcriture.Detail.Count = 0 then Exit;
      {18/09/06 : Nouvelle gestion Multi-sociétés}
      TobCompta := TobPieceCompta.Create('***', nil, -1);
      try
        {18/09/06 : Nouvelle intégration des écritures en comptabilité (gestion du multi-sociétés)}
        TRGenererPieceCompta(TobCompta, Obj.TobEcriture);
        if TobCompta.Detail.Count > 0 then begin
          {18/09/06 : Nouvelle fonction de traitement Multi-sociétés}
          TRIntegrationPieces(TobCompta, False);
        end
        else
          AfficheMessageErreur;
      finally
        FreeAndNil(TobCompta);
      end;
      {$ENDIF TRESO}

    except
      on E : Exception do begin
        Result := False;
        PGIError(E.Message, Ecran.Caption);
      end;
    end;
  finally
    FreeAndNil(Obj);
  end;
end;

{18/07/07 : FQ 21119 : Gestion du F6 pour la liste des relevés
{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQ.EcranOnKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
{---------------------------------------------------------------------------------------}
begin
  case Key of
    VK_F6     : ReleveOnClick(GetControl('BRELEVE'));
  else
    FEcranKeyDown(Sender, Key, Shift);
  end;
end;

{14/11/07 : Le pointage manuel concerne les natures MAN, NBQ et vide
{---------------------------------------------------------------------------------------}
function TOM_EEXBQ.EstPointageManuel : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := (GetField('EE_ORIGINERELEVE') = CODEPOINTAGENBQ) or
            (GetField('EE_ORIGINERELEVE') = CODEPOINTAGEMAN) or
            (GetField('EE_ORIGINERELEVE') = '');
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 20/02/2008
Modifié le ... :   /  /    
Description .. : - LG - 20/02/2008 - FB 22468 - on teste la valeur de retour 
Suite ........ : de genereEcriture pour stopper en cas d'erreur
Mots clefs ... : 
*****************************************************************}
procedure TOM_EEXBQ.OnError( sender : TObject ; Error : TRecError);
begin
 if ( trim(Error.RC_Message) <> '' ) then
  PGIInfo(Error.RC_Message + #10#13 + Error.RC_Methode , Ecran.Caption )
   else
    if Error.RC_Error <> RC_PASERREUR  then
      FMessCompta.Execute(Error.RC_Error) ;
end;

initialization
  RegisterClasses([TOM_EEXBQ]);

end.
