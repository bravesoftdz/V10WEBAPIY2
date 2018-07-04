{-------------------------------------------------------------------------------------
    Version   | Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
               18/03/03  GCO  Création de l'unité : mul sur la Table EEXBQ (Référence / Session de pointage)
 8.01.001.004  06/02/07  JP   Uniformisation de l'utilisation de la Table EEXBQ dans le pointage et le
                              rapprochement : ce mul sert à la fois d'entrée au pointage / rapprochement
                              ainsi qu'au dépointage
 8.01.001.011  16/04/07  JP   FQ 19380 : En communication BLine, si pointage Client, on désactive les menus principaux
 8.01.001.013  04/05/07  JP   FQ TRESO 10449 : On filtre sur les comptes de banque
 8.01.001.019  14/06/07  JP   Petites corrections concernant le pointage en cash pooling
 8.00.001.022  27/06/07  JP   FQ 20851 : gestion de la touche Suppr sur le KeyDown
 8.00.001.025  16/07/07  JP   FQ 21041 : On désactive le ctrl + A et le BSelectAll
 8.00.001.025  16/07/07  JP   FQ 19380 (suite) : on ne peut pas modifier une référence de pointage si en consultation
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
 8.10.001.011  01/10/07  JP   FQ 19380 (suite) : quelques finalisations sur les échanges BLine
 8.10.001.012  02/10/07  JP   Ajout d'un filtre sur les anciens relevés si pointage sur TRECRITURE
08.10.005.001  14/11/07  JP   Gestion des comptes pointables qui ne sont pas bancaires  
--------------------------------------------------------------------------------------}
unit uTOFPointageMul;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  Maineagl, eMul,
  {$ELSE}
  Fe_main, Mul, DB, Hdb,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  SysUtils, ComCtrls, HCtrls,
  {$IFDEF TRCONF}
  uLibConfidentialite,
  {$ELSE}
  UTOF,
  {$ENDIF TRCONF}
  Htb97, Menus;

type
  {$IFDEF TRCONF}
  TOF_POINTAGEMUL = class (TOFCONF)
  {$ELSE}
  TOF_POINTAGEMUL = class (TOF)
  {$ENDIF TRCONF}
    procedure OnArgument(S : string); override;
    procedure OnLoad                ; override;
  private
    FBoConsultation : Boolean;
    FBoNewPointage  : Boolean;
    FEcranKeyDown   : TKeyEvent;
    FBoPtgeNonBQ    : Boolean; {14/11/07}

    procedure InitControles;
    procedure InitEvenements;
    procedure InitAffichage(Arg : string);
    function  CSupprimeRefPointage(vCompteOuJournal, vRefPointage : string; vDatePointage : TDateTime) : Boolean;
  public
    {$IFDEF EAGLCLIENT}
    FListe : THGrid;
    {$ELSE}
    FListe : THDBGrid;
    {$ENDIF}

    TEE_GENERAL : THLabel;
    EE_GENERAL : THEdit;
    EE_DATEPOINTAGE : THEdit;
    EE_REFPOINTAGE : THEdit;

    BInsert : TToolBarButton97;
    BDelete : TToolBarButton97;
    BInfo : TToolBarButton97;
    BCherche : TToolBarButton97;
    BOuvrir : TToolBarButton97;
    BRechercher : TToolBarButton97;
    BEtatRappro : TToolBarButton97;

    POPF11 : TPopUpMenu;

    procedure OnDblClickFListe  (Sender : TObject);
    procedure OnClickBInsert    (Sender : TObject);
    procedure OnClickBDelete    (Sender : TObject);
    procedure OnClickBInfo      (Sender : TObject);
    procedure OnClickBEtatRappro(Sender : TObject);

    procedure OnElipsisClickEE_General(Sender : TObject);
    procedure OnElipsisClickEE_DatePointage(Sender : TObject);
    procedure OnElipsisClickEE_RefPointage(Sender : TObject);
    procedure OnExitEE_General(Sender : TObject);

    procedure OnKeyDownEcran(Sender : TObject; var Key : Word; Shift : TShiftState);

    procedure SuppressionReference(pCpte, pRef : string; pDatePtge : TDateTime);
  end;



procedure CPLanceFiche_PointageMul(vStParam : string = '');


implementation


uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  {$IFDEF VER150} Variants, {$ENDIF}
  {$IFDEF COMPTA} uTOFPointageEcr, {$ENDIF COMPTA}
  LookUp {LookUpList}, Forms {wsMaximized}, CpteSav{RecalculTotPointeNew1},
  HEnt1, HMsgBox, ParamSoc, Ent1, Windows {VK_F11 ...}, Constantes, 
  uTOMEEXBQ, ULibPointage, CPPOINTAGERAP_TOF, CPRAPPRODET_TOF;


{---------------------------------------------------------------------------------------}
procedure CPLanceFiche_PointageMul(vStParam : string = '');
{---------------------------------------------------------------------------------------}
begin
  if (ctxTreso in V_PGI.PGIContexte) and VH^.PointageJal then
    PGIError(TraduireMemoire('Le pointage sur journal est incompatible avec la Trésorerie'), TraduireMemoire('Pointage'))
  else
    AGLLanceFiche('CP', 'CPPOINTAGEMUL', '', '', vStParam);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_POINTAGEMUL.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRCONF}
  {08/08/07 : Pour le moment, seule la liste TREEXBQ gère BANQUECP}
  if EstPointageSurTreso then
    TypeConfidentialite := tyc_Banque + ';';
  {$ENDIF TRCONF}
  inherited;
  if ctxTreso in V_PGI.PGIContexte then Ecran.HelpContext := 7774000
                                   else Ecran.HelpContext := 7604000;
                                   
  {01/10/07 : FQ 19380 : mieux vaut initialiser le Boolean au début plutôt que dans InitAffichage}
  FBoConsultation := CEstPointageEnConsultationSurDossier;

  {Récupère les composants}
  InitControles;
  {Surcharge des évènements}
  InitEvenements;
  {Initialise les composants de la fiche}
  InitAffichage(S);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_POINTAGEMUL.OnLoad;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  {04/05/07 : FQ 10449 : exclusion des éventuels comptes divers pointables}
  if CtxTreso in V_PGI.PGIContexte then begin
    if not EstPointageSurTreso then
//      SetControlText('XX_WHERE', 'EE_GENERAL IN (SELECT BQ_CODE FROM BANQUECP)')
  //  else
      SetControlText('XX_WHERE', 'EE_GENERAL IN (SELECT BQ_GENERAL FROM BANQUECP)');
  end
  {02/10/07 : si on est en compta avec pointage sur TRECRITURE, on ne traitre que les "vieilles" sessions}
  else if EstPointageCache then
    {14/11/07 : Le filtre est trop strict, car on a pu créer des référence de pointage avec le nouveau pointage
                avant de passer en Cash pooling : avec le test sur EE_ORIGINERELEVE en compta et la requête de la
                liste qui se fait sur BQ_CODE en tréso, ces sessions de pointage auraient été invisibles dans les
                deux produits. Si on pointe sur TRECRITURE, on met EE_STATUTRELEVE à "TRE"
     SetControlText('XX_WHERE', 'NOT (EE_ORIGINERELEVE IN ("' + CODENEWPOINTAGE + '","' + CODEPOINTAGEMAN + '"))');}
    SetControlText('XX_WHERE', '(EE_STATUTRELEVE <> "' + SESSIONTRESO + '" OR EE_STATUTRELEVE IS NULL)');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_POINTAGEMUL.InitControles;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF EAGLCLIENT}
  FListe := THGrid(GetControl('FLISTE', True));
  {$ELSE}
  FListe := THDBGrid(GetControl('FLISTE', True));
  {$ENDIF}
  POPF11 := TPopUpMenu(GetControl('POPF11', True));

  TEE_GENERAL     := THLabel(GetControl('TEE_GENERAL'    , True));
  EE_GENERAL      := THEdit (GetControl('EE_GENERAL'     , True));
  EE_DATEPOINTAGE := THEdit (GetControl('EE_DATEPOINTAGE', True));
  EE_REFPOINTAGE  := THEdit (GetControl('EE_REFPOINTAGE' , True));

  BInsert     := TToolBarbutton97(GetControl('BINSERT'    , True));
  BDelete     := TToolBarbutton97(GetControl('BDELETE'    , True));
  BInfo       := TToolBarbutton97(GetControl('BINFO'      , True));
  BCherche    := TToolBarButton97(GetControl('BCHERCHE'   , True));
  BOuvrir     := TToolBarButton97(GetControl('BOUVRIR'    , True));
  BRechercher := TToolBarButton97(GetControl('BRECHERCHER', True));
  BEtatRappro := TToolBarButton97(GetControl('BETATRAPPRO', True));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_POINTAGEMUL.InitEvenements;
{---------------------------------------------------------------------------------------}
begin
  EE_GENERAL     .OnElipsisClick := OnElipsisClickEE_GENERAL;
  EE_DATEPOINTAGE.OnElipsisClick := OnElipsisClickEE_DATEPOINTAGE;
  EE_REFPOINTAGE .OnElipsisClick := OnElipsisClickEE_REFPOINTAGE;

  EE_General .OnExit     := OnExitEE_General;
  FListe     .OnDblClick := OnDblClickFListe;
  BInsert    .OnClick    := OnClickBInsert;
  BDelete    .OnClick    := OnClickBDelete;
  BInfo      .OnClick    := OnClickBInfo;
  BEtatRappro.OnClick    := OnClickBEtatRappro;

  {JP 16/04/07 : FQ 19380 : On désactive le menu F11 en Pointage Client avec BLine}
  POPF11.Items[0].Visible := not FBoConsultation;
  POPF11.Items[1].Visible := not FBoConsultation;
  
  POPF11.Items[0].OnClick := OnClickBInsert;
  POPF11.Items[1].OnClick := OnClickBDelete;
  POPF11.Items[2].OnClick := OnDblClickFListe;
  POPF11.Items[3].OnClick := OnClickBInfo;
  POPF11.Items[4].OnClick := OnClickBEtatRappro;

  AddMenuPop(PopF11, '', '');

  FEcranKeyDown := Ecran.OnKeyDown;
  Ecran.OnKeyDown := OnKeyDownEcran;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_POINTAGEMUL.InitAffichage(Arg : string);
{---------------------------------------------------------------------------------------}
var
  lStArg : string;
begin
  {$IFDEF IMP}
  BInfo  .Visible := False;
  BInsert.Visible := False;
  {$ENDIF IMP}
  FBoNewPointage := (Pos(CODENEWPOINTAGE, Arg) > 0) or EstPointageSurTreso;
  if FBoNewPointage then ReadTokenSt(Arg);

  if FBoConsultation then begin
    SetControlEnabled('BINSERT', False);
    SetControlEnabled('BDELETE', False);
    SetControlEnabled('BVALIDER', False);
  end;

  {14/06/07 : Si pointage sur TRECRITURE, on interdit la création de nouvelles références en compta
   14/11/07 : si, il faut autoriser pour les comptes non bancaires}
  if EstPointageCache and not (CtxTreso in V_PGI.PGIContexte) then begin
    //SetControlEnabled('BINSERT', False);
    FBoPtgeNonBQ := True;
  end
  else
    FBoPtgeNonBQ := False;

  if Arg <> '' then begin
    lStArg := ReadTokenSt(Arg);
    if lStArg <> '' then begin
      EE_General.Text := ReadTokenSt(lStArg);
      EE_General.Enabled := False;
      Ecran.WindowState := wsMaximized;
    end;
  end;

  if not VH^.PointageJal then begin
    Ecran      .Caption := TraduireMemoire('Pointage sur compte');
    if EstPointageSurTreso then begin
      TEE_GENERAL.Caption := TraduireMemoire('Compte bancaire') ;
      EE_GENERAL.MaxLength := 17;
      TFMul(Ecran).SetDBListe('TREEXBQ');
    end
    else begin
      TEE_GENERAL.Caption := TraduireMemoire('Compte général') ;
      EE_GENERAL.MaxLength := VH^.CPta[fbGene].Lg;
    end;
  end
  else begin
    lStArg := ReadTokenSt(Arg);
    if lStarg <> '' then EE_REFPOINTAGE.Text := lStArg;
    Ecran.Caption := TraduireMemoire('Pointage sur journal');
    TEE_GENERAL.Caption := TraduireMemoire('Journal');
    EE_GENERAL.MaxLength := 3;
    TFMul(Ecran).SetDBListe('CPEEXBQ2');
  end;

  UpDateCaption(Ecran);

  {JP 16/07/07 : FQ 21041 : On désactive le ctrl + A et le BSelectAll}
  SetControlVisible('BSELECTALL', False);
end;

{Suppression de la dernière référence de pointage
{---------------------------------------------------------------------------------------}
procedure TOF_POINTAGEMUL.OnClickBDelete(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  sCpte : string;
  sRefP : string;
  aDatP : TDateTime;
begin
  if FBoConsultation then Exit; {FQ 19380}

  if EE_GENERAL.Text <> '' then begin
    if VH^.PointageJal then sCpte := TraduireMemoire('du journal ') + EE_GENERAL.Text
                       else sCpte := TraduireMemoire('du compte ') + EE_GENERAL.Text;
  end;

  if PGIAsk(TraduireMemoire('Vous allez supprimer la session de pointage sélectionnée.') + #13 +
            TraduireMemoire('Seule la dernière session de pointage') + ' ' + sCpte + ' ' +
            TraduireMemoire('peut être supprimée.') + #13 +
            TraduireMemoire('Les écritures seront dépointées et les mouvements bancaires supprimés.') + #13#13 +
            TraduireMemoire('Souhaitez-vous poursuivre ?'), Ecran.Caption ) = mrYes then begin
    sCpte := VarToStr(GetField('EE_GENERAL'));
    sRefP := VarToStr(GetField('EE_REFPOINTAGE'));
    aDatP := VarToDateTime(GetField('EE_DATEPOINTAGE'));

    if Trim(sCpte) = '' then
      PGIError(TraduireMemoire('Impossible de récupérer la valeur du champ général.'), Ecran.Caption )

    else if aDatP <= iDate1900 then
      PGIError(TraduireMemoire('Impossible de récupérer la date de la session de pointage.'), Ecran.Caption )

    else if Trim(sRefP) = '' then
      PGIError(TraduireMemoire('Impossible de récupérer la référence de pointage.'), Ecran.Caption )

    else if not CanDeleteRefPtge(sCpte, aDatP) then
      PGIError(TraduireMemoire('Il existe une session postérieure à celle sélectionnée.') + #13 +
               TraduireMemoire('Il n''est pas possible de poursuivre le traitement.'), Ecran.Caption )

    else
      SuppressionReference(sCpte, sRefP, aDatP);
  end;
  TFMul(Ecran).BChercheClick(nil);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_POINTAGEMUL.SuppressionReference(pCpte, pRef : string; pDatePtge : TDateTime);
{---------------------------------------------------------------------------------------}
var
  Ok : Boolean;
begin
  Ok := True;
  BeginTrans;
  try
    {Détruit l'enregistrement de EEXBQ ...}
    if CSupprimeRefPointage(pCpte, pRef, pDatePtge) and
       {... Suppression du pointage dans (TR)ECRITURE et EEXBQLIG}
       CDepointeEcriture(pCpte, pRef, GetField('EE_NUMERO'), pDatePtge) then begin
      {Mise à jour des soldes}
      RecalculTotPointeNew1(pCpte);
      CommitTrans;
    end
    else begin
      RollBack;
      Ok := False;
    end;
  except
    on E : Exception do begin
      RollBack;
      Ok := False;
      PGIError(TraduireMemoire('Une erreur est intervenue lors de la suppression d''une référence') + #13 +
               TraduireMemoire('de pointage avec le message :') + #13#13 + E.Message);

    end;
  end;
  if Ok then PgiInfo(TraduireMemoire('Le traitement s''est correctement terminé.'), Ecran.Caption);
end;

{---------------------------------------------------------------------------------------}
function TOF_POINTAGEMUL.CSupprimeRefPointage(vCompteOuJournal, vRefPointage : string; vDatePointage : TDateTime) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := True;
  try
    ExecuteSQL('DELETE FROM EEXBQ WHERE EE_GENERAL = "' + vCompteOuJournal + '"' +
               ' AND EE_REFPOINTAGE = "' + vRefPointage + '"' +
               ' AND EE_DATEPOINTAGE = "' + UsDateTime(vDatePointage) + '"');
  except
    on E : Exception do begin
      PgiError(TraduireMemoire('Erreur SQL : ') + E.Message, TraduireMemoire('Fonction : CSupprimeRefPointage'));
      Result := False;
    end;
  end;
end;

{Creation d'une session de pointage
{---------------------------------------------------------------------------------------}
procedure TOF_POINTAGEMUL.OnClickBInsert(Sender : TObject);
{---------------------------------------------------------------------------------------}
{$IFNDEF IMP}
var
  lBoPresence : Boolean;
  St          : string;
{$ENDIF}
begin
  if FBoConsultation then Exit; {FQ 19380}
{$IFNDEF IMP}
  lBoPresence := False;
  if Trim(EE_General.Text) <> '' then begin
    if VH^.PointageJal then
      lBoPresence := Presence('JOURNAL', 'J_JOURNAL', EE_General.Text)
    else begin
      if EstPointageSurTreso then
        lBoPresence := Presence('BANQUECP', 'BQ_CODE', EE_General.Text)
      else begin
        if EE_General.Text <> '' then EE_General.Text := BourreEtLess(EE_General.Text, fbGene);
        lBoPresence := Presence('GENERAUX', 'G_GENERAL', EE_General.Text);
      end;
    end;

    if not lBoPresence then begin
      if VH^.PointageJal then
        PGIError(TraduireMemoire('Vous ne pouvez pas créer une session de pointage sur un journal inexistant.'), Ecran.Caption)
      else
        PGIError(TraduireMemoire('Vous ne pouvez pas créer une session de pointage sur un compte inexistant.'), Ecran.Caption);
      EE_General.SetFocus;
      EE_General.ElipsisClick(Sender);
    end;
  end;

  if (Trim(EE_General.Text) = '') or (lBoPresence) then begin
    St :=  'ACTION=CREATION;' + EE_GENERAL.Text + ';';
    {14/11/07 : gestion du pointage sur comptes non bancaires en compta si pointage sur TRECRITURE}
    if FBoPtgeNonBQ then St := St + CODEPOINTAGENBQ + ';'
                    else St := St + ';';
    CPLanceFiche_PointageFic('', St);
    TFMUL(Ecran).BChercheClick(nil);
  end;
{$ENDIF}
end;

{Lancement de la fiche de session de pointage (Tom de EEXBQ)
{---------------------------------------------------------------------------------------}
procedure TOF_POINTAGEMUL.OnClickBInfo(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {$IFNDEF IMP}
  if VarToStr(GetField('EE_GENERAL')) <> '' then begin
    {16/07/07 : FQ 19380 : on ne peut pas modifier une référence de pointage si en consultation}
    if FBoConsultation then
      CPLanceFiche_PointageFic(GetField('EE_GENERAL') + ';' + DateToStr(GetField('EE_DATEPOINTAGE')) + ';' + GetField('EE_REFPOINTAGE') + ';' + IntToStr(GetField('EE_NUMERO')), 'ACTION=CONSULTATION')
    else
      CPLanceFiche_PointageFic(GetField('EE_GENERAL') + ';' + DateToStr(GetField('EE_DATEPOINTAGE')) + ';' + GetField('EE_REFPOINTAGE') + ';' + IntToStr(GetField('EE_NUMERO')), 'ACTION=MODIFICATION');
    TFMUL(ECRAN).BChercheClick(nil);
  end;
  {$ENDIF}
end;

{Lancement de l'état de rapprochement bancaire
{---------------------------------------------------------------------------------------}
procedure TOF_POINTAGEMUL.OnClickBEtatRappro(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if VarToStr(GetField('EE_GENERAL')) <> '' then begin
    CC_LanceFicheEtatRapproDet(GetField('EE_GENERAL') + ';' +
                               DateToStr(GetField('EE_DATEPOINTAGE'))+ ';X');
    TFMUL(ECRAN).BChercheClick(nil);
  end;
end;

{Lancement de l'écran de rapprochement bancaire
{---------------------------------------------------------------------------------------}
procedure TOF_POINTAGEMUL.OnDblClickFListe(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  ActionFiche : string;
begin
  if VarToStr(GetField('EE_GENERAL')) <> '' then begin
    if FBoNewPointage then begin
      (* JP 09/10/07 : FQ 20196 : A ne surtout pas porter "la mise en commentaire" en V800

      {20/04/07 : FQ TRESO 10436 : L'ancien rappro ne gérait pas la devise dans EEXBQLIG}
      if VarToStr(GetField('EE_ORIGINERELEVE')) = ORIGINERELEVE then begin
        ExecuteSQL('UPDATE EEXBQLIG SET CEL_DEVISE = "' + VarToStr(GetField('EE_DEVISE')) +
                   '" WHERE CEL_REFPOINTAGE = "' +  VarToStr(GetField('EE_REFPOINTAGE')) +
                   '" AND CEL_GENERAL = "' +  VarToStr(GetField('EE_GENERAL')) +
                   '" AND CEL_NUMRELEVE = ' +  VarToStr(GetField('EE_NUMERO')) +
                   ' AND (CEL_DEVISE = "" OR CEL_DEVISE IS NULL)');
        ExecuteSQL('UPDATE EEXBQLIG SET CEL_REFPOINTAGE = ""' +
                   ' WHERE CEL_REFPOINTAGE = "' +  VarToStr(GetField('EE_REFPOINTAGE')) +
                   '" AND CEL_GENERAL = "' +  VarToStr(GetField('EE_GENERAL')) +
                   '" AND CEL_NUMRELEVE = ' +  VarToStr(GetField('EE_NUMERO')) +
                   ' AND (CEL_DATEPOINTAGE = "' + UsDateTime(iDate1900) + '" OR CEL_DATEPOINTAGE IS NULL)');
      end;
        *)

      {14/06/07 : Si pointage sur TRECRITURE, on interdit la modification des nouvelles références en compta
       01/10/07 : FQ 19380 : On met en lecture seule si l'on est en échange BLine}
      if FBoConsultation or (EstPointageCache and not (CtxTreso in V_PGI.PGIContexte) and not
         ((GetField('EE_ORIGINERELEVE') = ORIGINERELEVE) or (GetField('EE_ORIGINERELEVE') = '') or
          (GetField('EE_ORIGINERELEVE') = CODEPOINTAGENBQ) or (GetField('EE_STATUTRELEVE') <> SESSIONTRESO))) then
        ActionFiche := 'ACTION=CONSULTATION;'
      else
        ActionFiche := 'ACTION=MODIFICATION;';

      CPLanceFiche_PointageRappro(ActionFiche + GetField('EE_GENERAL') + ';' + DateToStr(GetField('EE_DATEPOINTAGE')) +
                                  ';' + GetField('EE_REFPOINTAGE') + ';' + IntToStr(GetField('EE_NUMERO')) + ';' +
                                  VarToStr(GetField('EE_ORIGINERELEVE')) + ';')
    end
    else begin
    {$IFDEF COMPTA}
      if (GetField('EE_ORIGINERELEVE') = ORIGINERELEVE) or (GetField('EE_ORIGINERELEVE') = CODENEWPOINTAGE) then begin
        PGIBox('Vous ne pouvez pas accéder au pointage manuel sur une référence de pointage automatique.'#13'Reportez vous à la commande Pointage automatique du menu Relevés.');
        CPLanceFiche_Pointage(GetField('EE_GENERAL') + ';' + DateToStr(GetField('EE_DATEPOINTAGE')) + ';' + GetField('EE_REFPOINTAGE') + ';' + IntToStr(GetField('EE_NUMERO')) + ';ACTION=CONSULTATION');
      end
      else
        CPLanceFiche_Pointage(GetField('EE_GENERAL') + ';' + DateToStr(GetField('EE_DATEPOINTAGE')) + ';' + GetField('EE_REFPOINTAGE') + ';' + IntToStr(GetField('EE_NUMERO')));
    {$ENDIF COMPTA}
    end;

    TFMUL(ECRAN).BChercheClick(nil);
  end;
end;

{Affichage des Comptes de Banque Pointable ou des Journaux de type Banque ayant leur
 compte de contrepartie Pointable
{---------------------------------------------------------------------------------------}
procedure TOF_POINTAGEMUL.OnElipsisClickEE_General(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  LookUpGenePtge(Sender);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_POINTAGEMUL.OnElipsisClickEE_DatePointage(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  LookUpList(THEdit(Sender), TraduireMemoire('Date de pointage'), 'EEXBQ', 'EE_DATEPOINTAGE',
      'EE_REFPOINTAGE', 'EE_GENERAL="' + EE_General.Text + '"', 'EE_DATEPOINTAGE DESC', True, 0);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_POINTAGEMUL.OnElipsisClickEE_RefPointage(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
    LookUpList(THEdit(Sender), TraduireMemoire('Référence de pointage'), 'EEXBQ', 'EE_REFPOINTAGE',
        'EE_DATEPOINTAGE', 'EE_GENERAL="' + EE_General.Text + '"', 'EE_DATEPOINTAGE DESC', True, 0);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_POINTAGEMUL.OnKeyDownEcran(Sender : TObject; var Key : Word; Shift : TShiftState);
{---------------------------------------------------------------------------------------}
var
  OldKey : Word;
begin
  OldKey := Key;
  Key := 0;

  case OldKey of
    {JP 16/07/07 : FQ 21041 : On désactive le ctrl + A et le BSelectAll
     Ord('A') : if Shift = [ssCtrl] then TToolBarButton97(GetControl('BSelectAll')).Click;}

    Ord('N') : if Shift = [ssCtrl] then BInsert.Click;

    Ord('F') : if Shift = [ssCtrl] then BRechercher.Click;

    VK_DELETE : if Shift = [ssCtrl] then BDelete.Click
                                    else Key := VK_DELETE; {27/06/07 : FQ 20851}

    VK_F5 : begin
                   if FListe.Focused then BOuvrir.Click
              else if EE_GENERAL.Focused then EE_GENERAL.ElipsisClick(EE_GENERAL)
              else if EE_DATEPOINTAGE.Focused then EE_GENERAL.ElipsisClick(EE_DATEPOINTAGE)
              else if EE_REFPOINTAGE.Focused then EE_GENERAL.ElipsisClick(EE_REFPOINTAGE);
            end;

    VK_F10 : BOuvrir.Click;

    VK_F11 : POPF11.Popup(Mouse.CursorPos.x, Mouse.CursorPos.y);

    VK_F12 : if FListe.Focused then TPageControl(GetControl('PAGES', False)).SetFocus
                               else FListe.SetFocus;
  else
    Key := OldKey;
  end;

  FEcranKeyDown(Sender, Key, Shift);
  if Key = VK_F9 then FListe.SetFocus;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_POINTAGEMUL.OnExitEE_General(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  lBoPresence : Boolean;
  lGeneral    : string;
begin
  if Trim(EE_General.Text) = '' then Exit;
  {27/06/07 : Pour éviter un message si on quitte la zone pour quitter la fiche}
  if not EE_General.CanFocus then Exit;

  if VH^.PointageJal then
    lBoPresence := Presence('JOURNAL', 'J_JOURNAL', EE_General.Text)
  else begin
    if EstPointageSurTreso then
      lBoPresence := Presence('BANQUECP', 'BQ_CODE', EE_General.Text)
    else begin
      {27/06/07 : Refonte de la présence du général}
      if EE_General.Text <> '' then lGeneral := BourreEtLess(EE_General.Text, fbGene);
      {Complétion de la saisie avec le caractère de bourrage}
      lBoPresence := Presence('GENERAUX', 'G_GENERAL', lGeneral);
      {Si la complétion existe, on remplit le composant}
      if lBoPresence then EE_General.Text := lGeneral;
    end;
  end;

  if not lBoPresence then begin
    EE_General.SetFocus;
    EE_General.ElipsisClick(Sender);
  end;
end;

initialization
    RegisterClasses([TOF_POINTAGEMUL]);

end.
