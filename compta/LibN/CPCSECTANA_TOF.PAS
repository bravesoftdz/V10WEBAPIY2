{***********UNITE*************************************************
Auteur  ...... : BELLARD Frédéric
Créé le ...... : 10/01/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPCSECTANA_VIE ()
Mots clefs ... : TOF;CPCSECTANA_VIE
*****************************************************************}
Unit CPCSECTANA_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Graphics,
     Forms,
     UTof,
     UTob,
     Sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Menus,           // TMenuItem, PopUpMenu
     LookUp,          // LookUpList
     HQry,            // HQuery
     Htb97,           // TToolBarButton97, TToolWindow97
     HRichOle,        // THRichEditOle
     HSysMenu,
     Filtre,          // LoadFiltre
     Windows,         // VK_
     Grids,           // TGridDrawState
     PARAMSOC,        // GetParamSoc
     AGLInit,         // TheData
     Ent1,            // VH^, TFichierBase
{$IFDEF EAGLCLIENT}
     MainEAgl,        // AGLLanceFiche
{$ELSE}
     Fe_Main,         // AGLLanceFiche
     Db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
{$IFDEF MODENT1}
     CPTypeCons,
{$ENDIF MODENT1}
     uTofViergeMul;   // Fiche Ancetre Vierge MUL

Type
  TOF_CPCSECTANA_VIE = Class (TOF_ViergeMul)

      TabTablesLibres  : TTabSheet;

      procedure OnClose; override ;
      procedure AfterShow ; override;
      procedure OnArgument (S : String ); override ;
      procedure OnLoad ; override ;
      procedure OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState); override ;
      procedure OnClickBBlocNote ( Sender : TObject ); override;
      procedure OnRowEnterFListe (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean); override;
      procedure OnPopUpPopF11(Sender : TObject); override;
    private
      FStSelect          : string;
      SAxe : THValComboBox;
      S_Section          : THEdit;
      S_Section_         : THEdit;
      FStArgumentTOF     : string;
      SModeSelection     : THValComboBox;

      SCorresp1          : THMultiValComboBox;

      S_DateDernMvt      : THEdit;
      S_DateDernMvt_     : THEdit;
      SSensNat           : THValComboBox;
      SSensReel          : THValComboBox;
      SSituationCompte   : THValComboBox;

      PopUpOutil         : TPopUpMenu;
      PopUpTraitement    : TPopUpMenu;
      PopUpUtilitaire    : TPopUpMenu;
      PopUpFctComp       : TPopUpMenu;
      PopUpEdition       : TPopUpMenu;

      BBlocNote          : TToolBarButton97;

      FColDebitE         : integer;
      FColCreditE        : integer;
{b fb 12/06/2006 FQ18301}
      FColDebitS         : integer;
      FColCreditS        : integer;
{e fb 12/06/2006 FQ18301}
      FColSoldeE         : integer;
      FColSoldeP         : integer;
      FColVariation      : integer;
      FColPourcentage    : integer;

      FStPrefixeE : string;
      FStPrefixeP : string;
      FStCodeExoE : string ;
      FStCodeExoP : string ;

      FToutAccorde          : Boolean; // Pour déterminer les droits de modif du USER

      function Axe : TFichierBase;
      function ChoixAxe(index : integer) : string;

      procedure OnElipsisClickS_Section( Sender : TObject );
      procedure OnExitS_Section( Sender : TObject );
      function GetMinMaxCompte(stCol, stDefaut : String) : String;
      procedure ListeYSituationCompte;

      procedure MiseAJourCaptionEcran;
      procedure IndiceColDebCre;
      procedure GetCellCanvasFListe( ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
      function  RecupWhereSituationCompte : string;
      function  RecupAutreWhere : string;

      // Initialisation des PopUp
      procedure InitPopUp ( vActivation : Boolean );

      // POPUPOUTIL
      procedure OnPopUpPopUpOutil(Sender : TObject);
      procedure OnPopUpPopUpTraitement(Sender: TObject);
      procedure OnPopUpPopUpUtilitaire(Sender: TObject);
      procedure OnPopUpPopUpFctComp(Sender: TObject);
      procedure OnPopUpPopUpEdition(Sender: TObject);

      // TTOOLBARBUTTON97
      procedure OnClickDetailSection( Sender : TObject );

      // Traitements Comptables
      procedure SaisieAnaClick(Sender : TObject); // Saisie analytique
      procedure ReVentAnaClick(Sender : TObject); // Re-ventilations analytique
      procedure ReImpAnaClick(Sender: TObject); // Re-imputations analytiques

      // Utilitaires
      procedure OnClickSectAna(Sender: TObject);

      // Fonctions complementaire
      procedure OnClickCommentSect(Sender : TObject);
      procedure CumulsSectClick(Sender : TObject); // Cumul de la section analytique

//      procedure OnClickMillesime(Sender : TObject); // Commentaire millésimé  //fb 19/05/2006 FQ17680
      procedure OnClickCubDecAna(Sender : TObject);
      procedure OnClickAnaStaAna(Sender : TObject);
      procedure OnClickAnaStaAnaGen(Sender : TObject);


      Procedure PrepareArgumentGLG ( vBoJustifSolde : Boolean );
      procedure OnClickGLAna(Sender: TObject);
      procedure LancementGLAnaGen( vFichierBase : TFichierBase );
      procedure OnClickGLAnaGen(Sender: TObject);
      procedure LancementGLGenAna( vFichierBase : TFichierBase );
      procedure OnClickGLGenAna(Sender: TObject);
      procedure PrepareArgumentBAL;
      procedure OnClickBalAna(Sender: TObject);
      procedure LancementBalAnaGen( vFichierBase : TFichierBase );
      procedure OnClickBalAnaGen(Sender: TObject);
      procedure LancementBalGenAna( vFichierBase : TFichierBase );
      procedure OnClickBalGenAna(Sender: TObject);
      function AxeAna(Alias : Boolean) : string;
    protected
      procedure InitControl ; override; // Init des composants de la fiche
      procedure RefreshFListe( vBoFetch : Boolean ) ; override;
      function  BeforeLoad : Boolean ; override;
      procedure RemplitATobFListe ; override;
      function  AjouteATobFListe( vTob : Tob ) : Boolean ; override;
    public
  end ;

function CPLanceFiche_ConsSectAna( vStParam : string = '' ) : string ;

Implementation

uses
  {$IFDEF MODENT1}
    CPProcMetier,
    CPProcGen,
  {$ENDIF MODENT1}
    // *** Fonctionnement de l'unité
    uLibWindows,      // AfficheDBCR  (testjoker)
    SaisComm,         // (GetO)
    SaiSUtil,         // pour le StrSO (RMVT)

     // *** Traitements comptables ***
    SaisODA,
    Reventilana,
    CPREIMPUTANA_TOF,
     // *** Fonctions complémentaires ***
    Cummens,          // CumulSectMensuel
     // *** Editions ***
    uTofCPGLAna,      // CPLanceFiche_CPGL
    CPBALANAL_TOF,    // CPLanceFiche_BalanceAnalytique
    CPBALANALGEN_TOF, // CPLanceFiche_BalanceAnalGen
    CPBALGENANAL_TOF, // CPLanceFiche_BalanceGenAnal
     // *** Utilitaires ***
    CPSECTION_TOM,    // fiche section
    CPGENERAUX_TOM;

const
  cFI_TABLE    = 'CPCSECTANA';
  cStChampsSup = ', S_TOTDEBP, S_TOTCREP, S_AXE, S_SENS SENS';
  cOrder = 'ORDER BY S_SECTION';

function CPLanceFiche_ConsSectAna( vStParam : string = '' ) : string ;
begin
{$IFDEF CCSTD}
  Result := AGLLanceFiche('CP', 'CPCSECTANA_VIE', '', '', vStParam);
{$ELSE}
  if CtxPcl in V_Pgi.PgiContexte then
  begin
    if (VH^.CpExoRef.Code = VH^.Encours.Code)
    or (VH^.CpExoRef.Code = VH^.Suivant.Code) then begin
      Result := AGLLanceFiche('CP', 'CPCSECTANA_VIE', '', '', vStParam);
     end
    else begin
      PgiInfo('Pour pouvoir utiliser cette fonction, l''exercice de référence doit ' +
              'être l''exercice en cours ou le suivant.', 'Consultation des comptes généraux');
      end;
  end else
    Result := AGLLanceFiche('CP', 'CPCSECTANA_VIE', '', '', vStParam);
{$ENDIF}
end;

procedure TOF_CPCSECTANA_VIE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPCSECTANA_VIE.AfterShow;
var
  lStArgument   : string ;
begin
  inherited;
  if FStArgumentTOF <> '' then begin
    lStArgument := FStArgumentTOF;
    S_Section.Text := ReadTokenSt(lStArgument);
    end;

  if S_Section.CanFocus then
    S_Section.SetFocus;
end;

procedure TOF_CPCSECTANA_VIE.InitControl;
begin
  // Init des valeurs
  S_Section.Text         := '';
  S_Section_.Text        := '';
{b fb 12/06/2006 FQ17678}
  SModeSelection.Value   := 'ALL';
    SModeSelection.Items[0]:=TraduireMemoire('Toutes les sections');
  SModeSelection.itemindex:=0;
{e fb 12/06/2006 FQ17678}

  SCorresp1.Value        := '<<' + TraduireMemoire('Tous') + '>>';

  S_DateDernMvt.Text     := DateToStr( iDate1900 );
  S_DateDernMvt_.Text    := DateToStr( iDate2099 );
  SSensNat.Value         := '';
  SSensReel.Value        := '0';
  SsituationCompte.Value := '';
end;

procedure TOF_CPCSECTANA_VIE.OnLoad     ;
begin
  LibellesTableLibre(TabTablesLibres, 'TS_TABLE', 'S_TABLE', 'S');
  Inherited ;
  InitPopUp( False );
  FToutAccorde := ExJaiLeDroitConcept(TConcept(ccSaisEcritures),False);
  InitAutoSearch ;
end ;

procedure TOF_CPCSECTANA_VIE.OnElipsisClickS_Section(Sender: TObject);
begin
  LookUpList( THEdit(Sender),
              'Sections analytiques',
              'SECTION',
              'S_SECTION' ,
              'S_LIBELLE',
              'S_AXE="A' + IntToStr(SAxe.itemindex+1) + '" AND ' + CGenereSQLConfidentiel('S'), //fb 19/05/2006 FQ17802
              'S_SECTION' ,
              True,
              0 );  //fb 19/05/2006 FQ17901
end;

function TOF_CPCSECTANA_VIE.Axe : TFichierBase;
begin
  Result := fbAxe1;
  Case SAxe.ItemIndex of
    0 : Result := fbAxe1;
    1 : Result := fbAxe2;
    2 : Result := fbAxe3;
    3 : Result := fbAxe4;
    4 : Result := fbAxe5;
    end;
end;

procedure TOF_CPCSECTANA_VIE.OnExitS_Section(Sender: TObject);
begin
  if (Trim(THEdit(Sender).Text) = '') or (TestJoker(THEdit(Sender).Text)) then
     Exit;

  if Length(THEdit(Sender).Text) < VH^.Cpta[Axe].Lg then
    THEdit(Sender).Text := BourreEtLess( THEdit(Sender).Text, Axe);
end;

procedure TOF_CPCSECTANA_VIE.ListeYSituationCompte;
begin
  SSituationCompte.Items[1] := 'Mouvementés sur l''exercice N mais sans mouvements sur N-1';
  SSituationCompte.Items[2] := 'Sans mouvement sur l''exercice N mais mouvementés sur N-1';
  SSituationCompte.Items[3] := 'Non soldés sur l''exercice N mais soldés sur N-1';
  SSituationCompte.Items[4] := 'Soldés sur l''exercice N mais non soldés sur N-1';
  SSituationCompte.Items[5] := 'N''ayant pas le même sens sur l''exercice N que sur N-1';
end;

procedure TOF_CPCSECTANA_VIE.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
 inherited;
  case Key of
    VK_F5 : if FListe.Focused then begin
              Key := 0;
              OnClickDetailSection(nil);
              end ;

{b fb 19/05/2006 FQ17680}
(*    // ALT + O  -> Commnentaire millésimé
    79 : if ssAlt in Shift then begin
           Key := 0;
           OnClickMillesime(nil);
           end ;*)
{e fb 19/05/2006 FQ17680}
    end;
end;

procedure TOF_CPCSECTANA_VIE.OnClickDetailSection(Sender: TObject);
var
  lTob : Tob;
begin
  lTob := GetO(FListe) ;
  if lTob = nil then
    Exit;
  AGLLanceFiche('CP', 'CPCECRANA_VIE','' , '', lTob.GetValue('S_SECTION')+';'+lTob.GetValue('S_AXE'));
end;

procedure TOF_CPCSECTANA_VIE.SaisieAnaClick(Sender: TObject);
{$IFDEF COMPTA}
var
  M : RMVT ;
{$ENDIF}
begin
{$IFDEF COMPTA}
  FillChar(M,Sizeof(M),#0) ;
  M.Simul := 'N' ;
  M.CodeD := V_PGI.DevisePivot ;
  M.DateC := V_PGI.DateEntree ;
  M.TauxD := 1 ;
  M.DateTaux := M.DateC ;
  M.Valide := False ;
  M.Etabl := VH^.ETABLISDEFAUT ;

  LanceSaisieODA(Nil,taCreat,M) ;
  BCherche.Click;
{$ENDIF}
end;

procedure TOF_CPCSECTANA_VIE.ReVentAnaClick(Sender: TObject);
var
  lTob : Tob;
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if lTob = nil then
    Exit;
  AGLLanceFiche('CP', 'CPREVENTILANA', '', '', ''+';'+
                                       ''+';'+
                                       'N;'+';'+
                                       ''+';'+
                                       ''+';'+''+';'+
                                       SAxe.Text);
  BCherche.Click;
{$ENDIF}
end;

procedure TOF_CPCSECTANA_VIE.ReImpAnaClick(Sender: TObject);
var
  lTob : Tob;
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if lTob = nil then
    Exit;
  AGLLanceFiche('CP', 'CPREIMPUTANA', '', '',SAxe.Text+';'+
                                       ''+';'+
                                       ''+';'+
                                       lTob.GetValue('S_SECTION')+';'+
                                       ''+';'+
                                       ''+';'+
                                       ''+';'+
                                       '');
{$ENDIF}
end;

procedure TOF_CPCSECTANA_VIE.OnClickSectAna(Sender: TObject);
var
  lTob : Tob;
  lAction : TActionFiche ;
begin
  lTob := GetO(FListe) ;
  if lTob = nil then
    Exit;
  lAction := TaModif;
  FicheSection(nil, ChoixAxe(SAxe.ItemIndex),
                    lTob.GetValue('S_SECTION'), lAction, 0);
  BCherche.Click;
end;

procedure TOF_CPCSECTANA_VIE.OnClickCommentSect(Sender: TObject);
var lTob : Tob;
    lAction : TActionFiche;
begin
  lTob := GetO(FListe) ;
  if lTob = nil then
    Exit;
  lAction := TaModif;
  FicheSection( nil, ChoixAxe(SAxe.ItemIndex),
                     lTob.GetValue('S_SECTION'), lAction, 2);
end;

procedure TOF_CPCSECTANA_VIE.CumulsSectClick(Sender: TObject);
var lTob : TOB ;
begin
  lTob := GetO(FListe) ;
  if lTob = nil then
    Exit;
  CumulCpteMensuel(Axe, lTob.GetValue('S_SECTION'), lTob.GetValue('S_LIBELLE'), VH^.Encours);
end;

{b fb 19/05/2006 FQ17680}
(*procedure TOF_CPCSECTANA_VIE.OnClickMillesime(Sender: TObject);
var
  lTob : Tob;
  lAction : TActionFiche;
begin
  lTob := GetO(FListe) ;
  if (lTob = nil) or (VH^.CPExoRef.Code <> VH^.Encours.Code) then
    Exit;
  lAction := TaModif;
  FicheSection( nil, ChoixAxe(SAxe.ItemIndex), lTob.GetValue('S_SECTION'), lAction, 0);
end; *)
{e fb 19/05/2006 FQ17680}

procedure TOF_CPCSECTANA_VIE.OnClickCubDecAna(Sender: Tobject);
begin
  AGLLanceFiche('CP','CPANALYTIQ_CUBE','','','') ;
end;

procedure TOF_CPCSECTANA_VIE.OnClickAnaStaAna(Sender: Tobject);
begin
  AGLLanceFiche('CP','CPANALYTIQ_TOBV','','','') ;
end;

procedure TOF_CPCSECTANA_VIE.OnClickAnaStaAnaGen(Sender: Tobject);
begin
  AGLLanceFiche('CP','CPANAGENE_TOBV','','','') ;
end;

Procedure TOF_CPCSECTANA_VIE.PrepareArgumentGLG ( vBoJustifSolde : Boolean );
begin
  if CtxPcl in V_PGI.PGIContexte then begin
    // Exercice
    ACritEdt.CritEdt.Exo.Code := VH^.CpExoRef.Code;
    // Début de la Date Comptable
    ACritEdt.CritEdt.Date1 := VH^.CPExoRef.Deb;
    // Fin de la Date Comptable
    ACritEdt.CritEdt.Date2 := VH^.CPExoRef.Fin;
   end
  else begin
    // Exercice
    ACritEdt.CritEdt.Exo.Code := VH^.EnCours.Code;
    // Début de la Date Comptable
    ACritEdt.CritEdt.Date1 := VH^.Encours.Deb;
    // Fin de la Date Comptable
    ACritEdt.CritEdt.Date2 := VH^.Encours.Fin;
    end;

  // Justificatif de Solde
  ACritEdt.CritEdt.GL.EnDateSituation := (vBoJustifSolde = True);

  // Type d'Ecritures
  ACritEdt.CritEdt.Qualifpiece := 'N;';

  // Ecritures Valides
  ACritEdt.CritEdt.Valide := '';
end;

procedure TOF_CPCSECTANA_VIE.OnClickGLAna(Sender: TObject);
{$IFDEF COMPTA}
var lTob : TOB ;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if lTob = nil then
    Exit;

  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentGLG( False );
  ACritEdt.CritEdt.Cpt1   := lTob.GetValue('S_SECTION');
  ACritEdt.CritEdt.Cpt2   := lTob.GetValue('S_SECTION');

  TheData := ACritEdt;
  CPLanceFiche_CPGLANA;
  TheData := nil;
{$ENDIF}
end;

procedure TOF_CPCSECTANA_VIE.LancementGLAnaGen( vFichierBase : TFichierBase );
{$IFDEF COMPTA}
var
  lTob : Tob ;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if lTob = nil then
    Exit;

  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentGLG( False );
  ACritEdt.CritEdt.Cpt1   := lTob.GetValue('S_SECTION');
  ACritEdt.CritEdt.Cpt2   := lTob.GetValue('S_SECTION');
  ACritEdt.CritEdt.GL.Axe := fbToAxe( vFichierBase );

  TheData := ACritEdt;
  CPLanceFiche_CPGLANAPARGENE;
  TheData := nil;
{$ENDIF}
end;

procedure TOF_CPCSECTANA_VIE.OnClickGLAnaGen(Sender: TObject);
begin
  LancementGLAnaGen(Axe);
end;

procedure TOF_CPCSECTANA_VIE.LancementGLGenAna( vFichierBase : TFichierBase );
{$IFDEF COMPTA}
var
  lTob : Tob ;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if lTob = nil then
    Exit;

  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentGLG( False );
  ACritEdt.CritEdt.sCpt1   := lTob.GetValue('S_SECTION');
  ACritEdt.CritEdt.sCpt2   := lTob.GetValue('S_SECTION');
  ACritEdt.CritEdt.GL.Axe := fbToAxe( vFichierBase );

  TheData := ACritEdt;
  CPLanceFiche_CPGLGENEPARANA;
  TheData := nil;
{$ENDIF}
end;

procedure TOF_CPCSECTANA_VIE.OnClickGLGenAna(Sender: TObject);
begin
  LancementGLGenAna(Axe);
end;

procedure TOF_CPCSECTANA_VIE.PrepareArgumentBAL;
begin
  if CtxPcl in V_PGI.PGIContexte then begin
    // Exercice
    ACritEdt.CritEdt.Exo.Code := VH^.CpExoRef.Code;
    // Début de la Date Comptable
    ACritEdt.CritEdt.Date1 := VH^.CPExoRef.Deb;
    // Fin de la Date Comptable
    ACritEdt.CritEdt.Date2 := VH^.CPExoRef.Fin;
   end
  else begin
    // Exercice
    ACritEdt.CritEdt.Exo.Code := VH^.EnCours.Code;
    // Début de la Date Comptable
    ACritEdt.CritEdt.Date1 := VH^.Encours.Deb;
    // Fin de la Date Comptable
    ACritEdt.CritEdt.Date2 := VH^.Encours.Fin;
    end;

  // Type d'Ecritures
  ACritEdt.CritEdt.Qualifpiece := 'N;';

  // Options
  ACritEdt.CritEdt.Bal.SansAnoANA := true;
  ACritEdt.CritEdt.AvecComparatif := false;
end;

procedure TOF_CPCSECTANA_VIE.OnClickBalAna(Sender: TObject);
{$IFDEF COMPTA}
var
  lTob  : Tob ;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if lTob = nil then
    Exit;

  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentBAL;
  ACritEdt.CritEdt.Cpt1 := lTob.GetValue('S_SECTION');
  ACritEdt.CritEdt.Cpt2 := lTob.GetValue('S_SECTION');

  ACritEdt.CritEdt.Bal.SansAnoANA := false;

  TheData := ACritEdt;
  CPLanceFiche_BalanceAnalytique;
  TheData := nil;
{$ENDIF}
end;

procedure TOF_CPCSECTANA_VIE.LancementBalAnaGen( vFichierBase : TFichierBase );
{$IFDEF COMPTA}
var
  lTob  : Tob ;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if lTob = nil then
    Exit;

  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentBAL;
  ACritEdt.CritEdt.Cpt1    := lTob.GetValue('S_SECTION');
  ACritEdt.CritEdt.Cpt2    := lTob.GetValue('S_SECTION');
  ACritEdt.CritEdt.BAL.Axe := fbToAxe( vFichierBase );

  TheData := ACritEdt;
  CPLanceFiche_BalanceAnalGen;
  TheData := nil;
{$ENDIF}
end;

procedure TOF_CPCSECTANA_VIE.OnClickBalAnaGen( Sender : TObject );
begin
  LancementBalAnaGen(Axe);
end;

procedure TOF_CPCSECTANA_VIE.LancementBalGenAna( vFichierBase : TFichierBase );
{$IFDEF COMPTA}
var
  lTob  : Tob ;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if lTob = nil then
    Exit;

  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentBAL;
  ACritEdt.CritEdt.sCpt1    := lTob.GetValue('S_SECTION');
  ACritEdt.CritEdt.sCpt2    := lTob.GetValue('S_SECTION');
  ACritEdt.CritEdt.BAL.Axe := fbToAxe( vFichierBase );

  TheData := ACritEdt;
  CPLanceFiche_BalanceGenAnal;
  TheData := nil;
{$ENDIF}
end;

procedure TOF_CPCSECTANA_VIE.OnClickBalGenAna( Sender : TObject );
begin
  LancementBalGenAna(Axe);
end;

procedure TOF_CPCSECTANA_VIE.OnPopUpPopF11(Sender: TObject);
begin
  PurgePopup( PopF11 );
  AddMenuPop( PopF11, '', '');
  InitPopUp( True );
  inherited;
end;

procedure TOF_CPCSECTANA_VIE.OnPopUpPopUpOutil(Sender: TObject);
begin
 InitPopUp( True );
end;

procedure TOF_CPCSECTANA_VIE.OnPopUpPopUpTraitement(Sender: TObject);
begin
 InitPopUp(True);
end;

procedure TOF_CPCSECTANA_VIE.OnPopUpPopUpUtilitaire(Sender: TObject);
begin
 InitPopUp(True);
end;

procedure TOF_CPCSECTANA_VIE.OnPopUpPopUpFctComp(Sender: TObject);
begin
  InitPopUp(True);
end;

procedure TOF_CPCSECTANA_VIE.OnPopUpPopUpEdition(Sender: TObject);
begin
  InitPopUp(True);
end;

procedure TOF_CPCSECTANA_VIE.InitPopUp ( vActivation : Boolean ) ;
var
  i : integer;
  lTob : Tob;
begin
  lTob := GetO(FListe) ;

  // POPUPOUTIL
  for i := 0 to PopUpOutil.Items.Count -1 do begin
    if PopUpOutil.Items[i].Name = 'DETAILSECT' then begin
      if vActivation then
        PopUpOutil.Items[i].Enabled := (lTob <> nil)
      else
        PopUpOutil.Items[i].OnClick := OnClickDetailSection;
      Continue;
      end;
    end; // for

  // POPUPTRAITEMENT
  for i := 0 to PopUpTraitement.Items.Count -1 do begin
    if PopUpTraitement.Items[i].Name = 'SAISANA' then begin
      if vActivation then
        PopUpTraitement.Items[i].Enabled := ExJaiLeDroitConcept(TConcept(ccSaisEcritures),False)
      else
        PopUpTraitement.Items[i].OnClick := SaisieAnaClick;
      Continue;
      end;

      if PopUpTraitement.Items[i].Name = 'REVENTANA' then begin
        if vActivation then
          PopUpTraitement.Items[i].Enabled := ExJaiLeDroitConcept(TConcept(ccSaisEcritures),False)
        else
          PopUpTraitement.Items[i].OnClick := ReVentAnaClick;
        Continue;
        end;

      if PopUpTraitement.Items[i].Name = 'REIMPANA' then begin
        if vActivation then
          PopUpTraitement.Items[i].Enabled := ExJaiLeDroitConcept(TConcept(ccSaisEcritures),False)
        else
          PopUpTraitement.Items[i].OnClick := ReImpAnaClick;
        Continue;
        end;
    end;

  // POPUP UTILITAIRE
  for i := 0 to PopUpUtilitaire.Items.Count -1 do begin
    if PopUpUtilitaire.Items[i].Name = 'PARSECTANA' then begin
      if vActivation then
        PopUpUtilitaire.Items[i].Enabled := (lTob <> nil)
      else
        PopUpUtilitaire.Items[i].OnClick := OnClickSectAna;
      Continue;
      end;
    end;

    // POPUP FONCTIONCOMPLEMENTAIRE
  for i := 0 to PopUpFctComp.Items.Count -1 do begin
    if PopUpFctComp.Items[i].Name = 'COMMENTSECT' then begin
      if vActivation then
        PopUpFctComp.Items[i].Enabled := (lTob <> nil)
      else
        PopUpFctComp.Items[i].OnClick := OnClickCommentSect;
      Continue;
      end;
    if PopUpFctComp.Items[i].Name = 'CUMULSECT' then begin
      if vActivation then
        PopUpFctComp.Items[i].Enabled := (lTob <> nil)
      else
        PopUpFctComp.Items[i].OnClick := CumulsSectClick;
      Continue;
      end;
{b fb 19/05/2006 FQ17680}
(*    if PopUpFctComp.Items[i].Name = 'MILLESIME' then begin
      if vActivation then
        PopUpFctComp.Items[i].Enabled := (lTob <> nil) and
                                         (VH^.CPExoRef.Code = VH^.Encours.Code)
      else
        PopUpFctComp.Items[i].OnClick := OnClickMillesime;
      Continue;
      end;*)
{e fb 19/05/2006 FQ17680}
    if PopUpFctComp.Items[i].Name = 'CUBDECANA' then begin
      if vActivation then
        PopUpFctComp.Items[i].Enabled := (lTob <> nil)
      else
        PopUpFctComp.Items[i].OnClick := OnClickCubDecAna;
      Continue;
      end;
    if PopUpFctComp.Items[i].Name = 'ANASTATANA' then begin
      if vActivation then
        PopUpFctComp.Items[i].Enabled := (lTob <> nil)
      else
        PopUpFctComp.Items[i].OnClick := OnClickAnaStaAna;
      Continue;
      end;
    if PopUpFctComp.Items[i].Name = 'ANASTATANAGEN' then begin
      if vActivation then
        PopUpFctComp.Items[i].Enabled := (lTob <> nil)
      else
        PopUpFctComp.Items[i].OnClick := OnClickAnaStaAnaGen;
      Continue;
      end;
    end;

  // POPUP EDITION
  for i := 0 to PopUpEdition.Items.Count - 1 do begin
    // Grand-livre analytique
    if PopUpEdition.Items[i].Name = 'GLANA' then begin
      PopUpEdition.Items[i].OnClick := OnClickGLAna;
      {$IFDEF COMPTA}
      if vActivation then
        PopUpEdition.Items[i].Enabled := (ATobFListe.Detail.Count > 0);
      {$ELSE}
        PopUpEdition.Items[i].Visible := False;
      {$ENDIF}
      Continue;
      end;

    // Grand-livre analytique par général
    if PopUpEdition.Items[i].Name = 'GLANAGEN' then begin
      PopUpEdition.Items[i].OnClick := OnClickGLAnaGen;
      {$IFDEF COMPTA}
      if vActivation then
        PopUpEdition.Items[i].Enabled := (ATobFListe.Detail.Count > 0);
      {$ELSE}
        PopUpEdition.Items[i].Visible := False;
      {$ENDIF}
      Continue;
      end;

    // Grand-livre général par analytique
    if PopUpEdition.Items[i].Name = 'GLGENANA' then begin
      PopUpEdition.Items[i].OnClick := OnClickGLGenAna;
      {$IFDEF COMPTA}
      if vActivation then
        PopUpEdition.Items[i].Enabled := (ATobFListe.Detail.Count > 0);
      {$ELSE}
        PopUpEdition.Items[i].Visible := False;
      {$ENDIF}
      Continue;
      end;

    // Balance analytique
    if PopUpEdition.Items[i].Name = 'BALANA' then begin
      PopUpEdition.Items[i].OnClick := OnClickBalAna;
      {$IFDEF COMPTA}
      if vActivation then
        PopUpEdition.Items[i].Enabled := (ATobFListe.Detail.Count > 0);
      {$ELSE}
        PopUpEdition.Items[i].Visible := False;
      {$ENDIF}
      Continue;
      end;

    // Balance analytique par général sur l'axe sélectionné
    if PopUpEdition.Items[i].Name = 'BALANAGEN' then begin
      PopUpEdition.Items[i].OnClick := OnClickBalAnaGen;
      {$IFDEF COMPTA}
      if vActivation then
        PopUpEdition.Items[i].Enabled := (ATobFListe.Detail.Count > 0);
      {$ELSE}
        PopUpEdition.Items[i].Visible := False;
      {$ENDIF}
      Continue;
      end;

    // Balance générale par analytique sur l'Axe sélectionné
    if PopUpEdition.Items[i].Name = 'BALGENANA' then begin
      PopUpEdition.Items[i].OnClick := OnClickBalGenAna;
      {$IFDEF COMPTA}
      if vActivation then
        PopUpEdition.Items[i].Enabled := (ATobFListe.Detail.Count > 0);
      {$ELSE}
        PopUpEdition.Items[i].Visible := False;
      {$ENDIF}
      Continue;
      end;
    end;
end;

procedure TOF_CPCSECTANA_VIE.OnClickBBlocNote(Sender: TObject);
var
  lTobAna : Tob;
begin
  inherited;
  HPB.Visible := BBlocNote.Down;
  if BBlocNote.Down then begin
    FBLocNote.Clear;
    lTobAna := GetO(FListe) ;
    if lTobAna = nil then
      Exit;
    StringToRich( FBlocNote, CPChargeBlocNoteMillesime('S', lTobAna.GetValue('S_SECTION'), VH^.Encours.Code));
    end;
end;

procedure TOF_CPCSECTANA_VIE.OnRowEnterFListe(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var
  lTobGene : Tob;
begin
  inherited;
  lTobGene := GetO(FListe) ;
  if lTobGene = nil then
    Exit;
  OnClickBBlocNote(nil);
end;

procedure TOF_CPCSECTANA_VIE.MiseAJourCaptionEcran;
begin
  Ecran.Caption := TraduireMemoire('Consultation des sections analytiques') + ' : ' + RechDom('TTEXERCICE', VH^.CPExoRef.Code, False);
  UpDateCaption(Ecran);
end;

procedure TOF_CPCSECTANA_VIE.RefreshFListe;
begin
  inherited;
  MiseAJourCaptionEcran;
end;

procedure TOF_CPCSECTANA_VIE.IndiceColDebCre ;
var
 lStChamps : string;
 lStListe  : string;
 lInIndex  : integer;
begin
  lStListe        := FStListeChamps;
  lInIndex        := 1;
  FColDebitE      := -1;
  FColCreditE     := -1;
  FColDebitS      := -1;
  FColCreditS     := -1;
  FColSoldeP      := -1;
  FColSoldeE      := -1;
  FColVariation   := -1;
  FColPourcentage := -1;

  while lStListe <> '' do begin
   lStChamps := READTOKENST(lStListe);
   if ( lStChamps = 'S_TOTDEBE' ) then
     FColDebitE := lInIndex
   else
     if ( lStChamps = 'S_TOTCREE' ) then
       FColCreditE := lInIndex
     else
       if ( lStChamps = 'S_TOTDEBS' ) then
         FColDebitS := lInIndex
       else
         if ( lStChamps = 'S_TOTCRES' ) then
           FColCreditS := lInIndex
         else
           if ( lStChamps = 'SOLDEP' ) then
             FColSoldeP := lInIndex
           else
             if ( lStChamps = 'SOLDEE' ) then
               FColSoldeE := lInIndex
             else
               if ( lStChamps = 'VARIATION' ) then
                 FColVariation := lInIndex
               else
                 if ( lStChamps = 'POURCENTAGE' ) then
                 FColPourcentage := lInIndex;

    Inc(lInIndex);
    end; // while
end;

procedure TOF_CPCSECTANA_VIE.GetCellCanvasFListe(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var
  lTob : Tob;
begin
  inherited;
  if ARow = 0 then
    Exit;
  if (FListe.Row <> ARow)then begin
    lTob := GetO(FListe, ARow) ;
    if lTob <> nil then begin
      if (ACol = FColSoldeP) or (ACol = FColSoldeE) then begin
        if lTob.GetValue('SENS') = 'D' then
          Canvas.Font.Color := IIF(Pos('C', FListe.Cells[ACol, ARow]) > 0, ClRed, ClGreen)
        else
          if lTob.GetValue('SENS') = 'C' then
            Canvas.Font.Color := IIF(Pos('C', FListe.Cells[ACol, ARow]) > 0, ClGreen, ClRed);
        end;
      end;
    end;
end;

function TOF_CPCSECTANA_VIE.AxeAna(Alias : Boolean) : string;
var
  Sql, Sql2 : string;
begin
  Sql:='';
  Sql2:='';
  if VH^.AnaCroisaxe then begin
    if GetParamSocSecur('SO_VENTILA'+IntToStr(SAxe.ItemIndex+1), false) then begin
      Sql := 'Y_SOUSPLAN'+IntToStr(SAxe.ItemIndex+1);
      if Alias then
        Sql2 := ' AS Y_SECTION';
      end;
   end
  else
    Sql := 'Y_SECTION';

  Result := Sql + Sql2;
end;

function TOF_CPCSECTANA_VIE.RecupWhereSituationCompte: string;
begin
  if SSituationCompte.Value = 'SI1' then begin // Comptes mouvementés sur l'exercice mais sans mouvements sur le précédent
    Result := ' AND S_SECTION IN (SELECT DISTINCT (Y1.' + AxeAna(false) + ') FROM ANALYTIQ Y1 WHERE' +
              ' Y1.Y_EXERCICE = "' + FStCodeExoE + '" AND' +
              ' Y1.Y_ECRANOUVEAU = "N" AND' +
              ' Y1.Y_QUALIFPIECE = "N" AND Y1.' + AxeAna(false) + ' NOT IN' +
              ' (SELECT Y2.' + AxeAna(false) + ' FROM ANALYTIQ Y2 WHERE' +
              ' Y2.Y_EXERCICE = "' + FStCodeExoP + '" AND' +
              ' Y2.Y_ECRANOUVEAU = "N" AND' +
              ' Y2.Y_QUALIFPIECE = "N" ))';
   end
  else if SSituationCompte.Value = 'SI2' then begin // Comptes sans mouvements sur l'exercice mais mouvementés sur le précédent
    Result := ' AND S_SECTION IN (SELECT DISTINCT (Y1.' + AxeAna(false) + ') FROM ANALYTIQ Y1 WHERE' +
              ' Y1.Y_EXERCICE = "' + FStCodeExoP + '" AND' +
              ' Y1.Y_ECRANOUVEAU = "N" AND' +
              ' Y1.Y_QUALIFPIECE = "N" AND Y1.' + AxeAna(false) + ' NOT IN' +
              ' (SELECT Y2.' + AxeAna(false) + ' FROM ANALYTIQ Y2 WHERE' +
              ' Y2.Y_EXERCICE = "' + FStCodeExoE + '" AND' +
              ' Y2.Y_ECRANOUVEAU = "N" AND' +
              ' Y2.Y_QUALIFPIECE = "N" ))';
   end
  else if SSituationCompte.Value = 'SI3' then begin // Comptes non soldés sur l'exercice mais soldés sur le précédent
    Result := ' AND (S_TOTDEB' + FStPrefixeE + ' <> S_TOTCRE' + FStPrefixeE + ')' +
              ' AND (S_TOTDEB' + FStPrefixeP + ' = S_TOTCRE' + FStPrefixeP + ')';
   end
  else if SSituationCompte.Value = 'SI4' then begin // Comptes soldés sur l'exercice mais son soldés sur le précédent
    Result := ' AND (S_TOTDEB' + FStPrefixeE + ' = S_TOTCRE' + FStPrefixeE + ')' +
              ' AND (S_TOTDEB' + FStPrefixeP + ' <> S_TOTCRE' + FStPrefixeP + ')';
   end
  else if SSituationCompte.Value = 'SI5' then begin //Comptes n'ayant pas le même sens sur l'exercice que sur le précédent
    Result := ' AND (((S_TOTDEB' + FStPrefixeE + ' - S_TOTCRE' + FStPrefixeE + ' > 0) AND ' +
                    '(S_TOTDEB' + FStPrefixeP + ' - S_TOTCRE' + FstPrefixeP + ' < 0)) OR ' +
                    '((S_TOTDEB' + FStPrefixeE + ' - S_TOTCRE' + FStPrefixeE + ' < 0) AND ' +
                    '(S_TOTDEB' + FStPrefixeP + ' - S_TOTCRE' + FstPrefixeP + ' > 0)))';
  end;

end;

function TOF_CPCSECTANA_VIE.RecupAutreWhere : string;
var
  lStValCorresp1 : string;
  lStLibCorresp1 : string;
begin
  Result := '';

  //Axe
  Result := ' AND (S_AXE = "' + ChoixAxe(SAxe.ItemIndex) + '")';

  // Mode de sélection des comptes
  // le Cas où Value = 'EXS' est géré à la création du texte SQL
  if SModeSelection.Value = 'EXT' then begin // Comptes mouvementés sur les deux exercices (avec ANO)
    Result := Result + ' AND ((S_TOTDEB' + FStPrefixeE + ' <> 0 OR S_TOTCRE' + FStPrefixeE + ' <> 0) OR' +
                       ' (S_TOTDEB' + FStPrefixeP + ' <> 0 OR S_TOTCRE' + FStPrefixeP + ' <> 0))';
   end
  else if SModeSelection.Value = 'NSL' then begin // Comptes non soldés sur les deux exercices
      Result := Result + ' AND ((S_TOTDEB' + FStPrefixeE + ' <> S_TOTCRE' + FStPrefixeE + ') OR' +
                         ' (S_TOTDEB' + FStPrefixeP + ' <> S_TOTCRE' + FStPrefixeP + '))';
    end;


  // Sens Réel du solde
  if SSensReel.Value = '1' then
    Result := Result + ' AND (S_TOTDEBE > S_TOTCREE)'
  else
    if SSensReel.Value = '2' then
      Result := Result + ' AND (S_TOTDEBE < S_TOTCREE)'
    else
      if SSensReel.Value = '3' then
        Result := Result + ' AND (S_TOTDEBE = S_TOTCREE)';

  // Gestion du S_CONFIDENTIEL
  if (V_Pgi.Confidentiel = '0') then begin
    Result := Result + ' AND ((S_CONFIDENTIEL = "0") OR (S_CONFIDENTIEL = "-")) ';
   end
  else begin
    Result := Result + ' AND (((S_CONFIDENTIEL = "-") OR (S_CONFIDENTIEL = "X")) OR ' +
                       ' (S_CONFIDENTIEL <= "' + V_PGI.Confidentiel + '"))';
    end;

  // Compte de Correspondance
  TraductionTHMultiValComboBox(SCorresp1, lStValCorresp1, lStLibCorresp1, 'S_CORRESP1', True );
  if lStLibCorresp1 <> '' then
    Result := Result + ' AND ' + lStValCorresp1;

end;

procedure TOF_CPCSECTANA_VIE.RemplitATobFListe;
var
  lStSelect   : string ;
  lStWhere    : string ;
  lSt         : string;
begin
  if (VH^.CpExoRef.Code = VH^.Encours.Code)
  or (VH^.CpExoRef.Code = VH^.Suivant.Code) then begin
    lSt := RecupWhereCritere(PageControl);
    if lSt <> '' then
      lStWhere := lSt + RecupWhereSituationCompte + RecupAutreWhere
    else

      lStWhere := ' WHERE ' + RecupWhereSituationCompte + RecupAutreWhere;

    // Cas particulier pour la requête générée, on recherche les comptes
    // mouvementés sur les deux exercices sans les ANO
    if SModeSelection.Value = 'EXS' then begin
      lStSelect := ' SELECT DISTINCT(Y.' + AxeAna(false) + '),' + CSqlTextFromList(FStListeChamps) + ' ' +
                   cStChampsSup + ' FROM ANALYTIQ Y ' +
                   ' LEFT JOIN SECTION ON S_SECTION = Y.' + AxeAna(false) + ' '+
                   ' LEFT JOIN JOURNAL ON J_JOURNAL = Y_JOURNAL ' +
                   lStWhere + ' AND Y_QUALIFPIECE = "N"' +
                   ' AND (Y_EXERCICE = "' + FStCodeExoE + '" OR Y_EXERCICE = "' + FStCodeExoP + '")' +
                   ' AND J_NATUREJAL <> "ANO" '  + cOrder;
     end
    else begin
      lStSelect := 'SELECT ' + CSqlTextFromList(FStListeChamps) + ' ' + cStChampsSup +
                 ' FROM SECTION ' + lStWhere +' ' + cOrder;
      end;

    // On envoie le texte SQL de la requête à la fiche ANCETRE pour qu'elle fasse le OPENSQL
    AStSqlTobFListe := lStSelect;
   end
  else begin
    PgiInfo('Pour pouvoir utiliser cette fonction, l''exercice de référence doit ' +
             'être l''exercice en cours ou le suivant.', 'Consultation des comptes généraux');

    AStSqlTobFListe := '';
    end;
end;

function TOF_CPCSECTANA_VIE.AjouteATobFListe(vTob : Tob) : Boolean;
var
  lRdSoldeE       : double ;
  lRdSoldeP       : double ;
  lTempSoldeE     : double ;
  lTempSoldeP     : double ;
  lNbMoisSoldeE   : word   ;
  lNbMoisSoldeP   : word   ;
  lPremMois       : word   ;
  lPremAnnee      : word   ;
  lVariation      : double ;
  lPourcentageVariation : double ;
  lStPourcentageVariation : string ;
begin
  Result:=true;
  if VH^.CPExoRef.Code = VH^.Suivant.Code then begin
    // Colonne de S_TOTDEBE et S_TOTCREE sont cachées
    if FColDebitE  <> -1 then FListe.ColWidths[FColDebitE]  := -1;
    if FColCreditE <> -1 then FListe.ColWidths[FColCreditE] := -1;

    // Nombre de mois des exercices pour le calcul de la variation
    NombreMois( VH^.Suivant.Deb, VH^.Suivant.Fin, lPremMois, lPremAnnee, lNbMoisSoldeE);
    NombreMois( VH^.Encours.Deb, VH^.Encours.Fin, lPremMois, lPremAnnee, lNbMoisSoldeP);

{b fb 12/06/2006 FQ18301}
    lRdSoldeE := vTob.GetValue('S_TOTDEBS') - vTob.GetValue('S_TOTCRES');
    lRdSoldeP := vTob.GetValue('S_TOTDEBE') - vTob.GetValue('S_TOTCREE');
   end
  else begin
    // Colonne de S_TOTDEBS et S_TOTCRES sont cachées
    if FColDebitS  <> -1 then FListe.ColWidths[FColDebitS]  := -1;
    if FColCreditS <> -1 then FListe.ColWidths[FColCreditS] := -1;
{e fb 12/06/2006 FQ18301}

    // Nombre de mois des exercices pour le calcul de la variation
    NombreMois( VH^.Encours.Deb, VH^.Encours.Fin, lPremMois, lPremAnnee, lNbMoisSoldeE);
    NombreMois( VH^.Precedent.Deb, VH^.Precedent.Fin, lPremMois, lPremAnnee, lNbMoisSoldeP);

    lRdSoldeE := vTob.GetValue('S_TOTDEBE') - vTob.GetValue('S_TOTCREE');
    lRdSoldeP := vTob.GetValue('S_TOTDEBP') - vTob.GetValue('S_TOTCREP');
    end;

  // Calcul de la Variation
{b fb 12/06/2006 FQ18301}
  lTempSoldeE := Abs(lRdSoldeE);
  lTempSoldeP := Abs(lRdSoldeP);

  if (vTob.GetString('SENS') = 'D') or (vTob.GetString('SENS') = 'M') then
  begin
    if (lRdSoldeE < 0) then lTempSoldeE := -1 * lTempSoldeE;
    if (lRdSoldeP < 0) then lTempSoldeP := -1 * lTempSoldeP;
  end
  else
  begin
    if (lRdSoldeE > 0) then lTempSoldeE := -1 * lTempSoldeE;
    if (lRdSoldeP > 0) then lTempSoldeP := -1 * lTempSoldeP;
  end;
{e fb 12/06/2006 FQ18301}

  lVariation := Abs((lTempSoldeE / lNbMoisSoldeE) * 12) - Abs((lTempSoldeP / lNbMoisSoldeP) * 12);

  // Calcul du Pourcentage
  if (lRdSoldeP = 0) and (lRdSoldeE <> 0) then // Division par 0
    lStPourcentageVariation := 'NS'
  else begin
    if (lRdSoldeP = 0) and (lRdSoldeE = 0) then
      lPourcentageVariation := 0
    else begin
      lPourcentageVariation := ((lRdSoldeE - lRdSoldeP) / lRdSoldeP) * 100;
      lPourcentageVariation := Arrondi(lPourcentageVariation, V_PGI.OkDecV);
      end;

    if (lPourcentageVariation > 300) or (lPourcentageVariation < -300) then
      lStPourcentageVariation := 'NS'
    else
      lStPourcentageVariation := StrFMontant( lPourcentageVariation, 15, V_PGI.OkDecV, '', True) + ' %';
    end;

  vTob.AddChampSupValeur('SOLDEE', lRdSoldeE, False);
  vTob.AddChampSupValeur('SOLDEP', lRdSoldeP, False);
  vTob.AddChampSupValeur('VARIATION',   lVariation,   False);
  vTob.AddChampSupValeur('POURCENTAGE', lStPourcentageVariation, False);
end;

function TOF_CPCSECTANA_VIE.BeforeLoad : Boolean;
begin
  FStPrefixeE := IIF(VH^.CPExoref.Code = VH^.Encours.Code, 'E', 'S');
  FStPrefixeP := IIF(VH^.CPExoref.Code = VH^.Encours.Code, 'P', 'E');
  FStCodeExoE := IIF( VH^.CpExoRef.Code = VH^.Encours.Code, VH^.Encours.Code, VH^.Suivant.Code);
  FStCodeExoP := IIF( VH^.CpExoRef.Code = VH^.Encours.Code, VH^.Precedent.Code, VH^.Encours.Code);

  Result := True;
end;

function TOF_CPCSECTANA_VIE.ChoixAxe(index : integer) : string;
begin
  Case index of
    0 : Result:='A1';
    1 : Result:='A2';
    2 : Result:='A3';
    3 : Result:='A4';
    4 : Result:='A5';
    end;
end;

function TOF_CPCSECTANA_VIE.GetMinMaxCompte(stCol, stDefaut : String) : String;
var
  Q : TQuery;
  stWhere : String;
begin
  stWhere := ' WHERE S_AXE="' + ChoixAxe(SAxe.itemindex) + '"';
      Q := OpenSQL('SELECT ' + stCol +' CODE FROM SECTION' + stWhere,True);
  if not Q.Eof then
    Result := Q.FindField('CODE').asString
  else
    Result := stDefaut;
  Ferme(Q);
end;

procedure TOF_CPCSECTANA_VIE.OnArgument (S : String ) ;
begin
  FFI_Table := cFI_TABLE;
  FStListeParam :='CPCSECTANA';
  FStArgumentTOF         := S;

  inherited ;
  Ecran.HelpContext := 7603300 ; {Lek 040706}
  // Elements de la fiche Ancêtre
  FListe.OnDblClick      := OnClickDetailSection;
  FListe.Options         := FListe.Options + [goRangeSelect] ;
  FListe.MultiSelect     := False;

  IndiceColDebCre;
  FStSelect    := FStListeChamps;
  FStSelect    := CSqlTextFromList(FStSelect) ;

  Ecran.OnKeyDown := OnKeyDownEcran;
  FListe.GetCellCanvas    := GetCellCanvasFListe;

  SAxe:= THValComboBox(GetControl('SAXE', True));

  S_Section       := THEdit(GetControl('S_SECTION', True));
  S_Section_      := THEdit(GetControl('S_SECTION_', True));
  S_Section.MaxLength       := VH^.Cpta[Axe].lg;
  S_Section_.MaxLength      := VH^.Cpta[Axe].lg;
  S_Section.OnElipsisClick  := OnElipsisClickS_Section;
  S_Section_.OnElipsisClick := OnElipsisClickS_Section;
  S_Section.OnExit          := OnExitS_Section;
  S_Section_.OnExit         := OnExitS_Section;
  SModeSelection  := THValComboBox(GetControl('SMODESELECTION', True));

  SCorresp1       := THMultiValComboBox(GetControl('SCORRESP1', True));

  S_DateDernMvt   := THEdit(GetControl('S_DATEDERNMVT', True));
  S_DateDernMvt_  := THEdit(GetControl('S_DATEDERNMVT_', True));
  SSensNat        := THValComboBox(GetControl('SSENSNAT', True));
  SSensReel       := THValComboBox(GetControl('SSENSREEL', True));
  SSituationCompte:= THValComboBox(GetControl('SSITUATIONCOMPTE', True));

  if SAxe<>nil then begin
    SAxe.itemindex:=0;
    end;

  if (GetControlText('S_SECTION')='') then
    SetControlText('S_SECTION',GetMinMaxCompte('MIN(S_SECTION)','0'));
  if (GetControlText('S_SECTION_')='') then
    SetControlText('S_SECTION_',GetMinMaxCompte('MAX(S_SECTION)','ZZZZZZZZZZZZZZZZZ'));

  PopUpOutil      := TPopUpMenu(GetControl('POPUPOUTIL', True));
  PopUpEdition    := TPopUpMenu(GetControl('POPUPEDITION', True));
  PopUpTraitement := TPopUpMenu(GetControl('POPUPTRAITEMENT', True));
  PopUpUtilitaire := TPopUpMenu(GetControl('POPUPUTILITAIRE', True));
  PopUpFctComp    := TPopUpMenu(GetControl('POPUPFCTCOMP', True));

  BBlocNote       := TToolBarButton97(GetControl('BBLOCNOTE', true));
  BBlocNote.visible:= false;

  TabTablesLibres := TTabSheet(GetControl('TABTABLESLIBRES', True));

  ListeYSituationCompte;
  InitControl; // Valeur par défaut de tous les controles de la fiche

  PopF11.OnPopup          := OnPopUpPopF11;
  PopUpOutil.OnPopup      := OnPopUpPopUpOutil;
  PopUpTraitement.OnPopUp := OnPopUpPopUpTraitement;
  PopUpUtilitaire.OnPopUp := OnPopUpPopUpUtilitaire;
  PopUpFctComp.OnPopUp    := OnPopUpPopUpFctComp;
  PopUpEdition.OnPopUp    := OnPopUpPopUpEdition;

end ;

Initialization
  registerclasses ( [ TOF_CPCSECTANA_VIE ] ) ;
end.


