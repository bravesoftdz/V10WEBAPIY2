{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 09/12/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : CPARAMGENER (CPARAMGENER)
Mots clefs ... : TOM;CPARAMGENER
*****************************************************************}
//XMG 20/04/04 début
//Nouvelle zone LIBELLESELECTION
//Order de tabulation PEmissions/InfosEdition
//XMG 20/04/04 fin
Unit CPARAMGENER_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     MaineAGL,
     eFiche,
{$ELSE}
     Fiche,
     db,
     {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
     hdb,
     FE_Main,
     EdtEtat,
     Saisutil, // Pour le EditDocumentS5S7
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     Ent1,
     UtilPGI,
     HMsgBox,
     UTOM,
     UTob,
     paramSoc,    // GetParamSoc
     ULibEncaDeca // pour TTypeEncaDeca
     ,ULibWindows // pour CDisableControl / CEnableControl  / GereNextControl
     ,Lookup      // pour LookUpValueExist
     ,LP_PARAM    // pour Param_LPTexte
     ;

Procedure CPLanceFiche_ParamGener ( vStRange, vStLequel, vStArgs : String ) ;

Type
  TOM_CPARAMGENER = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
  private

    // Controles
    BParamEtat   : TButton ;
    BParamExport : TButton ;
    BParamBor    : TButton ;
{
    CPG_REFINTERNE1   : THEdit ;
    CPG_LIBELLE1      : THEdit ;
    CPG_REFEXTERNE1   : THEdit ;
    CPG_REFLIBRE1     : THEdit ;
    CPG_REFINTERNE2   : THEdit ;
    CPG_LIBELLE2      : THEdit ;
    CPG_REFEXTERNE2   : THEdit ;
    CPG_REFLIBRE2     : THEdit ;
}
    // Variables
    Action     : String ;                // Stocke XXX de l'argument "ACTION=XXX"
    Operation  : TTypeEncaDeca ;         // Stocke les infos relatives au fonctionnement de l'opération

    // Infos Journal
    TobJournal : TOB ;
    TobGeneral : TOB ;

    // EVT fiche 3ème onglet
    procedure BParamEtatClick         ( Sender : TObject ) ;
    procedure BParamExportClick       ( Sender : TObject ) ;
    procedure BParamBorClick          ( Sender : TObject ) ;

    // EVT fiche 2ème onglet
    procedure BAssistClick(Sender: TObject);

    // Initialisations
    procedure AnalyseParametre( vStArg : String ) ;
    procedure InitComposants ;
    procedure InitEvenements ;
    Procedure InitParPays ;

    // gestion des erreurs
    procedure AfficheMessage( vInErr : Integer ) ;
    function  VerifieUnicite              : Integer ;
    function  VerifieValeursObligatoires  : Integer ;
    function  VerifieCoherenceValeurs     : Integer ;

    // Chargement des infos
    procedure ChargeOperation ;
    procedure ChargeJournal ;
    procedure ChargeGeneral ;

    // Gestion accès interface
    procedure majEcranOperation ;
    procedure majEcranEscompte ;
    procedure majEcranEscompteTVA ;
    procedure majEcranGeneraux ;
    procedure majEcranEdite ;
    procedure majEcranEtat;
    procedure majEcranCFONB ;
    procedure majEcranBOR ;
    procedure majEcranBOREXP ;
    procedure majEcranMethGener ;
    procedure majEcranGlobTiers ;
    procedure MajEcranFlux ;
    procedure MajEcranListeBOR ;
    procedure MajEcranMPForcer ;
    procedure MajEcranMultiSoc ;
    procedure MajEcranMultiEtab ;
    procedure majAccesTiersGlobalise ;
    procedure MajEcranMethodeMultiEtab ;

    // Gestion maj des champs
    procedure majChampsOperation ;
    procedure majChampsEscompte ;
    procedure majChampsEscompteTVA ;
    procedure majChampsGeneraux ;
    procedure majChampsEdite ;
    procedure majChampsEtat ;
    procedure majChampsCFONB ;
    procedure majChampsBOR ;
    procedure majChampsBOREXP ;
    procedure completeGeneral ( F : TField ) ;
    procedure majChampsMethGener ;
    procedure majChampsGlobTiers ;
    procedure MajChampsListeBOR ;
    procedure MajChampsMPForcer ;
    procedure MajChampsMultiSoc ;
    procedure MajChampsMultiEtab ;
    procedure MajChampsMethodeMultiEtab ;

    end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  GuidTool,
  Windows,
  Graphics,
  TypInfo ; // ChoixChampZone

Type TAdjustGraphicControl = Class(TGraphicControl)
      public
       property Canvas ;
      End ;
Type TAdjustCustomControl= Class(TGraphicControl)
      public
       property Canvas ;
      End ;

Function PropertyExists( Ctrl : TComponent ; PropertyName : String ) : Boolean ;
Begin
  Result:=FALSE ;
  if (Assigned(Ctrl)) and (Trim(PropertyName)<>'') then
     Result:=Assigned(GetPropInfo(Ctrl.ClassInfo,PropertyName)) ;
End ;

Function GetPropertyValue( Ctrl : TComponent ; PropertyName : String ) : Variant ;
Var PropInfo : PPropInfo ;
    PropType : PTypeInfo ;
Begin
   result:=#0 ;
   PropInfo:= GetPropInfo(Ctrl.ClassInfo,PropertyName);
   if Assigned(PropInfo) then
      Begin
      PropType := PropInfo^.PropType^;
      case PropType^.Kind of
        tkClass,
        tkInteger,
        tkEnumeration : Result:=GetOrdProp(Ctrl,PropInfo);
        tkFloat       : Result:=GetFloatProp(Ctrl,PropInfo);
        tkString,
        tkLString,
        tkWString     : Result:=GetStrProp(Ctrl,PropInfo);
        end;
      End ;
End ;

Procedure SetPropertyValue( Ctrl : TComponent ; PropertyName : String ; Value : Variant ) ;
Var PropInfo : PPropInfo ;
    PropType : PTypeInfo ;
Begin
   PropInfo:= GetPropInfo(Ctrl.ClassInfo,PropertyName);
   if Assigned(PropInfo) then
      Begin
      PropType := PropInfo^.PropType^;
      case PropType^.Kind of
        tkInteger,
        tkEnumeration : SetOrdProp(Ctrl,PropInfo,Value);
        tkFloat       : SetFloatProp(Ctrl,PropInfo,Value);
        tkString,
        tkLString,
        tkWString     : SetStrProp(Ctrl,PropInfo,String(Value));
        end;
      End ;
End ;
{***********A.G.L.***********************************************
Auteur  ...... : X.Maluenda
Créé le ...... : 12/08/2004
Modifié le ... : 12/08/2004
Description .. : On essaie d'ajuster  la taille (longeur) du control au contenu
Suite ........ : de la propriété Caption (ou text), si le conmtrol a la propriété
Suite ........ : AutoSize, on l'utilise, sinon on fait en sorte de la simuler.....
Suite ........ :
Suite ........ :
Suite ........ : C'est utili pour les controls qui ne disponsent pas de la
Suite ........ : propriété AutoSize (P.Ex. TCheckBox)
Mots clefs ... :
*****************************************************************}
Procedure ForcerAdjustSizeControl ( Ctrl : TControl ) ;
var NewTaille : Integer ;
    Caption   : String ;
    DC        : HDC ;
    Rect      : TRect ;
Begin
  if Assigned(Ctrl) then Begin
     if PropertyExists(Ctrl,'Autosize') then
        SetPropertyValue(Ctrl,'AutoSize',TRUE)
     else Begin
        Caption:='' ;
        Newtaille:=-1 ;
        //On recupere le texte ou caption...
        if PropertyExists(Ctrl,'Caption') then
           Caption:=GetPropertyValue(Ctrl,'Caption')
        else if PropertyExists(Ctrl,'Text') then
             Caption:=GetPropertyValue(Ctrl,'text')
        else
           ;
        //Si on a une caption à adapter, on cherche la nouvelle taille....
        if Trim(Caption)<>'' then Begin
           //... Par rapport aux propriétés disponibles ou au type de control....
           if PropertyExists(Ctrl,'Canvas') then
              NewTaille:=TCanvas(Longint(GetPropertyValue(Ctrl,'Canvas'))).TextWidth(Caption)
           else if (PropertyExists(Ctrl,'Handle')) or
                   (Ctrl is TWinControl)             then Begin
                   DC:=GetDC(0) ;
                   Rect:=Ctrl.ClientRect ;
                   DrawText(DC,PChar(Caption),length(Caption),Rect,DT_CALCRECT) ;
                   NewTaille:=Rect.Right ;
                   ReleaseDC(0,DC) ;
               End
           else if Ctrl is TCustomControl then
                NewTaille:=TAdjustGraphicControl(Ctrl).Canvas.TextWidth(Caption)
           else if Ctrl is TGraphicControl then
                NewTaille:=TAdjustGraphicControl(Ctrl).Canvas.TextWidth(Caption)
           else
              ;
           //Si on a touvé une nouvelle taille, on remet à jour la taille du contrôle....
           if NewTaille>0 then
              Ctrl.Width:=NewTaille ;
        End ;
     End ;
  End ;
End ; //XVI 24/02/2005    *)


//==============================================================================
Procedure CPLanceFiche_ParamGener ( vStRange, vStLequel, vStArgs : String ) ;
begin
  AGLLanceFiche('CP', 'CPPARAMGENER', vStRange, vStLequel, vStArgs) ;
end ;
//==============================================================================

procedure TOM_CPARAMGENER.OnNewRecord ;
begin
  Inherited ;
  SetField('CPG_BORNATURE', 'BOR' ) ;
end ;

procedure TOM_CPARAMGENER.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_CPARAMGENER.OnUpdateRecord ;
Var lInErr : Integer ; // Code Erreur
begin
  Inherited ;

  NextPrevControl(Ecran) ;

  // Unicite
  lInErr := VerifieUnicite ;
  if lInErr <> CPG_PASERREUR then
    begin
    AfficheMessage ( lInErr ) ;
    LastError := -1 ;
    Exit ;
    end ;

  // Valeurs obligatoires
  lInErr := VerifieValeursObligatoires ;
  if lInErr <> CPG_PASERREUR then
    begin
    AfficheMessage ( lInErr ) ;
    LastError := -1 ;
    Exit ;
    end ;

  // Cohérence des données
  lInErr := VerifieCoherenceValeurs ;
  if lInErr <> CPG_PASERREUR then
    begin
    AfficheMessage ( lInErr ) ;
    LastError := -1 ;
    Exit ;
    end ;

end ;

procedure TOM_CPARAMGENER.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_CPARAMGENER.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_CPARAMGENER.OnChangeField ( F: TField ) ;
begin
  Inherited ;

  // Flux
  if F.FieldName = 'CPG_FLUXENCADECA' then
    begin
    if VH^.PaysLocalisation=CodeISOES then
       majEcranFlux ; //XVI 24/02/2005
    // MAJ interface
    majEcranGeneraux;
    // MAJ Données
    majChampsGeneraux;
    // FQ 21090 SBO 07/08/2007 : blocage de l'escompte pour les encaissements
    majEcranOperation; // MAJ interface
    majChampsOperation; // MAJ Données
    end
  // opération
  else if F.FieldName = 'CPG_TYPEENCADECA' then
    begin
    ChargeOperation ;
    // MAJ interface
    majEcranOperation;
    // MAJ Données
    majChampsOperation;
    // Gestion du focus particulier du au enabled/disable qui précèdent
    if (VH^.PaysLocalisation=CodeISOES) and //XVI 24/02/2005
       ((DS.State=dsEdit) or
       ((DS.State=dsInsert) and (not AfterInserting))) then
       GereNextControl(Ecran,'CPG_TYPEENCADECA') ;
    end
  // Comptabilisation
  else if F.FieldName = 'CPG_COMPTABILISE' then
    begin
    majEcranBOREXP ;
    majChampsBOREXP ;
    end
  // Méthode de génération
  else if F.FieldName = 'CPG_GROUPEENCADECA' then
    begin
    majEcranGlobTiers ;
    majChampsGlobTiers ;
    MajEcranMethodeMultiEtab ;
    MajChampsMethodeMultiEtab ;
    end
  // Journal
  else if F.FieldName = 'CPG_JOURNAL' then
    begin
    ChargeJournal ;
    majEcranGeneraux ;
    majChampsGeneraux ;
    end
  // Escompte
  else if F.FieldName = 'CPG_ESCMETHODE' then
    begin
    majEcranEscompte ;
    majChampsEscompte ;
    end
  // Tva sur Escompte
  else if F.FieldName = 'CPG_ESCTVAAVEC' then
    begin
    majEcranEscompteTVA ;
    majChampsEscompteTVA ;
    end
  // Utilisation des états ou des anciens documents
  else if F.FieldName = 'CPG_AVECDOC' then
    begin
    majEcranEtat ;
    majChampsEtat ; //XVI 24/02/2005
    end
  // Indicateur d'édition
  else if F.FieldName = 'CPG_EDITE' then
    begin
    if VH^.PaysLocalisation=CodeISOES then
      begin
       MajEcranEdite ;
       MajChampsEdite ;
      end else
      begin
       majEcranEtat ;
       majChampsEtat ;
      End ; //XVI 24/02/2005
    end
  // Gestion d'emissions.....
  else if F.FieldName = 'CPG_BORDEREAUEXP' then
    begin
    majEcranBOR ;
    MajChampsBOR ;
    end
  // Gestion d'emissions.....
  else if F.FieldName = 'CPG_CFONBEXPORT' then
    begin
    majEcranCFONB ;
    MajChampsCFONB ;
    end
  // Bourrage du compte général
  else if F.FieldName='CPG_GENERAL' then
    begin
    completeGeneral ( F ) ;
    ChargeGeneral ;
    majEcranMethGener ;
    majChampsMethGener ;
    end
  // Bourrage du compte de sélection (ESP)
  else if F.FieldName = 'CPG_GENERALSEL' then
    begin
    completeGeneral ( F ) ;
    end
  else if (VH^.PaysLocalisation=CodeISOES) and (F.FieldName = 'CPG_FORCERMP') then //XVI 24/02/2005
    Begin
    SetControlEnabled('CPG_MODEPAIEMENT',(F.AsString='X')) ;
    if (DS.State=dsEdit) or
       ((DS.State=dsInsert) and (not AfterInserting)) then
       Begin
       GereNextControl(Ecran,'CPG_FORCERMP') ;
       if (GetField('CPG_MODEPAIEMENT')<>'') and (F.AsString<>'X') then SetField('CPG_MODEPAIEMENT','') ;
       End ;
    end
  // Sélection de la nature des états bordereau
  else if F.FieldName = 'CPG_BORNATURE' then
    begin
    // Seul les états BOR classique sont supporté dans les scénarii
    MajEcranListeBOR ;
    MajChampsListeBOR ;
    end
  else if F.FieldName = 'CPG_MODEPAIEMENT' then
    begin
    MajEcranMPForcer ;
    MajChampsMPForcer ;
    end
  // Gestion Multisoc
  else if F.FieldName = 'CPG_MULTISOC' then
    begin
    MajEcranMultiSoc ;
    MajChampsMultiSoc ;
    end
  else if F.FieldName = 'CPG_MULTIETAB' then
    begin
    MajEcranMultiEtab ;
    MajChampsMultiEtab ;
    end;

end ;

procedure TOM_CPARAMGENER.OnArgument ( S: String ) ;
begin

  Inherited ;

  // Analyse des paramètres
  AnalyseParametre( S ) ;

  // Initialisation de l'interface
  InitComposants ;
  InitEvenements ;

  //Adaptations Par Pays
  InitParPays ;

  // initialisation variable
  ChargeJournal ;
  ChargeGeneral ;

end ;

procedure TOM_CPARAMGENER.OnClose ;
begin
  Inherited ;
  FreeAndNil( TobJournal ) ;
  FreeAndNil( TobGeneral ) ;

end ;

procedure TOM_CPARAMGENER.OnCancelRecord ;
begin
  Inherited ;
end ;


// ==================================
// ==== INITIALISATION INTERFACE ====
// ==================================

procedure TOM_CPARAMGENER.AnalyseParametre(vStArg: String);
var lStChaine : String ;
begin

  while vStArg <> '' do
    begin
    lStChaine := ReadTokenSt ( vStArg ) ;

    // Stocke le XXX de "ACTION=XXX" dans la variable Action
    if Pos( 'ACTION=' , Uppercase( lStChaine ) ) > 0 then
      Action := Copy(lStChaine, 8, length( lStChaine ) ) ;
    end ;

end;

procedure TOM_CPARAMGENER.InitComposants;
begin

  BParamEtat          := TButton (      Getcontrol('BPARAMETAT',          True) ) ;
  BParamExport        := TButton (      Getcontrol('BPARAMEXPORT',        True) ) ;
  BParamBor           := TButton (      Getcontrol('BPARAMBOR',           True) ) ;

  {$IFDEF EAGLCLIENT}
  BParamEtat.Visible   := False ;
  BParamExport.Visible := False ;
  BParamBor.Visible    := False ;
  {$ELSE}
  BParamBOR.Visible := (VH^.PaysLocalisation<>CodeISOES) or (V_PGI.SAV) ; //XVI 24/02/2005
  {$ENDIF}

  // SBO 30/07/2004 Correction fiche FQ 13741 : Bizzarement des @ viennent s'intercaler dans le hint...???
  BParamEtat.Hint   := 'Paramétrage de l''état' ;
  BParamExport.Hint := 'Générateur de fichier' ;
  BParamBor.Hint    := 'Paramétrage de l''état' ;

  // Pour le moment, champ non accessible :
//  SetControlEnabled('CPG_AVECVERIFPIECE', False) ;
  SetControlEnabled('CPG_AVECVERIFETAT', False) ;

  // Accès multisoc
  if not EstMultiSoc then
    SetControlVisible('PMULTISOC',     EstMultiSoc ) ;
//    CDisableControl(  TCustomControl( GetControl('INFOSMULTISOC') ) ) ;
//  SetControlEnabled('INFOSMULTISOC', EstMultiSoc ) ;
//  TTabSheet( GetControl( 'PMULTISOC' ) ).

  // Accès multiEtab
  if not VH^.EtablisCpta then
    CDisableControl(  TCustomControl( GetControl('INFOSETAB') ) ) ;

  // Ajout assistant pour les formules
{
  CPG_REFINTERNE1   := THEdit( GetControl( 'CPG_REFINTERNE1', True ) ) ;
  CPG_LIBELLE1      := THEdit( GetControl( 'CPG_LIBELLE1', True ) ) ;
  CPG_REFEXTERNE1   := THEdit( GetControl( 'CPG_REFEXTERNE1', True ) ) ;
  CPG_REFLIBRE1     := THEdit( GetControl( 'CPG_REFLIBRE1', True ) ) ;
  CPG_REFINTERNE2   := THEdit( GetControl( 'CPG_REFINTERNE2', True ) ) ;
  CPG_LIBELLE2      := THEdit( GetControl( 'CPG_LIBELLE2', True ) ) ;
  CPG_REFEXTERNE2   := THEdit( GetControl( 'CPG_REFEXTERNE2', True ) ) ;
  CPG_REFLIBRE2     := THEdit( GetControl( 'CPG_REFLIBRE2', True ) ) ;
  CPG_REFINTERNE1.ElipsisButton := True ;
  CPG_LIBELLE1.ElipsisButton    := True ;
  CPG_REFEXTERNE1.ElipsisButton := True ;
  CPG_REFLIBRE1.ElipsisButton   := True ;
  CPG_REFINTERNE2.ElipsisButton := True ;
  CPG_LIBELLE2.ElipsisButton    := True ;
  CPG_REFEXTERNE2.ElipsisButton := True ;
  CPG_REFLIBRE2.ElipsisButton   := True ;
}

end;

procedure TOM_CPARAMGENER.InitEvenements;
begin
  BParamEtat.OnClick        := BParamEtatClick ;
  BParamExport.OnClick      := BParamExportClick ;
  BParamBor.OnClick         := BParamBorClick ;

{  CPG_REFINTERNE1.OnElipsisClick := BAssistClick  ;
  CPG_LIBELLE1.OnElipsisClick    := BAssistClick  ;
  CPG_REFEXTERNE1.OnElipsisClick := BAssistClick  ;
  CPG_REFLIBRE1.OnElipsisClick   := BAssistClick  ;
  CPG_REFINTERNE2.OnElipsisClick := BAssistClick  ;
  CPG_LIBELLE2.OnElipsisClick    := BAssistClick  ;
  CPG_REFEXTERNE2.OnElipsisClick := BAssistClick  ;
  CPG_REFLIBRE2.OnElipsisClick   := BAssistClick  ;
}
{$IFDEF EAGLCLIENT}
  THEdit( GetControl( 'CPG_REFINTERNE1', True ) ).OnElipsisClick := BAssistClick  ;
  THEdit( GetControl( 'CPG_LIBELLE1', True ) ).OnElipsisClick    := BAssistClick  ;
  THEdit( GetControl( 'CPG_REFEXTERNE1', True ) ).OnElipsisClick := BAssistClick  ;
  THEdit( GetControl( 'CPG_REFLIBRE1', True ) ).OnElipsisClick   := BAssistClick  ;
  THEdit( GetControl( 'CPG_REFINTERNE2', True ) ).OnElipsisClick := BAssistClick  ;
  THEdit( GetControl( 'CPG_LIBELLE2', True ) ).OnElipsisClick    := BAssistClick  ;
  THEdit( GetControl( 'CPG_REFEXTERNE2', True ) ).OnElipsisClick := BAssistClick  ;
  THEdit( GetControl( 'CPG_REFLIBRE2', True ) ).OnElipsisClick   := BAssistClick  ;
{$ELSE}
  THDBEdit( GetControl( 'CPG_REFINTERNE1', True ) ).OnElipsisClick := BAssistClick  ;
  THDBEdit( GetControl( 'CPG_LIBELLE1', True ) ).OnElipsisClick    := BAssistClick  ;
  THDBEdit( GetControl( 'CPG_REFEXTERNE1', True ) ).OnElipsisClick := BAssistClick  ;
  THDBEdit( GetControl( 'CPG_REFLIBRE1', True ) ).OnElipsisClick   := BAssistClick  ;
  THDBEdit( GetControl( 'CPG_REFINTERNE2', True ) ).OnElipsisClick := BAssistClick  ;
  THDBEdit( GetControl( 'CPG_LIBELLE2', True ) ).OnElipsisClick    := BAssistClick  ;
  THDBEdit( GetControl( 'CPG_REFEXTERNE2', True ) ).OnElipsisClick := BAssistClick  ;
  THDBEdit( GetControl( 'CPG_REFLIBRE2', True ) ).OnElipsisClick   := BAssistClick  ;
{$ENDIF EAGLCLIENT}

end;

// ==============================
// ==== EVENEMENTS INTERFACE ====
// ==============================

procedure TOM_CPARAMGENER.BParamEtatClick(Sender: TObject);
  {$IFDEF EAGLCLIENT}
  {$ELSE}
var
  lStOldEtat : String;
  {$ENDIF}
begin

  {$IFDEF EAGLCLIENT}
  {$ELSE}

  // Pour forcer la maj de la sélection de l'état,
  //  ( mal mis à jour si clique direct sur bouton de parmétrage ... )
  NextPrevControl(Ecran) ;

  if not ExJaiLeDroitConcept(ccParamEtat, True) then Exit;

  // Edition Etat / Document
  lStOldEtat := GetField('CPG_ModeleEncaDeca') ;
  if GetField('CPG_AVECDOC')<>'-'
     then EditDocumentS5S7('L',Operation.NatureDoc,lStOldEtat,True)
     else EditEtat('E', Operation.NatureEtat, lStOldEtat, TRUE, nil, '', Ecran.Caption);


  // Rechargement de la Combo
  THValComboBox(GetControl('CPG_ModeleEncaDeca')).ReLoad ;
  if lStOldEtat<>'' then
     SetControlText('CPG_ModeleEncaDeca', lStOldEtat ) ;

  {$ENDIF}

end;

// ===============================
// ==== TRAITEMENTS INTERFACE ====
// ===============================

procedure TOM_CPARAMGENER.ChargeOperation;
begin
  Operation := ChargeTypeEncaDeca( GetField('CPG_TYPEENCADECA') ) ;
end;

procedure TOM_CPARAMGENER.AfficheMessage(vInErr: Integer);
begin

  if vInErr = CPG_WNGMODIFCASCADE then
   PGIInfo('Attention ! La zone que vous venez de modifier en impacte d''autres. Veuillez vérifier votre scénario.', Ecran.Caption )

  else if vInErr = CPG_ERRUNICITE then
   PGIBox('Ce code est déjà utilisé dans un scénario existant.', Ecran.Caption )

  else if vInErr = CPG_ERRCODE then
    PGIBox('Vous devez renseigner un code sur 3 caractères.', Ecran.Caption )

  else if vInErr = CPG_ERRLIBELLE then
    PGIBox('Vous devez renseigner un libellé.', Ecran.Caption )

  else if vInErr = CPG_ERRMODELE then
    PGIBox('Vous devez renseigner un modèle d''état.', Ecran.Caption )

  else if vInErr = CPG_ERRTYPE then
    PGIBox('Vous devez renseigner une catégorie.', Ecran.Caption )

  else if vInErr = CPG_ERRFLUX then
    PGIBox('Vous devez renseigner un flux.', Ecran.Caption )

  else if vInErr = CPG_ERRJOURNAL then
    PGIBox('Vous devez renseigner un journal.', Ecran.Caption )

  else if vInErr = CPG_ERRGENERAL then
    PGIBox('Vous devez renseigner un compte de génération.', Ecran.Caption )

  else if vInErr = CPG_ERRGROUPE then
    PGIBox('Vous devez renseigner une méthode de génération.', Ecran.Caption )

  else if vInErr = CPG_ERRMETHANA then
    PGIBox('Vous devez renseigner un méthode de génération pour l''analytique.', Ecran.Caption )

  else if vInErr = CPG_ERRESCMETH then
    PGIBox('Vous devez renseigner un méthode de génération pour l''escompte.', Ecran.Caption )

  else if vInErr = CPG_ERRESCCPT then
    PGIBox('Vous devez renseigner un compte d''escompte.', Ecran.Caption )

  else if vInErr = CPG_ERRESCTAUX then
    PGIBox('Vous devez renseigner un taux d''escompte.', Ecran.Caption )

  else if vInErr = CPG_ERRESCTVACPT then
    PGIBox('Vous devez renseigner un compte de TVA pour la pièce d''escompte.', Ecran.Caption )

  else if vInErr = CPG_ERRESCTVATAUX then
    PGIBox('Vous devez renseigner un taux de TVA pour l''escompte.', Ecran.Caption )

  else if vInErr = CPG_ERRMODEPAIE then
    PGIBox('Vous devez renseigner un mode de paiement.', Ecran.Caption )

  else if vInErr = CPG_ERRCFONBFORMAT then
    PGIBox('Vous devez renseigner un format d''exportation CFONB.', Ecran.Caption )

  else if vInErr = CPG_ERRBORDEREAUMOD then
    PGIBox('Vous devez renseigner un modèle de bordereau.', Ecran.Caption )

  else if vInErr = CPG_ERRMSMETH then
    PGIBox('Vous devez renseigner une méthode de génération multi-sociétés.', Ecran.Caption )

  else if vInErr = CPG_ERRMSJAL then
    PGIBox('Vous devez renseigner un journal de génération multi-sociétés.', Ecran.Caption )

  else if vInErr = CPG_ERRMETHETAB then
    PGIBox('Vous devez renseigner une méthode de génération multi-établissements.', Ecran.Caption )

  else if vInErr = CPG_ERRSELECTETAB then
    PGIBox('Vous devez renseigner une méthode de sélection de l''établissement.', Ecran.Caption )

  else if (VH^.PaysLocalisation=CodeISOES) and (vInErr = CPG_ESP_ERRCNDREMINCOHERENTE) then  //XVI 24/02/2005
    PGIBox('Veuillez vérifier la cohérence entre la catégorie et le type d''encaissement de la condition de remise.', Ecran.Caption )
  else
    PGIBox('Votre saisie est incorrecte, veuillez vérifier.', Ecran.Caption ) ;

end;


function TOM_CPARAMGENER.VerifieCoherenceValeurs: Integer;
Var LeType : String ;
begin
  Result := CPG_PASERREUR ;

  // Vérif des infos de génération
  if GetField('CPG_Comptabilise')<>'-' then
    begin

    // Mode de paiement
    if (VH^.PaysLocalisation=CodeISOES) and //XVI 24/02/2005
       ( GetField('CPG_FORCERMP')='X') and
       (GetField('CPG_MODEPAIEMENT') = '' ) then
      begin
      Result := CPG_ERRMODEPAIE ;
      SetActiveTabSheet('PCOMPTABILISATION') ;
      SetFocusControl('CPG_MODEPAIEMENT') ;
      Exit ;
      end ;
    // Méthode d'escompte
    if GetField('CPG_ESCMETHODE') = '' then
      begin
      Result := CPG_ERRESCMETH ;
      SetActiveTabSheet('PCOMPTABILISATION') ;
      SetFocusControl('CPG_MODEPAIEMENT') ;
      Exit ;
      end ;

    // Vérif infos escompte
    if GetField('CPG_ESCMETHODE') <> 'RIE' then
      begin

      // Compte general d'escompte
      if GetField('CPG_ESCCPTGENE') = '' then
        begin
        Result := CPG_ERRESCCPT ;
        SetActiveTabSheet('PCOMPTABILISATION') ;
        SetFocusControl('CPG_ESCCPTGENE') ;
        Exit ;
        end ;

      // Taux d'escompte
      if GetField('CPG_ESCTAUX') = 0 then
        begin
        Result := CPG_ERRESCTAUX ;
        SetActiveTabSheet('PCOMPTABILISATION') ;
        SetFocusControl('CPG_ESCTAUX') ;
        Exit ;
        end ;

      // Vérif infos tva sur escompte
      if GetField('CPG_ESCTVAAVEC')<>'-' then
        begin

        // Compte de TVA
        if GetField('CPG_ESCTVACPT') = '' then
          begin
          Result := CPG_ERRESCTVACPT ;
          SetActiveTabSheet('PCOMPTABILISATION') ;
          SetFocusControl('CPG_ESCTVACPT') ;
          Exit ;
          end ;

        // Taux d'escompte
        if GetField('CPG_ESCTVATAUX') = 0 then
          begin
          Result := CPG_ERRESCTVATAUX ;
          SetActiveTabSheet('PCOMPTABILISATION') ;
          SetFocusControl('CPG_ESCTVATAUX') ;
          Exit ;
          end ;

        end ;
      end ;
    end;
  // Vérif des Condition de remise
  if (VH^.PaysLocalisation=CodeISOES) and (Trim(GetField('CPG_CONDREMISE'))<>'') then //XVI 24/02/2005
    begin
    LeType:='ENC' ;
    if pos(';'+GetField('CPG_TYPEENCADECA')+';',';REE;RJE;IME;')>0 then LeType:='ESC' ;
    if  not ExisteSQL('select CCB_CONDREMISE from CCONDREMISE where CCB_CONDREMISE="'+GetField('CPG_CONDREMISE')+'" and CCB_TYPEREMISE="'+LeType+'"') then
        Begin
        Result := CPG_ESP_ERRCNDREMINCOHERENTE ;
        SetActiveTabSheet('PGENERAL') ;
        SetFocusControl('CPG_CONDREMISE') ;
        Exit ;
        End ;
    End ;

end;

function TOM_CPARAMGENER.VerifieValeursObligatoires: Integer;
begin
  Result := CPG_PASERREUR ;

  // Test sur le code
  if Length( Trim ( GetField('CPG_CODE') ) ) < 3 then
    begin
    Result := CPG_ERRCODE ;
    SetActiveTabSheet('PGENERAL') ;
    SetFocusControl('CPG_CODE') ;
    Exit ;
    end ;

  // Test ssur le libellé
  if GetField('CPG_LIBELLE') = '' then
    begin
    Result := CPG_ERRLIBELLE ;
    SetActiveTabSheet('PGENERAL') ;
    SetFocusControl('CPG_LIBELLE') ;
    Exit ;
    end ;

  // Test sur le flux
  if GetField('CPG_FLUXENCADECA') = '' then
    begin
    Result := CPG_ERRFLUX ;
    SetActiveTabSheet('PGENERAL') ;
    SetFocusControl('CPG_FLUXENCEDECA') ;
    Exit ;
    end ;

  // Test sur l'opération
  if GetField('CPG_TYPEENCADECA') = '' then
    begin
    Result := CPG_ERRTYPE ;
    SetActiveTabSheet('PGENERAL') ;
    SetFocusControl('CPG_TYPEENCEDECA') ;
    Exit ;
    end ;

  // Test sur modèle d'état
  if GetField('CPG_EDITE')<>'-' then
    if GetField('CPG_ModeleEncaDeca') = '' then
      begin
      Result := CPG_ERRMODELE ;
      SetActiveTabSheet('PEMISSIONS') ;
      SetFocusControl('CPG_MODELEENCADECA') ;
      Exit ;
      end ;

  // Test sur exportCFONB
  if (VH^.PaysLocalisation=CodeISOES) then
    if GetField('CPG_CFONBEXPORT')<>'-' then
      if GetField('CPG_CFONBFORMAT') = '' then
        begin
        Result := CPG_ERRCFONBFORMAT ;
        SetActiveTabSheet('PEMISSIONS') ;
        SetFocusControl('CPG_CFONBFORMAT') ;
        Exit ;
        end ;

  // Test sur Bordereau
  if GetField('CPG_BORDEREAUEXP')<>'-' then
    if GetField('CPG_BORDEREAUMOD') = '' then
      begin
      Result := CPG_ERRBORDEREAUMOD ;
      SetActiveTabSheet('PEMISSIONS') ;
      SetFocusControl('CPG_BORDEREAUMOD') ;
      Exit ;
      end ;

  // Test sur Journal
  if not LookUpValueExist( GetControl('CPG_Journal') ) then
    begin
    Result := CPG_ERRJOURNAL ;
    SetActiveTabSheet('PGENERAL') ;
    SetFocusControl('CPG_JOURNAL') ;
    Exit ;
    end ;

  // Test sur compte de génération
  if not LookUpValueExist( GetControl('CPG_General',True) ) then
    begin
    Result := CPG_ERRGENERAL ;
    SetActiveTabSheet('PGENERAL') ;
    SetFocusControl('CPG_GENERAL') ;
    Exit ;
    end ;

  // Test sur méthode de génération
  if GetField('CPG_GROUPEENCADECA') = '' then
    begin
    Result := CPG_ERRGROUPE ;
    SetActiveTabSheet('PGENERAL') ;
    SetFocusControl('CPG_GROUPEENCADECA') ;
    Exit ;
    end ;

  // Test sur méthode de génération de l'analytique
  if GetField('CPG_METHODEANA') = '' then
    begin
    Result := CPG_ERRMETHANA ;
    SetActiveTabSheet('PGENERAL') ;
    SetFocusControl('CPG_METHODEANA') ;
    Exit ;
    end ;

  // Test sur compte de selection
  if (VH^.PaysLocalisation=CodeISOES) and (not LookUpValueExist(GetControl('CPG_GENERALSEL'))) then
    begin
    Result := CPG_ERRGENERALSEL ;
    SetActiveTabSheet('PGENERAL') ;
    SetFocusControl('CPG_GENERALSEL') ;
    Exit ;
    end ; //XVI 24/02/2005

  // Test sur multisoc
  if GetField('CPG_MULTISOC') = 'X' then
    begin
    if GetField('CPG_MULTISOCJAL') = '' then
      begin
      Result := CPG_ERRMSJAL ;
      SetActiveTabSheet('PMULTISOC') ;
      SetFocusControl('CPG_MULTISOCJAL') ;
      Exit ;
      end
    else if GetField('CPG_MULTISOCMETH') = '' then
      begin
      Result := CPG_ERRMSMETH ;
      SetActiveTabSheet('PMULTISOC') ;
      SetFocusControl('CPG_MULTISOCMETH') ;
      Exit ;
      end ;
    end ;

  // Vérif multi-établissement
  if GetField('CPG_MULTIETAB') = 'X' then
    begin
    if GetField('CPG_METHODEETAB') = '' then
      begin
      Result := CPG_ERRMETHETAB ;
      SetActiveTabSheet('PCOMPTABILISATION') ;
      SetFocusControl('CPG_METHODEETAB') ;
      Exit ;
      end
    else if GetField('CPG_SELECTETAB') = '' then
      begin
      Result := CPG_ERRSELECTETAB ;
      SetActiveTabSheet('PCOMPTABILISATION') ;
      SetFocusControl('CPG_SELECTETAB') ;
      Exit ;
      end ;
    end ;


end;

function TOM_CPARAMGENER.VerifieUnicite: Integer;
begin

  Result := CPG_PASERREUR ;
  if Action<>'CREATION' then Exit ;

  if ExisteSQL('SELECT CPG_CODE FROM CPARAMGENER WHERE CPG_CODE="' + GetControlText('CPG_CODE') + '"')
    then begin
         Result := CPG_ERRUNICITE ;
         SetActiveTabSheet('PGENERAL') ;
         SetFocusControl('CPG_CODE') ;
         End ;

end;

procedure TOM_CPARAMGENER.majChampsEscompteTVA;
begin
  if DS.State in [dsBrowse] then exit;

  // Valeur par défaut si non accessible
  if GetField('CPG_ESCTVAAVEC')='-' then
    begin
    SetField( 'CPG_ESCTVACPT' ,     '' ) ;
    SetField( 'CPG_ESCTVATAUX' ,    0  ) ;
    end ;
end;

procedure TOM_CPARAMGENER.majEcranOperation;
var lStListe : String ;
begin
   if VH^.PaysLocalisation=CodeISOES then //XVI 24/02/2005
    // Liste des états
      SetControlProperty('CPG_ModeleEncaDeca', 'Plus', Operation.NatureEtat ) ;

  // Coche "avec édition" + Nature Etat
  SetControlEnabled('CPG_Edite', Operation.AccesEdition = taPossible ) ;

  if VH^.PaysLocalisation<>CodeISOES then //XVI 24/02/2005
      majEcranEtat ;

  // Coche "avec comptabilisation"
  SetControlEnabled('CPG_COMPTABILISE', Operation.AccesCompta = taPossible );

  // MAJ des coches émission de bordereau et export de fichier
  majEcranBOREXP ;

  // Liste des journaux
  SetControlProperty('CPG_Journal', 'Datatype', Operation.AccesJournaux ) ;

  // Limitation des journaux
  majEcranGeneraux ;

  // Liste des mode de paiement
  if VH^.PaysLocalisation<>CodeISOES then //XVI 24/02/2005
     if GetControlText('CPG_TYPEENCADECA')<> ''
        then SetControlProperty('CPG_MODEPAIEMENT', 'Plus', GetControlText('CPG_TYPEENCADECA') )
        else SetControlProperty('CPG_MODEPAIEMENT', 'Plus', W_W ) ;

  // Méthode de génération de l'escompte
  // FQ 21090 SBO 07/08/2007 : blocage de l'escompte pour les encaissements
  SetControlEnabled('CPG_EscMethode',  (Operation.AccesEscompte <> taJamais) and (GetField('CPG_FLUXENCADECA') <> 'ENC') ) ;

  // Limitation des méthodes de génération
  if Operation.AccesMethode <> '' then
    begin
    lStListe := Operation.AccesMethode ;
    lStListe := FindEtReplace( lStListe, ';' , '","', True ) ;
    SetControlProperty('CPG_GroupeEncaDeca','Plus', ' AND CO_CODE IN ("' + lStListe + '") ' ) ;
    end ;

  //Adaptations Spécifiques Espagnoles
  if VH^.PaysLocalisation=CodeISOES then
     Begin
     SetControlEnabled('CPG_CONDREMISE',Operation.AccesCondRemise) ;
     End ; //XVI 24/02/2005

  // Gestion du champ E_NUMTRAITECHQ
  SetControlEnabled('CPG_AVECNUMCHEQUE', (Operation.Code = 'CHQ') ) ;

end;

procedure TOM_CPARAMGENER.majChampsOperation;
begin
  if DS.State in [dsBrowse] then exit;

  // Comptabilisation obligatoire ?
  if Operation.AccesCompta <> taPossible then
     if Operation.AccesCompta = taObligatoire
        then SetField('CPG_COMPTABILISE', 'X' )
        else SetField('CPG_COMPTABILISE', '-' ) ;

  // Edition obligatoire ?
  if Operation.AccesEdition <> taPossible then
     if Operation.AccesEdition = taObligatoire
        then SetField('CPG_EDITE', 'X' )
        else SetField('CPG_EDITE', '-' ) ;

  // Coche "avec export CFONB"
  if Operation.AccesCFONB <> taPossible then
     if Operation.AccesCFONB = taObligatoire
        then SetField('CPG_CFONBEXPORT', 'X' )
        else SetField('CPG_CFONBEXPORT', '-' ) ;

  // Coche "avec bordereau"
  if Operation.AccesBordereau <> taPossible then
     if Operation.AccesBordereau = taObligatoire
        then SetField('CPG_BORDEREAUEXP', 'X' )
        else SetField('CPG_BORDEREAUEXP', '-' ) ;

  // Escompte Obligatoire ?
  // FQ 21090 SBO 07/08/2007 : blocage de l'escompte pour les encaissements
  if (Operation.AccesEscompte = taJamais) or (GetField('CPG_FLUXENCADECA')='ENC')
     then SetField('CPG_ESCMETHODE', 'RIE' ) ;

  // MAJ du mode de paiement
  SetField('CPG_MODEPAIEMENT', '') ;

  // MAJ Nature d'Etat
  majChampsEtat ;

  // MAJ Généraux
  majChampsGeneraux;

  // Gestion du champ E_NUMTRAITECHQ
  if (Operation.Code <> 'CHQ') then
     SetField('CPG_AVECNUMCHEQUE', '-' ) ;

  //Adaptations Spécifiques Espagnoles
  if VH^.PaysLocalisation=CodeISOES then
     Begin
     if not Operation.AccesCondRemise then
        SetField('CPG_CONDREMISE','') ;
     End ; //XVI 24/02/2005

  // YMO  15/12/2005  Lettrage auto ; activé par défaut
  SetField('CPG_LETTRAGEAUTO', 'X') ;

end;

procedure TOM_CPARAMGENER.ChargeJournal;
var Q : TQuery ;
begin
  // Pour éviter les chargements multiples
  if (VH^.PaysLocalisation=CodeISOES) and
     {JP 19/07/05 : FQ 16247 : pour revenir à l'état avant suppression de IFDEF ESP}
     ((not Assigned(TobJournal)) or ( GetField('CPG_Journal') <> TobJournal.GetValue('J_JOURNAL'))) then
   begin

   // Init Tob
   if Assigned(TobJournal)
      then FreeAndNil( TobJournal ) ;

   TobJournal := TOB.Create('JOURNAL', nil, -1 ) ;

   // Chargement des infos utiles du journal dans la tob
   if length( Trim( GetField('CPG_Journal') ) ) > 0 then
     begin
     // Recherche nature et compte de contrepartie
     Q := OpenSQL('SELECT J_JOURNAL, J_LIBELLE, J_CONTREPARTIE, J_NATUREJAL, J_EFFET FROM JOURNAL WHERE J_JOURNAL="'
                   + GetField('CPG_Journal') + '" ' , TRUE ) ;
     TobJournal.SelectDB( '', Q ) ;
     Ferme(Q) ;
     end ;
   End else
   Begin
     // Pour éviter les chargements multiples
     if Assigned(TobJournal) and ( GetField('CPG_Journal') = TobJournal.GetValue('J_JOURNAL'))
      then Exit ;

     // Init Tob
     FreeAndNil( TobJournal ) ;
     TobJournal := TOB.Create('JOURNAL', nil, -1 ) ;

     // Chargement des infos utiles du journal dans la tob
     if length( Trim( GetField('CPG_Journal') ) ) > 0 then
       begin
       // Recherche nature et compte de contrepartie
       Q := OpenSQL('SELECT J_JOURNAL, J_LIBELLE, J_CONTREPARTIE, J_NATUREJAL, J_EFFET FROM JOURNAL WHERE J_JOURNAL="'
                     + GetField('CPG_Journal') + '" ' , TRUE ) ;
       TobJournal.SelectDB( '', Q ) ;
       Ferme(Q) ;
       end ;
   End ; //XMG 24/02/2005
end;

procedure TOM_CPARAMGENER.majChampsEscompte;
begin
  if DS.State in [dsBrowse] then exit;

  if GetField('CPG_EscMethode') = 'RIE' then
     begin
     SetField( 'CPG_ESCPRORATA' ,   '-' ) ;
     SetField( 'CPG_ESCCPTGENE' ,   ''    ) ;
     SetField( 'CPG_ESCTAUX' ,      0     ) ;
     SetField( 'CPG_ESCTVAAVEC' ,   '-' ) ;
     SetField( 'CPG_ESCTVACPT' ,    ''    ) ;
     SetField( 'CPG_ESCTVATAUX' ,   0     ) ;
     end ;

end;

procedure TOM_CPARAMGENER.majEcranEscompte;
begin
  // Accès paramètrage de l'escompte
  if GetField('CPG_EscMethode') = 'RIE'
     then CDisableControl(  TCustomControl( GetControl('INFOSESCOMPTE') ) )
     else
       begin
     { // en attente
       if GetParamSocSecur('SO_CPJALESCOMPTE','')='' then
         PGIInfo ( 'Le journal d''escompte n''est pas renseigné dans les paramètres société. Aucune escompte ne pourra être générée.'
                   , Ecran.Caption ) ;}
       CEnableControl(   TCustomControl( GetControl('INFOSESCOMPTE') ) ) ;
       majEcranEscompteTVA ;
       end ;
end;

procedure TOM_CPARAMGENER.majEcranEscompteTVA;
begin
  SetControlEnabled( 'CPG_ESCTVACPT' ,     GetField('CPG_ESCTVAAVEC')<>'-' ) ;
  SetControlEnabled( 'CPG_ESCTVATAUX' ,    GetField('CPG_ESCTVAAVEC')<>'-' ) ;
end;

procedure TOM_CPARAMGENER.majChampsEtat;
//var lStOldModele : String ;
begin
  if DS.State in [dsBrowse] then exit;

//  lStOldModele := GetField('CPG_MODELEENCADECA') ;

  // Paramètrage type d'état utilisé
  if GetField('CPG_AVECDOC')<>'-' then
    begin
    // Pour les documents
    if GetField('CPG_ETATNATURE') <> Operation.NatureDoc then
       begin
       SetField('CPG_ETATNATURE',      Operation.NatureDoc ) ;
       SetField('CPG_MODELEENCADECA',  '' ) ;
       end ;
    end
  else
    // Pour les états
    if GetField('CPG_ETATNATURE') <> Operation.NatureEtat then
       begin
       SetField('CPG_ETATNATURE',  Operation.NatureEtat ) ;
       SetField('CPG_MODELEENCADECA',  '' ) ;
       end ;

  if VH^.PaysLocalisation<>CodeISOES then //XVI 24/02/205
  Begin
    // Paramètrage avec édition
    if GetField('CPG_EDITE') <> 'X' then
      if GetField('CPG_MODELEENCADECA') <> '' then
        SetField('CPG_MODELEENCADECA',  '' ) ;
  End ;

  // MAJ choix de l'imprimante accessible unquement avec les documents
  if GetField('CPG_AVECDOC')='-' then
    SetField('CPG_MODELECHX', '-' ) ;

  // Avertissement // EN ATTENTE
{  if (lStOldModele <> '') and (lStOldModele<>GetField('CPG_MODELEENCADECA')) then
    begin
    AfficheMessage(CPG_WNGMODIFCASCADE) ;
    end ;
}

end;

procedure TOM_CPARAMGENER.majEcranEtat;
begin

  // Paramètrage type d'état utilisé
  if GetField('CPG_AVECDOC')<>'-' then
    begin
    // MAj de la liste des documents
    if Operation.NatureDoc <> ''
       then SetControlProperty('CPG_ModeleEncaDeca', 'Plus', Operation.NatureDoc )
       else SetControlProperty('CPG_ModeleEncaDeca', 'Plus', W_W ) ;
    // MAJ du libellé
    SetControlText( 'TCPG_MODELEENCADECA', 'Modèle de document' ) ;
    end
  else
    begin
    // MAj de la liste des états
    if Operation.NatureEtat<>''
       then SetControlProperty('CPG_ModeleEncaDeca', 'Plus', Operation.NatureEtat )
       else SetControlProperty('CPG_ModeleEncaDeca', 'Plus', W_W ) ;
    // MAJ du libellé
    SetControlText( 'TCPG_MODELEENCADECA', 'Modèle d''état' ) ;
    end ;

  // MAJ choix de l'imprimante accessible unquement avec les documents
  SetControlEnabled( 'CPG_MODELECHX', GetField('CPG_AVECDOC')='X' ) ;

  THValComboBox(GetControl('CPG_ModeleEncaDeca')).Reload ;

  if VH^.PaysLocalisation<>CodeISOES then //XVI 24/02/2005
  begin
     // Paramètrage avec édition
     SetControlEnabled('BPARAMETAT',   V_PGI.Superviseur ) ;
     EnableControls( TGroupBox( GetControl('INFOSEDITION', True) ),
                     GetField('CPG_EDITE')='X',
                     False ) ;
     // SBO 30/07/2004 Correction fiche FQ 13741 : Bizzarement des @ viennent s'intercaler dans le hint...???
     BParamEtat.Hint   := 'Paramétrage de l''état' ;
  end ; 
end;

procedure TOM_CPARAMGENER.majEcranGeneraux;
var lBoCptFixe : Boolean ;
    stPlus     : String ;
begin

  lBoCptFixe := ( TobJournal.GetValue('J_NATUREJAL') = 'BQE' ) or
                ( TobJournal.GetValue('J_NATUREJAL') = 'CAI' ) or
                ( TobJournal.GetValue('J_EFFET') = 'X' ) ;

  // Compte de contrepartie obligatoire pour les journaux de banque
  SetControlEnabled('CPG_General', not lBoCptFixe ) ;

  if VH^.PaysLocalisation=COdeISOES then
    Begin
    if lBoCptFixe then
       StPlus:='CCB_GENERAL="'+GetField('CPG_GENERAL')+'"'
      else
       StPlus:='' ;
    SetControlProperty('CPG_CONDREMISE','Plus',StPlus) ;
    End ; //XVI 24/02/2005

  // Restriction à la liste des comptes
  if lBoCptFixe then
     SetControlProperty('CPG_General', 'Datatype', 'TZGENERAL' )
  else if (VH^.PaysLocalisation<>CodeISOES) and (TobJournal.GetValue('J_NATUREJAL') = 'OD') then //XVI 24/02/2005
     begin
     if GetField('CPG_FLUXENCADECA') = 'ENC'
       then SetControlProperty('CPG_General', 'Datatype', 'TZGENCAIS' )
       else SetControlProperty('CPG_General', 'Datatype', 'TZGDECAIS' ) ;
     end
  else
     SetControlProperty('CPG_General', 'Datatype', Operation.AccesGeneraux ) ;

end;

procedure TOM_CPARAMGENER.majChampsGeneraux;
var lBoCptFixe : Boolean ;
begin

  if DS.State in [dsBrowse] then exit;

  lBoCptFixe := ( TobJournal.GetValue('J_NATUREJAL') = 'BQE' ) or
                ( TobJournal.GetValue('J_NATUREJAL') = 'CAI' ) or
                ( TobJournal.GetValue('J_EFFET') = 'X' ) ;

  // Mise en place du compte de contrepartie pour les journaux de type BQE ou CAI
  If lBoCptFixe
    then SetField('CPG_GENERAL', TobJournal.GetValue('J_CONTREPARTIE') )
  // Si un compte est renseigné, on vérifie sa validité par rapport à la tablette utilisé :
    else if GetField('CPG_General')<>'' then
            if not LookUpValueExist( GetControl('CPG_General',True) )
               then SetField('CPG_GENERAL', '') ;

end;

procedure TOM_CPARAMGENER.majEcranEdite;
begin
  SetControlEnabled('BPARAMETAT', V_PGI.Superviseur ) ;
  EnableControls( TGroupBox( GetControl('INFOSEDITION', True) ),
                  GetField('CPG_EDITE')='X',
                  False ) ;
  // SBO 30/07/2004 Correction fiche FQ 13741 : Bizzarement des @ viennent s'intercaler dans le hint...???
  BParamEtat.Hint   := 'Paramétrage de l''état' ;

end;

procedure TOM_CPARAMGENER.MajChampsEdite;
begin
  if DS.State in [dsBrowse] then exit;
  if GetField('CPG_EDITE')='X' then
     begin
     SetField('CPG_AVECDOC',         iif(operation.NatureDoc<>'','X','-') ) ;
     SetField('CPG_ETATNATURE',      '' ) ;
     SetField('CPG_MODELEENCADECA',  '' ) ;
     end
  else
    Begin
    SetField('CPG_AVECDOC',         '-' ) ;
    SetField('CPG_ETATNATURE',      ''  ) ;
    SetField('CPG_MODELEENCADECA',  ''  ) ;
    SetField('CPG_CODEBARRE',       '-' ) ;
    SetField('CPG_MODELECHX',       '-' ) ;
    End ;
End ;

procedure TOM_CPARAMGENER.majEcranCFONB;
begin
  // Globalisation des tiers accessible uniquement sans export ou bordereau
//  majAccesTiersGlobalise ;

  if (VH^.PaysLocalisation = CodeISOES) then
    begin
    SetControlEnabled('BPARAMEXPORT', V_PGI.Superviseur ) ;
    EnableControls( TGroupBox( GetControl('INFOSCFONB') ),
                    GetField('CPG_CFONBEXPORT') = 'X',
                    False ) ;
    // SBO 30/07/2004 Correction fiche FQ 13741 : Bizzarement des @ viennent s'intercaler dans le hint...???
    BParamExport.Hint := 'Générateur de fichier' ;
    end
  else
    SetControlVisible('INFOSCFONB', False) ;
end;

procedure TOM_CPARAMGENER.majChampsCFONB;
begin
  if DS.State in [dsBrowse] then exit;

  if GetField('CPG_CFONBEXPORT') = 'X'
    then //SetField('CPG_TIERSGLOBALISE',  '-' )
    else SetField('CPG_CFONBFORMAT',     '' ) ;

end;

procedure TOM_CPARAMGENER.majEcranBOR;
begin
  // Globalisation des tiers accessible uniquement sans export ou bordereau
//  majAccesTiersGlobalise ;

  SetControlEnabled('BPARAMBOR',    V_PGI.Superviseur ) ;
  EnableControls( TGroupBox( GetControl('INFOSBORDEREAU') ),
                  GetField('CPG_BORDEREAUEXP')='X',
                  False ) ;
  // SBO 30/07/2004 Correction fiche FQ 13741 : Bizzarement des @ viennent s'intercaler dans le hint...???
  BParamBor.Hint    := 'Paramétrage de l''état' ;
end;

procedure TOM_CPARAMGENER.majChampsBOR;
begin
  if DS.State in [dsBrowse] then exit;

  if GetField('CPG_BORDEREAUEXP') = 'X' then
//    SetField('CPG_TIERSGLOBALISE', '-')
  else
    begin
    SetField('CPG_BORDEREAUMOD',  '' ) ;
    SetField('CPG_BORDEREAUCHX',  '-' ) ;
    end ;

end;


procedure TOM_CPARAMGENER.InitParPays;
var lBoEsp      : Boolean ;
begin
  lBoEsp := (VH^.PaysLocalisation=CodeISOES) ; //XVI 24/02/2005

  //Affichage des Conditions de remise
  SetControlVisible('CPG_CONDREMISE',  lBoEsp ) ;
  SetControlVisible('TCPG_CONDREMISE', lBoEsp ) ;

  //Affichage du compte général de sélection
  SetControlVisible('CPG_GENERALSEL',  lBoEsp ) ;
  SetControlVisible('TCPG_GENERALSEL', lBoEsp ) ;
  SetControlVisible('LIBELLESELECTION', lBoEsp ) ; //XVI 24/02/2005

  //En Version Espagnole, on n'a pas gestion de la pièce d'escompte...
  SetControlVisible('INFOSESCOMPTE',   not lBoEsp ) ;
  SetControlVisible('CPG_ESCMETHODE',  not lBoEsp ) ;
  SetControlVisible('TCPG_ESCMETHODE', not lBoEsp ) ;

  //En Version Espagnole, on peut paramétrer le Bordereau....
  SetControlVisible('BPARAMBOR',lBoEsp); //XMG 20/04/04

  // SPECIF ESPAGNOL
  if lBoESP then //XVI 24/02/2005
    begin
    // Mode de paiement
    with THLabel(GetControl('TCPG_MODEPAIEMENT')) do
      Begin
      Visible:=FALSE ;
      SetControlProperty('CPG_FORCERMP','Left',Left) ;
      End ;
    ForcerAdjustSizeControl(GetControl('TCPG_MODEPAIEMENT')) ;
    with THDBValComboBox(GetControl('CPG_MODEPAIEMENT')) do
      Begin
      SetControlProperty('CPG_FORCERMP','TabOrder',TabOrder) ;
      Vide:=lBoEsp ;
      End ;
    SetControlProperty('CPG_FORCERMP','Caption','Forcer au mode de paiement') ;
    ForcerAdjustSizeControl(GetControl('CPG_FORCERMP')) ;
    SetControlProperty('CPG_MODEPAIEMENT','Datatype','TTMODEPAIE') ;
    end
  else
  // SPECIF FRANCAIS
    begin
    MajEcranFlux ;
    SetControlProperty('CPG_BORDEREAUMOD', 'DATATYPE', 'TTMODELEBOR') ;
    end ;

  // Corrections libellé champs multi-établissement FQ 17864
  SetControlProperty('TCPG_METHODEETAB','Caption','Méthode de génération ') ;
  SetControlProperty('TCPG_SELECTETAB','Caption','Echéances mouvementées sur ') ;

end;

procedure TOM_CPARAMGENER.completeGeneral ( F: TField ) ;
var lStCode : String ;
begin

  if DS.State in [dsBrowse] then exit;
  if F.AsString = '' then Exit ;

  lStCode := BourreLaDonc( F.AsString, fbGene ) ;
  if lStCode <> F.AsString
     then F.AsString := lStCode ;

end;

procedure TOM_CPARAMGENER.BParamExportClick(Sender: TObject);
var
  lStOldEtat : String;
begin

  // Pour forcer la maj de la sélection de l'état,
  //  ( mal mis à jour si clique direct sur bouton de parmétrage ... )
  NextPrevControl(Ecran) ;

  if not ExJaiLeDroitConcept(ccParamEtat, True) then Exit;

  // Edition Etat / Document
  lStOldEtat := GetField('CPG_CFONBFORMAT') ;

  Param_LPTexte( nil, 'F', 'FIC', lStOldEtat, True ) ;

  // Rechargement de la Combo
  THValComboBox(GetControl('CPG_CFONBFORMAT')).ReLoad ;
  if lStOldEtat<>'' then
     SetControlText('CPG_CFONBFORMAT', lStOldEtat ) ;

end;

procedure TOM_CPARAMGENER.majChampsBOREXP;
begin
  if DS.State in [dsBrowse] then exit;

  if GetField('CPG_COMPTABILISE')<>'X' then
    begin
    SetField( 'CPG_CFONBEXPORT',  '-' ) ;
    SetField( 'CPG_BORDEREAUEXP', '-' ) ;
    end ;

end;

procedure TOM_CPARAMGENER.majEcranBOREXP;
var lBoTest : Boolean ;
begin
  lBoTest := GetField('CPG_COMPTABILISE')='X' ;
  // Coche "avec bordereau"
  SetControlEnabled('CPG_BORDEREAUEXP', (Operation.AccesBordereau = taPossible) and lBoTest );
  // Coche "avec export CFONB"
  SetControlEnabled('CPG_CFONBEXPORT',  (Operation.AccesCFONB = taPossible) and lBoTest )
end;

procedure TOM_CPARAMGENER.BParamBorClick(Sender: TObject);
  {$IFDEF EAGLCLIENT}
  {$ELSE}
var
  lStOldEtat : String;
  {$ENDIF}
begin

  {$IFDEF EAGLCLIENT}
  {$ELSE}

  // Pour forcer la maj de la sélection de l'état,
  //  ( mal mis à jour si clique direct sur bouton de parmétrage ... )
  NextPrevControl(Ecran) ;

  if not ExJaiLeDroitConcept(ccParamEtat, True) then Exit;

  // Edition Etat / Document
  lStOldEtat := GetField('CPG_BORDEREAUMOD') ;
  EditEtat('E', 'BOR', lStOldEtat, TRUE, nil, '', Ecran.Caption);

  // Rechargement de la Combo
  THValComboBox(GetControl('CPG_BORDEREAUMOD')).ReLoad ;
  if lStOldEtat<>'' then
     SetControlText('CPG_BORDEREAUMOD', lStOldEtat ) ;

  {$ENDIF}

end;

procedure TOM_CPARAMGENER.majChampsMethGener;
var lBoCptColl : Boolean ;
begin

  if DS.State in [dsBrowse] then exit;

  lBoCptColl := Assigned(TobGeneral) and ( TobGeneral.GetValue('G_COLLECTIF') = 'X' ) ;

  if lBoCptColl and ( ( GetField('CPG_GROUPEENCADECA') = 'GLO' )
                   or ( GetField('CPG_GROUPEENCADECA') = 'ECH' ) )
    then SetField('CPG_GroupeEncaDeca','') ;

end;

procedure TOM_CPARAMGENER.majEcranMethGener;
var lBoCptColl : Boolean ;
begin
  lBoCptColl := Assigned(TobGeneral) and ( TobGeneral.GetValue('G_COLLECTIF') = 'X' ) ;

  if lBoCptColl
    then SetControlProperty('CPG_GroupeEncaDeca','Plus', ' AND CO_CODE IN ("DET","AUX","ECH") ')
    else SetControlProperty('CPG_GroupeEncaDeca','Plus', '' ) ;

end;

procedure TOM_CPARAMGENER.ChargeGeneral;
var Q : TQuery ;
begin

  // Pour éviter les chargements multiples
  if Assigned(TobGeneral) and ( GetField('CPG_General') = TobGeneral.GetValue('G_General'))
   then Exit ;

  // Init Tob
  FreeAndNil( TobGeneral ) ;
  TobGeneral := TOB.Create('GENERAUX', nil, -1 ) ;

  // Chargement des infos utiles du journal dans la tob
  if length( Trim( GetField('CPG_General') ) ) > 0 then
    begin
    // Recherche nature et compte de contrepartie
    Q := OpenSQL('SELECT G_GENERAL, G_LIBELLE, G_COLLECTIF, G_NATUREGENE FROM GENERAUX WHERE G_GENERAL="'
                  + GetField('CPG_General') + '" ' , TRUE ) ;
    TobGeneral.SelectDB( '', Q ) ;
    Ferme(Q) ;
    end ;

end;

procedure TOM_CPARAMGENER.majChampsGlobTiers;
begin
  if DS.State in [dsBrowse] then exit;

  if GetField('CPG_GROUPEENCADECA') = 'DET' then
    begin
    if ( GetField('CPG_TIERSGLOBALISE')='X' ) or
       ( GetField('CPG_MULTISOC')='X' ) then
       begin
       PGIInfo( 'Attention ! En mode "génération au détail de la contrepartie", '
              + 'le paramétrage multisociété et la globalisation des échéances tiers ne seront pas pris en compte', Ecran.Caption ) ;
       end ;
    SetField('CPG_TIERSGLOBALISE', '-') ;
    SetField('CPG_MULTISOCMETH',   'SOC') ;
    end ;

end;

procedure TOM_CPARAMGENER.majEcranGlobTiers;
begin
  majAccesTiersGlobalise ;
//  SetControlEnabled('CPG_TIERSGLOBALISE', GetField('CPG_GROUPEENCADECA') <> 'DET' ) ;
  SetControlEnabled('CPG_MULTISOCMETH',   GetField('CPG_GROUPEENCADECA') <> 'DET' ) ;
end;

procedure TOM_CPARAMGENER.MajEcranFlux;
begin
  //on ne prend que les type d'enca/Déca que corresponent....
  SetControlProperty('CPG_TYPEENCADECA','plus',ConditionTypeEncadeca(GetField('CPG_FLUXENCADECA'))) ;
end;

procedure TOM_CPARAMGENER.MajChampsListeBOR;
begin
  if DS.State in [dsBrowse] then exit;

  if (Getfield('CPG_BORNATURE')='') then //XVI 24/02/2005
    setField('CPG_BORNATURE','BOR') ;

//  SetField('CPG_BORDEREAUMOD', '' ) ;
end;

procedure TOM_CPARAMGENER.MajEcranListeBOR;
begin
  if (VH^.PaysLocalisation=CodeISOES) then
    SetControlProperty('CPG_BORDEREAUMOD', 'Plus', 'BOR') ;
end;

procedure TOM_CPARAMGENER.MajChampsMPForcer;
begin
  if DS.State in [dsBrowse] then exit;
  if GetField('CPG_MODEPAIEMENT')='' then
    SetField('CPG_FORCERMP', '-' ) ;
end;

procedure TOM_CPARAMGENER.MajEcranMPForcer;
begin
  SetControlEnabled('CPG_FORCERMP', GetField('CPG_MODEPAIEMENT')<>'' ) ;
end;

procedure TOM_CPARAMGENER.MajEcranMultiSoc;
begin
  SetControlEnabled('CPG_MULTISOCMETH', GetField('CPG_MULTISOC')='X' ) ;
  SetControlEnabled('CPG_MULTISOCJAL',  GetField('CPG_MULTISOC')='X' ) ;
end;

procedure TOM_CPARAMGENER.MajChampsMultiSoc;
begin
  if DS.State in [dsBrowse] then exit;
  if GetField('CPG_MULTISOC')<>'X' then
    begin
    SetField('CPG_MULTISOCJAL',  '' ) ;
    SetField('CPG_MULTISOCMETH', '' ) ;
    end ;

end;

procedure TOM_CPARAMGENER.MajChampsMultiEtab;
begin
  if DS.State in [dsBrowse] then exit;

  if GetField('CPG_MULTIETAB')<>'X' then
    begin
    SetField('CPG_METHODEETAB', '' ) ;
    SetField('CPG_SELECTETAB', '' ) ;
    end ;

end;

procedure TOM_CPARAMGENER.MajEcranMultiEtab;
begin
  SetControlEnabled('CPG_SELECTETAB',  GetField('CPG_MULTIETAB')='X' ) ;
  MajEcranMethodeMultiEtab ;
end;

procedure TOM_CPARAMGENER.majAccesTiersGlobalise;
begin
  SetControlEnabled('CPG_TIERSGLOBALISE', ( GetField('CPG_GROUPEENCADECA') <> 'DET' ) ) ;
{
  SetControlEnabled('CPG_TIERSGLOBALISE', ( GetField('CPG_CFONBEXPORT')<>'X' )
                                            and ( GetField('CPG_BORDEREAUEXP')<>'X' )
                                            and ( GetField('CPG_GROUPEENCADECA') <> 'DET' ) ) ;
}
end;

procedure TOM_CPARAMGENER.BAssistClick(Sender: TObject);
var lStFormule : string ;
    lStValC    : string ;
    lControl   : THEdit  ;
begin

  inherited;

  lControl   := THEdit(Sender) ;
  lStFormule := ChoixChampZone( 0, 'LIB') ;

  if lStFormule = '' then Exit;
  lStValC := lControl.Text;

  if Length(lStValC + ' ' + lStFormule) > 100 then
    PgiInfo('La formule saisie est trop longue. Tout ne sera pas retenu,vous devez la recomposer.', Ecran.Caption ) ;

  if DS.State=dsBrowse then DS.Edit ;

  if lControl.SelLength > 0 then
    Delete( lStValC, lControl.SelStart + 1, lControl.SelLength ) ;
  if lStValC = ''
    then lStValC := lStFormule
    else lStValC := lStValC + ' ' + lStFormule ;

  lControl.Text := lStValC ;

end;

procedure TOM_CPARAMGENER.MajEcranMethodeMultiEtab;
begin
  SetControlEnabled('CPG_METHODEETAB',  ( GetField('CPG_MULTIETAB')='X' ) and ( GetField('CPG_GROUPEENCADECA')<>'DET' ) ) ;
end;

procedure TOM_CPARAMGENER.MajChampsMethodeMultiEtab;
begin
  if DS.State in [dsBrowse] then exit;

  if GetField('CPG_GROUPEENCADECA')='DET' then
    begin
    SetField('CPG_METHODEETAB', 'DET' ) ;
    end ;

end;

Initialization
  registerclasses ( [ TOM_CPARAMGENER ] ) ;
end.
