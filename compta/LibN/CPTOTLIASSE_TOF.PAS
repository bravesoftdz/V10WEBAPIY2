{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 09/10/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPTOTLIASSE ()
Mots clefs ... : TOF;CPTOTLIASSE
*****************************************************************}
Unit CPTOTLIASSE_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
     Windows,  // VK_
{$IFDEF EAGLCLIENT}
     eMul,
     MaineAGL, // AGLLanceFiche
{$ELSE}
     FE_Main,  // AGLLanceFiche
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     HStatus,  // FiniMove
     ED_TOOLS, // InitMoveProgressForm
     HXlsPas,  // ExportGrid
     Dialogs,  //TSaveDialog;
     Menus,    // TPopUPMenu
     HSysMenu, // THSystemMenu
     ZCumul,   // TZCumul
     ParamSoc, // GetParamSocSecur
     uLibWindows,
     uTob,
     uTobDebug,
     UTOF ;

Type
  TOF_CPTOTLIASSE = Class (TOF)

    Pages   : TPageControl;
    Ctrl0   : THValComboBox;
    EXERCICE: THValComboBox;

    Etablissement : THValComboBox;
    Devise        : THValComboBox;
    QualifPiece   : THMultiValComboBox;
    CbTousAgregat : TCheckBox;

    Ctrl1   : THEdit;
    Ctrl1_  : THEdit;
    Date    : THEdit;
    Date_   : THEdit;

    FListe1 : THGrid;
    FListe2 : THGrid;

    BCherche   : TToolBarButton97;
    BDetail    : TToolBarButton97;
    BStop      : TToolBarButton97;
    BExport    : TToolBarButton97;
    BCptLiasse : TToolBarButton97;
    BRechercher: TToolBarButton97;

    PopZoom : TPopUpMenu;
    BGen    : TMenuItem;
    BRub    : TMenuItem;

    procedure OnArgument (S : String ) ; override ;
    procedure OnLoad                   ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;

    procedure OnChangeExercice  (Sender : TObject);
    procedure OnClickBCherche   (Sender : TObject);
    procedure OnClickBDetail    (Sender : TObject);
    procedure OnClickBStop      (Sender : TObject);
    procedure OnClickBExport    (Sender : TObject);
    procedure OnClickBImprimer  (Sender : TObject);
    procedure OnClickBRub       (Sender : TObject);
    procedure OnClickBGen       (Sender : TObject);
    procedure OnClickBCpt       (Sender : TObject);

    procedure OnChangeCtrl0     (Sender : TObject);
    procedure OnPopUpPopZoom    (Sender : TObject);
    procedure OnKeyDownEcran    (Sender : TObject; var Key : Word; Shift : TShiftState); virtual;
    procedure OnRowEnterFListe1 (Sender : TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure OnRowEnterFListe2 (Sender : TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);

    procedure OnClickBRechercher(Sender : Tobject);
    procedure OnFindFFindDialog (Sender : TObject);

  private

    FBoLoading : Boolean;
    FBoStop    : Boolean;

    FStArgument      : string;
    FStLiasseDossier : string;
    FStWhereRub      : string;
    FTobRub          : Tob;
    FZCumul          : TZCumul;
    FormKeyDown      : TKeyEvent;
    FFindFirst       : Boolean;
    FFindDialog      : TFindDialog;

    procedure InitFListe;
    procedure RemplitCriteres;
    procedure RemplitFourchettesCtrl1;

    function TraiteInfoCompte( vTobMere : Tob; vSt : string) : Double;

    procedure AfficheDetailRubrique;

    procedure OnAfterFormShow;
    procedure OnAfterSelectFiltre;

  end ;

procedure CPLanceFiche_CPTOTLiasse( vStArgument : string );

//procedure SelectFromRubrique( vTobRub: Tob vDateValidite : TDateTime );

Implementation

uses Ent1,             // VH
     RUBRIQUE_TOM,     // ParametrageRubrique
     CPGENERAUX_TOM,   // FicheGene
     CPTotalCPTLiasse_TOF, // CPLanceFiche_CPTotalCptLiasse
     CPStatutDossier_TOF, // SupprimeDoublonAgregat(FTobRub);
     uLibExercice;     // CExoRelatifToDates

const cColAgregat   = 1;
      cColLibelle   = 2;
      cColComptable = 3;
      cColFiscal    = 4;
      cColEcart     = 5;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/11/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
(*
procedure SelectFromRubrique( vStSql : string; vDateValidite : TDateTime ) : Tob;
var lTobRub : Tob;
    lTobTemp : Tob;
    lStSql : string;
    i : integer;
    lStRub : string;
begin
  try
    lTobRub := Tob.Create('LESRUBRIQUES', nil, -1);
    Result  := Tob.Create('LESRUBRIQUES', nil, -1);

    lStSql := vStSql + ' AND RB_DATEVALIDITE >= "'+ UsDateTime(vDateValidite) +
              ' ORDER BY RB_RUBRIQUE, RB_DATEVALIDITE ASC"';

    lTobRub.LoadDetailFromSQL( vStSql );

    lStRub := '';
    for i := 0 to lTobRub.Detail.Count -1 do
    begin
      lTobTemp := lTobRub.Detail[i];

      if lStRub <> lTobTemp.GetString('RB_RUBRIQUE') then
      begin

         
      end;




      lStRub := lTobTemp.GetString('RB_RUBRIQUE')




    end;

  finally
    FreeAndNil(lTobRub);

  end;
end;*)

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/03/2005
Modifié le ... :   /  /
Description .. : Procédure de lancement du Controle des Rubriques
Suite ........ : RUBRIQUE; BUDGET; CYCLE; LIASSE
Mots clefs ... :
*****************************************************************}
procedure CPLanceFiche_CPTOTLiasse( vStArgument : string );
begin
  AGLLanceFiche('CP', 'CPTOTLIASSE', '', '', vStArgument);
end ;



////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTOTLIASSE.OnArgument (S : String ) ;
begin
  Inherited ;
  FStArgument := S;

  // Récupérations des composants
  Pages    := TPageControl(GetControl('PAGES', True));
  FListe1  := THGrid(GetControl('FLISTE1', True));
  FListe2  := THGrid(GetControl('FLISTE2', True));

  Ctrl0    := THValComboBox(GetControl('CTRL0', True));
  Ctrl1    := THEdit(GetControl('CTRL1', True));
  Ctrl1_   := THEdit(GetControl('CTRL1_', True));

  EXERCICE := THValComboBox(GetControl('EXERCICE', True));
  CInitComboExercice(EXERCICE);

  Etablissement := THValComboBox(GetControl('ETABLISSEMENT', True));
  Devise        := THValComboBox(GetControl('DEVISE', True));
  QualifPiece   := THMultiValComboBox(GetControl('QUALIFPIECE', True));
  CbTousAgregat := TCheckBox(GetControl('CBTOUSAGREGAT', True));

  DATE        := THEdit(GetControl('DATE', True));
  DATE_       := THEdit(GetControl('DATE_', True));

  Bcherche    := TToolBarButton97(GetControl('BCHERCHE', True));
  BDetail     := TToolBarButton97(GetControl('BDETAIL', True));
  BStop       := TToolBarButton97(GetControl('BSTOP', True));
  BExport     := TToolBarButton97(GetControl('BEXPORT', True));
  BCptLiasse  := TToolBarButton97(GetControl('BCPTLIASSE', True));
  BRechercher := TToolBarButton97(GetControl('BRECHERCHER', True));

  PopZoom  := TPopUpMenu(GetControl('POPZOOM', True));
  BGen     := TPopupMenu(GetControl('POPZOOM')).Items[0];
  BRub     := TPopupMenu(GetControl('POPZOOM')).Items[1];

  // GCO - 30/11/2006 - FQ 19248
  FFindDialog := TFindDialog.Create(Ecran);
  FFindDialog.OnFind := OnFindFFindDialog;

  // Branchement des événements
  FormKeyDown       := Ecran.OnKeyDown;
  Ecran.OnKeyDown   := OnKeyDownEcran;
  Ctrl0.OnChange    := OnChangeCtrl0;
  Exercice.OnChange := OnChangeExercice;

  FListe1.OnRowEnter := OnRowEnterFListe1;
  FListe1.OnDblClick := OnClickBRub;

  FListe2.OnRowEnter := OnRowEnterFListe2;
  FListe2.OnDblClick := OnClickBGen;

  TToolbarButton97(GetControl('BIMPRIMER' )).OnClick := OnClickBImprimer;
  BCherche.OnClick    := OnClickBCherche;
  BDetail.OnClick     := OnClickBDetail;
  BStop.OnClick       := OnClickBStop;
  BExport.OnClick     := OnClickBExport;
  PopZoom.OnPopup     := OnPopUpPopZoom;
  BRub.OnClick        := OnClickBRub;
  BGen.OnClick        := OnCLickBGen;
  BCptLiasse.OnClick  := OnClickBCpt;
  BRechercher.OnClick := OnClickBRechercher;

  InitFListe;

  if CtxPCl in V_Pgi.PgiContexte then
  begin
    if VH^.CPExoRef.Code <> '' then
      Exercice.Value := CExerciceVersRelatif(VH^.CPExoRef.Code )
    else
      Exercice.Value := CExerciceVersRelatif(VH^.EnCours.Code );
  end
  else
    Exercice.Value := CExerciceVersRelatif(VH^.Entree.Code);

  FStLiasseDossier := GetParamSocSecur('SO_CPCONTROLELIASSE', '');
  if FStLiasseDossier <> '' then
    Ctrl0.Value := FStLiasseDossier
  else
    Ctrl0.ItemIndex := 0;

  Etablissement.ItemIndex := 0;
  Devise.ItemIndex := 0;

  BStop.Visible := False;
  BCptLiasse.Enabled := False;

  QualifPiece.SelectAll;
  QualifPiece.Text := TraduireMemoire('<<Tous>>');

  FTobRub := Tob.Create('', nil, -1);
  FZCumul := TZCumul.Create;
  FZCumul.ModeFonc := mfMemory;
  FZCumul.AvecLibelle := True;

  TFMul(Ecran).OnAfterFormShow     := OnAfterFormShow;
  TFMul(Ecran).OnAfterSelectFiltre := OnAfterSelectFiltre;

{$IFDEF EAGLCLIENT}
  SetControlVisible('BPREV', False);
  SetControlVisible('BNEXT', False);
{$ENDIF}

  FBoLoading := True;

  FStWhereRub := 'RB_NATRUB = "CPT" AND ' +
                 'RB_CLASSERUB = "CPT" AND ' +
                 'RB_TYPERUB = "GEN" AND ' +
                 '(RB_PREDEFINI <> "DOS" OR ' +
                 '(RB_PREDEFINI = "DOS" AND RB_NODOSSIER = "' + V_PGI.NoDossier + '"))';

end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTOTLIASSE.OnAfterFormShow;
var lStTemp : string;
begin
  FBoLoading := False;

  lStTemp := ReadTokenSt(FStArgument);
  if lStTemp <> '' then
  begin
    CTRL1.Text            := lStTemp;
    CTRL1_.Text           := CTRL1.Text;
    CbTousAgregat.Checked := True;
    BCherche.Click;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTOTLIASSE.OnLoad ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTOTLIASSE.OnDisplay () ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/09/2006
Modifié le ... : 02/10/2006
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTOTLIASSE.InitFListe;
begin
  FListe1.ColAligns[cColAgregat]   := TaCenter;
  FListe1.ColAligns[cColLibelle]   := TaLeftJustify;
  FListe1.ColAligns[cColComptable] := TaRightJustify;
  FListe1.ColAligns[cColFiscal]    := TaRightJustify;
  FListe1.ColTypes[cColEcart]      := 'B';
  FListe1.ColAligns[cColEcart]     := TaCenter;
  FListe1.ColFormats[cColEcart]    := IntToStr( Byte( csCheckBox ) );

  FListe1.RowCount := 2;

  FListe2.ColAligns[cColAgregat]   := TaCenter;
  FListe2.ColAligns[cColLibelle]   := TaLeftJustify;
  FListe2.ColAligns[cColComptable] := TaRightJustify;
  FListe2.ColAligns[cColFiscal]    := TaRightJustify;

  FListe2.ColTypes[cColEcart]      := 'B';
  FListe2.ColAligns[cColEcart]     := TaCenter;
  FListe2.ColFormats[cColEcart]     := IntToStr( Byte( csCheckBox ) );

  FListe2.RowCount := 2;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/09/2006
Modifié le ... : 02/10/2006
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTOTLIASSE.OnChangeCtrl0(Sender: TObject);
begin
  CTRL1.Plus  := ' RB_FAMILLES LIKE "' + Ctrl0.Value + '%" AND ' + FStWhereRub;
  CTRL1_.Plus := CTRL1.Plus;
  RemplitFourchettesCtrl1;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTOTLIASSE.RemplitCriteres;
begin
  if not TestJoker(Ctrl1.Text) then
  begin
    if (CTrl1.Text = '') and (Ctrl1_.Text = '') then
      RemplitFourchettesCtrl1
    else
    begin
      if (CTrl1.Text <> '') and (Ctrl1_.Text = '') then
        Ctrl1_.Text := Ctrl1.Text;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/10/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTOTLIASSE.RemplitFourchettesCtrl1;
var lQuery : TQuery;
begin
  try
    lQuery := OpenSQL('SELECT MIN(RB_RUBRIQUE) MIN, MAX(RB_RUBRIQUE) MAX ' +
                      'FROM RUBRIQUE WHERE ' +
                      'RB_FAMILLES LIKE "' + Ctrl0.Value + '%" AND ' + FStWhereRub,True);

    Ctrl1.Text  := lQuery.FindField('MIN').AsString;
    Ctrl1_.Text := lQuery.FindField('MAX').AsString;
  finally
    Ferme(lQuery);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTOTLIASSE.OnChangeExercice(Sender: TObject);
begin
  CExoRelatifToDates(Exercice.Value, Date, Date_);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/10/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTOTLIASSE.OnClickBCherche(Sender: TObject);
var i,j : integer;
    lTobTemp : Tob;
    LesCptes : TStringList;
    lSoldeComptable : Double;
    lSoldeFiscal    : Double;
    lDbEcart        : Double;
    lStEtab         : string;
    lStDevise       : string;
    lStQualifPiece  : string;
    lValZCumul      : Variant;
begin
  inherited;
  if FBoLoading then Exit;

  BStop.Visible := True;
  FBoStop := False;
  RemplitCriteres;
  Fliste1.BeginUpdate;
  LesCptes := TStringList.Create;
  try
    FListe1.VidePile(False);
    FListe2.VidePile(False);

    SelectAgregatDictionnaire('RB_RUBRIQUE, RB_NODOSSIER, RB_PREDEFINI, ' +
                              'RB_DATEVALIDITE, RB_LIBELLE, RB_SIGNERUB, ' +
                              '0 TOTALCOMPTABLE, 0 TOTALFISCAL, 0 ECART ',
                              ConvertitCaractereJokers(Ctrl1, Ctrl1_, 'RB_RUBRIQUE'),
                              StrToDate(Date_.Text), Ctrl0.Value, FTobRub);

    lStEtab := '';
    lStDevise := '';
    lStQualifPiece := '';
    lValZCumul := 0;

    if Etablissement.ItemIndex <> 0 then
      lStEtab := Etablissement.Value;

    if Devise.ItemIndex <> 0 then
      lStDevise := Devise.Value;

    if not QualifPiece.Tous then
      lStQualifPiece := Qualifpiece.Text;

    FZCumul.InitCriteres(lStqualifPiece, lStEtab , lStDevise , '('+Date.Text+')('+ Date_.Text+')', '');

    for i := 0 to FTobRub.DeTail.Count - 1 do
    begin
      Application.ProcessMessages;
      if FBoStop then
      begin
        if PgiAsk('Confirmez-vous l''arrêt du traitement en cours ?', Ecran.Caption) = MrYes then
          Break // Arrêt demandé par l'utilisateur
        else
          FBoStop := False;
        Application.ProcessMessages;
      end;

      lSoldeComptable := 0;
      lSoldeFiscal := 0;
      LesCptes.Clear;
      lTobTemp := FTobRub.Detail[i];

      lValZCumul := FZCumul.GetValeur('RUBRIQUE', lTobTemp.GetString('RB_RUBRIQUE'), LesCptes);
      if FZCumul.LastError = 0 then
      begin
        lSoldeFiscal := lValZCumul;
        lTobTemp.SetDouble('TOTALFISCAL', lSoldeFiscal);
      end
      else
      begin
        // Liasse en erreur
        PgiInfo('la liasse ' + lTobTemp.GetString('RB_RUBRIQUE') + ' a provoqué une erreur de calcul.', Ecran.Caption);
        Continue;
      end;

      // Parcours des informations des comptes
      for j := 0 to LesCptes.Count - 1 do
      begin
        lSoldeComptable := lSoldeComptable + TraiteInfoCompte( lTobTemp, LesCptes[j] );
      end;

      if lSoldeComptable >= 0 then
        lTobTemp.AddChampsUpvaleur('TOTALCOMPTABLE', StrFMontant(lSoldeComptable, 13, V_Pgi.OkDecV, '', True) + ' D')
      else
        lTobTemp.AddChampsUpvaleur('TOTALCOMPTABLE', StrFMontant(Abs(lSoldeComptable), 13, V_Pgi.OkDecV, '', True) + ' C');

      lTobTemp.AddChampSupValeur('TOTALFISCAL', StrFMontant(lSoldeFiscal, 13, V_Pgi.OkDecV, '', True));

      if lTobTemp.GetString('RB_SIGNERUB') = 'POS' then
        lDbEcart := lSoldeComptable - lSoldeFiscal
      else
        lDbEcart := lSoldeComptable - (-lSoldeFiscal);

      lDbEcart := Arrondi( lDbEcart, 2);
      lTobTemp.AddChampSupValeur('ECART', IIF(lDBEcart <> 0, 'X', '-'));
    end;

    if FTobRub.Detail.Count = 0 then
      FListe1.RowCount := 2
    else
    begin
      for i := 0 to FTobRub.DeTail.Count - 1 do
      begin
        lTobTemp := FTobRub.Detail[i];
        // Affichage
        if (CbTousAgregat.Checked) or
           ((not CbTousAgregat.Checked) and (lTobTemp.GetString('ECART')<> '-')) then
        begin
          FListe1.Cells[cColAgregat,   FListe1.RowCount-1] := lTobTemp.GetString('RB_RUBRIQUE');
          FListe1.Cells[cColLibelle,   FListe1.RowCount-1] := lTobTemp.GetString('RB_LIBELLE');
          FListe1.Cells[cColComptable, FListe1.RowCount-1] := lTobTemp.GetString('TOTALCOMPTABLE');
          FListe1.Cells[cColFiscal,    FListe1.RowCount-1] := lTobTemp.GetString('TOTALFISCAL');
          FListe1.Cells[cColEcart,     FListe1.RowCount-1] := lTobTemp.GetString('ECART');
          FListe1.RowCount := FListe1.RowCount + 1;
        end;
      end;
      if FListe1.RowCount > 2 then
        FListe1.RowCount := FListe1.RowCount - 1;
    end;

  finally
    BStop.Visible := False;
    LesCptes.Free;
    FListe1.Row := 1;
    AfficheDetailRubrique;
    THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(FListe1);
    FListe1.EndUpdate ;
    if FListe1.CanFocus then FListe1.SetFocus;
    TFMUL(Ecran).CritModified := False;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPTOTLIASSE.TraiteInfoCompte(vTobMere: Tob; vSt: string): Double;
var lStTemp    : string;
    lStGeneral : string;
    lStLibelle : string;
    lDbSM    : Double;
    lDbSD    : Double;
    lDbSC    : Double;
    lDbTD    : Double;
    lDbTC    : Double;
    lDbEcart : Double;

    lInCodification : integer;

    lTobFille : Tob;

begin
  lStTemp := vSt;
  lStGeneral := ReadTokenPipe(lStTemp, ':');
  lDbSM := Valeur( ReadTokenPipe(lStTemp, ':') );           // Sens Mixte
  lDbSC := Valeur( ReadTokenPipe(lStTemp, ':') );           // Sens Créditeur
  lDbSD := Valeur( ReadTokenPipe(lStTemp, ':') );           // Sens Débiteur
  lDbTC := Valeur( ReadTokenPipe(lStTemp, ':') );           // Total Crédit du Compte
  lDbTD := Valeur( ReadTokenPipe(lStTemp, ':') );           // Total Débit du Compte
  lInCodification := StrToInt(ReadTokenPipe(lStTemp, ':')); // Codification
  lStLibelle := Copy(lStTemp, 5, Length(lStTemp));

  if Arrondi(lDbTD - lDbTC,2) = 0 then
  begin
    Result := 0;
    Exit;
  end;

  lTobFille := Tob.Create('', vTObMere, -1);
  lTobFille.AddChampSupValeur('GENERAL', lStGeneral);
  lTobFille.AddChampSupValeur('LIBELLE', lStLibelle);

  if (lDbTD - lDbTC) >= 0 then
    lTobFille.AddChampSupValeur('TOTALCOMPTABLE', StrFMontant(lDbTD - lDbTC, 13, V_Pgi.OkDecV, '', True) + ' D')
  else
    lTobFille.AddChampSupValeur('TOTALCOMPTABLE', StrFMontant(Abs(lDbTD - lDbTC), 13, V_Pgi.OkDecV, '', True) + ' C');

  Result := lDbTD - lDbTC;

  lDbEcart := 0;
  case lInCodification of
    7 : begin
          lDbEcart := lDbSM;
          lTobFille.AddChampSupValeur('TOTALFISCAL', StrFMontant(lDbSM, 13, V_Pgi.OkDecV, '', True) + ' (SM)');
        end;

    6 : begin
          lDbEcart := lDbSC;
          lTobFille.AddChampSupValeur('TOTALFISCAL', StrFMontant(lDbSC, 13, V_Pgi.OkDecV, '', True) + ' (SC)');
        end;

    5 : begin
          lDbEcart := lDbSD;
          lTobFille.AddChampSupValeur('TOTALFISCAL', StrFMontant(lDbSD, 13, V_Pgi.OkDecV, '', True) + ' (SD)');
        end;

    3 : begin
          lDbEcart := lDbTC;
          lTobFille.AddChampSupValeur('TOTALFISCAL', StrFMontant(lDbTC, 13, V_Pgi.OkDecV, '', True) + ' (TC)');
        end;

    2 : begin
          lDbEcart := lDbTD;
          lTobFille.AddChampSupValeur('TOTALFISCAL', StrFMontant(lDbTD, 13, V_Pgi.OkDecV, '', True) + ' (TD)');
        end;
  end;

  if vTobMere.GetString('RB_SIGNERUB') = 'POS' then
    lDbEcart := lDbTD - lDbTC - lDBEcart
  else
    lDbEcart := lDbTD - lDbTC - (-lDBEcart);

  lDBEcart := Arrondi(lDBEcart, 2);
  lTobFille.AddChampSupValeur('ECART', IIF(lDBEcart <> 0, 'X', '-'));
end;


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTOTLIASSE.OnClickBDetail(Sender: TObject);
begin
  FListe2.Visible    := (not FListe2.Visible);
  
  if FListe2.Visible then
    THSystemMenu(GetControl('HMTRAD')).ResizeGridColumns(FListe2);

  AfficheDetailRubrique;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/10/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTOTLIASSE.OnClickBStop(Sender: TObject);
begin
  FBoStop := True;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/10/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTOTLIASSE.OnClickBExport(Sender: TObject);
var SD : TSaveDialog;
begin
  if not ExJaiLeDroitConcept(ccExportListe,True) then Exit;
  SD := TSaveDialog.Create(Ecran);
  SD.DefaultExt := 'XLS';
  SD.Filter := 'Fichier Texte (*.txt)|*.txt|Fichier Excel (*.xls)|*.xls|Fichier ' +
               'Ascii (*.asc)|*.asc|Fichier Lotus (*.wks)|*.wks|Fichier HTML (*.' +
               'html)|*.html|Fichier XML (*.xml)|*.xml';
  SD.Options := [ofOverwritePrompt, ofHideReadOnly, ofPathMustExist, ofNoReadOnlyReturn, ofNoLongNames];
  SD.Title   := 'Export';
  try
    if SD.Execute then
    begin
      if (FListe2.Visible) and (FListe2.Focused) then
        ExportGrid(FListe2, nil, SD.FileName, SD.FilterIndex, True)
      else
        ExportGrid(FListe1, nil, SD.FileName, SD.FilterIndex, True);
    end;
  finally
    SD.Free;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTOTLIASSE.OnClickBImprimer(Sender: TObject);
begin

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTOTLIASSE.OnClickBGen(Sender: TObject);
var lStTemp : string;
begin
  lStTemp := (FListe2.Cells[cColAgregat, FListe2.Row]);
  if lStTemp <> '' then
    FicheGene(nil, '', lStTemp, taConsult, 0);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/11/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTOTLIASSE.OnClickBCpt(Sender: TObject);
var lStTemp : string;
begin
  lStTemp := (FListe2.Cells[cColAgregat, FListe2.Row]);
  if lStTemp <> '' then
    CPLanceFiche_CPTotalCptLiasse(lStTemp + ';' + Date.Text + ';' + Date_.Text);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/10/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTOTLIASSE.OnClickBRub(Sender: TObject);
var lStTemp : string;
    lTobTemp : Tob;
begin
  lStTemp := (FListe1.Cells[cColAgregat, FListe1.Row]);
  if lStTemp <> '' then
  begin
    lTobTemp := FTobRub.FindFirst(['RB_RUBRIQUE'],[lStTemp], False);
    if lTobTemp <> nil then
      ParametrageRubrique(lTobTemp.GetString('RB_RUBRIQUE') + ';' +
                          lTobTemp.GetString('RB_NODOSSIER') + ';' +
                          lTobTemp.GetString('RB_PREDEFINI') + ';' +
                          DateToStr(lTobTemp.GetDateTime('RB_DATEVALIDITE')), taConsult, CtxRubrique);

  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTOTLIASSE.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key = VK_F11) then
  begin
    PopZoom.Popup(Mouse.CursorPos.x, Mouse.CursorPos.y);
  end
  else
  if (Key = VK_F12) then
  begin
    if FListe1.Focused then
    begin
      if FListe2.CanFocus then FListe2.SetFocus;
    end
    else
      if FListe2.Focused then
      begin
        Ctrl0.SetFocus;
      end
      else
        if FListe1.CanFocus then
          FListe1.SetFocus;
  end
  else
  if (Key = VK_F5) and (FListe1.Focused) then
    FListe1.OnDblClick(nil)
  else
  if (Key = 70) and (Shift = [ssCtrl]) then // GCO - 30/11/2006 - FQ 19248
    BRechercher.Click
  else
    FormKeyDown(Sender, Key, Shift); // Pour touches standard AGL
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTOTLIASSE.OnPopUpPopZoom(Sender: TObject);
var lStGeneral  : string;
    lStRubrique : string;
begin
  BGen.Enabled := False;
  BRub.Enabled := False;

  lStRubrique := (FListe1.Cells[cColAgregat, FListe1.Row]);

  lStGeneral := '';
  if FListe2.Visible then
    lStGeneral  := (FListe2.Cells[cColAgregat, FListe2.Row]);

  if lStRubrique <> '' then
  begin
    if ExisteSQL('SELECT RB_RUBRIQUE FROM RUBRIQUE WHERE RB_RUBRIQUE = "' + lStRubrique + '" AND ' +
                'RB_FAMILLES LIKE "' + Ctrl0.Value + '%" AND ' + FStWhereRub) then
      BRub.Enabled := True;
  end;

  if lStGeneral <> '' then
  begin
    if ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL = "' + lStGeneral + '"') then
      BGen.Enabled := True;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/10/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTOTLIASSE.OnClose ;
begin
  FreeAndNil(FTobRub);
  FreeAndNil(FZCumul);
  FreeAndNil(FFindDialog);
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/10/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTOTLIASSE.OnRowEnterFListe1(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  AfficheDetailRubrique;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/11/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTOTLIASSE.OnRowEnterFListe2(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var lStTemp : string;
begin
  lStTemp := FListe2.Cells[cColEcart, FListe2.Row];
  BCptLiasse.Enabled := lStTemp = 'X';
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/11/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTOTLIASSE.OnClickBRechercher(Sender: Tobject);
begin
  FFindFirst := True;
  FFindDialog.Execute;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/11/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTOTLIASSE.OnFindFFindDialog(Sender: TObject);
begin
  Rechercher(FListe1, FFindDialog, FFindFirst);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/10/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTOTLIASSE.AfficheDetailRubrique;
var i : integer;
    lTobLigne : Tob;
    lTobTemp  : Tob;

begin
  if FListe2.Visible then
  begin
    FListe1.ALign  := alTop;
    FListe1.Height := 214;
    FListe2.Align  := alClient;
  end
  else
  begin
    FListe1.ALign := alClient;
  end;

  if not FListe2.Visible then Exit;

  if FTobRub.Detail.Count = 0 then Exit;

  lTobLigne := FTobRub.FindFirst(['RB_RUBRIQUE'],[FListe1.Cells[cColAgregat, FListe1.Row]], False);
  if lTobLigne <> nil then
  begin
    if lTobLigne.Detail.Count > 0 then
    begin
      FListe2.RowCount := lTobLigne.Detail.Count + 1;
      for i := 0 to lTobLigne.Detail.Count -1 do
      begin
        lTobTemp := lTobLigne.Detail[i];
        FListe2.Cells[cColAgregat, i+1]   := lTobTemp.GetString('GENERAL');
        FListe2.Cells[cColLibelle, i+1]   := lTobTemp.GetString('LIBELLE');
        FListe2.Cells[cColComptable, i+1] := lTobTemp.GetString('TOTALCOMPTABLE');
        FListe2.Cells[cColFiscal, i+1]    := lTobTemp.GetString('TOTALFISCAL');
        FListe2.Cells[cColEcart, i+1]     := lTobTemp.GetString('ECART');
      end;
      BCptLiasse.Enabled := FListe2.Cells[cColEcart, FListe2.Row] = 'X';
    end
    else
    begin
      FListe2.RowCount := 2;
      BCptLiasse.Enabled := False;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTOTLIASSE.OnAfterSelectFiltre;
//var lStTemp : string;
begin
  (*lStTemp := TMPCtrl0.Text;
  CTRL0.Value := Copy(lStTemp, 2, Length(lStTemp)-1);

  lStTemp := TMPCtrl1.Text;
  Ctrl1.Text := Copy(lStTemp, 2, Length(lStTemp)-1);

  lStTemp := TMPCtrl1_.Text;
  Ctrl1_.Text := Copy(lStTemp, 2, Length(lStTemp)-1);*)
end;

////////////////////////////////////////////////////////////////////////////////

Initialization
  registerclasses ( [ TOF_CPTOTLIASSE ] ) ;
end.
