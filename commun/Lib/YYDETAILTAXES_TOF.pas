{***********UNITE*************************************************
Auteur  ...... : Sylvie DE ALMEIDA
Créé le ...... : 26/02/2007
Modifié le ... : 26/02/2007
Description .. : Source TOF de la FICHE : YYDETAILTAXES ()
Mots clefs ... : TOF;YYDETAILTAXES
*****************************************************************}
Unit YYDETAILTAXES_TOF ;

Interface

Uses ed_formu,
     ed_tools,
     StdCtrls,
     Controls,
     Classes,
     UtilPGI,
{$IFNDEF EAGLCLIENT}
     db,dbTables,Fiche,FE_main,FichList, EdtREtat, HDB,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eFiche,maineagl,eFichList,HPdfPrev,UtileAGL,
     eMul,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HMsgBox, 
     UTOF,
     UTOB,
     Windows,
     CPCODEPOSTAL_TOF,
     HTB97,
     {$IFDEF TAXES}
     YYInternationalTax,
     {$ENDIF}
     Messages,
     wcommuns,
     uiutil,
     SaisieList,
     Filtre,     // VideFiltre
     uTableFiltre,
     HPanel,
     commun,
     paramsoc,
     utobdebug;

{$IFDEF TAXES}
procedure YYLanceFiche_Detailtaxes;
procedure TESTFONCTION;
Type
  TOF_YYDETAILTAXES = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
   pPasDonnee : boolean;
   FListe : THGrid;
   CodeModele: THEdit;
   CbCategorie: THValComboBox;
   CodeCategorie : THEdit;
   CbType : THValComboBox;
   CbRegime : THValComboBox;
   CbPays : THEdit;
   CbRegion : THEdit;
   CbCodePostal : THEdit;
   CbVille : THEdit;
   Formule : THEdit;
   TT : TTreeview;
   MtMini : THNumEdit;
   PerDeb : THEdit;
   Perfin : THEdit;
   Assfin : THEdit;
   AssDeb : THEdit;
   TxMontantAchat : THNumEdit;
   TxMontantVente : THNumEdit;
   CompteAchat : THEdit;
   CompteVente : THEdit;
   CompteAchatEncaissement : THEdit;
   CompteVenteEncaissement : THEdit;
   BFirst: TToolbarButton97;
   BPrev : TToolbarButton97;
   BNext: TToolbarButton97;
   BLast: TToolbarButton97;
   BInsert: TToolbarButton97;
   BDelete: TToolbarButton97;
   BDupliquer: TToolbarButton97;
   BInitialise: TToolbarButton97;
   BTree: TToolbarButton97;
   TTModele : THTreeview;
   LibCat : THLabel;
   Pancritere : THPanel;
   //objets nouvelle sélection
   //BCherche: TToolbarButton97;
   BEnvoyer : TToolbarButton97;
   EdMod: THEdit;
   cbMod: THValComboBox;
   CbCat: THValComboBox;
   CbTyp: THValComboBox;
   CbReg: THValComboBox;
   Perdeb1 : THEdit;
   Perfin1 : THEdit;
   Assdeb1 : THNumEdit;
   Assfin1 : THNumEdit;
   ToolCritere : TToolWindow97;
   strCritereMod : string;
   strCritereCat : string;
   strCritereTyp : string;
   strCritereReg : string;
   Fvalidation: TNotifyEvent;
   FFerme : TNotifyEvent;
   FF : TFSaisieList;
   TF : tTableFiltre;
   Page1 : THTabSheet;
   procedure DoClick(Sender: TObject);
   procedure AfficheModele (Sender: TObject);
   procedure OnChangeCategorie (Sender: TObject);
   procedure OnChangeModele (Sender: TObject);
   procedure OnChangeCat (Sender: TObject);
   procedure onChangePerdeb1 (Sender: TObject);
   procedure onChangePerFin1 (Sender: TObject);
   procedure onChangeAssdeb1 (Sender: TObject);
   procedure onChangeAssfin1 (Sender: TObject);
   procedure DoChangePays(Sender: TObject);
   procedure OnDblClickCodePostalVille(Sender: TObject);
   procedure Cherche (Sender: TObject);
   procedure AppelFormule (Sender : TObject);
   Procedure CreateTV(TN : ttreeNode; Var TobFinale : Tob ; Var TV : tTreeView) ;
   procedure MAJComboCat;
   procedure MAJComboType;
   procedure MAJComboRegime;
   procedure MAJComboCatFiltre;
   procedure MAJComboTypFiltre;
   procedure MAJComboRegFiltre;
   procedure InitialiserTableDetails (Sender: TObject);
   procedure ChercheAvecCritere(strCritereMod, strCritereCat, strCritereTyp, strCritereReg : string);
   procedure GriserElements(bValeur : boolean);
   procedure GereModeleFerme(strModele : string);
   function  GetMinAnneeExercice : TDateTime ;
   function  bDonnee (strModele : string) : boolean;
   function  GetColIndex    ( vStColName : String ) : Integer ;
   function  bVide (strModele : string) : boolean;
   procedure OnbpostClick(Sender: TObject);
   procedure OnFermeClick (Sender : TObject);
   function  bChercheChevauchementAssiette (strModele, strCategorie, strType, strRegime, strAssDeb, strAssFin : string; Tperdeb, Tperfin : TDatetime; bDebut : boolean) : boolean;
   procedure ChargeLibCat (strmodele : string);
   procedure FormKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
  end ;
{$ENDIF}
Implementation

uses Grids, TntComCtrls, TntStdCtrls, TntGrids, TntWideStrings;
{$IFDEF TAXES}
Const ColonneMod : Integer = 1;// Index de la colonne modèle de taxe
      ColonneCat : Integer = 2;// Index de la colonne catégorie de taxe
      ColonneTyp : Integer = 3;// Index de la colonne type de taxe
      ColonneReg : Integer = 4;// Index de la colonne régime de taxe
      MessageListe : Array[0..17] of String =// Message d'erreurs
					('Merci de paramétrer les types de taxes et les régimes gérés.',
           'Merci de renseigner un libellé.',
					 'L''enregistrement est inaccessible.',
           'Merci de renseigner la catégorie.',
           'Merci de renseigner le type.',
           'Merci de renseigner le régime.',
           'La date de fin d''application doit être supérieure à la date de début.',
           'L''assiette de fin doit être supérieure à l''assiette début.',
           'Ce compte n''existe pas.',
           'Les données ont été alimentées avec succès. #13Merci de relancer la fenêtre pour prise en compte des données.',
           'L''assiette de début est incluse dans une assiette précédemment renseignée.',
           'L''assiette de fin est incluse dans une assiette précédemment renseignée.',
           'Merci de renseigner la catégorie',
           'Merci de renseigner le type',
           'Merci de renseigner le régime',
           'Le pays saisi n''existe pas. Merci de modifier votre saisie.',
           'La région saisie n''existe pas. Merci de modifier votre saisie.',
           'Le traitement a été effectué avec succès.'
           );

procedure YYLanceFiche_Detailtaxes;
begin
    AglLanceFiche('YY','YYDETAILTAXES','','','') ;
end;

procedure TESTFONCTION ;
var
 //TESTS  FONCTION INTERNATIONAL
   oInt : TInternationalTax;
   TobQ, TobResult : Tob ;
   //Fin TEST
   //TESTS fonction creation ligne tob question
   strMod, strCategorie, strType, strRegime, strEtablissement, strFlux, strCP, strVille, strRegion, strPays  : string ;
   tDatepiece : TDatetime;
   iQuantiteLigne : integer;
   dMontantLigne : double ;
   TobLigneQuestion2 : TOB;
   strdate2 : string;
   //Fin TESTS fonction creation ligne tob question
begin
  //TESTS FONCTION INTERNATIONAL

    oInt := TInternationalTax.Create;
    //TobDebug (oInt.FTOBPARAMETRESTAXES);

    //strType := oInt.tstrGetLibelleTypesTaxesModele('def','TX1');
    //strType := oInt.tstrGetCodeTypesTaxesModele('def','TX1');
    //strRegime := oInt.tstrGetLibelleRegimesTaxesModele('def','TX1');
    //strRegime := oInt.tstrGetCodeRegimesTaxesModele('def','TX1');

    oInt.FTOBPARAMETRESTAXES.Load(true);
    TobDebug (oInt.FTOBPARAMETRESTAXES);

    //tests fonction creation ligne tob question
    {strMod := 'DEF';
    strCategorie := 'TX1';
    strType := 'NOR';
    strRegime := 'titi';
    strEtablissement := '';
    strFlux := 'ACH';
    strCP := '';
    strVille := '';
    strRegion := '';
    strPays := 'FRA';
    iQuantiteLigne := 10;
    dMontantLigne := 300;

    strDate2 := '10/01/2007';
    tDatepiece := strtodate(strDate2); }
    TobQ := TOB.Create('Question',nil,-1);
    {TobLigneQuestion := TOB.Create('LigneQuestion',TobQ,-1);
    oInt.tCreeLigneQuestion(TobLigneQuestion, strMod, strCategorie, strType, strRegime, strEtablissement, strFlux, strCP, strVille, strRegion, strPays, tDatepiece,iQuantiteLigne, dMontantLigne );
    TobDebug (TobQ); }
    strMod := 'DEF';
    strCategorie := 'TX1';
    strType := 'NOR';
    strRegime := 'CEE';
    strEtablissement := '002';
    strFlux := 'ACH';
    strCP := '';
    strVille := '';
    strRegion := '';
    strPays := 'AUS';
    iQuantiteLigne := 2;
    dMontantLigne := 4000;

    strDate2 := '01/01/1999';
    tDatepiece := strtodate(strDate2);
    TobLigneQuestion2 := TOB.Create('LigneQuestion2',TobQ,-1);
    oInt.tCreeLigneQuestion(TobLigneQuestion2, strMod, strCategorie, strType, strRegime, strEtablissement, strFlux, strCP, strVille, strRegion, strPays, tDatepiece,iQuantiteLigne, dMontantLigne );
    //i := tobQ.Detail.count;
    TobDebug (tobQ);

    TobResult := oInt.GetInternationalTax(tobQ, true);
     TobDebug (TobResult);
    //Fin tests  fonction creation ligne tob question

end;


procedure TOF_YYDETAILTAXES.OnNew ;
var
  tDate : tdatetime;
  strModele : string;
begin
Inherited ;
  if TTmodele.Selected.Text = TraduireMemoire('Modèles de taxes') then
      strModele := TTModele.items[1].text
  else
      strModele := trim(copy(TTModele.Selected.text,1,pos('(', TTModele.Selected.text) - 1));


  setControlText ('CCT_CODEMODELE', strModele);

  cbCategorie.Enabled := True;
  CbType.Enabled := True;
  CbRegime.Enabled := True;
  MtMini.Enabled := True;
  PerDeb.Enabled := True;
  Perfin.Enabled := True;
  AssDeb.Enabled := True;
  AssFin.Enabled := True;
  TxMontantAchat.Enabled := True;
  TxMontantVente.Enabled := True;
  cbPays.Enabled := True;
  cbRegion.Enabled := True;
  cbCodePostal.Enabled := True;
  cbVille.Enabled := True;
  CompteAchat.Enabled := True;
  CompteVente.Enabled := True;
  CompteAchatEncaissement.Enabled := True;
  CompteVenteEncaissement.Enabled := True;
  Formule.Enabled := True;

  if cbCategorie.Items.Count > 0 then
    cbCategorie.ItemIndex := 0;
  MAJComboCat;
  MAJComboType;
  MAJComboRegime;
  //Positionnement des dates par défaut
  tDate := GetMinAnneeExercice;
  PerDeb.text := DateToStr(tDate);
  Page1.Visible := True;
  PerDeb.SetFocus;
  Perfin.Text :=  DateToStr(iDate2099);
  Perfin.SetFocus;
  CbCategorie.SetFocus;
end ;

procedure TOF_YYDETAILTAXES.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_YYDETAILTAXES.OnUpdate ;
begin
   Inherited ;
   CbCategorie.enabled := false;
end ;


procedure TOF_YYDETAILTAXES.OnLoad ;
var
  strModele, strMask : string;
  tDate : TDatetime;
  iNbDecimal, i : integer;
begin
  Inherited ;
  SetControlVisible('PCOMPLEMENT',False) ;

  if TTModele.Items.Count = 0 then exit;
  TTModele.Items[0].Text := TraduireMemoire('Modèles de taxes');
  if pPasDonnee then
  begin
    Ecran.Close;
    if IsInside (Ecran) then
      CloseInsidePanel (Ecran) ;
    exit;
  end;

  TTModele.SetFocus;
  if strCritereMod <> '' then
  begin
    strModele := strCritereMod;
    TTModele.Items[1].Selected := true;
  end
  else
  begin
    TTModele.Items[1].Selected := true;
    //strModele := copy(TTModele.Selected.text,1,3);
    strModele := trim(copy(TTModele.Selected.text,1,pos('(', TTModele.Selected.text) - 1))
  end;
  GereModeleFerme (strModele);
  cbcategorie.Enabled := false;

  if Fliste.RowCount > 0 then
    ChargeLibCat (strModele);

  EdMod.Text := strModele;
  tDate := GetMinAnneeExercice;
  PerDeb1.text := DateToStr(tDate);
  Perfin1.Text := DateToStr(iDate2099);
  if cbcat.Items.Count > 0 then
    cbCat.ItemIndex := 0;
  if cbTyp.Items.Count > 0 then
    cbTyp.ItemIndex := 0;
  if cbReg.Items.Count > 0 then
    cbReg.ItemIndex := 0;

  MAJComboCat;  //combo catégories de taxe
  MAJComboType;  //Combo types de taxe
  MAJComboRegime;  //Combo régimes de taxe
  cbRegion.plus := 'RG_PAYS = "' + cbPays.Text + '"';
  cbCodePostal.Plus := 'O_PAYS = "' + cbPays.Text + '"';
  cbVille.Plus := 'O_PAYS = "' + cbPays.Text + '"';
  CompteAchat.Plus := 'G_NATUREGENE = "DIV" ';
  CompteVente.Plus := 'G_NATUREGENE = "DIV"';
  CompteAchatEncaissement.Plus :=  'G_NATUREGENE = "DIV" ';
  CompteVenteEncaissement.Plus :=  'G_NATUREGENE = "DIV" ';

  Formule.ElipsisButton := TRUE;
  
  if (strCritereMod <> '')
     and (strCritereCat <> '')
     and (strCritereTyp <> '')
     and (strCritereReg <> '') then
  begin
    ChercheAvecCritere (strCritereMod, strCritereCat, strCritereTyp, strCritereReg);
    bTree.Visible := False;
    pancritere.Enabled := False;
    pancritere.Height := 0;
    pancritere.Visible := False;

  end;


  if bVide (strModele) then
  begin
    cbCategorie.Enabled := false;
    CbType.Enabled := false;
    CbRegime.Enabled := false;
    MtMini.Enabled := false;
    PerDeb.Enabled := false;
    Perfin.Enabled := false;
    AssDeb.Enabled := false;
    AssFin.Enabled := false;
    TxMontantAchat.Enabled := false;
    TxMontantVente.Enabled := false;
    cbPays.Enabled := false;
    cbRegion.Enabled := false;
    cbCodePostal.Enabled := false;
    cbVille.Enabled := false;
    CompteAchat.Enabled := false;
    CompteVente.Enabled := false;
    CompteAchatEncaissement.Enabled := false;
    CompteVenteEncaissement.Enabled := false;
    Formule.Enabled := false;
  end;
  iNbDecimal := CalcDecimaleDevise(GetParamSoc('SO_DEVISEPRINC'));
  strMask := '0.';
  for i := 1 to iNbDecimal do
    strMask := strMask + '0';
  FListe.ColTypes[7]   := 'R' ;         // Montant
  FListe.ColFormats[7] := '#,#0'+ strMask ;  // Montant

  FListe.ColTypes[8]   := 'R' ;         // Montant
  FListe.ColFormats[8] := '#,#0'+ strMask ;  // Montant


end ;

procedure TOF_YYDETAILTAXES.OnArgument (S : String ) ;
var
  Critere : string;
  x : integer;
  Arg, Val, strMask: string;
  iNbDecimal, i : integer;
begin
  Inherited ;
  pPasDonnee := false;
  // Récupération des contrôles

  // Gestion des arguments
  repeat
    Critere := uppercase(Trim(ReadTokenSt(S)));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        Arg := copy(Critere, 1, x - 1);
        Val := copy(Critere, x + 1, length(Critere));
        if Arg = 'CODEMOD' then strCritereMod := Val;
        if Arg = 'CODECAT' then strCritereCat := Val;
        if Arg = 'CODETYP' then strCritereTyp := Val;
        if Arg = 'CODEREG' then strCritereReg := Val;
      end;
    end;
  until Critere = '';

  FListe := THGrid(GetControl('FListe'));
  FListe.OnClick := DoClick;

  LibCat := THLabel (GetControl('LIBCAT'));
  CodeModele := THEdit(GetControl('CCT_CODEMODELE'));

  CbCategorie := THValComboBox(GetControl('CCT_CODECAT'));
  CbCategorie.OnChange := OnChangeCategorie;
  CodeCategorie := THEdit(GetControl('CODECAT'));

  CbType := THValComboBox(GetControl('CCT_CODETYP'));
  CbRegime := THValComboBox(GetControl('CCT_CODEREG'));

  CbPays := THEdit(GetControl('CCT_PAYS'));
  CbPays.OnChange := DoChangePays;
  CbRegion := THEdit(GetControl('CCT_REGION'));
  CbCodePostal := THEdit(GetControl('CCT_CODEPOSTAL',true));
  CbCodePostal.OnChange := OnDblClickCodePostalVille;
  CbVille := THEdit(GetControl('CCT_VILLE',true));
  CbVille.OnChange := OnDblClickCodePostalVille;
  MtMini := THNumEdit(GetControl('CCT_MTMINI',true));
  iNbDecimal := CalcDecimaleDevise(GetParamSoc('SO_DEVISEPRINC'));
  MtMini.Decimals := iNbDecimal;
  TxMontantAchat := THNumEdit(GetControl('CCT_TXMTACHAT',true));
  TxMontantAchat.Decimals := iNbDecimal;
  TxMontantVente := THNumEdit(GetControl('CCT_TXMTVENTE',true));
  TxMontantVente.Decimals := iNbDecimal;
  CompteAchat := THEdit(GetControl('CCT_CPTACHAT',true));
  CompteAchat.ElipsisButton := TRUE;
  CompteVente := THEdit(GetControl('CCT_CPTVENTE',true));
  CompteVente.ElipsisButton := TRUE;
  CompteAchatEncaissement := THEdit(GetControl('CCT_ENCACHAT',true));
  CompteAchatEncaissement.ElipsisButton := TRUE;
  CompteVenteEncaissement := THEdit(GetControl('CCT_ENCVENTE',true));
  CompteVenteEncaissement.ElipsisButton := TRUE;

  //Formule
  Formule := THEdit(GetControl('CCT_FORMULE',true));
  Formule.ElipsisButton := TRUE;
  Formule.OnElipsisClick := AppelFormule;
  //Gestion insertion d'un nouvel enregistrement
  TTModele := THTreeview (GetControl('TreeEntete',true));
  TTModele.OnClick := DoClick;
  TTModele.OnDblClick := AfficheModele;

  //Période fin et assiette fin
  Perfin := THEdit(GetControl('CCT_PERFIN',true));
  Assfin := THEdit(GetControl('CCT_ASSFIN',true));
  strMask := '0.';
  for i := 1 to iNbDecimal do
    strMask := strMask + '0';
  Assfin.DisplayFormat := strMask;
  AssDeb := THEdit(GetControl('CCT_ASSDEB',true));
  AssDeb.DisplayFormat := strMask;
  PerDeb := THEdit(GetControl('CCT_PERDEB',true));

  //Objets nouvelle sélection
  EdMod:= THEdit(GetControl('YMT_CODE__'));
  EdMod.OnChange := OnChangeModele;

  cbCat := THValComboBox(GetControl('FCODECAT'));
  //cbCat.OnChange := OnChangeCat;
  cbTyp := THValComboBox(GetControl('FCODETYP'));
  cbReg := THValComboBox(GetControl('FCODEREG'));
  Perdeb1 := THEdit(GetControl('CCT_PERDEB1',true));
  Perdeb1.OnExit := onChangePerdeb1;
  Perfin1 := THEdit(GetControl('CCT_PERDEB1_',true));
  Perfin1.OnExit := onChangePerfin1;
  Assdeb1 := THNumEdit(GetControl('CCT_ASSDEB1',true));
  AssDeb1.OnExit := onChangeAssdeb1;
  //AssDeb1.Decimals := CalcDecimaleDevise(GetParamSoc('SO_DEVISEPRINC'));
  Assfin1 := THNumEdit(GetControl('CCT_ASSFIN1',true));
  AssFin1.OnExit := onChangeAssFin1;
  //AssFin1.Decimals := CalcDecimaleDevise(GetParamSoc('SO_DEVISEPRINC'));

  BEnvoyer := TToolbarButton97 (GetControl('BEnvoyer'));
  BEnvoyer.OnClick := Cherche;
  ToolCritere := TToolWindow97 (GetControl('ToolCritere'));
  BInsert := TToolbarButton97 (GetControl('BInsert'));
  BDelete := TToolbarButton97 (GetControl('BDelete'));
  BDupliquer := TToolbarButton97 (GetControl('BDupliquer'));
  BTree:= TToolbarButton97(GetControl('BTree'));

  Fvalidation := TToolbarButton97(GetControl('Bpost')).OnClick;
  TToolbarButton97(GetControl('Bpost')).Onclick := OnbpostClick;

  //FFerme := TToolbarButton97(GetControl('BFerme')).OnClick;
  //TToolbarButton97(GetControl('BFerme')).Onclick := OnFermeClick;

  BInitialise := TToolbarButton97 (GetControl('bInitialise'));
  BInitialise.OnClick := InitialiserTableDetails;

  Pancritere := THPanel (GetControl('Pancritere'));

  BFirst := TToolbarButton97 (GetControl('BFirst'));
  BPrev:= TToolbarButton97 (GetControl('BPrev'));
  BNext := TToolbarButton97 (GetControl('BNext'));
  BLast := TToolbarButton97 (GetControl('BLast'));

  Ecran.OnKeyDown := FormKeyDown;
  Page1 := THTabSheet (GetControl('Page1'));

  if strCritereMod <> '' then //Mode consultation
  begin
    FF := TFSaisieList(ecran);
    FF.TypeAction := taconsult;
    BInitialise.Enabled := false;
    Perfin.Enabled := False;
    CompteAchat.Enabled := False;
    CompteVente.Enabled := False;
    CompteAchatEncaissement.Enabled := False;
    CompteVenteEncaissement.Enabled := False;
    cbPays.Enabled := false;
    cbRegion.Enabled := false;
    cbCodePostal.Enabled := false;
    cbVille.Enabled := false;
    formule.Enabled := false;
    setcontrolVisible('BFirst', false);
    setcontrolVisible('BPrev', false);
    setcontrolVisible('BNext', false);
    setcontrolVisible('BLast', false);

  end;
end ;



procedure TOF_YYDETAILTAXES.DoClick(Sender: TObject);
var
   strModele : string;
begin
   try
   if pPasDonnee then exit;
   if pos('(', TTModele.Selected.text) > 0 then
    strModele := trim(copy(TTModele.Selected.text,1,pos('(', TTModele.Selected.text) - 1))
   else
    exit;

   ChargeLibCat (strModele);
   //MajComboCatFiltre;

   if not bDonnee (strModele) then
    bInsert.Enabled := false
   else
    if strCritereMod = '' then //Mode consultation
      bInsert.Enabled := true;

   EdMod.Text := strModele;

   MAJComboCat;
   MAJComboType;
   MAJComboRegime;
   //on teste si le modèle est fermé : si modèle fermé écran en consultation
   //strModele := copy(TTModele.Selected.text,1,3);
   strModele := trim(copy(TTModele.Selected.text,1,pos('(', TTModele.Selected.text) - 1));
   GereModeleFerme (strModele);
   except
   end;

end;
procedure TOF_YYDETAILTAXES.GereModeleFerme(strModele : string);
var
   St : string;
   strSQL : string;
   OQuery : TQuery ;
begin
   strSQL := 'SELECT YMODELETAXE.YMT_FERME FROM YMODELETAXE WHERE YMODELETAXE.YMT_CODE = "'+strModele+'"';
   OQuery := OpenSql (strSQL, TRUE);
   try
   if not OQuery.Eof then
   begin
    St := OQuery.Fields[0].AsString;
    if St = 'X' then // le modèle est fermé
      GriserElements (false)
    else
      GriserElements (true);
   end;
   finally
    Ferme(OQuery);
   end;
end;

procedure TOF_YYDETAILTAXES.OnChangeCategorie(Sender: TObject);
begin

   //CodeCategorie.Text := cbcategorie.Values [cbcategorie.itemindex];
   //FListe.Cells[ColonneCatCode,Fliste.row] := CodeCategorie.Text;
   //FListe.Cells[ColonneCat,Fliste.row] := cbCategorie.Text;
   MAJComboType;
   MAJComboRegime;
   if strCritereMod = '' then //Mode consultation
   begin
    cbType.Enabled := true;
    cbRegime.Enabled := true;
   end;

end;

procedure TOF_YYDETAILTAXES.OnChangeModele(Sender: TObject);
begin
   MAJComboCatFiltre;
   MAJComboTypFiltre;
   MAJComboRegFiltre;
  if cbcat.Items.Count > 0 then
    cbCat.ItemIndex := 0;
  if cbTyp.Items.Count > 0 then
    cbTyp.ItemIndex := 0;
  if cbReg.Items.Count > 0 then
    cbReg.ItemIndex := 0;
end;

procedure TOF_YYDETAILTAXES.OnChangeCat(Sender: TObject);
begin
   MAJComboTypFiltre;
   MAJComboRegFiltre;
end;

procedure TOF_YYDETAILTAXES.DoChangePays(Sender: TObject);
begin
   //Si on change le pays, on remet à niveau la liste des régions
   cbRegion.plus := 'RG_PAYS = "' + cbPays.Text + '"';
   //Changement pays : on initialise région, cp, ville
   cbRegion.Text := '';
   cbCodePostal.Plus := 'O_PAYS = "'+ cbPays.Text + '"';
   cbCodePostal.Text := '';
   cbVille.Plus := 'O_PAYS = "'+ cbPays.Text + '"';
   cbVille.Text := '';
end;

procedure TOF_YYDETAILTAXES.OnDblClickCodePostalVille(Sender: TObject);
begin
  if cbCodePostal.Text <> '' then
    VerifCodePostal(nil,THEdit(GetControl('CCT_CODEPOSTAL',true)),THEdit(GetControl('CCT_VILLE',true)),TRUE);
end;


procedure TOF_YYDETAILTAXES.OnClose ;
begin
  Inherited ;
  FreeAndNil (TT);
end ;

procedure TOF_YYDETAILTAXES.OnDisplay () ;
begin
  Inherited ;

end ;

procedure TOF_YYDETAILTAXES.OnCancel () ;
begin
  Inherited ;
  if TToolbarButton97(GetControl('Bdefaire')).Enabled = True then
    TToolbarButton97(GetControl('Bdefaire')).click;
end ;

procedure TOF_YYDETAILTAXES.MAJComboCat;
var
   strSQL, St, St1, va, it : string;
   OQuery : TQuery ;
   id_cat, i : integer;
   strModele : hstring;
begin

   //Mise à jour combo catégorie
   cbCategorie.Clear; //on initialise la combo des catégories
   if (FListe.Cells[ColonneMod,Fliste.row]='') then
      //strModele := copy(TTModele.Selected.text,1,3)
      strModele := trim(copy(TTModele.Selected.text,1,pos('(', TTModele.Selected.text) - 1))
   else
      strModele := FListe.Cells[ColonneMod,Fliste.row];

   strSQL := 'SELECT "TX1", YMODELETAXE.YMT_LIBCAT1 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT1 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT1 <> '' AND YMODELETAXE.YMT_CODE = "'+strModele+'"';
   strSql := strSQL + 'AND "TX1" IN (SELECT YCY_CODECAT FROM YMODELECATTYP WHERE YCY_TYPGERE = "X") AND "TX1" IN (SELECT YCR_CODECAT FROM YMODELECATREG WHERE YCR_REGGERE = "X")';
   strSQL := strSQL + ' UNION ';
   strSQL := strSQL + 'SELECT "TX2", YMODELETAXE.YMT_LIBCAT2 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT2 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT2 <> '' AND YMODELETAXE.YMT_CODE = "'+strModele+'"';
   strSql := strSQL + 'AND "TX2" IN (SELECT YCY_CODECAT FROM YMODELECATTYP WHERE YCY_TYPGERE = "X") AND "TX2" IN (SELECT YCR_CODECAT FROM YMODELECATREG WHERE YCR_REGGERE = "X")';
   strSQL := strSQL + ' UNION ';
   strSQL := strSQL + 'SELECT "TX3", YMODELETAXE.YMT_LIBCAT3 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT3 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT3 <> '' AND YMODELETAXE.YMT_CODE = "'+strModele+'"';
   strSql := strSQL + 'AND "TX3" IN (SELECT YCY_CODECAT FROM YMODELECATTYP WHERE YCY_TYPGERE = "X") AND "TX3" IN (SELECT YCR_CODECAT FROM YMODELECATREG WHERE YCR_REGGERE = "X")';
   strSQL := strSQL + ' UNION ';
   strSQL := strSQL + 'SELECT "TX4", YMODELETAXE.YMT_LIBCAT4 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT4 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT4 <> '' AND YMODELETAXE.YMT_CODE = "'+strModele+'"';
   strSql := strSQL + 'AND "TX4" IN (SELECT YCY_CODECAT FROM YMODELECATTYP WHERE YCY_TYPGERE = "X") AND "TX4" IN (SELECT YCR_CODECAT FROM YMODELECATREG WHERE YCR_REGGERE = "X")';
   strSQL := strSQL + ' UNION ';
   strSQL := strSQL + 'SELECT "TX5", YMODELETAXE.YMT_LIBCAT5 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT5 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT5 <> '' AND YMODELETAXE.YMT_CODE = "'+strModele+'"';
   strSql := strSQL + 'AND "TX5" IN (SELECT YCY_CODECAT FROM YMODELECATTYP WHERE YCY_TYPGERE = "X") AND "TX5" IN (SELECT YCR_CODECAT FROM YMODELECATREG WHERE YCR_REGGERE = "X")';

  //OQuery := OpenSQL(strSql, true, -1, '', true, '', false);
  OQuery := OpenSQL( strSql , false);
  try
  i := 0;
  id_cat := 0;
  While not OQuery.Eof do
  begin
    St := OQuery.Fields[0].AsString;
    if St = FListe.Cells[ColonneCat,Fliste.row] then
      id_cat := i;
    St1 := OQuery.Fields[1].AsString;
    va:=ReadTokenSt(St) ;
    it:=TraduireMemoire(ReadTokenSt(St1)) ;
    cbCategorie.Values.Add(va) ;
    cbcategorie.Items.Add(it) ;
    i := i + 1;
    OQuery.next;
  end;
  finally
    Ferme(OQuery);
  end;
  //Positionnement combo catégorie
  if not bVide (strModele) then
  begin
  if cbcategorie.Items.Count > 0 then
  begin
    cbcategorie.itemindex := id_cat;

  if cbcategorie.Enabled = true then
  begin
    if cbcategorie.itemindex > -1 then
      CbCategorie.Value := cbcategorie.Values [cbcategorie.itemindex];
  end;

  end;
  end;
end;

procedure TOF_YYDETAILTAXES.MAJComboType;
var
   strSQL, St, St1, va, it : string;
   OQuery : TQuery ;
   id_typ, i : integer;
   strModele : hstring;
begin
   //Mise à jour combo types de taxes dans le cas d'une création d'enregistrement
   cbType.Clear; //on initialise la combo des types de taxes
   cbType.Values.Clear; //on initialise la combo des types de taxes

   if (FListe.Cells[ColonneMod,Fliste.row]='') then
      //strModele := copy(TTModele.Selected.text,1,3)
      strModele := trim(copy(TTModele.Selected.text,1,pos('(', TTModele.Selected.text) - 1))
   else
      strModele := FListe.Cells[ColonneMod,Fliste.row];

  if bVide (strModele) and (cbCategorie.Text = '') then exit;

  if (TTModele.Selected.parent = nil) then exit;
  strSQL := 'SELECT YMODELECATTYP.YCY_CODETYP, YTYPETAUX.TYP_LIBELLE FROM YMODELECATTYP INNER JOIN YTYPETAUX ON YMODELECATTYP.YCY_CODETYP = YTYPETAUX.TYP_CODE WHERE YMODELECATTYP.YCY_TYPGERE = "X" AND YMODELECATTYP.YCY_CODECAT  = "'+CbCategorie.values[CbCategorie.itemindex]+'" AND YMODELECATTYP.YCY_CODEMODELE = "'+strModele+'"';
  OQuery := OpenSql( strSQL, true, -1, '', true, '', false);
  try
  i := 0;
  id_typ := 0;

  While not OQuery.Eof do
  begin
    St := OQuery.Fields[0].AsString;
    if St = FListe.Cells[ColonneTyp,Fliste.row] then
      id_typ := i;
    St1 := OQuery.Fields[1].AsString;
    va:=ReadTokenSt(St) ;
    it:=TraduireMemoire(ReadTokenSt(St1)) ;
    cbType.Values.Add(va) ;
    cbType.Items.Add(it) ;
    i := i + 1;
    OQuery.next;
  end;
  finally
    Ferme(OQuery);
  end;
  //Positionnement combo types
  if cbType.Items.Count > 0 then
  begin
  cbType.itemindex := id_typ;
  if cbType.Enabled = true then
    cbType.Value := cbType.Values [cbType.itemindex];
  end;

end;

procedure TOF_YYDETAILTAXES.MAJComboRegime;
var
   strSQL, St, St1, va, it : string;
   OQuery : TQuery ;
   id_reg, i : integer;
   strModele : hstring;
begin
   //Mise à jour combo types de taxes dans le cas d'une création d'enregistrement
   cbRegime.Clear; //on initialise la combo des régimes de taxes
   cbRegime.Values.Clear; //on initialise la combo des régimes de taxes


   if (FListe.Cells[ColonneMod,Fliste.row]='') then
      //strModele := copy(TTModele.Selected.text,1,3)
      strModele := trim(copy(TTModele.Selected.text,1,pos('(', TTModele.Selected.text) - 1))
   else
      strModele := FListe.Cells[ColonneMod,Fliste.row];

  if bVide (strModele) and (cbCategorie.Text = '') then exit;
  
  if (TTModele.Selected.parent = nil) then exit;

  strSQL := 'SELECT YMODELECATREG.YCR_CODEREG, CHOIXCOD.CC_LIBELLE FROM YMODELECATREG INNER JOIN CHOIXCOD ON YMODELECATREG.YCR_CODEREG = CHOIXCOD.CC_CODE WHERE YMODELECATREG.YCR_REGGERE = "X" AND CHOIXCOD.CC_TYPE = "RTV" AND YMODELECATREG.YCR_CODECAT  = "'+CbCategorie.values[CbCategorie.itemindex]+'" AND YMODELECATREG.YCR_CODEMODELE = "'+strModele+'"';
  OQuery := OpenSql( strSQL, true, -1, '', true, '', false);
  try
  i := 0;
  id_reg := 0;
  While not OQuery.Eof do
  begin
    St := OQuery.Fields[0].AsString;
    if St = FListe.Cells[ColonneReg,Fliste.row] then
      id_reg := i;
    St1 := OQuery.Fields[1].AsString;
    va:=ReadTokenSt(St) ;
    it:=TraduireMemoire(ReadTokenSt(St1)) ;
    cbRegime.Values.Add(va) ;
    cbRegime.Items.Add(it) ;
    i := i + 1;
    OQuery.next;
  end;
  finally
    Ferme(OQuery);
  end;
  //Positionnement combo types
  if cbRegime.Items.Count > 0 then
  begin
  cbRegime.itemindex := id_reg;
  if cbRegime.Enabled = true then
    cbRegime.Value := cbRegime.Values [cbRegime.itemindex];
  end;


end;

procedure TOF_YYDETAILTAXES.MAJComboCatFiltre;
var
   strSQL, St, St1, va, it : string;
   OQuery : TQuery ;
   strModele : hstring;
begin

   //Mise à jour combo catégorie
   cbCat.Clear; //on initialise la combo des catégories
   //strModele := cbMod.values[cbMod.itemindex];
   strModele := Edmod.Text;

   strSQL := 'SELECT "TX1", YMODELETAXE.YMT_LIBCAT1 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT1 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT1 <> '' AND YMODELETAXE.YMT_CODE = "'+strModele+'"';
   strSQL := strSQL + ' UNION ';
   strSQL := strSQL + 'SELECT "TX2", YMODELETAXE.YMT_LIBCAT2 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT2 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT2 <> '' AND YMODELETAXE.YMT_CODE = "'+strModele+'"';
   strSQL := strSQL + ' UNION ';
   strSQL := strSQL + 'SELECT "TX3", YMODELETAXE.YMT_LIBCAT3 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT3 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT3 <> '' AND YMODELETAXE.YMT_CODE = "'+strModele+'"';
   strSQL := strSQL + ' UNION ';
   strSQL := strSQL + 'SELECT "TX4", YMODELETAXE.YMT_LIBCAT4 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT4 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT4 <> '' AND YMODELETAXE.YMT_CODE = "'+strModele+'"';
   strSQL := strSQL + ' UNION ';
   strSQL := strSQL + 'SELECT "TX5", YMODELETAXE.YMT_LIBCAT5 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT5 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT5 <> '' AND YMODELETAXE.YMT_CODE = "'+strModele+'"';

  OQuery := OpenSql (strSQL, TRUE);
  try
  va := '<<Tous>>';
  it := TraduireMemoire('<<Tous>>');
  cbCat.Values.Add(va);
  cbcat.Items.Add(it);
  While not OQuery.Eof do
  begin
    St := OQuery.Fields[0].AsString;
    St1 := OQuery.Fields[1].AsString;
    va:=ReadTokenSt(St) ;
    it:=TraduireMemoire(ReadTokenSt(St1)) ;
    cbCat.Values.Add(va) ;
    cbcat.Items.Add(it) ;
    OQuery.next;
  end;
  finally
    Ferme(OQuery);
  end;
  MAJComboTypFiltre;
  MAJComboRegFiltre;
end;

procedure TOF_YYDETAILTAXES.MAJComboTypFiltre;
var
   strSQL, St, St1, va, it : string;
   OQuery : TQuery ;
    i : integer;
   strModele : hstring;
begin
   //Mise à jour combo types de taxes dans le cas d'une création d'enregistrement
   cbTyp.Clear; //on initialise la combo des types de taxes
   //strModele := cbMod.values[cbMod.itemindex];
   strModele := Edmod.Text;
  //if CbCat.itemindex > -1 then
  //  strSQL := 'SELECT YMODELECATTYP.YCY_CODETYP, YTYPETAUX.TYP_LIBELLE FROM YMODELECATTYP INNER JOIN YTYPETAUX ON YMODELECATTYP.YCY_CODETYP = YTYPETAUX.TYP_CODE WHERE YMODELECATTYP.YCY_TYPGERE = "X" AND YMODELECATTYP.YCY_CODECAT  = "'+CbCat.values[CbCat.itemindex]+'" AND YMODELECATTYP.YCY_CODEMODELE = "'+strModele+'"'
  //else
  strSQL := 'SELECT YMODELECATTYP.YCY_CODETYP, YTYPETAUX.TYP_LIBELLE FROM YMODELECATTYP INNER JOIN YTYPETAUX ON YMODELECATTYP.YCY_CODETYP = YTYPETAUX.TYP_CODE WHERE YMODELECATTYP.YCY_TYPGERE = "X" AND YMODELECATTYP.YCY_CODEMODELE = "'+strModele+'"';
  OQuery := OpenSql (strSQL, TRUE);
  try
  i := 0;
  va := '<<Tous>>';
  it := TraduireMemoire('<<Tous>>');
  cbTyp.Values.Add(va);
  cbTyp.Items.Add(it);
  While not OQuery.Eof do
  begin
    St := OQuery.Fields[0].AsString;
    St1 := OQuery.Fields[1].AsString;
    va:=ReadTokenSt(St) ;
    it:=TraduireMemoire(ReadTokenSt(St1)) ;
    cbTyp.Values.Add(va) ;
    cbTyp.Items.Add(it) ;
    i := i + 1;
    OQuery.next;
  end;
  finally
    Ferme(OQuery);
  end;
end;

procedure TOF_YYDETAILTAXES.MAJComboRegFiltre;
var
   strSQL, St, St1, va, it : string;
   OQuery : TQuery ;
   i : integer;
   strModele : hstring;
begin
   //Mise à jour combo types de taxes dans le cas d'une création d'enregistrement
   cbReg.Clear; //on initialise la combo des types de taxes
   //strModele := cbMod.values[cbMod.itemindex];
   strModele := Edmod.Text;

  //if CbCat.itemindex > -1 then
  //  strSQL := 'SELECT YMODELECATREG.YCR_CODEREG, CHOIXCOD.CC_LIBELLE FROM YMODELECATREG INNER JOIN CHOIXCOD ON YMODELECATREG.YCR_CODEREG = CHOIXCOD.CC_CODE WHERE YMODELECATREG.YCR_REGGERE = "X" AND CHOIXCOD.CC_TYPE = "RTV" AND YMODELECATREG.YCR_CODECAT  = "'+CbCat.values[CbCat.itemindex]+'" AND YMODELECATREG.YCR_CODEMODELE = "'+strModele+'"'
  //else
  strSQL := 'SELECT YMODELECATREG.YCR_CODEREG, CHOIXCOD.CC_LIBELLE FROM YMODELECATREG INNER JOIN CHOIXCOD ON YMODELECATREG.YCR_CODEREG = CHOIXCOD.CC_CODE WHERE YMODELECATREG.YCR_REGGERE = "X" AND CHOIXCOD.CC_TYPE = "RTV" AND YMODELECATREG.YCR_CODEMODELE = "'+strModele+'"';
  OQuery := OpenSql (strSQL, TRUE);
  try
  i := 0;
  va := '<<Tous>>';
  it := TraduireMemoire('<<Tous>>');
  cbReg.Values.Add(va);
  cbReg.Items.Add(it);
  While not OQuery.Eof do
  begin
    St := OQuery.Fields[0].AsString;
    St1 := OQuery.Fields[1].AsString;
    va:=ReadTokenSt(St) ;
    it:=TraduireMemoire(ReadTokenSt(St1)) ;
    cbReg.Values.Add(va) ;
    cbReg.Items.Add(it) ;
    i := i + 1;
    OQuery.next;
  end;
  finally
    Ferme(OQuery);
  end;

end;


procedure TOF_YYDETAILTAXES.AppelFormule(Sender: TObject);
var
   tTob : TOB;
   TN : TTreeNode ;
   TN1 : TTreeView ;
begin
   TT := TTreeview.Create(ecran);
   TT.Parent:=ecran;
   TT.Visible := false;

   tTob := TOB.Create('Formule',nil,-1);
   tTob.ClearDetail;
   tTob.LoadDetailDBFromSql ('', 'SELECT CO_CODE, CO_LIBELLE FROM COMMUN WHERE CO_TYPE = "FOR"');

   TN1 := TTreeView.create(nil);
   CreateTV(TN,tTob,TN1) ;

   Formule.Text := EditeFormule(TT, Formule.Text);

   FreeAndNil (TN1);

end;

Procedure TOF_YYDETAILTAXES.CreateTV(TN : ttreeNode; Var TobFinale : Tob ; Var TV : tTreeView ) ;
Var i : Integer ;
    OEdt: TOedt;
    p, n: TTreeNode;
    Liste, Liste2 : TStrings;

BEGIN

   p := HTreeNode.Create (nil);
   n := HTreeNode.Create (nil);
   p := TT.Items.Add(nil, ' Champs disponibles');
   Liste := TstringList.Create;
   Liste2 := TstringList.Create;

   for i := 0 to TobFinale.Detail.Count - 1 do
   begin
      Liste.add (TobFinale.Detail[i].GetValue('CO_CODE'));
      Liste2.add(TobFinale.Detail[i].GetValue('CO_LIBELLE'));
   end;

   for  i := 0 to Liste.Count - 1  do
   begin
   try
     OEdt := TOedt.Create;
     OEdt.Quoi := oeChamp;

     OEdt.Code := Liste.Strings[i];
     OEdt.Libelle := Liste2.Strings[i];
     OEdt.TypeChamp := ChampToType(OEdt.Code);
     if (OEdt.Libelle = '') or (OEdt.Libelle = '??') then OEdt.Libelle := OEdt.Code;
      OEdt.Controle := nil;
      n := TT.Items.AddChildObject(p, OEdt.Libelle, OEdt);
      n.ImageIndex := 2;
      n.SelectedIndex := n.ImageIndex;
   finally

   end;
end;
END ;

procedure TOF_YYDETAILTAXES.Cherche(Sender: TObject);
var
   strModele, strCat, strTyp, strReg : string;
   strDateDeb, strDateFin : TDatetime;
   strDate1 : string;
   bVal : boolean;
   strMontantDeb, strMontantFin : string;
   dbMontantDeb, dbMontantFin : double;
begin

  strModele := '';
  strCat := '';
  strTyp := '';
  strReg := '';
  bVal := false;
  if cbCat.itemindex <> -1 then
    strCat := cbCat.values[cbCat.itemindex];
  if cbTyp.itemindex <> -1 then
    strTyp := cbTyp.values[cbTyp.itemindex];
  if cbReg.itemindex <> -1 then
    strReg := cbReg.values[cbReg.itemindex];
  if (strModele <> '') or (strCat <> '') or (strTyp <> '') or (strReg <> '') then
    bVal := true;

  strDate1 := Perdeb1.Text;
  if not isDateVide (strDate1) then
    strDateDeb := strtodate(strDate1);

  strDate1 := Perfin1.Text;
  if not isDateVide (strDate1) then
    strDateFin := strtodate(strDate1);

  strMontantDeb := GetControltext('CCT_ASSDEB1');
  if strMontantDeb = '' then
    strMontantDeb := '0';
  dbMontantDeb := strTofloat (strMontantDeb);
  strMontantDeb := StringReplace(strMontantDeb,',','.',[rfReplaceAll]);

  strMontantFin := GetControltext('CCT_ASSFIN1');
  if strMontantFin = '' then
    strMontantFin := '0';
  dbMontantFin := strTofloat (strMontantFin);
  strMontantFin := StringReplace(strMontantFin,',','.',[rfReplaceAll]);

  try
  if (Ecran <> Nil) and ( Ecran is TFSaisieList) then
  begin
  TF := TFSaisieList(Ecran).LeFiltre;
  TF.Changed := True;
  //TFSaisieList(Ecran).LeFiltre.Edit;
  if Trim(TF.WhereTable) <> 'WHERE CCT_CODEMODELE=:YMT_CODE' then
    TF.WhereTable := '';
  TFSaisieList(Ecran).LeFiltre.Post;
  TF.Changed := False;
  //TFSaisieList(Ecran).LeFiltre.Edit;
  if Trim(TF.WhereTable) = '' then
    TF.WhereTable := 'WHERE CCT_CODEMODELE=:YMT_CODE';
  if (strCat <> '<<Tous>>') and (strCat <> '') then
    TF.WhereTable := TF.WhereTable + ' and CCT_CODECAT = "'+strCat+'"';
  if (strTyp <> '<<Tous>>') and (strTyp <> '') then
    TF.WhereTable := TF.WhereTable + ' and CCT_CODETYP = "'+strTyp+'"';
  if (strReg <> '<<Tous>>') and (strReg <> '') then
    TF.WhereTable := TF.WhereTable + ' and CCT_CODEREG = "'+strReg+'"';
  if (strDateDeb > 0) and (strDateDeb <> iDate1900) then
    TF.WhereTable := TF.WhereTable + ' and CCT_PERDEB >= "'+ USDateTime(strDateDeb)+ '"';
  if (strDateFin > 0) and (strDateFin <> iDate1900) then
    TF.WhereTable := TF.WhereTable + ' and CCT_PERFIN <= "'+ USDateTime(strDateFin)+ '"';
  if dbMontantDeb > 0 then
    TF.WhereTable := TF.WhereTable + ' and CCT_ASSDEB >= '+ strMontantDeb;
  if dbMontantFin > 0 then
    TF.WhereTable := TF.WhereTable + ' and CCT_ASSFIN <= '+  strMontantFin;
  TFSaisieList(Ecran).LeFiltre.Post;
  TF.Changed := False;
  TF.RefreshLignes;
  end
  except

  end;

  {Fliste.BeginUpdate;
  strSQL := 'SELECT YDETAILTAXES.CCT_CODEMODELE, YDETAILTAXES.CCT_CODECAT, ';
  strSQL := strSQL + 'YDETAILTAXES.CCT_CODETYP, YDETAILTAXES.CCT_CODEREG, ';
  strSQL := strSQL + 'YDETAILTAXES.CCT_PERDEB, YDETAILTAXES.CCT_PERFIN, ';
  strSQL := strSQL + 'YDETAILTAXES.CCT_ASSDEB, YDETAILTAXES.CCT_ASSFIN ';
  strSQL := strSQL + 'FROM YMODELETAXE';
  strSQL := strSQL + ' INNER JOIN YDETAILTAXES ON YMT_CODE = CCT_CODEMODELE';
  strSQL := strSQL + ' WHERE CCT_CODEMODELE = "'+strModele+'"';  }


  {if strModele <> '' then
      //strSQL := strSQL + ' WHERE CCT_CODEMODELE = "'+strModele+'"';
      strSQL := strSQL + ' WHERE YMT_CODE = "'+strModele+'"';
  if (strCat <> '') and (strModele <> '') then
      strSQL := strSQL + ' AND CCT_CODECAT = "'+strCat+'"';
  if (strCat <> '') and (strModele = '') then
      strSQL := strSQL + ' CCT_CODECAT = "'+strCat+'"';
  if (strTyp <> '') then
      strSQL := strSQL + ' AND CCT_CODETYP = "'+strTyp+'"';
  if (strReg <> '') then
      strSQL := strSQL + ' AND CCT_CODEREG = "'+strReg+'"'; }


 { strDate1 := Perdeb1.Text;
  if not isDateVide (strDate1) then
    strDateDeb := strtodate(strDate1);

  if strDateDeb > 0 then bVal := true;

  if (bVal) and (strDateDeb > 0) then
      strSQL := strSQL + ' AND CCT_PERDEB >= "'+ USDateTime(strDateDeb)+ '"';
  if (not bVal) and (strDateDeb > 0) then
      strSQL := strSQL + ' CCT_PERDEB >= "'+USDateTime(strDateDeb)+'"';

  strDate1 := Perfin1.Text;
  if not isDateVide (strDate1) then
    strDateFin := strtodate(strDate1);

  if (bVal) and (strDateFin > 0 )  then
      strSQL := strSQL + ' AND CCT_PERFIN <= "'+ USDateTime(strDateFin)+ '"';
  if (not bVal) and (strDateFin > 0) then
      strSQL := strSQL + ' CCT_PERFIN <= "'+ USDateTime(strDateFin)+ '"'; }

 { strMontantDeb := GetControltext('CCT_ASSDEB1');
  dbMontantDeb := strTofloat (strMontantDeb);
  strMontantDeb := StringReplace(strMontantDeb,',','.',[rfReplaceAll]);

  strMontantFin := GetControltext('CCT_ASSFIN1');
  dbMontantFin := strTofloat (strMontantFin);
  strMontantFin := StringReplace(strMontantFin,',','.',[rfReplaceAll]); }

  {if (dbMontantDeb > 0) and (bVal) or (strDateDeb > 0) or (strDateFin > 0)  then
      strSQL := strSQL + ' AND CCT_ASSDEB >= '+ strMontantDeb;
  if (dbMontantDeb > 0) and (not bVal) and (strDateDeb = 0) and (strDateFin = 0)  then
      strSQL := strSQL + ' CCT_ASSDEB >= '+ strMontantDeb;
  if (dbMontantFin > 0) and (bVal) and (strDateDeb > 0) or (strDateFin > 0) or (dbMontantDeb > 0)  then
      strSQL := strSQL + ' AND CCT_ASSFIN <= '+  strMontantFin;
  if (dbMontantFin > 0) and (not bVal) and (strDateDeb = 0) and (strDateFin = 0) and (dbMontantDeb = 0)  then
      strSQL := strSQL + ' CCT_ASSFIN <= '+  strMontantFin; }


  {FListe.VidePile(False) ;
  OQuery := OpenSql (strSQL, TRUE);
  try
  While Not OQuery.Eof do
   BEGIN
   for i := 0 to 7 do
    FListe.Cells[i+1,FListe.RowCount-1]:=OQuery.Fields[i].AsString ;

   FListe.RowCount:=FListe.RowCount+1 ;
   OQuery.Next ;
   END ;
  if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount-1 ;
  finally
    Ferme(OQuery) ;
  end;
  Fliste.EndUpdate; }
  //TF.DataSet.Refresh;

  //ChargeLibCat (strModele);
  //TPageControl( GetControl('PCPied', True)).Refresh;

end;

procedure TOF_YYDETAILTAXES.InitialiserTableDetails (Sender: TObject);
var
    strSQL, strModele, strCategorie, strType, strRegime : string;
    OQuery : TQuery ;
    strDate : TDatetime;
    nb, nbFin,  i: integer;
    bInsertion : boolean;
begin
  bInsertion := false;

  StrSQL := 'SELECT YMODELECATTYP.YCY_CODEMODELE, YMODELECATTYP.YCY_CODECAT, YMODELECATTYP.YCY_CODETYP, YMODELECATREG.YCR_CODEREG';
  strSQL := strSQL + ' FROM YMODELECATTYP INNER JOIN YMODELECATREG ';
  strSQL := strSQL + ' ON YMODELECATTYP.YCY_CODEMODELE = YMODELECATREG.YCR_CODEMODELE AND ';
  strSQL := strSQL + ' YMODELECATTYP.YCY_CODECAT = YMODELECATREG.YCR_CODECAT ';
  strSQL := strSQL + ' AND YMODELECATTYP.YCY_TYPGERE = "X" ';
  strSQL := strSQL + ' AND YMODELECATREG.YCR_REGGERE = "X" ';



  OQuery := OpenSql (strSQL, TRUE);
  try
  if OQuery.Eof then
    begin
      PGIBox(TraduireMemoire(MessageListe[0]),TraduireMemoire(Ecran.caption)); //Message utilisateur
      //Ferme(OQuery);
      pPasDonnee := true;
      exit;
    end;

  strDate := GetMinAnneeExercice;
  nb := 0;
  nbFin := 999999999;
  While not OQuery.Eof do
  begin
    strModele := OQuery.Fields[0].AsString;
    strCategorie := OQuery.Fields[1].AsString;
    strType := OQuery.Fields[2].AsString;
    strRegime := OQuery.Fields[3].AsString;
    if not ExisteSQL ('SELECT 1 FROM YDETAILTAXES WHERE YDETAILTAXES.CCT_CODEMODELE = "'+strModele+'" AND YDETAILTAXES.CCT_CODECAT = "'+strCategorie+'" AND YDETAILTAXES.CCT_CODETYP = "'+strType+'" AND YDETAILTAXES.CCT_CODEREG = "'+strRegime+'"') then
    begin
    //Insertion ligne dite principale
    ExecuteSQL('INSERT INTO YDETAILTAXES (CCT_CODEMODELE, CCT_CODECAT, CCT_CODETYP, CCT_CODEREG,  CCT_PERDEB, CCT_PERFIN, CCT_ASSDEB, CCT_ASSFIN, CCT_TXMTACHAT, CCT_TXMTVENTE, CCT_MTMINI) VALUES ( "'
                        + strModele + '","'
                        + strCategorie + '","'
                        + strType + '","'
                        + strRegime + '","'
                        + USDateTime(strDate) + '","'
                        + USDateTime(iDate2099) + '","'
                        + inttostr(nb) + '","'
                        + inttostr(nbFin) + '","'
                        + inttostr(nb) + '","'
                        + inttostr(nb) + '","'
                        + inttostr(nb) + '") ' );
    bInsertion := true;
    end;
    OQuery.next;
  end;
  finally
    Ferme(OQuery);
  end;

  if binsertion then
  begin
    FListe.VidePile(False) ;
    strSQL := 'SELECT YDETAILTAXES.CCT_CODEMODELE, YDETAILTAXES.CCT_CODECAT, YDETAILTAXES.CCT_CODETYP, YDETAILTAXES.CCT_CODEREG, YDETAILTAXES.CCT_PERDEB, YDETAILTAXES.CCT_ASSDEB, YDETAILTAXES.CCT_TXMTACHAT, YDETAILTAXES.CCT_TXMTVENTE, YDETAILTAXES.* FROM YDETAILTAXES';
    strSql := strSql + ' WHERE YDETAILTAXES.CCT_CODEMODELE = "'+strModele+'"';
    OQuery := OpenSql (strSQL, TRUE);
    try
    While Not OQuery.Eof do
    BEGIN
    for i := 0 to 7 do
      FListe.Cells[i+1,FListe.RowCount-1]:=OQuery.Fields[i].AsString ;

      FListe.RowCount:=FListe.RowCount+1 ;
      OQuery.Next ;
     END ;
    if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount-1 ;
    finally
      Ferme(OQuery) ;
      PGIBox(TraduireMemoire(MessageListe[9]),TraduireMemoire(Ecran.caption)); //Message utilisateur
      Ecran.Close;
      if IsInside (Ecran) then
        CloseInsidePanel (Ecran) ;
    end;

    cbCategorie.Enabled := false;
    cbType.Enabled := false;
    cbRegime.Enabled := false;
  end
  else
    PGIBox(TraduireMemoire(MessageListe[17]),TraduireMemoire(Ecran.caption)); //Message utilisateur
end ;

procedure TOF_YYDETAILTAXES.ChercheAvecCritere(strCritereMod, strCritereCat, strCritereTyp, strCritereReg : string);
begin

  try
  TF := TFSaisieList(Ecran).LeFiltre;
  TF.Changed := True;
  TFSaisieList(Ecran).LeFiltre.Edit;
  TF.WhereTable := '';
  if Trim(TF.WhereTable) = '' then
    TF.WhereTable := 'WHERE CCT_CODEMODELE=:YMT_CODE';
  if strCritereCat <> '' then
    TF.WhereTable := TF.WhereTable + ' and CCT_CODECAT = "'+strCritereCat+'"';
  if strCritereTyp <> '' then
    TF.WhereTable := TF.WhereTable + ' and CCT_CODETYP = "'+strCritereTyp+'"';
  if strCritereReg <> '' then
    TF.WhereTable := TF.WhereTable + ' and CCT_CODEREG = "'+strCritereReg+'"';
  TFSaisieList(Ecran).LeFiltre.Post;
  TF.RefreshLignes;
  except
  end;

  {try
    Fliste.VidePile(False);
    for j := 1 to Fliste.RowCount - 1 do
      Fliste.Rows[j].Clear;
  finally
  end;
  strSQL := 'SELECT YDETAILTAXES.CCT_CODEMODELE, YDETAILTAXES.CCT_CODECAT, ';
  strSQL := strSQL + ' YDETAILTAXES.CCT_CODETYP, YDETAILTAXES.CCT_CODEREG, ';
  strSQL := strSQL + ' YDETAILTAXES.CCT_PERDEB, YDETAILTAXES.CCT_PERFIN,';
  strSQL := strSQL + ' YDETAILTAXES.CCT_ASSDEB, YDETAILTAXES.CCT_ASSFIN, YDETAILTAXES.*';
  strSQL := strSQL + ' FROM YDETAILTAXES ';
  strSQL := strSQL + ' WHERE YDETAILTAXES.CCT_CODEMODELE = "'+strCritereMod + '"';
  strSQL := strSQL + ' AND YDETAILTAXES.CCT_CODECAT = "' + strCritereCat + '"';
  strSQL := strSQL + ' AND YDETAILTAXES.CCT_CODETYP = "' + strCritereTyp + '"';
  strSQL := strSQL + ' AND YDETAILTAXES.CCT_CODEREG = "' + strCritereReg + '"';

  OQuery := OpenSql (strSQL, TRUE);
  try
  While Not OQuery.Eof do
   BEGIN
   for i := 0 to 7 do
    FListe.Cells[i+1,FListe.RowCount-1]:=OQuery.Fields[i].AsString ;

   FListe.RowCount:=FListe.RowCount+1 ;
   OQuery.Next ;
   END ;
  if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount-1 ;
  finally
    Ferme(OQuery) ;
  end;

  ToolCritere.Visible := TRUE;
  cbMod.value := strCritereMod;
  cbCat.value := strCritereCat;
  cbTyp.value := strCritereTyp;
  cbReg.value := strCritereReg;

   T := ToolCritere.Handle;

    for i := 0 to TToolWindow97(GetControl('ToolCritere',true)).ControlCount-1 do
    begin
        if (TToolWindow97(GetControl('ToolCritere',true)).Controls[i] is TToolBarButton97) then
        begin
            Bouton := TToolBarButton97(TToolWindow97(GetControl('ToolCritere',true)).Controls[i]);
            szName := Uppercase(Bouton.Name);
            if (szName = 'BCHERCHE') then
            begin

              PostMessage(
              T, //Handle de la fenêtre mère
              WM_COMMAND, //message WM_COMMAND
              MAKELONG(Bouton.tag, BN_CLICKED),//type de commande: clic sur bouton CtrlID
              GetDlgItem(T, bouton.tag) //handle du bouton
              );
              ToolCritere.Visible := false;
            end;
        end;
    end;
    Fliste.OnClick(nil);
    TPageControl( GetControl('PCPied', True)).Refresh; }

end;

procedure TOF_YYDETAILTAXES.GriserElements(bValeur : boolean);
begin
  //Premier onglet
  MtMini.Enabled := bValeur;
  if strCritereMod = '' then
    Perfin.Enabled := bValeur;
  Assfin.Enabled := bValeur;
  TxMontantAchat.Enabled := bValeur;
  TxMontantVente.Enabled := bValeur;

  //Deuxième onglet
  if strCritereMod = '' then
  begin
    CbPays.Enabled := bValeur;
    CbRegion.Enabled := bValeur;
    CbCodePostal.Enabled := bValeur;
    CbVille.Enabled := bValeur;
  end;

  //Troisième onglet
  if strCritereMod = '' then
  begin
    CompteAchat.Enabled := bValeur;
    CompteVente.Enabled := bValeur;
    CompteAchatEncaissement.Enabled := bValeur;
    CompteVenteEncaissement.Enabled := bValeur;
  end;

  //Quatrième onglet
  if strCritereMod = '' then
    Formule.Enabled := bValeur;

  //Boutons d'actions
  if strCritereMod = '' then //Mode consultation
  begin
    BInsert.Visible := bValeur;
    BDelete.Visible := bValeur;
    BDupliquer.Visible := bValeur;
  end;

end;

function TOF_YYDETAILTAXES.GetMinAnneeExercice : TDateTime ;
var
   SQL : String;
   Q : TQuery;
begin
   result := 1900;
   SQL := 'SELECT MIN(EX_DATEDEBUT) FROM EXERCICE';
   try
     Q := OpenSQL(SQL,true);
     If Not Q.Eof then
     begin
       result := Q.Fields[0].AsDateTime;
     end;
   finally
     ferme(Q);
   end;
end;


function TOF_YYDETAILTAXES.bDonnee(strModele: string): boolean;
var
   SQL : String;
   Q : TQuery;
begin
   result := false;
   //SQL := 'SELECT 1 FROM YDETAILTAXES WHERE CCT_CODEMODELE = "' + strModele + '"';
   SQL := 'SELECT 1 FROM YMODELECATTYP INNER JOIN YMODELECATREG ON ';
   SQL := SQL + 'YCY_CODEMODELE = YCR_CODEMODELE ';
   SQL := SQL + 'WHERE YCY_CODEMODELE = "' + strModele + '" AND YCY_TYPGERE = "X"    ';
   SQL := SQL + 'AND YCR_CODEMODELE = "' + strModele + '" AND YCR_REGGERE = "X"';
   try
     Q := OpenSQL(SQL,true);
     If Not Q.Eof then
     begin
       result := true;
     end;
   finally
     ferme(Q);
   end;
end;

function TOF_YYDETAILTAXES.GetColIndex(vStColName: String): Integer;
var i : Integer ;
begin
  result := -1 ;
  for i := 0 to FListe.ColCount do
    begin
    if FListe.ColNames[i] = vStColName then
      begin
      result := i ;
      Exit ;
      end ;
    end ;
end;

procedure TOF_YYDETAILTAXES.OnbpostClick(Sender: TObject);
var
  strDateDeb, strDateFin : string;
  TDateDeb, TDateFin : TDateTime;
  dAssDeb, dAssFin : double;
  strAssDeb, strAssFin, strCompte : string;
  bTrouveAssiette : Boolean;
begin
  Inherited ;
  //Contrôle des éléments saisis
  //Contrôle de la date de fin d'application
  strDateFin := Perfin.Text;
  if not isDateVide (strDateFin) then
    TDateFin := strtodate(strDateFin);
  strDateDeb := PerDeb.Text;
  if not isDateVide (strDateDeb) then
    TDateDeb := strtodate(strDateDeb);
  if TDateDeb > TDateFin then
  begin
    PGIBox(MessageListe[6],Ecran.Caption);  //message d'erreur
    SetFocusControl('CCT_PERFIN');
    exit;
  end;
  //Contrôle des assiettes
  SetFocusControl('CCT_TXMTVENTE');
  strAssDeb := AssDeb.Text;
  dAssDeb := strtofloat(strAssDeb);

  strAssFin := AssFin.Text;
  dAssFin := strtofloat(strAssFin);
  if dAssDeb > dAssFin then
  begin
    PGIBox(MessageListe[7],Ecran.Caption);  //message d'erreur
    SetFocusControl('CCT_ASSFIN');
    exit;
 end;

  //Contrôle catégorie vide
  if trim (cbCategorie.Value) = '' then
  begin
    PGIBox(MessageListe[12],Ecran.Caption);  //message d'erreur
    SetFocusControl('CCT_CODECAT');
    exit;
  end;

  //Contrôle type vide
  if trim (cbType.Value) = '' then
  begin
    PGIBox(MessageListe[13],Ecran.Caption);  //message d'erreur
    SetFocusControl('CCT_CODETYP');
    exit;
  end;

  //Contrôle régime vide
  if trim (cbRegime.Value) = '' then
  begin
    PGIBox(MessageListe[14],Ecran.Caption);  //message d'erreur
    SetFocusControl('CCT_CODEREG');
    exit;
  end;

  //Contrôle "chevauchement" assiette avec donnée existante en BD
  try
  bTrouveAssiette := bChercheChevauchementAssiette (CodeModele.Text, cbcategorie.Values [cbcategorie.itemindex],cbType.Values [cbType.itemindex],cbRegime.Values [cbRegime.itemindex], strAssDeb, strAssFin, TDateDeb, TDateFin, True);
  if bTrouveAssiette then
  begin
    PGIBox(MessageListe[10],Ecran.Caption);  //message d'erreur
    SetFocusControl('CCT_ASSDEB');
  end;
  except
  end;

  try
  bTrouveAssiette := bChercheChevauchementAssiette (CodeModele.Text, cbcategorie.Values [cbcategorie.itemindex],cbType.Values [cbType.itemindex],cbRegime.Values [cbRegime.itemindex], strAssDeb, strAssFin, TDateDeb, TDateFin, False);
  if bTrouveAssiette then
  begin
    PGIBox(MessageListe[11],Ecran.Caption);  //message d'erreur
    SetFocusControl('CCT_ASSFIN');
  end;
  except
  end;

  //Contrôle pays saisi
  if trim (cbpays.Text) <> '' then
  begin
  if not ExisteSQL ('SELECT 1 FROM PAYS WHERE PY_PAYS = "'+cbPays.text+'"') then
  begin
    PGIBox(MessageListe[15],Ecran.Caption);  //message d'erreur
    SetFocusControl('CCT_PAYS');
    exit;
  end
  end;

  //Contrôle région saisie
  if trim (cbRegion.Text) <> '' then
  begin
  if not ExisteSQL ('SELECT 1 FROM REGION WHERE RG_REGION = "'+cbRegion.text+'"') then
  begin
    PGIBox(MessageListe[16],Ecran.Caption);  //message d'erreur
    SetFocusControl('CCT_REGION');
    exit;
  end
  end;

  //Contrôle des comptes saisis
  strCompte := CompteAchat.Text;
  if trim (strCompte) <> '' then
  begin
  if not ExisteSQL ('SELECT 1 FROM GENERAUX WHERE G_GENERAL = "'+strCompte+'"') then
  begin
    PGIBox(MessageListe[8],Ecran.Caption);  //message d'erreur
    SetFocusControl('CCT_CPTACHAT');
    exit;
  end
  end;

  strCompte := CompteVente.Text;
  if trim (strCompte) <> '' then
  begin
  if not ExisteSQL ('SELECT 1 FROM GENERAUX WHERE G_GENERAL = "'+strCompte+'"') then
  begin
    PGIBox(MessageListe[8],Ecran.Caption);  //message d'erreur
    SetFocusControl('CCT_CPTVENTE');
    exit;
  end
  end;

  strCompte := CompteAchatEncaissement.Text;
  if trim (strCompte) <> '' then
  begin
  if not ExisteSQL ('SELECT 1 FROM GENERAUX WHERE G_GENERAL = "'+strCompte+'"') then
  begin
    PGIBox(MessageListe[8],Ecran.Caption);  //message d'erreur
    SetFocusControl('CCT_ENCACHAT');
    exit;
  end
  end;

  strCompte := CompteVenteEncaissement.Text;
  if trim (strCompte) <> '' then
  begin
  if not ExisteSQL ('SELECT 1 FROM GENERAUX WHERE G_GENERAL = "'+strCompte+'"') then
  begin
    PGIBox(MessageListe[8],Ecran.Caption);  //message d'erreur
    SetFocusControl('CCT_ENCVENTE');
    exit;
  end
  end;
  FValidation(Sender);

end;

function TOF_YYDETAILTAXES.bVide(strModele: string): boolean;
begin
  Result := false;
  if not ExisteSQL ('SELECT 1 FROM YDETAILTAXES WHERE YDETAILTAXES.CCT_CODEMODELE = "'+strModele+'"') then
    result := true;

end;

procedure TOF_YYDETAILTAXES.AfficheModele(Sender: TObject);
var strModele : string;
begin
//strModele := TTModele.Selected.text;
strModele := trim(copy(TTModele.Selected.text,1,pos('(', TTModele.Selected.text) - 1));
if strModele <> TraduireMemoire('Modèles de taxes') then
  AGLLanceFiche('YY','YYMODELETAXE','', '', 'CODEMOD='+ strModele);
end;

function TOF_YYDETAILTAXES.bChercheChevauchementAssiette (strModele, strCategorie, strType, strRegime, strAssDeb, strAssFin : string; Tperdeb, Tperfin : TDatetime; bDebut : boolean) : boolean;
var
   strSQL , strMontantDeb, strMontantFin : string;
   dbMontantDeb, dbMontantFin : double;
   OQuery : TQuery ;
begin
   result := false;

   dbMontantDeb := strTofloat (strAssDeb);
   strMontantDeb := StringReplace(strAssDeb,',','.',[rfReplaceAll]);

   dbMontantFin := strTofloat (strAssFin);
   strMontantFin := StringReplace(strAssFin,',','.',[rfReplaceAll]);


   strSQL := 'SELECT 1 FROM YDETAILTAXES WHERE ';
   strSQL := strSQL + 'CCT_CODEMODELE = "'+strModele+'"';
   strSQL := strSQL + ' AND CCT_CODECAT = "'+strCategorie+'"';
   strSQL := strSQL + ' AND CCT_CODETYP = "'+strType+'"';
   strSQL := strSQL + ' AND CCT_CODEREG = "'+strRegime+'"';
   if bDebut then
   begin
    strSQL := strSQL + ' AND CCT_PERDEB <= "'+ USDateTime(TperDeb)+ '"';
    strSQL := strSQL + ' AND CCT_PERFIN >= "'+ USDateTime(TperDeb)+ '"';
    strSQL := strSQL + ' AND CCT_ASSDEB <= '+ strMontantDeb;
    strSQL := strSQL + ' AND CCT_ASSFIN >= '+  strMontantDeb;
   end
   else
   begin
    strSQL := strSQL + ' AND CCT_PERDEB <= "'+ USDateTime(TperFin)+ '"';
    strSQL := strSQL + ' AND CCT_PERFIN >= "'+ USDateTime(TperFin)+ '"';
    strSQL := strSQL + ' AND CCT_ASSDEB <= '+ strMontantFin;
    strSQL := strSQL + ' AND CCT_ASSFIN >= '+  strMontantFin;
   end;
   strSql := strSql + ' AND CCT_PERDEB';
   strSql := strSql + ' NOT IN (SELECT CCT_PERDEB FROM YDETAILTAXES ';
   strSql := strSql + ' WHERE CCT_CODEMODELE = "'+strModele+'"';
   strSql := strSql + ' AND CCT_CODECAT = "'+strCategorie+'"';
   strSQL := strSQL + ' AND CCT_CODETYP = "'+strType+'"';
   strSQL := strSQL + ' AND CCT_CODEREG = "'+strRegime+'"';
   strSQL := strSQL + ' AND CCT_PERDEB = "'+ USDateTime(TperDeb)+ '"';
   strSQL := strSQL + ' AND CCT_PERFIN = "'+ USDateTime(TperFin)+ '"';
   strSQL := strSQL + ' AND CCT_ASSDEB = '+ strMontantDeb;
   //strSQL := strSQL + ' AND CCT_ASSFIN = '+ strMontantFin;
   strSQL := strSQL + ' ) AND CCT_ASSDEB ';
   strSQL := strSQL + ' NOT IN (SELECT CCT_ASSDEB FROM YDETAILTAXES';
   strSql := strSql + ' WHERE CCT_CODEMODELE = "'+strModele+'"';
   strSql := strSql + ' AND CCT_CODECAT = "'+strCategorie+'"';
   strSQL := strSQL + ' AND CCT_CODETYP = "'+strType+'"';
   strSQL := strSQL + ' AND CCT_CODEREG = "'+strRegime+'"';
   strSQL := strSQL + ' AND CCT_PERDEB = "'+ USDateTime(TperDeb)+ '"';
   strSQL := strSQL + ' AND CCT_PERFIN = "'+ USDateTime(TperFin)+ '"';
   strSQL := strSQL + ' AND CCT_ASSDEB = '+ strMontantDeb;
   //strSQL := strSQL + ' AND CCT_ASSFIN = '+ strMontantFin;
   strSQL := strSQL + ')';

   OQuery := OpenSql (strSQL, TRUE);
   try
   if not OQuery.Eof then
      result := True;
   finally
    Ferme(OQuery);
   end;
end;


procedure TOF_YYDETAILTAXES.ChargeLibCat (strModele : string);
var
   strLib, strLibCat, strLibCat2: string;
   oInt : TInternationalTax;
   x, i : integer;
   Arg : string;

begin
//Chargement libellé catégorie dans grille

  try
  strLibCat := oInt.tstrGetLibelleTaxesModele(strModele);
  strLibCat2 := strLibCat;
  libcat.Caption := '';
  i := 0;
  repeat
    strLib := uppercase(Trim(ReadTokenSt(strLibCat)));
    if strLib <> '' then
    begin
      x := pos(';',strLibCat2);
      if x <> 0 then
      begin
        Arg := strLib;
        if i = 0 then
          LibCat.caption := 'Concernant la colonne "Catégorie" : TX1 = ' + Arg
        else
          LibCat.caption := LibCat.caption + '; ' + 'TX' + inttostr(i+1) + ' = ' + Arg;
        i := i + 1;

      end;
    end;
  until strLib = '';
  finally
    FreeAndNil (oInt);
  end;
  LibCat.caption := traduireMemoire (LibCat.caption);


end;

procedure TOF_YYDETAILTAXES.OnFermeClick(Sender: TObject);
begin
  if (strCritereMod <> '')
     and (strCritereCat <> '')
     and (strCritereTyp <> '')
     and (strCritereReg <> '') then
  begin
  TToolbarButton97(GetControl('Bdefaire')).Enabled := True;
  TToolbarButton97(GetControl('Bdefaire')).click;

  if IsInside (Ecran) then
    CloseInsidePanel (Ecran)
  else
    Ecran.Close;
  exit;
  end;
  FFerme(sender);
end;

procedure TOF_YYDETAILTAXES.onChangePerdeb1(Sender: TObject);
var
  strDateDeb, strDateFin : string;
  TDateDeb, TDateFin : TDateTime;
begin
  if isDateVide (PerFin1.Text) then
    Perfin1.Text := DateToStr(iDate1900);
    
  if (Perfin1.Text = DateToStr(iDate1900)) then
    perfin1.Text := PerDeb1.Text;

  strDateDeb := PerDeb1.Text;
  TDateDeb := strtodate(strDateDeb);

  strDateFin := PerFin1.Text;
  TDateFin := strtodate(strDateFin);

  if TDateFin < TDateDeb then
    PerFin1.Text := PerDeb1.Text;

end;

procedure TOF_YYDETAILTAXES.onChangePerFin1(Sender: TObject);
var
  strDateDeb, strDateFin : string;
  TDateDeb, TDateFin : TDateTime;
begin

  //if (PerDeb1.Text = DateToStr(iDate1900)) then
  //  PerDeb1.Text := Perfin1.Text;

  strDateDeb := PerDeb1.Text;
  TDateDeb := strtodate(strDateDeb);

  strDateFin := PerFin1.Text;
  TDateFin := strtodate(strDateFin);

  if TDateDeb > TDatefin then
    PerDeb1.Text := Perfin1.Text;


end;

procedure TOF_YYDETAILTAXES.onChangeAssdeb1(Sender: TObject);
var
  strMontantDeb, strMontantFin : string;
  dbMontantDeb, dbMontantFin : double;
begin

  strMontantDeb := GetControltext('CCT_ASSDEB1');
  dbMontantDeb := strTofloat (strMontantDeb);

  strMontantFin := GetControltext('CCT_ASSFIN1');
  dbMontantFin := strTofloat (strMontantFin);

  if dbMontantFin < dbMontantDeb then
    AssFin1.Text := AssDeb1.Text;


end;

procedure TOF_YYDETAILTAXES.onChangeAssfin1(Sender: TObject);
var
  strMontantDeb, strMontantFin : string;
  dbMontantDeb, dbMontantFin : double;
begin

  strMontantDeb := GetControltext('CCT_ASSDEB1');
  dbMontantDeb := strTofloat (strMontantDeb);

  strMontantFin := GetControltext('CCT_ASSFIN1');
  dbMontantFin := strTofloat (strMontantFin);

  if dbMontantDeb > dbMontantFin then
    AssDeb1.Text := Assfin1.Text;

end;

procedure TOF_YYDETAILTAXES.FormKeyDown(Sender : TObject ; var Key : Word ; Shift : TShiftState);
begin
  case Key of
    VK_F1: //lance l'aide
      begin
        Key := 0;
        Ecran.HelpContext := 0;   // en attendant l'implémentation dans l'aide
      end;

    VK_F9: // lance la recherche
      begin
        Key := 0;
        Cherche(nil);
      end;
    17: // CTRL + Z  initialisation des zones du filtre
      begin
        if Shift = [ssCtrl] then
        begin
          Key := 0;
          perdeb1.Text := DateToStr(iDate1900);
          Perfin1.Text := DateToStr(iDate2099);
          cbCat.ItemIndex := 0;
          cbTyp.ItemIndex := 0;
          cbReg.ItemIndex := 0;
          assdeb1.Text := inttostr(0);
          assfin1.Text := inttostr(0);
        end;
      end;
  end;
end;


Initialization
  registerclasses ( [ TOF_YYDETAILTAXES ] ) ;

{$ENDIF} 
end.

 