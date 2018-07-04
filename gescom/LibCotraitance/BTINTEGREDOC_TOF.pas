{***********UNITE*************************************************
Auteur  ...... : Franck VAUTRAIN
Créé le ...... : 19/07/2011
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTINTEGREDOC ()
Mots clefs ... : TOF;BTINTEGREDOC
*****************************************************************}
Unit BTINTEGREDOC_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     UTOF,
     EtudePiece,
     Dialogs,
     Paramsoc,
     Splash,
     Variants,
     UtilsRapport,
     HRichOLE,
     uEntCommun,
     Udefexport,
     SAISUTIL,
     PiecesRecalculs
     ;

Type
  TTypeErreur = (TteWARNING,TteERROR);
  TTypeImport = (TTImport,Ttddeprix,TtSsTraite,TtTarifFr);

  TOF_BTINTEGREDOC = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  private
    fTTypeImport : TTYpeImport;

  	ArtEcart  : string;
    BValider  : TToolbarButton97;
    BTExcel   : TToolbarButton97;
    BTImprime : TToolbarButton97;
    ChkRapport: THCheckbox;
    XLSName   : THEdit;
    BrowseXLS : TOpenDialog;
    Rapport   : TGestionRapport;
    //
    Splash    : TFsplashScreen;
    //
    Tobdoc    : TOB;
    TobOuvrages: TOB;
    TobErreur : TOB;
    TOBDET    : TOB;
    TOBXLS    : TOB;
    TOBLXLS   : TOB;
    //
    WinXLS    : OleVariant;
    Workbook  : OleVariant;
    WinNew    : Boolean;
    ActiveWB      : Variant;
    SheetXLS      : OleVariant;
    SheetName     : string;
    CellValue     : string;
    //..
    //
    procedure AffichageSplash;
    procedure ChargementDdeByXLS;
    procedure ChargementDocByXLS(LaPiece: TRecalculPiece);
    procedure ChargeTOBXLSSsTraite(TOBXLS: TOB; CleLigXls: R_CLEXLS);
    procedure ChargeTOBXLSDdePrix (TOBXLS: TOB; CleLigXls: R_CLEXLS);
    procedure ChercheFichierXLS(sender: TObject);
    procedure ControleChamp(Champ, Valeur: String);
    procedure ControleCritere(Critere: String);
    procedure ControleFichierXLS;
    procedure FinTraiteImport;
    procedure GestionErreur(LibError: string);
    procedure GestionRapport(sender: TObject);
    procedure ImportFichierXLS(sender: TObject);
    procedure ImprimeFichierXLS(sender: TObject);
    procedure LogErreur(numErreur : integer; TypeErreur : TTypeErreur);
    procedure OuvrirFichierXLS(sender: TObject);
    //
    function AddligneXls (TOBXLS : TOB) : TOB;
    function ControlePiece(cledoc: r_cledoc): integer;
    function FindDoc(TOBXLS: TOB): TOB;
    function OuvertureApplicationExcel: boolean;
    function RechercheRange(Etiquette: string; var LibEtiq: string; var LigEtiq, ColEtiq: Integer): Boolean;
    //
  end ;

Implementation
  uses  UtilXLSBTP,
        UtilTOBPiece,
        factTOB,
        FactVariante,
        FactOuvrage,
        FactureBtp,
        UCotraitance,
				UtilArticle
         ;


procedure TOF_BTINTEGREDOC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTINTEGREDOC.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTINTEGREDOC.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTINTEGREDOC.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTINTEGREDOC.OnArgument (S : String ) ;
Var critere : String;
    X       : Integer;
    Champ   : string;
    Valeur  : string;
begin
  Inherited ;

  fTTypeImport := TtSsTraite;

  //Récupération valeur de argument
  Critere:=(Trim(ReadTokenSt(S)));

  while (Critere <> '') do
  begin
    if Critere <> '' then
      begin
      X := pos (';', Critere) ;
      if x = 0 then
        X := pos ('=', Critere) ;
      if x <> 0 then
        begin
        Champ := copy (Critere, 1, X - 1) ;
        Valeur:= Copy (Critere, X + 1, length (Critere) - X) ;
        ControleChamp(champ, valeur);
        end
      end;
    ControleCritere(Critere);
    Critere   := (Trim(ReadTokenSt(S)));
  end;

  TOBDET := TOB.create ('LES LIGNES EXCEL',nil,-1);

  if Assigned(GetControl('BTINTEGRATION')) then
  begin
    BValider := TToolbarButton97(ecran.FindComponent('BTINTEGRATION'));
    BValider.OnClick := ImportFichierXLS;
    BValider.Enabled := false;
  end;

  if Assigned(GetControl('BTEXCEL')) then
  begin
    BTExcel := TToolbarButton97(ecran.FindComponent('BTEXCEL'));
    BTExcel.OnClick := OuvrirFichierXLS;
  end;

  if Assigned(GetControl('BIMPRIMER')) then
  begin
    BTImprime := TToolbarButton97(ecran.FindComponent('BIMPRIMER'));
    BTImprime.OnClick := ImprimeFichierXLS;
    Btimprime.visible := False;
  end;

  if Assigned(GetControl('CHKRAPPORT')) then
  begin
    ChkRapport := THCheckbox(ecran.FindComponent('CHKRAPPORT'));
    ChkRapport.OnClick := GestionRapport;
  end;

  if Assigned(GetControl('XLSNAME')) then
  begin
    XLSName := THedit(ecran.FindComponent('XLSNAME'));
    XLSName.OnElipsisClick := ChercheFichierXLS;
  end;

  TobErreur := TOB.Create('ERREURS', Nil, -1);
  TobOuvrages := TOB.Create ('LES OUVRAGES',nil,-1);
  //
  Rapport := TGestionRapport.Create(Ecran);

  if fTTypeImport = TtSsTraite     then
  begin
    Tobdoc := TOB.Create('DOCUMENT', nil, -1);
    Rapport.Titre   := 'Intégration XLS dans Document';
  end
  else if fTTypeImport = Ttddeprix then
  begin
    //Tobdoc := TOB.Create('DDEPRIX', nil, -1);
    Rapport.Titre   := 'Intégration XLS dans demande de Prix';
  end;

  Rapport.Close   := True;
  Rapport.Sauve   := False;
  Rapport.Print   := False;
  //
end ;

procedure TOF_BTINTEGREDOC.OnClose ;
begin
  Inherited ;

  //Fermeture de l'instance Excel
  if fTTypeImport = TtSsTraite then FreeAndNil(Tobdoc);
  if assigned(TobErreur) Then FreeAndNil(TobErreur);
  if assigned(TobOuvrages) Then FreeAndNil(TobOuvrages);
  if assigned(TobDet) Then FreeAndNil (TOBDET);

  //Fin de traitement
  if Assigned(Rapport ) then FreeAndNil(Rapport);

end ;

procedure TOF_BTINTEGREDOC.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTINTEGREDOC.OnCancel () ;
begin
  Inherited ;
end ;

function TOF_BTINTEGREDOC.ControlePiece (cledoc : r_cledoc) : integer;
var TOBPiece : TOB;
begin
  TOBPiece := TOB.Create ('PIECE',nil,-1);
  result := 0;
  TRY
  	LoadPiece (cledoc, TOBPiece);
    if TOBPiece.getValue('AFF_ETATAFFAIRE')='TER' then result := -1
    else if TOBPiece.getValue('GP_VIVANTE')='-' then result := -2
    else if TOBPiece.getValue('AFF_ETATAFFAIRE')='ACP' then result := 1;
  FINALLY
  TOBPiece.free;
	END;
end;

procedure TOF_BTINTEGREDOC.ImportFichierXLS(sender : TObject);
var LaPiece : TRecalculPiece;
begin

  if fTTypeImport = TtSsTraite then
  begin
    ArtEcart := trim(GetParamsoc('SO_BTECARTPMA'));
    if not isExistsArticle(ArtEcart) then
    begin
      PgiBox(TraduireMemoire('L''article d''écart est invalide ou non renseigné#13#10Veuillez le définir'), Traduirememoire('Gestion d''écart'));
      exit;
    end;
  end;

  if not OuvertureApplicationExcel then exit;

  if fTTypeImport = TtSsTraite then
  begin
    LaPiece := TRecalculPiece.create;
    LaPiece.CreateEnv;
  end;

  TOBXLS := TOB.Create ('LES LIGNES XLS',nil,-1);

  TRY
    //Affichage d'un splash de traitement pour patienter...
    AffichageSplash;

    //Contrôle si le fichier existe et si c'est bien un fichier XLS
    ControleFichierXLS;
    //
    if fTTypeImport = TtSsTraite then ChargementDocByXLS(LaPiece)
    else if fTTypeImport = TtddePrix then ChargementDdeByXLS;

    //Fin de traitement
    FinTraiteImport;
  FINALLY
    if fTTypeImport = TtSsTraite then
    begin
      LaPiece.DestroyEnv;
      LaPiece.free;
    end;
    TOBXLS.free;
    if not VarIsEmpty(WinXLS) then WinXLS.Quit;
    WinXLS := unassigned;
  END;

end;

function TOF_BTINTEGREDOC.OuvertureApplicationExcel : boolean;
begin

  Result := True;

  if not OpenExcel (True,WinXLS, WinNew) then
  begin
    GestionErreur(TraduireMemoire('Liaison Excel impossible'));
    XLSName.Text := '';
    Result := false;
  end;

end;


Procedure TOF_BTINTEGREDOC.AffichageSplash;
begin

  splash := TFsplashScreen.Create (ecran);

  if fTTypeImport = TtSsTraite then
    splash.Label1.Caption := 'Renvoi des données Excel vers Document PGI'
  else if fTTypeImport = TtddePrix then
    splash.Label1.Caption := 'Renvoi des données Excel vers Demande de Prix';

  splash.Show;
  splash.Refresh;

  //Gestion de l'entête de rapport
  Rapport.SauveLigMemo(Splash.Label1.Caption);


end;

procedure TOF_BTINTEGREDOC.ControleFichierXLS;
Var Ligne       : Integer;
    Colonne     : Integer;
    //
    CleLigXls   : R_CLEXLS;
    //
    Ind         : Integer;
    INbvide     : Integer;
    LigRefPiece : Integer;
    ColRefpiece : Integer;
    ColUnite    : Integer;
    ColQte      : Integer;
    ColPu       : Integer;
    ColRefArt   : Integer;
    ColNbJour   : Integer;
    //
    RefFourn    : string;
    RefPiece    : String;
    //
    StZone      : string;
    ligWithoutPrice : Integer;
begin

  WorkBook := OpenWorkBook(XLSName.Text, WinXLS, False);
  ActiveWB:= WinXLS.ActiveWorkBook;

  INbVide := 0;
  ligWithoutPrice := 0;

  For ind := 1 to ActiveWB.Sheets.Count do
  begin
    SheetName := ActiveWB.Sheets[ind].Name;
    SheetXLS  := SelectSheet (Workbook, SheetName);
    Splash.Label2.Caption := 'Traitement de la feuille : ' + SheetName;
    Rapport.SauveLigMemo(Splash.Label2.Caption);

    //contrôle de l'existence des etiquette de référence du document XLS
    if not RechercheRange('REFFOURNISSEUR', RefFourn, Ligne, Colonne)     then BEGIN LogErreur(1,TteERROR); Exit; END;
    if not RechercheRange('REFLIGNE', RefPiece, LigRefPiece, ColRefPiece) then BEGIN LogErreur(6,TteERROR); Exit; END;

    if fTTypeImport = TtDdePrix then
    Begin
      if not RechercheRange('REFARTLIGNE', CellValue, ligne, ColRefArt) then Begin LogErreur(5,TteERROR); Exit; End;
      if not RechercheRange('NBJOUR', CellValue, ligne, ColNbJour)      then Begin LogErreur(4,TteERROR); Exit; End;
    End;

    if not RechercheRange('UNITELIGNE', CellValue, Ligne, ColUnite)     then BEGIN LogErreur(7,TteWARNING); Exit; END;
    if not RechercheRange('QTELIGNE', CellValue, Ligne, ColQte)         then BEGIN LogErreur(2,TteERROR); Exit; END;
    if not RechercheRange('PULIGNE', CellValue, Ligne, ColPU)           then BEGIN LogErreur(3,TteERROR); Exit; END;

    //si la reference fournisseur est renseignée on enlève les guillemets en début et fin...
    if RefFourn <> '' then
    begin
      refFourn := StringReplace(RefFourn,'"','',[rfReplaceAll]);
      Rapport.SauveLigMemo('Traitement fournisseur '+ RefFourn);
    end;

    //parcours des lignes de la feuille XLS
    inc(LigRefPiece);

    repeat
      //découpage de la cellule pour récupération Cledoc
      RefPiece := GetExcelFormated(Workbook.ActiveSheet, LigRefPiece, ColRefPiece);
      if RefPiece <> '' then
      begin
        if fTTypeImport = TtSsTraite then
          CellValue := Trim(GetExceltext(Workbook.ActiveSheet, LigRefPiece, ColQte))
        else if fTTypeImport = TtDdePrix then
          CellValue := Trim(GetExceltext(Workbook.ActiveSheet, LigRefPiece, ColPU));
        if (CellValue <> '0') and (CellValue <> '') then //si la qté ou le Pu est à zéro on ne charge pas la ligne => gain de temps...
        begin
          DecodeclefXLS(RefPiece, CleLigXls);
          if fTTypeImport = TtSsTraite then
            ChargeTOBXLSSsTraite(TOBXLS, CleLigXLS)
          else if fTTypeImport = TtDdePrix then
            ChargeTOBXLSDdePrix(TOBXLS, CleLigXLS);
          //
          TOBLXLS.SetString('FOURNISSEUR', RefFourn);

          if ColRefArt <> 0 then TOBLXLS.SetString('REFARTFRS', GetExceltext(Workbook.ActiveSheet, LigRefPiece, ColRefArt));

          if ColUnite <> 0 then TOBLXLS.SetString('UNITE', GetExceltext(Workbook.ActiveSheet, LigRefPiece, ColUnite));
          if ColQte   <> 0 then TOBLXLS.SetDouble('QTE', StrToFloat(Trim(GetExceltext(Workbook.ActiveSheet, LigRefPiece, ColQte))));

          //chargement du prix fiche xls dans tob
          if ColPU <> 0 then
          begin
            StZone := GetExceltext(Workbook.ActiveSheet, LigRefPiece, ColPU);
            if isNumeric(Stzone) then TOBLXLS.SetDouble ('PU', StrToFloat(StZone));
          end;

          //chargement nombre de jour fiche xls dans TOB
          if ColNbjour <> 0 then
          begin
            StZone := GetExceltext(Workbook.ActiveSheet, LigRefPiece, ColNbjour);
            if isNumeric(Stzone) then TOBLXLS.SetDouble  ('NBJOUR',   StrToFloat(StZone));
          end;
          //
          TOBLXLS.SetString  ('TRAITEOK','-');
          //
          INbvide := 0;
        end
        else
          inc(LigWithoutPrice);
      end
      else
        Inc(INbVide);
      inc(LigRefPiece);
    until INBvide > 5;
  end;

  if ligWithoutPrice <> 0 then Rapport.SauveLigMemo(intToStr(LigWithoutPrice) + ' ligne(s) sans prix et/ou sans quantité');
    
end;

Procedure TOF_BTINTEGREDOC.ChargeTOBXLSSsTraite(TOBXLS : TOB; CleLigXls : R_CLEXLS);
begin

  TOBLXLS := AddligneXls (TOBXLS);

  TOBLXLS.SetString  ('NATUREPIECEG', CleLigXls.NaturePiece);
  TOBLXLS.SetString  ('SOUCHE',       CleLigXls.Souche);
  TOBLXLS.SetInteger ('NUMERO',       CleLigXls.NumeroPiece);
  TOBLXLS.SetInteger ('INDICEG',      CleLigXls.Indice);
  TOBLXLS.SetInteger ('NUMORDRE',     CleLigXls.NumOrdre);
  TOBLXLS.SetInteger ('UNIQUEBLO',    CleLigXls.UniqueBlo);

end;

Procedure TOF_BTINTEGREDOC.ChargeTOBXLSDdePrix(TOBXLS : TOB; CleLigXls : R_CLEXLS);
begin

  TOBLXLS := AddligneXls (TOBXLS);

  TOBLXLS.SetString  ('NATUREPIECEG', CleLigXls.NaturePiece);
  TOBLXLS.SetString  ('SOUCHE',       CleLigXls.Souche);
  TOBLXLS.SetInteger ('NUMERO',       CleLigXls.NumeroPiece);
  TOBLXLS.SetInteger ('INDICEG',      CleLigXls.Indice);
  TOBLXLS.SetInteger ('UNIQUE',       CleLigXls.NumOrdre);
  TOBLXLS.SetInteger ('UNIQUELIG',    CleLigXls.UniqueBlo);

end;

Procedure TOF_BTINTEGREDOC.ChargementDocByXLS(LaPiece : TRecalculPiece);
var Cledoc      : R_CLEDOC;
    //
    Ind         : Integer;
    Indice      : Integer;
    ResControle : Integer;
    //
    TOBL        : TOB;
    //
    Infoligne   : RinfoLigne;
begin

  //
  TOBXLS.detail.sort ('NATUREPIECEG;SOUCHE;NUMERO;INDICEG;NUMORDRE;UNIQUEBLO');
  repeat
	  TOBLXLS := TOBXLS.detail[0];
    TOBDOC := FindDoc(TOBLXLS);
    TOBLXLS.ChangeParent(TOBDOC,-1);
  until TOBXLS.detail.count = 0;

  FillChar(cledoc,sizeof(cledoc),0);

  for Ind := 0 to TOBDET.detail.count -1 do
  begin
    TOBDOC := TOBDET.detail[Ind];
    cledoc.NaturePiece  := TOBDOC.getString('NATUREPIECEG');
    cledoc.Souche       := TOBDOC.getString('SOUCHE');
    cledoc.NumeroPiece  := TOBDOC.getInteger('NUMERO');
    cledoc.Indice       := TOBDOC.getInteger('INDICEG');
    Rapport.SauveLigMemo('Traitement du document '+Cledoc.naturePiece+' N° :'+IntToStr(Cledoc.numeroPiece));
    // Chargement de la nouvelle pièce
    LaPiece.ReInitEnv;
    if LaPiece.ChargeLapiece (cledoc) < 0 then
    begin
      Rapport.SauveLigMemo('ERREUR : Le document n''existe pas');
      continue;
    end;
    LaPiece.RecalcPv := true;
    ResControle  :=  ControlePiece (cledoc);
    begin
      if ResControle = -1 then
      begin
        Rapport.SauveLigMemo('ERREUR : Ce document est indiqué comme terminé');
        continue;
      end else if ResControle = -2 then
      begin
        Rapport.SauveLigMemo('ERREUR : Ce document est indiqué comme non vivant');
        continue;
      end else if ResControle = 1 then
      begin
        Rapport.SauveLigMemo('ATTENTION : Ce document est indiqué comme accepté. Seul les prix d''achats seront impactés, ainsi que la marge.');
        LaPiece.RecalcPv := false;
      end;
    end;
    // La on traite le contenu d'un document
    for Indice := 0 to TOBDOC.detail.count -1 do
    begin
      TOBL := TOBDoc.detail[Indice];
      InfoLigne.Ligne := TOBL.GetInteger ('NUMORDRE');
      InfoLigne.UniqueBlo := TOBL.GetInteger ('UNIQUEBLO');
      if TOBL.GetDouble ('QTE') <> 0 then
      begin
        if LaPiece.MajInfoLigne (InfoLigne,TTiQte,TOBL.GetDouble ('QTE')) < 0 then
        begin
          Rapport.SauveLigMemo('ATTENTION : La ligne ');
          continue;
        end;
      end;
      if TOBL.GetDouble ('PU') <> 0 then
      begin
        if LaPiece.MajInfoLigne (InfoLigne,TTipua,TOBL.GetDouble ('PU')) < 0 then
        begin
          Rapport.SauveLigMemo('ATTENTION : La ligne ');
          continue;
        end;
      end;
      if TOBL.GetDouble ('UNITE') <> 0 then
      begin
        if LaPiece.MajInfoLigne (InfoLigne,TTiUnite,TOBL.GetString ('UNITE')) < 0 then
        begin
          Rapport.SauveLigMemo('ATTENTION : La ligne ');
          continue;
        end;
      end;
    end;
    LaPiece.ReCalculeLaPiece;
    if LaPiece.Result <> TrrOk then
    begin
      Rapport.SauveLigMemo('ERREUR : une anomalie s''est produite lors du calcul de la pièce ');
      break;
    end;
    LaPiece.EnregistreLaPiece;
    if LaPiece.Result <> TrrOk then
    begin
      Rapport.SauveLigMemo('ERREUR : une anomalie s''est produite lors de l''enregistrement de la pièce ');
    end else
    begin
      Rapport.SauveLigMemo('Nbre de ligne(s) traitée(s)     : ' + IntToStr(TOBDOC.detail.count));
      Rapport.SauveLigMemo('document enregistré avec succès.');
    end;
  end;

end;

//si click sur elipsis appel recherche fichier XLS sur le disque
procedure TOF_BTINTEGREDOC.ChercheFichierXLS(sender : TObject);
begin

  if ChkRapport.Checked then Rapport.InitRapport;

  BrowseXLS := TOpenDialog.Create(ecran);
  BrowseXLS.Title := 'Recherche fichier XLS à intégrer';
  BrowseXLS.Filter := 'Fichiers Excel|*.XLS;*.XLSX';
  BrowseXLS.InitialDir := GetParamSocSecur('SO_BTDIREXPORTS', 'C:\');

  if ChkRapport.Checked then Rapport.SauveLigMemo(BrowseXLS.Title);

  if BrowseXLS.Execute then
  begin
    XLSName.Text := BrowseXLS.FileName;
    BValider.Enabled := True;
  end
  else
    XLSName.Text := '';


  FreeAndNil(BrowseXLS);

end;

procedure TOF_BTINTEGREDOC.ChargementDdeByXLS;
var iInd      : Integer;
    Cledoc    : R_Cledoc;
    Unique    : Integer;
    UniqueLig : Integer;
    RefFrs    : String;
    StSQL     : String;
    LigMAJ    : Integer;
    LigNonMAJ : Integer;
    StPUAchat : String;
    StNbJour  : String;
begin

  LigMaj    := 0;
  LigNonMaj := 0;

  for iInd := 0 to TOBXLS.detail.count -1 do
  begin
    TOBDOC := TOBXLS.detail[iInd];
    cledoc.NaturePiece  := TOBDOC.getString('NATUREPIECEG');
    cledoc.Souche       := TOBDOC.getString('SOUCHE');
    cledoc.NumeroPiece  := TOBDOC.getInteger('NUMERO');
    cledoc.Indice       := TOBDOC.getInteger('INDICEG');
    Unique              := TOBDOC.getInteger('UNIQUE');
    UniqueLig           := TOBDOC.getInteger('UNIQUELIG');
    RefFrs              := TOBDOC.getString('FOURNISSEUR');
    StPuAchat           := StrfPoint(TOBDOC.GetDouble('PU'));
    StNbJour            := StrfPoint(TOBDOC.GetDouble('NBJOUR'));
    // Chargement de la ligne fournisseur demande de prix
    StSQL := 'UPDATE FOURLIGDEMPRIX SET';
    StSQL := StSQL + ' BD1_REFERENCE="'    + TOBDOC.GetString('REFARTFRS')         + '", ';
    StSQL := StSQL + ' BD1_PRIXACH='       + StPuAchat                             + ' , ';
    StSQL := StSQL + ' BD1_NBJOUR='        + StNbJour                              + ' WHERE';
    StSQL := StSQL + ' BD1_NATUREPIECEG="' + Cledoc.NaturePiece                    + '"  AND';
    StSQL := StSQL + ' BD1_SOUCHE="'       + Cledoc.Souche                         + '"  AND';
    StSQL := StSQL + ' BD1_NUMERO='        + IntToStr(Cledoc.NumeroPiece)          + '   AND';
    StSQL := StSQL + ' BD1_INDICEG='       + IntToStr(Cledoc.Indice)               + '   AND';
    StSQL := StSQL + ' BD1_UNIQUE='        + IntToStr(Unique)                      + '   AND';
    StSQL := StSQL + ' BD1_UNIQUELIG='     + IntToStr(UniqueLig)                   + '   AND';
    StSQl := StSQL + ' BD1_TIERS="'        + RefFrs + '"';
    if ExecuteSQL(StSQL) = 1 then
    begin
      Rapport.SauveLigMemo('Traitement demande de prix N° '+ IntToStr(Unique) +'-'+ IntToStr(UniqueLig) + ' Traitée');
      Inc(LigMaj);
    end
    else
    begin
      Rapport.SauveLigMemo('Traitement demande de prix N° '+ IntToStr(Unique) +'-'+ IntToStr(UniqueLig) + ' Non Traitée');
      Inc(LigNonMaj);
    end
  end;

  if (LigMaj <> 0) And (LigNonMaj = 0) then
  Begin
    Rapport.SauveLigMemo('Nbre de ligne(s) traitée(s)     : ' + IntToStr(LigMaj));
    Rapport.SauveLigMemo('Document enregistré avec succès.');
  end
  else if (LigMaj <> 0) And (LigNonMaj <> 0) then
  begin
    Rapport.SauveLigMemo('Nbre de ligne(s) traitée(s)     : ' + IntToStr(LigMaj));
    Rapport.SauveLigMemo('Nbre de ligne(s) non traitée(s) : ' + IntToStr(LigNonMaj));
    Rapport.SauveLigMemo('Demande de Prix enregistré avec des erreurs, vérifier le fichier XLS importé.');
  end
  else if (LigMaj = 0) And (LigNonMaj = 0) then
    Rapport.SauveLigMemo('Aucune ligne à traiter, veuillez vérifier le fichier ' + XLSName.Text + ' intégré')
  else
  begin
    Rapport.SauveLigMemo('Nbre de ligne(s) non traitée(s) : ' + IntToStr(LigNonMaj));
    Rapport.SauveLigMemo('Demande de prix non enregistré.');
  end;

end;


procedure TOF_BTINTEGREDOC.OuvrirFichierXLS(sender : TObject);
begin
  OpenFicXLS(XLSName.Text);
end;

procedure TOF_BTINTEGREDOC.GestionRapport(sender: TObject);
begin
  Rapport.Affiche := ChkRapport.Checked;
end;

procedure TOF_BTINTEGREDOC.ImprimeFichierXLS(sender: TObject);
begin

end;

function TOF_BTINTEGREDOC.AddligneXls  (TOBXLS : TOB) : TOB;
begin
	result := TOB.Create ('UNE LIGNE XLS',TOBXLS,-1);
  result.AddChampSupValeur ('NATUREPIECEG','');
  result.AddChampSupValeur ('SOUCHE','');
  result.AddChampSupValeur ('NUMERO',0);
  result.AddChampSupValeur ('INDICEG',0);
	result.AddChampSupValeur ('FOURNISSEUR','');
	result.AddChampSupValeur ('NUMORDRE',0);
	result.AddChampSupValeur ('UNIQUEBLO',0);
  result.AddChampSupValeur ('UNIQUE', 0);
  result.AddChampSupValeur ('UNIQUELIG',0);
  result.AddChampSupValeur ('REFARTFRS','');
	result.AddChampSupValeur ('UNITE',0);
	result.AddChampSupValeur ('QTE',0);
	result.AddChampSupValeur ('PU',0);
  result.AddChampSupValeur ('NBJOUR',0);
	result.AddChampSupValeur ('TRAITEOK','-');
end;

procedure TOF_BTINTEGREDOC.LogErreur(numErreur: integer; TypeErreur : TTypeErreur);
var MessErreur : string;
begin
  if TypeErreur = TteERROR then
  begin
    Case Numerreur of
      1: MessErreur := 'ERREUR : Zones obligatoire [REFFOURNISSEUR] manquantes,' + #13#10 + 'Feuille Abandonnée';
      2: MessErreur := 'ERREUR : Zones obligatoire [QTELIGNE] manquantes ,' + #13#10 + 'Feuille Abandonnée';
      3: MessErreur := 'ERREUR : Zones obligatoire [PULIGNE]  manquantes, ' + #13#10 + 'Feuille Abandonnée';
      4: MessErreur := 'ERREUR : Zones obligatoire [NBJOUR]   manquantes, ' + #13#10 + 'Feuille Abandonnée';
      5: MessErreur := 'ERREUR : Zones obligatoire [REFARTLIGNE] manquantes, ' + #13#10 + 'Feuille Abandonnée';
      6: MessErreur := 'ERREUR : Zones obligatoire [REFLIGNE] manquantes, ' + #13#10 + 'Feuille Abandonnée';
      7: MessErreur := 'ERREUR : Zones obligatoire [UNITELIGNE] manquantes, ' + #13#10 + 'Feuille Abandonnée';
      else MessErreur := 'ERREUR';
    end;
  end else if TypeErreur = TteWARNING then
  begin
    Case Numerreur of
      1: MessErreur := 'ATTENTION : la zone unité n''est pas définie';
      else MessErreur := 'ERREUR';
    end;
  end;
	Rapport.SauveLigMemo(MessErreur);
end;

function TOF_BTINTEGREDOC.FindDoc(TOBXLS : TOB): TOB;
begin
  result := TOBDET.findFirst(['NATUREPIECEG','SOUCHE','NUMERO','INDICEG'],
           [TOBXLS.getString('NATUREPIECEG'),
            TOBXLS.getString('SOUCHE'),
            TOBXLS.GetInteger('NUMERO'),
            TOBXLS.GetInteger('INDICEG')],true);

  if result = nil then
  begin
		result := TOB.Create ('UN DOCUMENT',TOBDET,-1);
    result.AddChampSupValeur ('NATUREPIECEG',TOBXLS.getString('NATUREPIECEG'));
    result.AddChampSupValeur ('SOUCHE',TOBXLS.getString('SOUCHE'));
    result.AddChampSupValeur ('NUMERO',TOBXLS.GetInteger('NUMERO'));
    result.AddChampSupValeur ('INDICEG',TOBXLS.GetInteger('INDICEG'));
  end;

end;

procedure TOF_BTINTEGREDOC.ControleChamp(Champ, Valeur: String);
begin

  if      (Champ = 'TYPEIMPORT') AND (Valeur = 'DDE') then fTTypeImport := Ttddeprix
  Else if (Champ = 'TYPEIMPORT') AND (Valeur = 'SST') then fTTypeImport := TtSsTraite
  Else if (Champ = 'TYPEIMPORT') AND (Valeur = 'TAF') then fTTypeImport := TtTarifFr
  Else fTTypeImport := TTImport;

end;

procedure TOF_BTINTEGREDOC.ControleCritere(Critere: String);
begin
end;


Procedure TOF_BTINTEGREDOC.FinTraiteImport;
Begin

  XLSName.Text := '';
  BValider.Enabled := false;
  Splash.close;
  FreeAndNil(splash);

  //affichage du rapport d'intégration
  if ChkRapport.Checked then Rapport.AfficheRapport;    //affichage du rapport d'intégration

end;

function TOF_BTINTEGREDOC.RechercheRange(Etiquette : string; Var LibEtiq : string; Var LigEtiq, ColEtiq : Integer) : Boolean;
var SheetName : string;
    NBCols    : Integer;
begin

  result := false;

  SheetName := Workbook.ActiveSheet.Name;

  ExcelFindRange(Workbook, SheetName, Etiquette, LigEtiq, ColEtiq, Nbcols);

  if ColEtiq = 0 then
    GestionErreur('L''étiquette' + Etiquette + ' n''existe pas dans la feuille ' + SheetName)
  else
  begin
    result  := True;
    LibEtiq := GetExcelText(Workbook.ActiveSheet, LigEtiq, ColEtiq);
  end;

end;

Procedure TOF_BTINTEGREDOC.GestionErreur(LibError : string);
begin

  if ChkRapport.Checked then
    Rapport.SauveLigMemo(LibError)
  else
    PGIInfo(TraduireMemoire(Liberror), 'Erreur integration XLS');

end;


Initialization
  registerclasses ( [ TOF_BTINTEGREDOC ] ) ;
end.


