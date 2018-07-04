{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 12/03/2015
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : ETIQUETTE ()
Mots clefs ... : TOF;ETIQUETTE
*****************************************************************}
Unit Utof_BTETIQUETTE ;

Interface

Uses StdCtrls, 
     Controls,
     Classes, 
{$IFNDEF EAGLCLIENT}
     FE_Main,
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Hpanel,
     uTOB,
     UTOF,
     DateUtils,
{$IFDEF EAGLCLIENT}
      utileAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ENDIF}Fiche,
{$IFDEF V530}
      EdtEtat,
{$ELSE}
      EdtREtat,
{$ENDIF}
{$ENDIF}
      AglInit,
      EntGC,
      HTB97,
      HQry,
      AffaireUtil,
      vierge,
      Menus,
      Dialogs,
      ShellAPI,
      Windows;

Type
  TOF_ETIQUETTE = Class (TOF)
    procedure OnArgument (S : String )  ; override ;
    Procedure OnUpdate                  ; Override ;
    procedure OnClose                   ; override ;

  private
    NatureCAB         : THValComboBox;
    QualifCAB         : THValComboBox;
    //
    BSelect1          : TToolbarbutton97;
    BSelect2          : TToolBarButton97;
    BEfface           : TToolBarButton97;
    BGenereWord       : TToolBarButton97;
    BValider          : TToolBarButton97;
    //
    CodeArticle       : THEdit;
    CodeArticle_      : THEdit;
    //
    CodeTiers         : THEdit;
    CodeTiers_        : THEdit;
    //
    ModeleWord        : THEdit;
    DocGenere         : THEdit;
    //
    MnnewModele       : TmenuItem;
    MnModifModele     : TmenuItem;
    MnSupModele       : TmenuItem;
    //
    LIBREFOURNISSEUR  : THPanel;
    LIBREARTICLE      : THPanel;
    //
    Affaire           : THCritMaskEdit;
    Part0             : THCritMaskEdit;
    Part1             : THCritMaskEdit;
    Part2             : THCritMaskEdit;
    Part3             : THCritMaskEdit;
    Avenant           : THCritMaskEdit;
    //
    Affaire_          : THCritMaskEdit;
    Part0_            : THCritMaskEdit;
    Part1_            : THCritMaskEdit;
    Part2_            : THCritMaskEdit;
    Part3_            : THCritMaskEdit;
    Avenant_          : THCritMaskEdit;
    //
    OnlyCB            : TCheckBox;
    EditViaWord       : TCheckBox;
    Principal         : TCheckBox;
    //
    PAvances          : TTabSheet;
    PContact          : TTabSheet;
    PFournisseur      : TTabSheet;
    PArticle          : TTabSheet;
    PAffaire          : TTabSheet;
    PTableLibre       : TTabSheet;
    PTableLibre_      : TTabSheet;
    PEdition          : TTabSheet;
    //
    Pages             : THPageControl2;
    //
    GBEDITETIQUETTEWORD : TGroupBox;
    //
    TobEtiquette      : TOB;
    //
    StSQL             : String;
    StWhere           : String;
    C_StWhere         : String;
    //
    TEtat             : THLabel;
    FEtat             : THValComboBox;
    BParamEtat        : TToolbarButton97;
    //
    fRepServeur       : String;
    fRepArchiveS      : String; //Repertoire archive sur Serveur
    fRepArchiveL      : String; //Répertoire Archive en Local
    //
    procedure ArticleOnElipsisClick(Sender: TObject);
    procedure ArticleOnElipsisClick_(Sender: TObject);
    procedure BEffaceOnClick(Sender: TObject);
    procedure BSelect1OnClick(Sender: TObject);
    procedure BSelect2OnClick(Sender: TObject);
    procedure ChangingOnglet(Sender: Tobject);
    procedure ChargeRequete;
    procedure CreateTOB;
    procedure Controlechamp(Champ, Valeur: String);
    procedure ConstitueListe(Prefixe, TABLE: string; var ListeChamps, ListChampsSql: string);
    procedure ControleRepEtiquette;
    Function  CreationRepertoire(NomRep: String) : Boolean;
    procedure GenereWordClick(Sender: TObject);
    procedure GestionEcranAffaire;
    procedure GestionEcranArticle;
    procedure GestionEcranFournisseur;
    procedure GetObjects;
    procedure ModifModeleClick(Sender: Tobject);
    procedure NatureCABOnChange(Sender: TObject);
    procedure NewModeleClick(Sender: Tobject);
    function  Recupwhere: String;
    function  RecupWhereAffaires: String;
    function  RecupWhereFournisseurs: String;
    function  RecupWhereMarchandises: String;
    function  RecupZoneCheck(Zone1: TCheckBoxState; LibZone: String): String;
    function  RecupZoneDate(Zone1, Zone2 : TdateTime; LibZone : String) : String;
    function  RecupZoneListe(Zone1, LibZone: String): String;
    function  RecupZoneMultiple(Zone1, Zone2, LibZone: String): String;
    function  RecupZoneSimple(Zone1, LibZone: String): String;
    procedure SelectDocWord(Sender: TObject);
    procedure SelectModele(Sender: TObject);
    procedure SuppModeleClick(Sender: Tobject);
    procedure SelectTypeEdition(Sender: TObject);
    procedure SetScreenEvents;
    procedure TiersOnElipsisClick(Sender: TObject);
    procedure TiersOnElipsisClick_(Sender: TObject);
    function  TransformeValeur(valeurOri: string): String;
    function RecupWhereContact: String;
    //

  end ;

Implementation

uses  TntStdCtrls,
      ParamSoc,
      UtilWord,
      UtilFichiers,
      FactUtil;

procedure TOF_ETIQUETTE.OnArgument (S : String ) ;
var i       : Integer;
    Critere : string;
    Champ   : string;
    Valeur  : string;
begin
  Inherited ;

  Inherited ;
  //
  //Chargement des zones ecran dans des zones programme
  GetObjects;
  //
  CreateTOB;
  //
  Critere := S;
  //
  While (Critere <> '') do
  BEGIN
    i:=pos(':',Critere);
    if i = 0 then i:=pos('=',Critere);
    if i <> 0 then
       begin
       Champ:=copy(Critere,1,i-1);
       Valeur:=Copy (Critere,i+1,length(Critere)-i);
       end
    else
       Champ := Critere;
    Controlechamp(Champ, Valeur);
    Critere:=(Trim(ReadTokenSt(S)));
  END;

  Part0.Text  := 'A';

  Part0_.Text := 'A';

  SetScreenEvents;

  ChargeCleAffaire (Part0, Part1, Part2, Part3, AVENANT, BSelect1, TaCreat, Affaire.text, False);
  ChargeCleAffaire (Part0_, Part1_, Part2_, Part3_, AVENANT_, BSelect2, TaCreat, Affaire_.text, False);  

  NatureCAB.Value := 'MAR';

  SelectTypeEdition(self);
  
  GestionEcranArticle;

end ;

Procedure TOF_ETIQUETTE.ControleRepEtiquette;
var YY  : Word;
    DD  : Word;
    MM  : Word;
    H   : Word;
    M   : Word;
    S   : Word;
    MS  : Word;
    Ext : string;
begin

  fRepServeur := '';

  fRepServeur     := GetParamSocSecur('SO_BTREPMODELEETIQUETTE','');
  ModeleWord.text := GetParamSocSecur('SO_MODELEETIQUETTE','');
  fRepArchiveS    := GetParamSocSecur('SO_BTREPDOCETIQUETTE','');

  if fRepServeur  = '' then
  begin
    PGIError('Veuillez renseigner le repertoire de sauvegarde des modèles dans les paramètres !','Edition des Etiquettes');
    GBEDITETIQUETTEWORD.Visible := False;
  end;

  if fRepArchiveS = '' then
  begin
    PGIError('Veuillez renseigner le repertoire de sauvegarde des modèles dans les paramètres !','Edition des Etiquettes');
    GBEDITETIQUETTEWORD.Visible := False;
  end;

  //On vérifie si le nom du fichier modèle est à blanc
  if ModeleWord.text = '' then PGIError('Attention vous n''avez aucun modèle par défaut dans les paramètres !','Edition des Etiquettes');

  //gestion répertoires Modèles Serveurs
  if not CreationRepertoire(fRepServeur) then exit;

  //fRepServeur := fRepServeur + '\Etiquettes';
  //if not CreationRepertoire(fRepServeur) then exit;

  //fRepServeur := fRepServeur + '\Modeles';
  //if not CreationRepertoire(fRepServeur) then exit;

  //gestion répertoires Archives Serveurs
  if not CreationRepertoire(fRepArchiveS) then exit;

  //fRepArchiveS := fRepArchiveS + '\Etiquettes';
  //if not CreationRepertoire(fRepArchiveS) then exit;

  //fRepArchiveS := fRepArchiveS + '\Archives';
  //if not CreationRepertoire(fRepArchiveS) then exit;

  //On vérifie s'il existe sur le serveur...
  if Not FileExists(fRepServeur + '\' +  ModeleWord.text) then
  begin
    PGIError('Attention le modèle sélectionné n''est pas dans le répertoire des modèle : ' + fRepServeur, 'Edition des Etiquettes');
    ModeleWord.text := '';
  end;

  //Gestion répertoire Archives Local
  fRepArchiveL := 'C:\PGI01';
  if Not CreationRepertoire(fRepArchiveL) then exit;

  fRepArchiveL := fRepArchiveL + '\Etiquettes';
  if Not CreationRepertoire(fRepArchiveL) then exit;

  fRepArchiveL := fRepArchiveL + '\Archives';
  If not CreationRepertoire(fRepArchiveL) then exit;

  ModeleWord.text := frepServeur + '\' + ModeleWord.text;

  DecodeDateTime(Now, YY, MM, DD, H,M,S,MS);
  Ext := IntToStr(DD) + IntToStr(MM) + IntToStr(YY) + IntToStr(H) + IntToStr(M) + IntToStr(S);
  DocGenere.Text := fRepArchiveL + '\' + 'Etiquette_' + Ext + '.doc';

end;


procedure TOF_ETIQUETTE.OnClose ;
Var Info     : TSearchRec;
    Nomfile  : string;
begin
  Inherited ;

  StSQL := 'DELETE FROM BTTMPETQ WHERE BZD_UTILISATEUR = "'+ V_PGI.USer +'"';
  ExecuteSQL(StSQL);

  //copie du fichier doc dans le répertoire serveur...
  if (fRepArchiveL <> '') And (fRepArchiveS <> '' ) then
  begin
    fRepArchiveL := IncludeTrailingPathDelimiter(fRepArchiveL);
    fRepArchiveS := IncludeTrailingPathDelimiter(fRepArchiveS);
    //On charge le répertoire document
    if FindFirst (fRepArchiveL + '*.*', faAnyFile, Info) = 0 then
    begin
      repeat
        If Not((Info.Attr And faDirectory)=0)  then
        begin
        end
        else
        begin
          Nomfile := ExtractFileName(Info.Name);
          CopieFichier(fRepArchiveL  + NomFile,  fRepArchiveS + NomFile);
          DeleteFichier(fRepArchiveL + NomFile);
        end;
      Until FindNext(Info)<>0;
      //FindClose(Info);
    end;
  end;

  FreeAndNil(TobEtiquette);

end ;

Procedure TOF_ETIQUETTE.Controlechamp(Champ, Valeur : String);
begin

  if Champ = 'ID'   then NatureCAB.Value := Valeur;

end;

Procedure TOF_ETIQUETTE.GetObjects;
begin
  //
  Affaire := THCritMaskEdit(GetControl('AFF_AFFAIRE'));
  Part0   := THCritMaskEdit(GetControl('AFF_AFFAIRE0'));
  Part1   := THCritMaskEdit(GetControl('AFF_AFFAIRE1'));
  Part2   := THCritMaskEdit(GetControl('AFF_AFFAIRE2'));
  Part3   := THCritMaskEdit(GetControl('AFF_AFFAIRE3'));
  Avenant := THCritMaskEdit(GetControl('AFF_AVENANT'));
  //
  Affaire_:= THCritMaskEdit(GetControl('AFF_AFFAIRE_'));
  Part0_  := THCritMaskEdit(GetControl('AFF_AFFAIRE0_'));
  Part1_  := THCritMaskEdit(GetControl('AFF_AFFAIRE1_'));
  Part2_  := THCritMaskEdit(GetControl('AFF_AFFAIRE2_'));
  Part3_  := THCritMaskEdit(GetControl('AFF_AFFAIRE3_'));
  Avenant_:= THCritMaskEdit(GetControl('AFF_AVENANT_'));
  //
  CodeArticle       := THEdit(GetControl('GA_CODEARTICLE'));
  CodeArticle_      := THEdit(GetControl('GA_CODEARTICLE_'));
  //
  CodeTiers         := THEdit(GetControl('T_TIERS'));
  CodeTiers_        := THEdit(GetControl('T_TIERS_'));
  //
  NatureCAB         := THValComboBox(GetControl('BCB_NATURECAB'));
  QualifCAB         := THValComboBox(GetControl('BCB_QUALIFCODEBARRE'));
  //
  BSelect1          := TToolBarButton97(GetControl('BSELECT1'));
  BSelect2          := TToolBarButton97(GetControl('BSELECT2'));
  Befface           := TToolBarButton97(GetControl('BEFFACE'));
  //
  LIBREFOURNISSEUR  := THPanel(Getcontrol('LIBREFOURNISSEUR'));
  LIBREARTICLE      := THPanel(Getcontrol('LIBREARTICLE'));
  //
  ModeleWord        := THEdit(GetControl('MODELEWORD'));
  DocGenere         := THEdit(GetControl('DOCGENERE'));
  //
  MnnewModele       := TmenuItem(GetControl('MnnewModele'));
  MnModifModele     := TmenuItem(GetControl('MnModifModele'));
  MnSupModele       := TmenuItem(GetControl('MnSupModele'));
  //
  PAvances          := TTabSheet(GetControl('Avances'));
  PFournisseur      := TTabSheet(GetControl('PAGEFOURNISSEUR'));
  PAffaire          := TTabSheet(GetControl('Standards'));
  PArticle          := TTabSheet(GetControl('PAGEARTICLE'));
  PContact          := TTabSheet(Getcontrol('PCONTACT'));
  PTableLibre       := TTabSheet(GetControl('PTABLESLIBRES'));
  PTableLibre_      := TTabSheet(GetControl('PTABLESLIBRES_'));
  PEdition          := TTabSheet(GetControl('MiseEnPage'));
  //
  Pages             := THPageControl2(GetControl('PAGES'));
  //
  GBEDITETIQUETTEWORD := TGroupBox(GetControl('GBEDITETIQUETTEWORD'));
  //
  EditViaWord       := TCheckBox(GetControl('CHKEDITVIAWORD'));
  OnlyCB            := TCheckBox(GetControl('CHKONLYCB'));
  Principal         := TCheckBox(GetControl('C_PRINCIPAL'));
  //
  FEtat             := THValComboBox(GetControl('FETAT'));
  TEtat             := THLabel(GetControl('TETAT'));
  BParamEtat        := TToolbarButton97(GetControl('BPARAMETAT'));
  //
  Bvalider          := TtoolBarButton97(GetControl('BVALIDER'));
  BGenereWord       := TtoolBarButton97(GetControl('BGENEREWORD'));

end;

Procedure TOF_ETIQUETTE.CreateTOB;
begin

  TOBEtiquette := TOB.Create('LES ETIQUETTES',Nil, -1);

end;
Procedure TOF_ETIQUETTE.SetScreenEvents;
begin

  NatureCAB.OnChange := NatureCABOnChange;
  //
  CodeArticle.OnElipsisClick  := ArticleOnElipsisClick;
  CodeArticle_.OnElipsisClick := ArticleOnElipsisClick_;
  //
  CodeTiers.OnElipsisClick    := TiersOnElipsisClick;
  CodeTiers_.OnElipsisClick   := TiersOnElipsisClick_;
  //
  BSelect1.OnClick            := BSelect1OnClick;
  BSelect2.OnClick            := BSelect2OnClick;
  Befface.OnClick             := BEffaceOnClick;
  //
  Pages.OnChange              := ChangingOnglet;
  //
  // menu des modèles
  MODELEWORD.OnElipsisClick   := SelectModele;
  DOCGENERE.OnElipsisClick    := SelectDocWord;
  MnnewModele.onclick         := NewModeleClick;
  MnModifModele.onclick       := ModifModeleClick;
  MnSupModele.onclick         := SuppModeleClick;
  //
  EditViaWord.OnClick         := SelectTypeEdition;
  //
  BGenereWord.OnClick         := GenereWordClick;
end;

procedure TOF_ETIQUETTE.ChangingOnglet(Sender : Tobject);
begin

  if Pages.ActivePage.Name = 'MiseEnPage' then
  begin
    if EditViaWord.Checked then ControleRepEtiquette;
  end;

end;

Procedure TOF_ETIQUETTE.ArticleOnElipsisClick(Sender : TObject);
Var SWhere   : String;
    StChamps : String;
begin

	sWhere := ' AND ((GA_TYPEARTICLE="MAR") AND GA_TENUESTOCK="X")';

  StChamps := CodeArticle.Text;

  if CodeArticle.Text <> '' then
    sWhere := 'GA_CODEARTICLE=' + Trim(Copy(CodeArticle.text, 1, 18)) + ';XX_WHERE=' + sWhere
  else
    sWhere := 'XX_WHERE=' + sWhere;

	CodeArticle.text := AGLLanceFiche('BTP', 'BTARTICLE_RECH', '', '', sWhere +';RECHERCHEARTICLE');
  CodeArticle.text := Trim(Copy(CodeArticle.text, 1, 18));

end;
Procedure TOF_ETIQUETTE.ArticleOnElipsisClick_(Sender : TObject);
Var SWhere   : String;
    StChamps : String;
begin

	sWhere := ' AND ((GA_TYPEARTICLE="MAR") AND GA_TENUESTOCK="X")';

  StChamps := CodeArticle_.Text;

  if CodeArticle_.Text <> '' then
    sWhere := 'GA_CODEARTICLE=' + Trim(Copy(CodeArticle_.text, 1, 18)) + ';XX_WHERE=' + sWhere
  else
    sWhere := 'XX_WHERE=' + sWhere;

	CodeArticle_.text := AGLLanceFiche('BTP', 'BTARTICLE_RECH', '', '', sWhere +';RECHERCHEARTICLE');
  CodeArticle_.text := Trim(Copy(CodeArticle_.text, 1, 18));

end;

Procedure TOF_ETIQUETTE.TiersOnElipsisClick(Sender : TObject);
Var StChamps  : string;
begin

  StChamps  := CodeTiers.Text;

  CodeTiers.Text := AGLLanceFiche('GC','GCFOURNISSEUR_MUL','T_TIERS=' + StChamps +';T_NATUREAUXI=FOU','','SELECTION');

end;

Procedure TOF_ETIQUETTE.TiersOnElipsisClick_(Sender : TObject);
Var StChamps  : string;
begin

  StChamps  := CodeTiers_.Text;

  CodeTiers_.Text := AGLLanceFiche('GC','GCFOURNISSEUR_MUL','T_TIERS=' + StChamps +';T_NATUREAUXI=FOU','','SELECTION');

end;

Procedure TOF_ETIQUETTE.BSelect1OnClick(Sender : TObject);
Var StChamps  : String;
begin

  StChamps  := Affaire.Text;

  if GetAffaireEnteteSt(Part0, Part1, Part2, Part3, Avenant, nil, StChamps, false, false, false, True, true,'') then Affaire.text := StChamps;

  ChargeCleAffaire (Part0,Part1,Part2,Part3,Avenant, BSelect1, TaCreat, Affaire.Text, False);

end;

Procedure TOF_ETIQUETTE.BSelect2OnClick(Sender : TObject);
Var StChamps  : String;
begin

  StChamps  := Affaire_.Text;

  if GetAffaireEnteteSt(Part0_, Part1_, Part2_, Part3_, Avenant_, nil, StChamps, false, false, false, True, true,'') then Affaire_.text := StChamps;

  ChargeCleAffaire (Part0_,Part1_,Part2_,Part3_,Avenant_, BSelect2, TaCreat, Affaire_.Text, False);

end;

Procedure TOF_ETIQUETTE.BEffaceOnClick(Sender : TObject);
begin

  Affaire.Text  := '';
  Part0.Text    := '';
  Part1.Text    := '';
  Part2.Text    := '';
  Part3.Text    := '';
  Avenant.Text  := '';

  Affaire_.Text := '';
  Part0_.Text   := '';
  Part1_.Text   := '';
  Part2_.Text   := '';
  Part3_.Text   := '';
  Avenant_.Text := '';

end;


Procedure TOF_ETIQUETTE.NatureCABOnChange(Sender : TObject);
begin

  if      NatureCAB.Value = 'FOU' then GestionEcranFournisseur
  else if NatureCAB.Value = 'MAR' then GestionEcranArticle
  else if NatureCAB.Value = 'AFF' then GestionEcranAffaire
  else                                 GestionEcranArticle;

end;

Procedure TOF_Etiquette.GestionEcranFournisseur;
begin

  QualifCAB.Value           := '128';
  QualifCAB.Enabled         := False;

  PFournisseur.TabVisible   := True;
  PContact.TabVisible       := True;
  PTableLibre_.TabVisible   := True;

  PArticle.TabVisible       := False;
  PAffaire.TabVisible       := False;
  PTableLibre.TabVisible    := False;

  PAvances.TabVisible       := False;

  Pages.ActivePage := PFournisseur;

  Principal.Checked := True;
end;

Procedure TOF_Etiquette.GestionEcranArticle;
begin

  QualifCAB.Value           := '';
  QualifCAB.Enabled         := True;

  PFournisseur.TabVisible   := False;
  PContact.TabVisible       := False;
  PTableLibre_.TabVisible   := False;

  PArticle.TabVisible       := True;
  PAffaire.TabVisible       := False;
  PTableLibre.TabVisible    := True;

  PAvances.TabVisible       := False;

  Pages.ActivePage := PArticle;

end;

Procedure TOF_Etiquette.GestionEcranAffaire;
begin

  QualifCAB.Value           := '128';
  QualifCAB.Enabled         := False;

  PFournisseur.TabVisible   := False;
  PContact.TabVisible       := False;
  PTableLibre_.TabVisible   := False;

  PArticle.TabVisible       := False;
  PAffaire.TabVisible       := True;
  PTableLibre.TabVisible    := False;

  PAvances.TabVisible       := False;

  Pages.ActivePage := PAffaire;

end;


procedure TOF_ETIQUETTE.OnUpdate;
Var IP      : Integer;
begin
  inherited;

  If NatureCAB.Value = 'AFF' then
  begin
    Affaire.text  := DechargeCleAffaire(Part0, Part1, Part2, Part3, Avenant, '', Taconsult, False, True, false, IP);
    Affaire_.text := DechargeCleAffaire(Part0_, Part1_, Part2_, Part3_, Avenant_, '', Taconsult, False, True, false, IP);
  end;

  ChargeRequete;

end;

Procedure TOF_ETIQUETTE.ChargeRequete;
Var QEtiq       : TQuery;
    TOBL        : TOB;
    NbEtiq      : Integer;
    NbEnreg     : Integer;
    indice      : Integer;
    i           : Integer;
    RegimePrix  : String;
    LibEnreg    : String;
    LibCompl    : String;
begin

  TOBEtiquette.ClearDetail;

  ExecuteSQL('DELETE FROM BTTMPETQ WHERE BZD_UTILISATEUR = "'+ V_PGI.USer +'"');

  StWhere := '';
      
  If NatureCAB.Value = 'AFF' then
  begin
    if Affaire.text > Affaire_.text then
    begin
      PGIError('Les codes Affaire saisis ne sont pas cohérents.', 'Erreur Sélection');
      Exit;
    end;
  end
  else If NatureCAB.Value = 'FOU' then
  begin
    if CodeTiers.text > CodeTiers_.text then
    begin
      PGIError('Les codes Fournisseur saisis ne sont pas cohérents.', 'Erreur Sélection');
      Exit;
    end;
  end
  else If NatureCAB.Value = 'MAR' then
  begin
    if CodeArticle.text > CodeArticle_.text then
    begin
      PGIError('Les codes Article saisis ne sont pas cohérents.', 'Erreur Sélection');
      Exit;
    end;
  end;

  if GetControlText('RHT') = 'X' then RegimePrix := 'HT' else RegimePrix := 'TTC';

  NbEtiq := 0;

  StWhere := RecupWhere;

  //Lecture pour charger précisément la base temporaire !!!
  If NatureCAB.Value = 'FOU' then
  begin
    StSQL  := 'SELECT COUNT(T_TIERS) FROM TIERS ';
    if C_StWhere <> '' then C_StWhere := 'AND ' + C_StWhere;
    StSQL  := StSQL + 'LEFT JOIN CONTACT    ON T_AUXILIAIRE=C_AUXILIAIRE ' + C_Stwhere;
    StSQL  := StSQL + 'LEFT JOIN TIERSCOMPL ON T_AUXILIAIRE=YTC_AUXILIAIRE ';
    StSQL  := StSQL + 'LEFT JOIN BTIERS ON BT1_AUXILIAIRE=YTC_AUXILIAIRE ';
    if stwhere <> '' then
      StWhere := StWhere + 'AND T_NATUREAUXI = "FOU"'
    else
      StWhere := ' T_NATUREAUXI = "FOU"';
    NbEtiq := StrToInt(GetControlText('NBEXEMPLAIRE2'));
  end
  else if NatureCAB.Value = 'MAR' then
  begin
    NbEtiq := StrToInt(GetControlText('NBEXEMPLAIRE1'));
    StSQL := 'SELECT COUNT(GA_ARTICLE) FROM ARTICLE ';
  end
  else if NatureCAB.Value = 'AFF' then
  begin
    NbEtiq := StrToInt(GetControlText('NBEXEMPLAIRE'));
    StSQL  := 'SELECT COUNT(AFF_AFFAIRE) FROM AFFAIRE ';
    StSQL := StSQL + 'LEFT JOIN TIERS ON AFF_TIERS=T_TIERS AND T_NATUREAUXI="CLI" ';
    if stwhere <> '' then
      StWhere := StWhere + 'AND AFF_AFFAIRE0 = "A"'
    else
      StWhere := ' AFF_AFFAIRE = "A"';
  end;

  if stWhere <> '' then StSQL := StSQL + ' WHERE ' + stWhere;

  QEtiq := OpenSQL(StSQL, true);

  NbEnreg := QEtiq.Fields[0].AsInteger;

  Ferme(QEtiq);
  if NbEnreg = 0 then exit;

  If NatureCAB.Value = 'FOU' then
  begin
    StSQl := 'SELECT T_NATUREAUXI AS NATURECAB, T_AUXILIAIRE AS CODECAB, ';
    StSQL := StSQL + '"Code128" AS QUALIFCAB, T_TIERS AS CODEBARRE, ';
    StSQL := StSQL + 'T_LIBELLE AS LIBELLE, C_NUMEROCONTACT, C_NOM, C_PRENOM FROM TIERS ';
    StSQL  := StSQL +'LEFT JOIN CONTACT    ON T_AUXILIAIRE=C_AUXILIAIRE ' + C_StWhere;
    StSQL  := StSQL +'LEFT JOIN TIERSCOMPL ON T_AUXILIAIRE=YTC_AUXILIAIRE ';
    StSQL := StSQL + ' WHERE ' + Stwhere;
  end
  else if NatureCAB.Value = 'MAR' then
  begin
    StSQl := 'SELECT "' + NatureCAB.Value + '" AS NATURECAB, BCB_IDENTIFCAB AS CODECAB, ';
    StSQL := StSQL + 'BCB_QUALIFCODEBARRE AS QUALIFCAB, BCB_CODEBARRE as CODEBARRE, ';
    StSQL := StSQL + 'GA_LIBELLE AS LIBELLE,GA_LIBCOMPL AS LIBCOMPL,GA_PVHT AS PVHT, ';
    StSQL := StSQL + 'GA_PVTTC AS PVTTC, GA_CODEARTICLE,GA_STATUTART ';
    StSQL := StSQL + 'FROM ARTICLE LEFT JOIN BTCODEBARRE ON GA_ARTICLE=BCB_IDENTIFCAB ';
    StSQL := StSQL + ' WHERE ' + Stwhere; 
  end
  else if NatureCAB.Value = 'AFF' then
  begin
    StSQl := 'SELECT "' + NatureCAB.Value + '" AS NATURECAB, SUBSTRING(AFF_AFFAIRE,2,14) AS CODECAB, ';
    StSQL := StSQL + '"Code128" AS QUALIFCAB, SUBSTRING(AFF_AFFAIRE,2,14) AS CODEBARRE, ';
    StSQL := StSQL + 'AFF_LIBELLE AS LIBELLE, T_LIBELLE AS LIBCOMPL ';
    StSQl := StSQL + 'FROM AFFAIRE ';
    StSQl := StSQL + 'LEFT JOIN TIERS ON AFF_TIERS=T_TIERS AND T_NATUREAUXI="CLI" ';
    StSQL := StSQL + ' WHERE ' + Stwhere;
  end;

  QEtiq := OpenSQL(StSQL,true);

  If not QEtiq.eof then
  begin
    TobEtiquette.LoadDetailDB('Etiquette','','',QEtiq,false,true);
    For I := 0 to TobEtiquette.detail.count -1 do
    begin
      TOBL        := TobEtiquette.detail[I];
      LibEnreg    := TobL.GetString('LIBELLE');
      LibEnreg    := StringReplace(LibEnreg, '"', '""',[rfReplaceAll]);
      LibCompl    := TOBL.GetString('LIBCOMPL');
      LibCompl    := StringReplace(LibCompl, '"', '""',[rfReplaceAll]);
      for indice  :=0 to NbEtiq - 1 do
      begin
        if NatureCAB.Value = 'MAR' then
        begin
          if TOBL.GetValue('CODEBARRE') = '' then
          begin
            TOBL.PutValue('NATURECAB', NatureCAB.Value);
            TOBL.PutValue('CODEBARRE', TOBL.GetValue('GA_CODEARTICLE'));
            TOBL.PutValue('CODECAB',   TOBL.GetValue('GA_CODEARTICLE'));
            TOBL.PutValue('QUALIFCAB', 'Code128');
          end;
        end;
        StSQL :='INSERT INTO BTTMPETQ ';
        StSQL := Stsql + '(BZD_UTILISATEUR, BZD_CODE,      BZD_LIBELLE, BZD_LIBCOMPL, ';
        StSQL := StSQL + ' BZD_REGIMEPRIX,  BZD_NBETIQ,    BZD_PVHT,    BZD_PVTTC, ';
        StSQL := StSQL + ' BZD_DEVISEPRINC, BZD_CODEBARRE, BZD_INDICE, ';
        StSQL := StSQL + ' BZD_NATURECAB,   BZD_QUALIFCAB) values (';
        StSQL := StSQL + ' "' + V_PGI.User                + '",'; //Utilisateur
        StSQL := StSQL + ' "' + TOBL.GetString('CODECAB');        //code...
        //
        If NatureCAB.Value = 'FOU' then
        begin
          if (Principal.State = cbGrayed) or (Principal.State = cbUnchecked) then //tous
          begin
            if TOBL.GetString('C_NUMEROCONTACT') <> '0' then StSQL := StSQL + '/' + TOBL.GetString('C_NUMEROCONTACT')
          end;
          if TOBL.GetString('C_PRENOM') <> '' then LibCompl := TOBL.GetString('C_PRENOM') + ' ';
          if TOBL.GetString('C_NOM')    <> '' then LibCompl := LibCompl + TOBL.GetString('C_NOM');
        end;
        //
        StSQL := StSQL + '", "' + Libenreg + '",';                    //Libellé principal
        //
        StSQL := StSQL + ' "' + LibCompl + '",';
        //
        if TOBL.GetString('NATURECAB') = 'MAR' then
        Begin
          //
          StSQL := StSQL + ' "' + RegimePrix + '",';
          StSQL := StSQL + IntToStr(NbEtiq)  + ',';
          StSQL := StSQL + StringReplace(TOBL.GetString('PVHT'),',','.',[rfReplaceAll])  + ',';
          StSQL := StSQL + StringReplace(TOBL.GetString('PVTTC'),',','.',[rfReplaceAll]) + ',';
        end
        else
        begin
          StSQL := StSQL + '"",';
          StSQL := StSQL + IntToStr(NbEtiq) + ',';
          StSQL := StSQL + '0,';
          StSQL := StSQL + '0,';
        end;
        StSQL := StSQL   + '"' + V_PGI.DevisePivot + '",';
        StSQL := StSQL   + '"' + TOBL.GetString('CODEBARRE') + '",';
        //
        StSQL := StSQL   + '"' + IntToStr(indice+1) + '", ';
        StSQL := StSQL   + '"' + NatureCAB.Value    + '",';
        StSQL := StSQL   + '"' + TOBL.GetString('QUALIFCAB') + '")';
        //
        if ExecuteSQL(StSQL) = 0 then PGIError('Enreg N°' + IntToStr(i) + '  ' + StSQL);
      end;
    end;
  end;

end;


function TOF_ETIQUETTE.Recupwhere : String;
begin

  If NatureCAB.Value = 'MAR' then
  begin
    Result := RecupWhereMarchandises;
  end
  else If NatureCAB.Value = 'FOU' then
  begin
    Result := RecupWhereFournisseurs;
    C_StWhere := RecupWhereContact;
  end
  Else If NatureCAB.Value = 'AFF' then
  begin
    Result := RecupWhereAffaires;
  end;

end;

Function TOF_ETIQUETTE.RecupWhereMarchandises : String;
Var Zone1 : String;
    Zone2 : String;
    Zone3 : TCheckBoxState;
    Zone4 : TDateTime;
    Zone5 : TDateTime;
    ind   : Integer;
begin

  //
  Zone1  := GetControlText('GA_CODEARTICLE');
  Zone2  := GetControlText('GA_CODEARTICLE_');
  Result := RecupZoneMultiple(zone1, Zone2, 'GA_CODEARTICLE');
  //
  Zone1  := GetControlText('GA_FAMILLENIV1');
  Result := RecupZoneListe(Zone1, 'GA_FAMILLENIV1');
  //
  Zone1  := GetControlText('GA_FAMILLENIV2');
  Result := RecupZoneListe(Zone1, 'GA_FAMILLENIV2');
  //
  Zone1  := GetControlText('GA_FAMILLENIV3');
  Result := RecupZoneListe(Zone1, 'GA_FAMILLENIV3');
  //
  Zone1  := GetControlText('GA_STATUTART');
  Result := RecupZoneListe(Zone1, 'GA_STATUTART');
  //
  Zone3  := THCheckbox(GetControl('GA_FERME')).State;
  Result := RecupZoneCheck(Zone3, 'GA_FERME');
  //
  Zone3  := THCheckbox(GetControl('GA_TENUESTOCK')).State ;
  Result := RecupZoneCheck(Zone3, 'GA_TENUESTOCK');
  //
  Zone4  := StrToDate(GetControlText('GA_DATECREATION'));
  Zone5  := StrToDate(GetControlText('GA_DATECREATION_'));
  Result := RecupZoneDate(Zone4, zone5, 'GA_DATECREATION');
  //
  Zone4  := StrToDate(GetControlText('GA_DATEMODIF'));
  Zone5  := StrToDate(GetControlText('GA_DATEMODIF_'));
  Result := RecupZoneDate(Zone4, zone5, 'GA_DATEMODIF');
  //
  //Tabsheet ==> Table Libres
  //
  For Ind:=1 to 9 do
  begin
    Zone1  := GetControlText('GA_LIBREART' + IntToSTr(Ind));
    Result := RecupZoneListe(Zone1, 'GA_LIBREART' + IntToStr(Ind));
  end;

  if OnlyCB.Checked then
  begin
    If Result <> '' then Result := Result + ' AND ';
    Result := ' GA_CODEBARRE <> ""';
  end;

end;

Function TOF_ETIQUETTE.RecupWhereAffaires : String;
Var Zone0   : String;
    Zone1   : String;
    Zone2   : string;
    Zone3   : string;
    ZoneA   : string;
    Zone4   : TDateTime;
    Zone5   : TDateTime;
begin

  //
  Zone0   := Part0.Text;
  Zone1   := Part1.Text;
  Zone2   := Part2.Text;
  Zone3   := Part3.Text;
  ZoneA   := Avenant.text;

  if (Zone1 <> '') OR (Zone2 <> '') OR (zone3 <> '') then
    Affaire.text := CodeAffaireRegroupe(Zone0, zone1, Zone2, zone3, ZoneA, taModif, True, False, True)
  else
    Affaire.text := '';
  //
  //
  Zone0   := Part0_.Text;
  Zone1   := Part1_.Text;
  Zone2   := Part2_.Text;
  Zone3   := Part3_.Text;
  ZoneA   := Avenant_.text;

  if (Zone1 <> '') OR (Zone2 <> '') OR (zone3 <> '') then
    Affaire_.text := CodeAffaireRegroupe(Zone0, zone1, Zone2, zone3, ZoneA, taModif, True, False, True)
  else
    Affaire_.text := '';
  //
  Result := RecupZoneMultiple(Affaire.text, Affaire_.text, 'AFF_AFFAIRE');
  //
  Zone1  := GetControlText('AFF_LIBELLE');
  Result := RecupZoneSimple(Zone1, 'AFF_LIBELLE');
  //
  Zone1  := GetControlText('AFF_TIERS');
  Result := RecupZonesimple(Zone1, 'AFF_TIERS');
  //
  Zone1  := GetControlText('AFF_ETABLISSEMENT');
  Result := RecupZoneListe(Zone1, 'AFF_ETABLISSEMENT');
  //
  Zone1  := GetControlText('AFF_ETATAFFAIRE');
  Result := RecupZoneListe(Zone1, 'AFF_ETATAFFAIRE');
  //
  Zone1  := GetControlText('AFF_RESPONSABLE');
  Result := RecupZoneSimple(Zone1, 'AFF_RESPONSABLE');
  //
  Zone4  := StrToDate(GetControlText('AFF_DATEDEBUT'));
  Zone5  := StrToDate(GetControlText('AFF_DATEDEBUT_'));
  Result := RecupZoneDate(Zone4, zone5, 'AFF_DATEDEBUT');
  //
  Zone4  := StrToDate(GetControlText('AFF_DATEFIN'));
  Zone5  := StrToDate(GetControlText('AFF_DATEFIN_'));
  Result := RecupZoneDate(Zone4, zone5, 'AFF_DATEFIN');

  if OnlyCB.Checked then
  begin
    If Result <> '' then Result := Result + ' AND ';
    Result := ' AFF_CODEBARRE <> ""';
  end;


end;

Function TOF_ETIQUETTE.RecupWhereFournisseurs : String;
Var Zone1 : String;
    Zone2 : String;
    Zone3 : TCheckBoxState;
    Zone4 : TDateTime;
    Zone5 : TDateTime;
    ind   : Integer;
begin

  //
  Zone1  := GetControlText('T_TIERS');
  Zone2  := GetControlText('T_TIERS_');
  Result := RecupZoneMultiple(zone1, Zone2, 'T_TIERS');
  //
  Zone1  := GetControlText('T_LIBELLE');
  Result := RecupZoneSimple(zone1, 'T_LIBELLE');
  //
  Zone1  := GetControlText('T_CODEPOSTAL');
  Result := RecupZoneSimple(zone1, 'T_CODEPOSTAL');
  //
  Zone1  := GetControlText('T_VILLE');
  Result := RecupZoneSimple(zone1, 'T_VILLE');
  //
  Zone1  := GetControlText('T_PAYS');
  Result := RecupZoneSimple(zone1, 'T_PAYS');
  //
  Zone1  := GetControlText('T_APE');
  Result := RecupZoneSimple(zone1, 'T_APE');
  //
  Zone1  := GetControlText('T_SECTEUR');
  Result := RecupZoneListe(Zone1, 'T_SECTEUR');
  //
  Zone3  := THCheckbox(GetControl('T_FERME')).State;
  Result := RecupZoneCheck(Zone3, 'T_FERME');
  //
  Zone4  := StrToDate(GetControlText('T_DATECREATION'));
  Zone5  := StrToDate(GetControlText('T_DATECREATION_'));
  Result := RecupZoneDate(Zone4, zone5, 'T_DATECREATION');
  //
  Zone4  := StrToDate(GetControlText('T_DATEMODIF'));
  Zone5  := StrToDate(GetControlText('T_DATEMODIF_'));
  Result := RecupZoneDate(Zone4, zone5, 'T_DATEMODIF');
  //

  //
  //Tabsheet ==> TableLibres
  //
  For Ind:=1 to 3 do
  begin
    Zone1  := GetControlText('YTC_TABLELIBREFOU' + IntToSTr(Ind));
    Result := RecupZoneListe(Zone1, 'YTC_TABLELIBREFOU' + IntToStr(Ind));
  end;

  if OnlyCB.Checked then
  begin
    If Result <> '' then Result := Result + ' AND ';
    Result := ' BT1_CODEBARRE <> ""';
  end;

end;

Function TOF_ETIQUETTE.RecupWhereContact : String;
Var Zone1 : String;
    Zone3 : TCheckBoxState;
begin

  //Tabsheet ==> contacts
  Result := 'C_NATUREAUXI="FOU" ';
  //
  Zone1  := GetControlText('C_NOM');
  Result := RecupZoneSimple(zone1, 'C_NOM');
  //
  Zone1  := GetControlText('C_FONCTION');
  Result := RecupZoneListe(Zone1, 'C_FONCTION');
  //
  Zone3  := THCheckbox(GetControl('C_PRINCIPAL')).State;
  Result := RecupZoneCheck(Zone3, 'C_PRINCIPAL');

end;

function TOF_ETIQUETTE.RecupZoneMultiple(Zone1, Zone2, LibZone : String) : String;
begin

  if zone1 = Zone2 then
  begin
    if Pos(zone1,'%') <> 0 then
    begin
      If Result <> '' then Result := Result + ' AND ';
      Result := result + Libzone + ' LIKE "' + Zone1 + '" '
    end
    else
    begin
      If Result <> '' then Result := Result + ' AND ';
      If Zone1 <> ''  then Result := Result + LibZone + '="' + Zone1 + '" '
    end;
  end
  else
  begin
    If Result <> '' then Result := Result + ' AND ';
    Result := Result + Libzone + '>="' + Zone1 + '" AND ' + Libzone + '<="' + Zone2 + '" ';
  end;

end;

function TOF_ETIQUETTE.RecupZoneSimple(Zone1, LibZone : String) : String;
Begin

  If Zone1 <> '' then
  begin
    If Result <> '' then Result := Result + ' AND ';
    Result := Result + LibZone + '="' + Zone1 + '" ';
  end;

end;

function TOF_ETIQUETTE.RecupZoneListe(Zone1, LibZone : String) : String;
Begin

  Zone1  := TransformeValeur(Zone1);
  if (Zone1 = '') or (Zone1 = '<<Tous>>') then
    exit
  else
  begin
    If Result <> '' then Result := Result + ' AND ';
    Result := Result + Libzone + ' IN ("'+ Zone1 + '") '
  end;

end;

function TOF_ETIQUETTE.RecupZoneCheck(Zone1 : TCheckBoxState; LibZone : String) : String;
Begin

  If Zone1 = CbGrayed then Exit;

  If Zone1 = CbChecked then 
  begin
    If Result <> '' then Result := Result + ' AND ';
    Result := Result + LibZone + ' ="X" '
  end
  Else
  begin
    If Result <> '' then Result := Result + ' AND ';
    Result := Result + LibZone + ' ="-" ';
  end;

end;

function TOF_ETIQUETTE.RecupZoneDate(Zone1, Zone2 : TdateTime; LibZone : String) : String;
var aa,mm,jj : Word;
    D1,D2 : TDateTime;
Begin

  If Result <> '' then Result := Result + ' AND ';
  D1 := Zone1;
  DecodeDate(D1,aa,mm,jj);
  D1 := EncodeDate(aa,mm,jj);

  D2 := IncDay(Zone2); 
  DecodeDate(D2,aa,mm,jj);
  D2 := EncodeDate(aa,mm,jj);

  Result := Result + LibZone + ' >="' + UsDateTime(D1) + '" AND ' + Libzone + ' <="' + UsDateTime(D2) + '" ';

end;

function TOF_ETIQUETTE.TransformeValeur(ValeurOri : string) : String;
var StVal       : string;
    valeurListe : string;
begin
  result:='';

  if valeurOri='' then exit;

  valeurListe := valeurOri;

  repeat
    stVal:=ReadToKenSt(valeurListe);
    if stVal <> '' then
    begin
      if result = '' then
        result :=stVal
      else
        result:=result+'","'+stVal;
    end;
  until stVal='';

end;

Procedure TOF_ETIQUETTE.SelectTypeEdition(Sender: TObject);
begin

  GBEDITETIQUETTEWORD.Visible := EditViaWord.Checked;
  BGenereWord.Visible := EditViaWord.Checked;

  if EditViaWord.Checked then ControleRepEtiquette;

  BValider.Visible  := not EditViaWord.Checked;
  FEtat.Visible     := not EditViaWord.Checked;
  TEtat.Visible     := not EditViaWord.Checked;
  BParamEtat.Visible:= not EditViaWord.Checked;

end;

//Bouton elipsis sur zoine modèle Etiquette Word
procedure TOF_ETIQUETTE.SelectModele(Sender: TObject);
var TT : TOpenDialog;
begin

	TT := TOpenDialog.Create(self.Ecran);

  TRY
    TT.DefaultExt := '.dotx';
    TT.Filter     := 'Modèle de document word (*.dotx)|*.dotx';
    TT.InitialDir := fRepServeur;
    if TT.Execute then
    begin
      //On copie du rep. Serveur vers le rep. Local
      ModeleWord.text := ExtractFileName(TT.FileName);
      if ModeleWord.text = '' then
        PGIError('Le modèle par défaut est obligatoire', 'Edition des etiquettes')
      else
        ModeleWord.text := TT.FileName;
    end;
  FINALLY
  	TT.Free;
  end;

end;

procedure TOF_ETIQUETTE.SelectDocWord(Sender: TObject);
var TT : TOpenDialog;
begin

	TT := TOpenDialog.Create(self.Ecran);

  TRY
    TT.DefaultExt := '.doc';
    TT.Filter     := 'Document word (*.doc)|*.doc';
    TT.InitialDir := GetParamSocSecur('SO_BTREPDOCETIQUETTE', '');
    if TT.Execute then DocGenere.text := TT.FileName;
  FINALLY
  	TT.Free;
  end;

end;


procedure TOF_ETIQUETTE.NewModeleClick(Sender: Tobject);
var TT          : TOB;
		ListeChamps : string;
    NomFile     : string;
    ListSqlEtq  : string;
begin

  ListeChamps := '';

  TT      := TOB.Create('BTTMPETQ',nil,-1);

  TRY
		NomFile := fRepServeur + '\NewDoc.docx';
    ConstitueListe('Etiquette', 'BTTMPETQ', ListeChamps, ListSqlEtq);
    LancePublipostage('NEW',NomFile,'',TT,ListeChamps,nil,False);
  finally
    //On copie le modèle sur le serveur une fois
    TT.Free;
  end;

end;

procedure TOF_ETIQUETTE.ModifModeleClick(Sender: Tobject);
var TT          : TOB;
		ListeChamps : string;
    ListSqlEtq  : String;
    NomFile     : string;
begin

  NomFile := ExtractFileName(ModeleWord.text);
  if (NomFile = '') or (NomFile = '*.dotx') then
  begin
  	PGIError('Vous devez renseigner un nom de modèle');
    Exit;
  end;

  ListeChamps := '';

  TT      := TOB.Create('BTTMPETQ',nil,-1);

  TRY
    ConstitueListe('Etiquette', 'BTTMPETQ', ListeChamps, ListSqlEtq);
    LancePublipostage('OPEN',ModeleWord.text,'',TT,ListeChamps,nil,False);
  finally
    TT.Free;
  end;

end;

procedure TOF_ETIQUETTE.SuppModeleClick(Sender: Tobject);
var NomFile : string;
begin

  NomFile := ExtractFileName(MODELEWORD.text);
  if (NomFile = '') or (NomFile = '*.dotx') then
  begin
  	PGIError('Vous devez renseigner un nom de modèle');
    Exit;
  end;

  if not FileExists(MODELEWORD.text) then
  begin
  	PGIError('Ce modèle n''existe pas');
    Exit;
  end;

  if PGIAsk('Désirez-vous réellement supprimer ce modèle ?',ecran.caption) = mryes then
  begin
    //Suppression du modèle sur le serveur
    DeleteFichier(MODELEWORD.text);
		ModeleWord.text := '';
  end;

end;

procedure TOF_ETIQUETTE.GenereWordClick(Sender: TObject);
var TT          : TOB;
		ListeChamps : string;
    ListSqlEtq  : string;
    NomFile     : string;
    NomDocx     : string;
    QQ          : TQuery;
    YY  : Word;
    DD  : Word;
    MM  : Word;
    H   : Word;
    M   : Word;
    S   : Word;
    MS  : Word;
    Ext : string;
begin

  NomFile := ExtractFileName(ModeleWord.text);

  if (NomFile = '')  or (NomFile = '*.dotx') then
  begin
  	PGIError('Vous devez renseigner un nom de modèle');
    Exit;
  end;

  if not FileExists(MODELEWORD.text) then
  begin
  	PGIError('Ce modèle n''existe pas');
    Exit;
  end;

  //on génère un autre fichier doc...
  DecodeDateTime(Now, YY, MM, DD, H,M,S,MS);
  Ext := IntToStr(DD) + IntToStr(MM) + IntToStr(YY) + IntToStr(H) + IntToStr(M) + IntToStr(S);
  NomDocx := 'Etiquette_' + Ext + '.doc';
  DocGenere.text := fRepArchiveL + '\' + NomDocx;

  //
  if (NomDocx = '') or (NomDocx = '*.doc') then
  begin
  	PGIError('Vous devez renseigner un nom de document de sortie');
    Exit;
  end;

  ListeChamps := '';

  //On lit les enregs et on les charge dans la table temporaire
  ChargeRequete;
  //
  TT      := TOB.Create('BTTMPETQ',nil,-1);

  TRY
    //Lecture de la table tempo
    StSQL   := 'SELECT * FROM BTTMPETQ WHERE BZD_UTILISATEUR = "' + V_PGI.USer + '"';
    QQ      := OpenSQL(StSQL, False);
    //chargement de la TOb
    TT.LoadDetailDB('BTTMPETQ','','',QQ,True,true);
    //
    ConstitueListe('Etiquette', 'BTTMPETQ', ListeChamps, ListSqlEtq);
    //
  	LancePublipostage('TOB',MODELEWORD.text,DocGenere.text,TT,ListeChamps,nil,False);
  finally
    TT.Free;
  end;

end;

procedure TOF_ETIQUETTE.ConstitueListe(Prefixe,TABLE : string; var ListeChamps,ListChampsSql : string);
var QQ : TQuery;
		Sql : string;
begin

  Sql := 'SELECT DH_NOMCHAMP,DH_LIBELLE FROM DECHAMPS WHERE DH_PREFIXE = "BZD" ORDER BY DH_NOMCHAMP';

	QQ := OpenSql(Sql,True,0,'',true);

  if Not QQ.Eof then
  begin
    QQ.first;
		While not QQ.eof do
    begin
      if ListChampsSql <> '' then ListChampsSql := ListChampsSql + ',';
      ListChampsSql := ListChampsSql +QQ.Fields [0].AsString;
			if ListeChamps <> '' then ListeChamps := ListeChamps + ';';
			ListeChamps := ListeChamps + Prefixe + '_' + QQ.Fields [1].AsString;

      QQ.Next;
    end;
  end;

  ferme (QQ);

end;

Function TOF_ETIQUETTE.CreationRepertoire(NomRep : String) : Boolean;
begin

  Result := True;

  if not DirectoryExists(NomRep) then
  begin
    if not CreationDir(NomRep) then
    begin
      PGIBox('le répertoire ' + NomRep + ' n''existe pas.', 'Gestion des métrés');
      Result := False;
      Exit;
    end;
  end;

end;

Initialization
  registerclasses ( [ TOF_ETIQUETTE ] ) ;
end.

