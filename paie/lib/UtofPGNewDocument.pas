{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 13/12/2006
Modifié le ... : 15/12/2006
Description .. : Source TOF de la FICHE : PGNEWDOCUMENT ()
Suite ........ : Fiche AGL reprise depuis RTnewDocument
Suite ........ : Fiche sauvegarde document Ged lié selon contexte à un 
Suite ........ : tiers
Mots clefs ... : TOF;PGNEWDOCUMENT
*****************************************************************}
{
PT1   : 11/06/2007 VG V_72 Adaptation nouvelles méthodes CBPPath
}
Unit UtofPGNewDocument ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_Main,           //AGLLanceFiche
{$else}
     eMul,
     MaineAGL,          //AGLLanceFiche
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HRichOle,          //THRichEditOle
     utilTom,           //tomtofgetcontrol
     wcommuns,          //GetNewGUID
     ParamSoc,          //Getparamsocsecur
     Dialogs,           //Topendialog
     UGedFiles,         //Openfile
     UtilGed,           //GetFileNameGed
     LookUp,            //lookuplist
     HTB97,             //TToolBarbutton97
     UTOF,
     PGRepertoire ;

Type
  TOF_PGNEWDOCUMENT = Class (TOF)
    private
    RefTiers            : String;
    TobDoc              : Tob;        //TOb YDOCUMENTS
    TobDocGed           : tob;         //TOb RTDOCUMENTS
    ModifDoc            : boolean;

    procedure LetiersOnElipsisClick ( Sender : Tobject );
    function RechRefTiers (CodeTiers : String) : String;
    procedure ActualiseTitre;
    procedure TitreDocOnChange ( Sender : Tobject );
    procedure LeFichierOnElipsisClick ( Sender : Tobject );
    procedure LesMotsClesOnExit ( Sender : Tobject );
    procedure bValiderOnClick ( Sender : Tobject );
    function SauveDocGed : boolean;
    procedure FormCloseQuery ( Sender : Tobject ; var CanClose : Boolean );
    procedure LesMotsClesOnChange ( Sender : Tobject );
    procedure LeBlocNoteOnChange ( Sender : Tobject );

    public

    bValider            : TToolbarButton97;
    bFerme              : TToolbarButton97;
    bDelete             : TToolbarButton97;
    bInsert             : TToolbarButton97;
    LaNature            : THValComboBox;
    THLibGed1           : THLabel;
    THLibGed2           : THLabel;
    THLibGed3           : THLabel;
    TGed1               : THValComboBox;
    TGed2               : THValComboBox;
    TGed3               : THValComboBox;
    LeTiers             : THEdit;
    LibelleTiers        : THEdit;
    TitreDoc            : THEdit;
    LeFichier           : THEdit;
    DateReception       : THedit;
    LesMotsCles         : THMemo;
    LeBlocNote          : THRichEditOle;


    FSDocGUID           : String;
    FNoDossier          : String;
    FCodeGed            : String;
    FSFileGUID          : String;
    FTypeGed            : String;
    FDefName            : String;
    FObjet              : String;
    FInfos              : String;
    FModifLien          : Boolean;


    procedure OnArgument (Arguments : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

  TParamGedDoc        = record
    SDocGUID          : String;
    NoDossier         : String;
    CodeGed           : String;
    SFileGUID         : String;
    TypeGed           : String;
    DefName           : String;
    Objet             : String;
    Infos             : String;
    ModifLien         : Boolean;
  end;

  function ShowNewDocument ( ParamGedDoc:TParamGedDoc ) : String;
  procedure GereTableLibreGed ( ParamCtx : String ; FF : TForm );

Implementation

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 15/12/2006
Modifié le ... :   /  /
Description .. : Gestion des tables libres Ged selon contexte
Mots clefs ... :
*****************************************************************}
procedure GereTableLibreGed ( ParamCtx : String ; FF : TForm );
var
  i                     : integer;
  TobLibGed             : Tob;
  TobLibGed1            : Tob;
  TobLibGed2            : Tob;
  TobLibGed3            : Tob;
  TobG                  : Tob;
  StrSql                : String;
  CombG                 : THValComboBox;
  LibG                  : THLabel;

  //affiche les combos et libellés si ok
  procedure AffLib(LaTob: Tob ; ComboGed : THValComboBox ; LibGed : THLabel);
  begin
    LibGed.Visible      := True;
    ComboGed.Visible    := True;
    LibGed.Caption      := LaTob.GetValue('CC_LIBELLE');
  end;

  //rempli la combo avec les valeurs du contexte
  procedure RplTable(LaTob: Tob; ComboGed: THValComboBox);
  var
    i                     : integer;
  begin
    ComboGed.Items.Add('<<Aucun>>');
    ComboGed.Values.Add('');
    for i := 0 to LaTob.Detail.Count -1 do
    begin
      ComboGed.Items.Add(LaTob.Detail[i].GetValue('CC_LIBELLE'));
      ComboGed.Values.Add(LaTob.Detail[i].GetValue('CC_CODE'));
    end;
  end;

begin
  //Chargement des tables libres
  TobLibGed             := Tob.Create('CHOIXCOD', nil, -1);
  TobLibGed1            := Tob.Create('CHOIXCOD', nil, -1);
  TobLibGed2            := Tob.Create('CHOIXCOD', nil, -1);
  TobLibGed3            := Tob.Create('CHOIXCOD', nil, -1);
  StrSql              := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE = "PG0" AND CC_LIBRE LIKE ' + ParamCtx;
  TobLibGed.LoadDetailFromSQL(StrSql);
  StrSql              := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE = "PG1" AND CC_LIBRE LIKE ' + ParamCtx;
  TobLibGed1.LoadDetailFromSQL(StrSql);
  StrSql              := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE = "PG2" AND CC_LIBRE LIKE ' + ParamCtx;
  TobLibGed2.LoadDetailFromSQL(StrSql);
  StrSql              := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE = "PG3" AND CC_LIBRE LIKE ' + ParamCtx;
  TobLibGed3.LoadDetailFromSQL(StrSql);

  for i := 0 to TobLibGed.Detail.Count -1 do
  begin
    if i > 2 then     // s'il y a plus de 3 tables .... pas normal
      Continue;
    tobG              := TobLibGed.Detail[i];
    CombG             := THValComboBox(TOMTOFGetControl(FF, 'RTD_TABLELIBREGED'+IntToStr(i+1)));
    LibG              := THLabel(TOMTOFGetControl(FF, 'TRTD_TABLELIBREGED'+IntToStr(i+1)));
    AffLib(TobG, CombG, LibG);
  end;

  CombG               := THValComboBox(TOMTOFGetControl(FF, 'RTD_TABLELIBREGED1'));
  if CombG <> nil then CombG.DataType := 'PGLIBGED1';
//  RplTable(TobLibGed1, CombG);
  CombG               := THValComboBox(TOMTOFGetControl(FF, 'RTD_TABLELIBREGED2'));
  if CombG <> nil then CombG.DataType := 'PGLIBGED2';
//  RplTable(TobLibGed2, CombG);
  CombG               := THValComboBox(TOMTOFGetControl(FF, 'RTD_TABLELIBREGED3'));
  if CombG <> nil then CombG.DataType := 'PGLIBGED3';
//  RplTable(TobLibGed3, CombG);
  TobLibGed.Free;
  TobLibGed1.Free;
  TobLibGed2.Free;
  TobLibGed3.Free;

end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 15/12/2006
Modifié le ... :   /  /
Description .. : Appel de la fiche après le scan ou bien directement depuis
Suite ........ : le mul de recherche
Mots clefs ... :
*****************************************************************}
function ShowNewDocument ( ParamGedDoc:TParamGedDoc ) : String;
var
  StrArg                : String;

begin
  With ParamGedDoc do
  begin
    StrArg              := SDocGUID + ';' + SFileGUID + ';' + NoDossier + ';' + CodeGed + ';' + TypeGed + ';' + DefName +
                           ';' + Infos + ';' + Objet + ';' + BoolToStr(ModifLien);
  end;
  Result                := AGLLanceFiche('PAY', 'PGNEWDOCUMENT', '', '', StrArg);
end;



procedure TOF_PGNEWDOCUMENT.OnArgument (Arguments : String ) ;
var
  Critere               : String;
  ChampMul              : STring;
  ValMul                : String;
  x                     : integer;
  THTiers               : THLabel;
  Q                     : TQuery;
  ParamCtx              : String;


begin
  Inherited ;

  //Récupère des infos
  Critere               := UpperCase_(Trim(ReadTokenSt(Arguments)));
  FSDocGUID             := Critere;
  Critere               := UpperCase_(Trim(ReadTokenSt(Arguments)));
  FSFileGUID            := Critere;
  Critere               := UpperCase_(Trim(ReadTokenSt(Arguments)));
  FNoDossier            := Critere;
  Critere               := UpperCase_(Trim(ReadTokenSt(Arguments)));
  FCodeGed              := Critere;
  Critere               := UpperCase_(Trim(ReadTokenSt(Arguments)));
  FTypeGed              := Critere;
  Critere               := UpperCase_(Trim(ReadTokenSt(Arguments)));
  FDefName              := Critere;
  Critere               := UpperCase_(Trim(ReadTokenSt(Arguments)));
  FInfos                := Critere;
  Critere               := UpperCase_(Trim(ReadTokenSt(Arguments)));
  FObjet                := Critere;
  Critere               := UpperCase_(Trim(ReadTokenSt(Arguments)));
  FModifLien            := StrToBool_(Critere);

  x                     := Pos ('=', FInfos);
  ChampMul              := copy(FInfos, 1, x-1);
  ValMul                := copy(FInfos, x+1, length(FInfos));
  if UpperCase_(ChampMul) = 'SALARIE' then
    RefTiers            := ValMul;


  // les propriétés
  bValider              := TToolbarButton97(GetControl('BVALIDER'));
  bFerme                := TToolbarButton97(GetControl('BFERME'));
  bDelete               := TToolbarButton97(GetControl('BDELETE'));
  bInsert               := TToolbarButton97(GetControl('BINSERT'));
  LaNature              := THValComboBox(GetControl('YDO_NATDOC'));
  TGed1                 := THValComboBox(GetControl('RTD_TABLELIBREGED1'));
  TGed2                 := THValComboBox(GetControl('RTD_TABLELIBREGED2'));
  TGed3                 := THValComboBox(GetControl('RTD_TABLELIBREGED3'));
  THLibGed1             := THLabel(GetControl('TRTD_TABLELIBREGED1'));
  THLibGed2             := THLabel(GetControl('TRTD_TABLELIBREGED2'));
  THLibGed3             := THLabel(GetControl('TRTD_TABLELIBREGED3'));
  THTiers               := THLabel(GetControl('TRTD_TIERS'));
  LeTiers               := THEdit(GetControl('RTD_TIERS'));
  LibelleTiers          := THEdit(GetControl('LIBELLETIERS'));
  TitreDoc              := THEdit(GetControl('YDO_LIBELLEDOC'));
  LeFichier             := THEdit(GetControl('LEFICHIER'));
  Datereception         := THEdit(GetControl('RTD_DATERECEPTION'));
  LesMotsCles           := THMemo(GetControl('YDO_MOTSCLES'));
  LeBlocNote            := THRichEditOLE(GetControl('YDO_BLOCNOTE'));

  TForm(Ecran).OnCloseQuery       := FormCloseQuery;
  bValider.OnClick      := bValiderOnClick;
  LeTiers.ElipsisButton := True;
  LeTiers.OnElipsisClick          := LetiersOnElipsisClick;
  TitreDoc.OnChange     := TitreDocOnChange;
  LeFichier.ElipsisButton         := True;
  LeFichier.OnElipsisClick        := LeFichierOnElipsisClick;
  LaNature.DataType     := 'YYNATDOC';            //nature des documents
  LaNature.DataTypeParametrable   := True;
  TGed1.Vide            := True;
  TGed1.VideString      := '<<Aucun>>';
  TGed1.Visible         := False;
  THLibGed1.Visible     := False;
  TGed2.Vide            := True;
  TGed2.VideString      := '<<Aucun>>';
  TGed2.Visible         := False;
  THLibGed2.Visible     := False;
  TGed3.Vide            := True;
  TGed3.VideString      := '<<Aucun>>';
  TGed3.Visible         := False;
  THLibGed3.Visible     := False;
  LesMotsCles.OnExit    := LesMotsClesOnExit;
  LesMotsCles.OnChange  := LesMotsClesOnChange;
  LeBlocNote.OnChange   := LeBlocNoteOnChange;

  if ctxPaie in V_PGI.PGIContexte then            // Contexte PAIE
  begin
    THTiers.Caption     := 'Salarié';
    ParamCtx            := '"%PAIE%"';
  end
  else if ctxCompta in V_PGI.PGIContexte then     // Contexte COMPTA
  begin
    THTiers.Caption     := 'Tiers';
    ParamCtx            := '"%COMPTA%"';
  end;

  GereTableLibreGed(ParamCtx, TForm(Ecran));      // maj des libellés et combos pour tables libres Ged

  // les tob pour le document
  TobDoc                := Tob.Create('YDOCUMENTS', nil, -1);
  TobDocGed             := Tob.Create('RTDOCUMENT', nil, -1);

  // Si le document existe
  if FSDocGUID <> '' then
  begin
    TobDoc.PutValue('YDO_DOCGUID', FSDocGUID);
    TobDocGed.PutValue('RTD_DOCGUID', FSDocGUID);
    TobDocGed.LoadDB;
    TobDoc.LoadDB;
    TobDoc.PutEcran(TForm(Ecran));
    TobDocGed.PutEcran(TForm(Ecran));
    LeTiers.Text        := TobDocGed.GetValue('RTD_TIERS');
    Q                   := OpenSQL('SELECT YDF_FILEGUID FROM YDOCFILES WHERE YDF_DOCGUID= "' + FSDocGUID + '" AND YDF_FILEROLE="PAL"', True);
    if Not Q.Eof then
      FSFileGUID        := Q.FindField('YDF_FILEGUID').AsString;
    Ferme(Q);
  end
  else
  // Si c'est un nouveau document
  begin
    ModifDoc            := True;
    if RefTiers <> '' then
    begin
      LeTiers.Text      := RefTiers;
    end;

    if FDefName <> '' then
      TitreDoc.Text     := FDefName
    else
      TitreDoc.Text     := LeFichier.Text;

  end;

  if RefTiers <> '' then
    LibelleTiers.Text := RechRefTiers(RefTiers);

  if FSFileGUID <> '' then
  begin
    LeFichier.Enabled   := False;
    LeFichier.Text      := GetFileNameGed(FSFileGUID);
  end;

  if not FModifLien then
  begin
    LeTiers.Enabled     := False;
  end;

  ActualiseTitre;

  //si le document existe déjà
  if FSDocGUID <> '' then
    ModifDoc            := False;

end ;


procedure TOF_PGNEWDOCUMENT.OnClose ;
begin
  Inherited ;
  TobDoc.Free;
  TobDocGed.Free;

end ;


{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 14/12/2006
Modifié le ... :   /  /    
Description .. : Choix du Tiers
Suite ........ : lookup différent selon contexte (Paie, compta)
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGNEWDOCUMENT.LetiersOnElipsisClick(Sender: Tobject);

begin
  if ctxPaie in V_PGI.PGIContexte then
  begin
    if lookuplist(LeTiers, 'Liste des salariés', 'SALARIES', 'PSA_AUXILIAIRE', 'PSA_LIBELLE', '', 'PSA_AUXILIAIRE', True, -1) then
    begin
      LibelleTiers.Text := RechRefTiers(LeTiers.Text);
      RefTiers          := LeTiers.Text;
      ModifDoc          := True;
    end;
  end;

  if ctxCompta in V_PGI.PGIContexte then
  begin
    if lookuplist(LeTiers, 'Liste des Tiers', 'TIERS', 'T_AUXILIAIRE', 'T_LIBELLE', 'T_NATUREAUXI="CLI" OR T_NATUREAUXI="FOU"', 'T_AUXILIAIRE', True, -1) then
    begin
      LibelleTiers.Text := RechRefTiers(LeTiers.Text);
      RefTiers          := LeTiers.Text;
      ModifDoc          := True;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 14/12/2006
Modifié le ... :   /  /
Description .. : retourne le code selon contexte
Suite ........ : SAL_AUXILIAIRE -> code salarié paie
Mots clefs ... :
*****************************************************************}
function TOF_PGNEWDOCUMENT.RechRefTiers(CodeTiers: String): String;
Var
  Q                     : TQuery;

begin
  Result                := '';
  if ctxPaie in V_PGI.PGIContexte then
  begin
    Q                   := OpenSQL('SELECT PSA_LIBELLE, PSA_PRENOM FROM SALARIES WHERE PSA_AUXILIAIRE="' + CodeTiers + '"', True);
    if not Q.Eof then
    begin
      Result            := Trim(Trim(Q.FindField('PSA_LIBELLE').asstring) + ' ' + Trim(Q.FindField('PSA_PRENOM').asstring));
      ModifDoc          := True;
    end;
    Ferme(Q);
  end;

  if ctxCompta in V_PGI.PGIContexte then
  begin
    Q                   := OpenSQL('SELECT T_LIBELLE FROM TIERS WHERE T_AUXILIAIRE="' + CodeTiers + '"', True);
    if not Q.Eof then
    begin
      Result            := Trim(Q.FindField('T_LIBELLE').asstring);
      ModifDoc          := True;
    end;
    Ferme(Q);
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 14/12/2006
Modifié le ... :   /  /    
Description .. : mise à jour de la barre de titre de la fenêtre
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGNEWDOCUMENT.ActualiseTitre;
var
  Titre                 : String;

begin
  Titre                 := TraduireMemoire('Document');
  if FSDocGUID = '' then
    Titre               := TraduireMemoire('Nouveau ') + LowerCase_(Titre);
  if TitreDoc.Text <> '' then
    Titre               := Titre + ' - ' + TitreDoc.Text;

  TForm(Ecran).Caption  := Titre;
  UpdateCaption(TForm(Ecran));

end;

procedure TOF_PGNEWDOCUMENT.TitreDocOnChange(Sender: Tobject);
begin
  ModifDoc              := True;
  ActualiseTitre;
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 14/12/2006
Modifié le ... :   /  /    
Description .. : boite de dialogue de sélection des fichiers
Mots clefs ... :
*****************************************************************}
procedure TOF_PGNEWDOCUMENT.LeFichierOnElipsisClick(Sender: Tobject);
var
  OdFichier             : Topendialog;
  StrFiltre             : String;

begin
  OdFichier             := TOpenDialog.Create(Application);
  Try
  begin
    OdFichier.Title     := 'Sélectionner un fichier';
    StrFiltre           := 'Fichier Word (*.doc)|*.doc|';
    StrFiltre           := StrFiltre + 'Fichier Excel (*.xls)|*.xls|';
    StrFiltre           := StrFiltre + 'Fichier Pdf (*.pdf)|*.pdf|';
    StrFiltre           := StrFiltre + 'Images (gif, bmp, jpg, ico)|*.gif;*.bmp;*.jpg;*.ico|';
    StrFiltre           := StrFiltre + 'Fichier texte (*.txt)|*.txt|';
    StrFiltre           := StrFiltre + 'Tous les fichier (*.*)|*.*';
    OdFichier.Filter    := StrFiltre;
{PT1
    OdFichier.InitialDir  := GetParamSocSecur('SO_PGCHEMINRECH', 'C:\TEMP', True);    //chemin de recherche
}
    OdFichier.InitialDir:= VerifieCheminPG (GetParamSocSecur('SO_PGCHEMINRECH', 'C:\TEMP', True));    //chemin de recherche
//FIN PT1
    OdFichier.FilterIndex := 6;                                                       //par défaut, tous les fichiers
    If OdFichier.Execute then
    begin
      LeFichier.Text    := OdFichier.FileName;
      ModifDoc          := True;
    end;
      
  end;
  finally
    OdFichier.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 14/12/2006
Modifié le ... :   /  /    
Description .. : gestion des mots clés
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGNEWDOCUMENT.LesMotsClesOnExit(Sender: Tobject);
begin
  LesMotsCles.Text      := UpperCase_(LesMotsCles.Text);
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 14/12/2006
Modifié le ... :   /  /
Description .. : Validation et enregistrement du document
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGNEWDOCUMENT.bValiderOnClick(Sender: Tobject);
begin
  if not SauveDocGed then
    TForm(Ecran).ModalResult    := mrNone
  else
    inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 14/12/2006
Modifié le ... :   /  /    
Description .. : enregistrement du document dans les tables Ged
Mots clefs ... : 
*****************************************************************}
function TOF_PGNEWDOCUMENT.SauveDocGed: boolean;
var
  TobDocFiles           : tob;
  Nouveau               : Boolean;

begin
  Result                := False;
  Nouveau               := False;

  if RechRefTiers(LeTiers.Text) = '' then
  begin
    PGIInfo('Le code tiers est inexistant', TForm(Ecran).Caption);
    LeTiers.SetFocus;
    Exit;
  end;

  if TitreDoc.Text = '' then
  begin
    PGIInfo('Il faut obligatoirement un titre', TForm(Ecran).Caption);
    TitreDoc.SetFocus;
    Exit;
  end;

  if FSFileGUID = '' then
  begin
    if LeFichier.Text = '' then
    begin
      PGIInfo('Aucun fichier à insérer', TForm(Ecran).Caption);
      LeFichier.SetFocus;
      Exit;
    end;

    if not FileExists(LeFichier.Text) then
    begin
      PGIInfo(Format(TraduireMemoire('Le fichier %s n''existe plus.'),[LeFichier.Text]), TForm(Ecran).Caption);
      LeFichier.SetFocus;
      Exit;
    end;

    FSFileGUID          := V_GedFiles.Import(LeFichier.Text);
    Nouveau             := True;
    if FSFileGUID = '' then
    begin
      PGIInfo(Format(TraduireMemoire('Le fichier %s ne peut pas être importé dans la GED.'),[LeFichier.Text]), TForm(Ecran).Caption);
      Exit;
    end;
  end;

  if FSDocGUID = '' then
  begin
    FSDocGUID           := AGLGetGuid;
    Nouveau             := True;
  end;

  TobDoc.GetEcran(TForm(Ecran));
  TobDocGed.GetEcran(TForm(Ecran));

  TobDoc.PutValue('YDO_DOCGUID', FSDocGUID );
  TobDoc.PutValue('YDO_DOCID', -2);
  TobDoc.PutValue('YDO_ANNEE', FormatDateTime('yyyy', StrToDateTime(DateReception.Text)));
  TobDoc.PutValue('YDO_MOIS', FormatDateTime('mm', StrToDateTime(DateReception.Text)));
  TobDoc.PutValue('YDO_DOCREF', FSFileGUID + ' ' + FSDocGUID );
  TobDoc.PutValue('YDO_AUTEUR',V_PGI.UserName );

  TobDocGed.PutValue('RTD_DOCGUID', FSDocGUID );
  TobDocGed.PutValue('RTD_TIERS', LeTiers.Text);
  TobDocGed.PutValue('RTD_DOCID', -2);

  if Nouveau then
  begin
    TobDocFiles           := Tob.Create('YDOCFILES', nil, -1);
    TobDocFiles.PutValue('YDF_DOCGUID', FSDocGUID);
    TobDocFiles.PutValue('YDF_FILEGUID', FSFileGUID);
    TobDocFiles.PutValue('YDF_DOCID', -2);
    TobDocFiles.LoadDB;
    TobDocFiles.PutValue('YDF_FILEROLE', 'PAL');
    TobDocFiles.InsertOrUpdateDB(False);
    TobDocFiles.Free;
  end;

  Try
    TobDoc.InsertOrUpdateDB(False);
    TobDocGed.InsertOrUpdateDB(False);
    Result              := True;
    ModifDoc            := False;
  except
    PGIInfo('Erreur lors de la sauvegarde', TForm(Ecran).Caption);
  end;


end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 14/12/2006
Modifié le ... :   /  /    
Description .. : sortie de la fenêtre
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGNEWDOCUMENT.FormCloseQuery(Sender: Tobject; var CanClose: Boolean);
var
  Rep                   : integer;

begin
  if ModifDoc then
  begin
    Rep                 :=  PGIAskCancel('Voulez vous enregistrer les modifications ?', TForm(Ecran).Caption);
    if Rep = mrCancel then
      CanClose          := False
    else if Rep = mrNo then
    begin
      ModifDoc          := False;
      if FSFileGUID <> '' then
        V_GedFiles.Erase(FSFileGUID);
    end
    else if Rep = mrYes then
      CanClose          := SauveDocGed;
  end;
end;



procedure TOF_PGNEWDOCUMENT.LeBlocNoteOnChange(Sender: Tobject);
begin
  ModifDoc              := True;
end;

procedure TOF_PGNEWDOCUMENT.LesMotsClesOnChange(Sender: Tobject);
begin
  ModifDoc              := True;
end;

Initialization
  registerclasses ( [ TOF_PGNEWDOCUMENT ] ) ;
end.
