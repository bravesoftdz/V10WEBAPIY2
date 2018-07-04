unit YNewDocument;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, StdCtrls, ComCtrls, HRichEdt, HRichOLE, Mask, Hctrls, ExtCtrls,
  HSysMenu, HTB97, UGedFiles,

{$IFDEF EAGLCLIENT}
  MaineAGL,
{$ELSE}
  FE_Main,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  UIUtil, UTob, HPanel, HMsgBox, HEnt1, DBCtrls, HDB, ImgList, CBPPath;

Const
  //--- Liste image
  CImgActif       = 0;
  CImgVerrouille  = 1;
  CImgPerime      = 2;

  //--- TypeDocument
  CEXTAUTRE       = 0;
  CEXTDOC         = 1;
  CEXTXLS         = 2;

type

  TParamGedDoc = Record
                  SDocGUID   : String;
                  SFileGUID  : String;
                  NoDossier  : String;
                  CodeGed    : String;
                  TypeGed    : String;
                  DefName    : String;
                  strUrl     : String;
                  Annee      : String;
                  Mois       : String;
                 end;

  TFNewDocument = class(TFVierge)
    PNewDoc: TPageControl;
    TSCaracteristiques: TTabSheet;
    LblEmplacement: THLabel;
    LblDossier: THLabel;
    BEfface3: TToolbarButton97;
    Dossier: THCritMaskEdit;
    LblArmoire: TLabel;
    BEfface1: TToolbarButton97;
    CODEGED1: THValComboBox;
    LblClasseur: TLabel;
    BEfface2: TToolbarButton97;
    CODEGED2: THValComboBox;
    ImgEws: TImage;
    LblBrancheEws: THLabel;
    LblProprietes: THLabel;
    LblFichier: THLabel;
    Fichier: THCritMaskEdit;
    LblLibelleDoc: THLabel;
    YDO_LIBELLEDOC: TEdit;
    LblAuteur: THLabel;
    YDO_AUTEUR: TEdit;
    LblAnnee: THLabel;
    YDO_ANNEE: TEdit;
    LblDateDeb: THLabel;
    YDO_DATEDEB: THCritMaskEdit;
    LblDateFin: THLabel;
    YDO_DATEFIN: THCritMaskEdit;
    LblMois: THLabel;
    YDO_MOIS: TEdit;
    LblBlocnote: THLabel;
    YDO_BLOCNOTE: THRichEditOLE;
    LblTheme: TLabel;
    YMO_THEMEMODELE: THValComboBox;
    TSRecherche: TTabSheet;
    LblTablesLibres: THLabel;
    Ltablelibre1: TLabel;
    BEffaceLib1: TToolbarButton97;
    YDO_LIBREGED1: THValComboBox;
    Ltablelibre2: TLabel;
    BEffaceLib2: TToolbarButton97;
    YDO_LIBREGED2: THValComboBox;
    Ltablelibre3: TLabel;
    BEffaceLib3: TToolbarButton97;
    YDO_LIBREGED3: THValComboBox;
    Ltablelibre4: TLabel;
    BEffaceLib4: TToolbarButton97;
    YDO_LIBREGED4: THValComboBox;
    Ltablelibre5: TLabel;
    BEffaceLib5: TToolbarButton97;
    YDO_LIBREGED5: THValComboBox;
    LblNatDocument: THLabel;
    LblMotsCles: THLabel;
    BEfface4: TToolbarButton97;
    YDO_NATDOC: THValComboBox;
    YDO_MOTSCLES: TMemo;
    TSInformation: TTabSheet;
    LblAutresCriteres: THLabel;
    HLDroits: THLabel;
    YDO_DROITGED: THValComboBox;
    LargeImages: TImageList;
    LblCreateur: THLabel;
    LblEtatDocument: THLabel;
    LblDroitDocument: THLabel;
    HLIntegrePar: THLabel;
    HIntegrePar: THLabel;
    HLIntegreLe: THLabel;
    HIntegreLe: THLabel;
    TImageDocEtat: TImage;
    HDocEtat: THLabel;
    HLExtraitPar: THLabel;
    HLExtraitLe: THLabel;
    HExtraitLe: THLabel;
    HExtraitPar: THLabel;
    DPD_ARBOGED: THRadioGroup;
    YDO_PRIVE: THCheckbox;
    BRemplacer: TToolbarButton97;
    BEfface5: TToolbarButton97;
    LblIntercal: THLabel;
    CodeGed3: THValComboBox;
    BVisualiser: TToolbarButton97;

    procedure FormDestroy(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure OnDocChange(Sender: TObject);
    procedure YDO_LIBELLEDOCExit(Sender: TObject);
    procedure YDO_MOTSCLESExit(Sender: TObject);
    procedure CODEGED1Change(Sender: TObject);
    procedure CODEGED2Change(Sender: TObject);
    procedure InitialiserClasseur;
    procedure InitialiserIntercalaire;
    procedure BEfface1Click(Sender: TObject);
    procedure BEfface2Click(Sender: TObject);
    procedure BEfface3Click(Sender: TObject);
    procedure BEfface4Click(Sender: TObject);
    procedure BDocumentClick(Sender: TObject);
    procedure DossierElipsisClick(Sender: TObject);
    procedure FichierElipsisClick(Sender: TObject);
    procedure YDO_DATEDEBExit(Sender: TObject);
    procedure YDO_DATEFINExit(Sender: TObject);
    procedure DossierChange(Sender: TObject);
    procedure BEffaceLib1Click(Sender: TObject);
    procedure BEffaceLib2Click(Sender: TObject);
    procedure BEffaceLib3Click(Sender: TObject);
    procedure BEffaceLib4Click(Sender: TObject);
    procedure BEffaceLib5Click(Sender: TObject);
    procedure DPD_ARBOGEDClick(Sender: TObject);
    procedure BRemplacerClick(Sender: TObject);
    procedure BEfface5Click(Sender: TObject);
    procedure CodeGed3Change(Sender: TObject);

  private
    { Déclarations privées }
    FIconFiles   : TIconFiles;
    SDocGUID, SDocGUIDASupprimer : string;
    TypeGed    : String;
    TobDocGed  : TOB;      // Tob de l'enregistrement DPDOCUMENT ou YMODELE (tables métier)
    TobDoc     : TOB;      // Tob de l'enregistrement YDOCUMENTS (gestion documentaire)
    SFileGUID     : string;  // Id si le fichier à archiver est déjà dans YFILES
    DefName    : String;   // Nom de fichier par défaut
    bDuringLoad, bChange, bDocChange_f : Boolean ; // Chgt dans les données
    bEwsActif   : Boolean; // Est-ce que eWS est actif pour ce cabinet (=VH_DP.EwsActif)
    bDossierEws : Boolean; // Est-ce que le dossier en cours est actif sur eWS

    function  DonnerNomFichier : String;
    function  SauveEnregDocument : Boolean;
    procedure ActualiseTitre;
    function  ControleDatesValidite : Boolean;
    procedure ActualiseBrancheEws;
  public
    { Déclarations publiques }
    UnFichierRemplacer : String;
    bModifie    : Boolean;
    NoDossier   : String;
    CodeGed     : String;
    bModele     : Boolean; // Gestion d'un modèle de doc
    SArboGed    : String;
    TypeFichier : Integer;
    Annee       : String;
    Mois        : String;
    bNoChangeCodeGed : Boolean; // $$$ JP 28/09/04 - pour interdir sur demande la modification dans la hiérarchie GED
  end;

function ShowNewDocument(ParamGedDoc:TParamGedDoc; bInside:Boolean=False; bModele:Boolean=False; bNoChangeCodeGed:boolean=FALSE; SArboGed : String='D'; TypeFichier : Integer=CEXTAUTRE):string;
{                        iDocId : Integer ; sNoDossier, sCodeGed : String ;
                         iFileId : Integer=0 ; sTypeGed : String='DOC' ;
                         bInside : Boolean=False; sDefName: String='') : String;
}

////////// IMPLEMENTATION /////////////
implementation


uses
     {$IFDEF EWS}
     UtileWS,
     {$ENDIF}
     UtilDossier, UtilGed, galOutil, ParamSoc;


//------------------------
//--- Nom : FormDestroy
//------------------------
procedure TFNewDocument.FormDestroy(Sender: TObject);
begin
 if FIconFiles<>Nil then begin FIconFiles.Free; FIconFiles := Nil; end;
 inherited;
end;

//----------------------------
//--- Nom : ShowNewDocument
//----------------------------
function ShowNewDocument (ParamGedDoc:TParamGedDoc; bInside:Boolean=False; bModele:Boolean=False; bNoChangeCodeGed:boolean=FALSE; SArboGed : String='D'; TypeFichier : Integer=CEXTAUTRE):string;
// Retourne :
// - ##NONE## si l'enregistrement n'est pas validé
// - le NoDossier;CodeGed si validé (rq : le CodeGed peut être '')
var
  F : TFNewDocument;
  PP : THPanel;
begin
  Result := '##NONE##';
  F := TFNewDocument.Create(Application);
  with ParamGedDoc do
  begin
    F.SDocGUID    := SDocGUID;
    F.SFileGUID   := SFileGUID;
    F.NoDossier   := NoDossier;
    F.CodeGed     := CodeGed;
    if TypeGed='' then
      F.TypeGed := 'DOC'
    else
      F.TypeGed := TypeGed;
    F.DefName := DefName;
    F.Annee       := Annee;
    F.Mois        := Mois;
  end;
  F.bModele := bModele;
  // si choix d'un élément de l'arborescence incompatible avec la branche choisie
  if (F.CodeGed<>'') and (Copy(F.CodeGed, 1, 1)<>SArboGed) then
    F.SArboGed := Copy(F.CodeGed, 1, 1)
  else if SArboGed<>'' then
    F.SArboGed := SArboGed
  else
    F.SArboGed := 'D';
  if F.CodeGed='' then F.CodeGed := F.SArboGed;
  F.TypeFichier := TypeFichier;

  PP := Nil;
  if bInside then
    begin
    // Ferme la précédente fiche éventuelle
    if PrepareInside then PP := FindInsidePanel;
    end;
  if PP<>Nil then
    begin
    InitInside(F,PP);
    F.Show;
    end
  else
    begin
    try
      // pour affichage des commentaires en inside, elle est sizeable, et là on veut pas
      F.BorderStyle      := bsDialog;
      F.bNoChangeCodeGed := bNoChangeCodeGed;
      F.ShowModal ;
    finally
      // MD 09/10/07 - un seul type de code retour quels que soient les paramètres
      // if (TypeFichier<>CEXTAUTRE) and (F.ModalResult<>MrCancel) then Result := F.SDocGuid;
      if (F.bModifie) and (F.ModalResult<>MrCancel) then
        Result := F.NoDossier+';'+F.CodeGed+';'+F.SDocGuid;
      F.Free ;
      end;
    end ;
end;

{$R *.DFM}

procedure TFNewDocument.FormShow(Sender: TObject);
var chemged, ged1, ged2, ged3 : String;
    Q                         : TQuery;
    strLibLibre               : String;
    C1, C2                    : TControl;
    I1, I2                    : Integer;
begin
  inherited;
  BDuringLoad:=True;

  FIconFiles := TIconFiles.Create;
  FIconFiles.LargeImageList := LargeImages;

  // Gestion d'un modèle de document : pas la même table métier
  if bModele then
    begin
    LblDossier.Visible := False;
    Dossier.Visible := False;
    LblArmoire.Visible := False;
    LblClasseur.Visible := False;
    LblIntercal.Visible :=False;
    CODEGED1.Visible := False;
    bEfface1.Visible := False;
    CODEGED2.Visible := False;
    bEfface5.Visible := False;
    CODEGED3.Visible := False;
    bEfface2.Visible := False;
    bEfface3.Visible := False;
    LblAnnee.Visible := False;
    YDO_ANNEE.Visible := False;
    LblMois.Visible := False;
    YDO_MOIS.Visible := False;
    LblMotsCles.Visible := False;
    YDO_MOTSCLES.Visible := False;
    LblNatDocument.Visible := False;
    YDO_NATDOC.Visible := False;
    bEfface4.Visible := False;
    LblBlocnote.Visible := False;
    YDO_BLOCNOTE.Visible := False;

    LblTheme.Visible := True;
    YMO_THEMEMODELE.Visible := True;

    // $$$ JP 05/04/06
    TSRecherche.TabVisible:=FALSE;
    BRemplacer.Visible:=True;
    BVisualiser.Visible:=True;
    end;

  // Toujours un dossier de rattachement
  // if NoDossier='' then NoDossier := '000000';
  // On l'affiche
  if (NoDossier<>'') then
   Dossier.Text := NoDossier + ' ' + GetNomDossier(NoDossier);

  // #### inutile car traité par DossierChange ?
  {$IFDEF EWS}
  bDossierEws := bEwsActif and IsDossierEws(NoDossier);
  {$ENDIF}
  // Affiche l'armoire/classeur/intercalaire correspondant au CodeGed appelant

  chemged := GetCheminGed(CodeGed, False);
  ged1 := ReadTokenSt(chemged);
  ged2 := ReadTokenSt(chemged);
  ged3 := ReadTokenSt(chemged);
  if ged1<>'' then CODEGED1.Value := ged1;
  InitialiserClasseur;
  if ged2<>'' then CODEGED2.Value := ged2;
  InitialiserIntercalaire;
  if ged3<>'' then CODEGED3.Value := ged3;

  TobDoc := Tob.Create('YDOCUMENTS', nil, -1);
  if bModele then TobDocGed := Tob.Create('YMODELES', nil, -1)
             else TobDocGed := Tob.Create('DPDOCUMENT', nil, -1);

  // Nouveau document
  if SDocGUID = '' then
    begin
    TobDocGed.InitValeurs;
    TobDoc.InitValeurs;
    if Annee<>'' then
      YDO_ANNEE.Text := Annee
    else
      YDO_ANNEE.Text := FormatDateTime('yyyy', Date);
    if Mois<>'' then
      YDO_MOIS.Text := Mois
    else
      YDO_MOIS.Text := FormatDateTime('mm', Date);
    YDO_AUTEUR.Text := V_PGI.UserName;
    end
  // Document existant
  else
    begin
    // sélectionne l'enregistrement via la clé primaire
    TobDocGed.SelectDb('"'+SDocGUID+'"', Nil);
    TobDoc.SelectDB('"'+SDocGUID+'"', Nil);
    TobDocGed.PutEcran(Self);
    TobDoc.PutEcran(Self);
    SFileGUID := '';
    if SDocGUID <> '' then
      begin
      Q := OpenSQL('SELECT YDF_FILEGUID FROM YDOCFILES WHERE YDF_DOCGUID = "' + SDocGUID
                  + '" AND YDF_FILEROLE="PAL"', True);
      if Not Q.Eof then SFileGUID := Q.FindField('YDF_FILEGUID').AsString;
      Ferme(Q);
      end;
    if bModele then BVisualiser.Visible := True;
    end;
  ActualiseTitre;

  // Fichier déjà présent dans YFILES
  if SFileGUID <> '' then
   begin
    Fichier.Enabled := False;
    Fichier.ElipsisAutoHide := True;
    Fichier.Text := GetFileNameGed(SFileGUID);
   end;

  // Description par défaut du document
  if (SDocGUID = '') and (YDO_LIBELLEDOC.Text='') then
    begin
    if DefName<>'' then YDO_LIBELLEDOC.Text := DefName
    else YDO_LIBELLEDOC.Text := Fichier.Text;
    end;

  // Commentaire
  if (TypeGed='COM') or (TypeGed='URL') then
  begin
     // Pas de fichier associé au document, donc zones inutiles
    if (TypeGed='COM') then
     begin
      LblFichier.Visible := False;
      Fichier.Visible := False;
      LblProprietes.Caption := 'Propriétés du commentaire';
      LblProprietes.Top := LblFichier.Top ;
     end;

    if (TypeGed='URL') then
     begin
      if sDocGuid <> '' then
       Fichier.Text := TobDoc.GetString ('YDO_URL')
      else
       Fichier.Text := 'http://';

      lblFichier.Caption := 'Url';
      Fichier.ElipsisButton := FALSE;
     end;

    LblEtatDocument.Visible := False;
    TImageDocEtat.Visible := False;
    HDocEtat.Visible := False;
    HLExtraitPar.Visible := False;
    HExtraitPar.Visible := False;

    HLExtraitLe.Visible := False;
    HExtraitLe.Visible := False;
    LblDroitDocument.Visible := False;
    YDO_PRIVE.Visible := False;
    HLDROITS.Visible := False;
    YDO_DROITGED.Visible := False;
    BRemplacer.Visible := False;
  end;

  //--- CAT : 02/01/2007
  if (TypeFichier<>CEXTAUTRE) then
   begin
    // MD FQ11519 - Zone inaccessible
    Fichier.Enabled := False;
    if (SFileGUID='') then
     begin
      Fichier.Text := ''; // au lieu de 'DOCUMENT'
      LblFichier.Visible := False;
      Fichier.Visible := False;
     end
    else
     Fichier.Text:=GetFileNameGed(SFileGUID);
    Fichier.ElipsisButton := FALSE;
   end;

  // Retaillage
  //if Not BDocument.Visible then Fichier.Width := YDO_LIBELLEDOC.Width;

  // $$$ JP 28/09/04 - interdir si demandé la modification hiérarchique GED
  if bNoChangeCodeGed = TRUE then
  begin
       LblDossier.Enabled    := FALSE;
       Dossier.Enabled       := FALSE;
       LblArmoire.Enabled    := FALSE;
       CODEGED1.Enabled      := FALSE;
       LblClasseur.Enabled   := FALSE;
       CODEGED2.Enabled      := FALSE;
       LblIntercal.Enabled   := FALSE;
       CODEGED3.Enabled      := FALSE;
       BEfface1.Enabled      := FALSE;
       BEfface2.Enabled      := FALSE;
       BEfface3.Enabled      := FALSE;
       BEfface5.Enabled      := FALSE;
  end;

  // $$$ JP 05/04/06 - tables libres ged
  strLibLibre := Trim (RechDom ('YYZONELIBREGED', 'ZL1', FALSE));
  if strLibLibre <> '' then
      Ltablelibre1.Caption := strLibLibre
  else
  begin
       LtableLibre1.Visible  := FALSE;
       beffaceLib1.visible   := FALSE;
       YDO_LIBREGED1.Visible := FALSE;
  end;
  strLibLibre := Trim (RechDom ('YYZONELIBREGED', 'ZL2', FALSE));
  if strLibLibre <> '' then
      Ltablelibre2.Caption := strLibLibre
  else
  begin
       LtableLibre2.Visible  := FALSE;
       beffaceLib2.visible   := FALSE;
       YDO_LIBREGED2.Visible := FALSE;
  end;
  strLibLibre := Trim (RechDom ('YYZONELIBREGED', 'ZL3', FALSE));
  if strLibLibre <> '' then
      Ltablelibre3.Caption := strLibLibre
  else
  begin
       LtableLibre3.Visible  := FALSE;
       beffaceLib3.visible   := FALSE;
       YDO_LIBREGED3.Visible := FALSE;
  end;
  strLibLibre := Trim (RechDom ('YYZONELIBREGED', 'ZL4', FALSE));
  if strLibLibre <> '' then
      Ltablelibre4.Caption := strLibLibre
  else
  begin
       LtableLibre4.Visible  := FALSE;
       beffaceLib4.visible   := FALSE;
       YDO_LIBREGED4.Visible := FALSE;
  end;
  strLibLibre := Trim (RechDom ('YYZONELIBREGED', 'ZL5', FALSE));
  if strLibLibre <> '' then
      Ltablelibre5.Caption := strLibLibre
  else
  begin
       LtableLibre5.Visible  := FALSE;
       beffaceLib5.visible   := FALSE;
       YDO_LIBREGED5.Visible := FALSE;
  end;

  //--- CAT le  13/11/2006
  if (BModele) then
   DPD_ARBOGED.Visible:=False
  else
   begin
    if (SDocGuid='') then
     begin
      CodeGed1.Plus:=SArboGed;
      if (SArboGed='D') then
       DPD_ARBOGED.itemIndex:=0
      else
       DPD_ARBOGED.itemIndex:=1;
     end
    else
     begin
      CodeGed1.Plus:=TobDocGed.GetValue ('DPD_ARBOGED');
      if (TobDocGed.GetValue ('DPD_ARBOGED')='D') then
       DPD_ARBOGED.itemIndex:=0
      else
       DPD_ARBOGED.itemIndex:=1;
     end;
   end;

  //--- CAT le 15/11/2006 : Initialisation de l'onglet Information
  YDO_DROITGED.Text:=RechDom ('YYDROITGED',TobDoc.GetValue ('YDO_DROITGED'),False); // FQ 11828
  HIntegrePar.Caption:=TobDoc.GetValue ('YDO_CREATEUR');
  HIntegreLe.Caption:=DateToStr (TobDoc.GetValue ('YDO_DATECREATION'));

  if (TobDoc.GetValue ('YDO_DOCSTATE')<>'') then // Fichier en cours de modification
   begin
    HLExtraitPar.Visible:=True;
    HExtraitPar.Visible:=True;
    HLExtraitLe.Visible:=True;
    HExtraitLe.Visible:=True;
    TImageDocEtat.Canvas.FillRect(TImageDocEtat.Picture.BitMap.Canvas.ClipRect);
    FIconFiles.LargeImageList.GetBitMap (cImgVerrouille, TImageDocEtat.Picture.BitMap);
    HDocEtat.Caption:='En cours de modification';
    HExtraitPar.Caption:=TobDoc.GetValue ('YDO_UTILISATEUR');
    HExtraitLe.Caption:=TobDoc.getValue ('YDO_DATEMODIF');
   end
  else
   if (TobDoc.GetValue ('YDO_DATEFIN')<Now) and (SDocGuid<>'') then // Fichier périmé
    begin
     TImageDocEtat.Canvas.FillRect(TImageDocEtat.Picture.BitMap.Canvas.ClipRect);
     FIconFiles.LargeImageList.GetBitMap (cImgPerime, TImageDocEtat.Picture.BitMap);
     HDocEtat.Caption:='Périmé';
    end
   else
    begin
     HLExtraitPar.Visible:=False;
     HExtraitPar.Visible:=False;
     HLExtraitLe.Visible:=False;
     HExtraitLe.Visible:=False;
     TImageDocEtat.Canvas.FillRect(TImageDocEtat.Picture.BitMap.Canvas.ClipRect);
     FIconFiles.LargeImageList.GetBitMap (CImgActif, TImageDocEtat.Picture.BitMap);
     HDocEtat.Caption:='Actif';
    end;

  //--- CAT le 28/11/2006
  UnFichierRemplacer:='';
  // Bouton "Remplacer" visible uniquement si droits ouverts
  if (TobDoc.GetValue ('YDO_DOCSTATE')<>'EXT') and (TobDoc.GetValue ('YDO_DROITGED')='')
  // et qu'on est sur un document existant et que c'est pas un commentaire
  and (SDocGuid<>'') and (TypeGed<>'COM') and (TypeGed<>'URL') then
   BRemplacer.Visible := True;

  //--- CAT le 09/11/2006
  if (TobDoc.GetValue ('YDO_DOCSTATE')='EXT') or ((TobDoc.GetValue ('YDO_DROITGED')='RO') and (TobDoc.GetValue ('YDO_CREATEUR')<>V_PGI.USER)) then
   begin
    Self.Caption:=Self.Caption + ' (Lecture seule)';

    //Désactive les controls de la fiche
    for i1 := 0 to Self.ControlCount-1 do
     begin
      C1 := Self.Controls[i1];
      if (C1 is TControl) and not (C1 is TImage) then
       begin
        if ( (C1 is TCustomGroupBox) or (C1 is TDock97) or (C1 is TCustomPanel) or (C1 is TToolWindow97) or (C1 is TPageControl) or (C1 is TTabSheet) ) and (C1.Visible = True) then
         begin
          for i2 := 0 to TWinControl(C1).ControlCount - 1 do
           begin
            C2 := TWinControl(C1).Controls[i2];
            C2.Enabled := False;
           end;
         end
        else
          C1.Enabled := False;
       end;
     end;
    Dock971.Enabled  := True;
    PBouton.Enabled  := True;
    BValider.Enabled := False;
    BRemplacer.Enabled := False;
    BVisualiser.Enabled := False;
    SourisNormale;
   end
  else
   begin
    if (TobDoc.GetValue ('YDO_CREATEUR')<>V_PGI.USER) then
     begin
      YDO_PRIVE.Enabled:=False;
      HLDROITS.Enabled:=False;
      YDO_DROITGED.Enabled:=False;
     end;
   end;

  // Chargement terminé
  bDuringLoad := False;
end;


procedure TFNewDocument.FormCreate(Sender: TObject);
begin
  inherited;
  bDuringLoad  := True;
  bChange      := False;
  bDocChange_f := False;
  bModifie     := False;
  bEwsActif    := GetParamsocDpSecur('SO_NE_EWSACTIF', False);
  bDossierEws  := False;
end;


procedure TFNewDocument.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  inherited;
  if TobDocGed<>Nil then TobDocGed.Free;
  if TobDoc<>Nil then TobDoc.Free;
end;


procedure TFNewDocument.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
var rep: Integer;
begin
  inherited;
  // Changement demandé
  if bChange then
    begin
    if bDocChange_f then
      rep := PGIAskCancel('Le document lié a changé. Voulez-vous enregistrer les modifications ?', Self.Caption)
    else
      rep := PGIAskCancel('Voulez-vous enregistrer les modifications ?', Self.Caption);
    if rep=mrCancel then
      CanClose := False
    else if rep=mrNo then
      begin
      bChange := False;
      bDocChange_f := false;
      end
    else if rep=mrYes then
      CanClose := SauveEnregDocument;
    end;
end;

function TFNewDocument.DonnerNomFichier : String;
begin
 Result := Fichier.Text;
 if Result = '' then Result := YDO_LIBELLEDOC.Text;
 // Elimine tout ce qui suit le premier "." trouvé
 if (Result<>'') then
  if (Pos ('.',Result)>0) then
   Result := copy (Result,1,pos ('.',Result)-1);
 // Fournit une extension valide
 if (TypeFichier=CEXTDOC) then Result:=Result+'.doc';
 if (TypeFichier=CEXTXLS) then Result:=Result+'.xls';
end;

procedure TFNewDocument.BValiderClick(Sender: TObject);
var UnGuid     : String;
begin
 // Par défaut, le bouton BValider referme la fiche
 if Not SauveEnregDocument then
  ModalResult := mrNone // donc on l'en empêche si SauveEnregDocument n'a pas pu valider
 else
  begin
   if (UnFichierRemplacer<>'') then
    begin
     //--- Remplacement du fichier dans la base
     UnGuid:=V_GedFiles.Import (UnFichierRemplacer);
     if UnGuid <> '' then
      begin
       ExecuteSql ('UPDATE YDOCFILES SET YDF_FILEGUID="'+UnGuid+'" WHERE YDF_DOCGUID="'+TobDoc.GetValue ('YDO_DOCGUID')+'"');
       V_GedFiles.Erase (SFileGuid);
       ExecuteSql ('UPDATE YDOCUMENTS SET YDO_DOCSTATE="", YDO_DATEMODIF="'+UsDateTime (Now)+'", YDO_UTILISATEUR="'+V_PGI.USER+'" WHERE YDO_DOCGUID="'+TobDoc.GetValue ('YDO_DOCGUID')+'"');
      end;
     UnFichierRemplacer:='';
    end;
   inherited;
  end;
end;

procedure TFNewDocument.OnDocChange(Sender: TObject);
begin
  if not bDuringLoad then bChange := True;
  //if (TControl(Sender).Name='YDO_DATEDEB') or (TControl(Sender).Name='YDO_DATEFIN') then ControleDatesValidite;
end;


procedure TFNewDocument.YDO_LIBELLEDOCExit(Sender: TObject);
begin
  // Descriptif "minimum" du document
  if (YDO_LIBELLEDOC.Text='') and (Fichier.Text<>'') then
    if TypeGed='URL' then
      begin
      if Copy(Fichier.Text, 1, 7) = 'http://' then
        YDO_LIBELLEDOC.Text := Copy(Fichier.Text, 8, Length(Fichier.Text)-7 )
      else
        YDO_LIBELLEDOC.Text := Fichier.Text;
      end
    else
      YDO_LIBELLEDOC.Text := ChangeFileExt(ExtractFileName(Fichier.Text), '');
  ActualiseTitre;
end;


procedure TFNewDocument.YDO_MOTSCLESExit(Sender: TObject);
begin
  YDO_MOTSCLES.Text := UpperCase(YDO_MOTSCLES.Text);
end;


procedure TFNewDocument.CODEGED1Change(Sender: TObject);
begin
 OnDocChange(Sender);

 //--- le classement choisi a changé...
 if Not bDuringLoad then
  begin
   if (CodeGed1.Value<>'') then
    CodeGed := CODEGED1.Value
   else
    begin
     if (DPD_ARBOGED.ItemIndex=0) then CodeGed:='D';
     if (DPD_ARBOGED.ItemIndex=1) then CodeGed:='A';
    end;

   InitialiserClasseur;
   InitialiserIntercalaire;
  end;

 ActualiseBrancheEws;
end;


procedure TFNewDocument.CODEGED2Change(Sender: TObject);
begin
  OnDocChange(Sender);
  // le classement choisi a changé...
  if Not bDuringLoad then
   begin
    if (CodeGed2.Value<>'') then
     CodeGed:=CODEGED2.Value
    else
     if (CodeGed1.Value<>'') then
      CodeGed:=CODEGED1.Value
     else
      begin
       if (DPD_ARBOGED.ItemIndex=0) then CodeGed:='D';
       if (DPD_ARBOGED.ItemIndex=1) then CodeGed:='A';
      end;
    InitialiserIntercalaire;
   end;

  ActualiseBrancheEws;
end;

procedure TFNewDocument.CodeGed3Change(Sender: TObject);
begin
 OnDocChange(Sender);
 // le classement choisi a changé...

 if Not bDuringLoad then
  begin
   if (CodeGed3.Value<>'') then
    CodeGed:=CODEGED3.Value
   else
    if (CodeGed2.Value<>'') then
     CodeGed:=CODEGED2.Value
    else
     if (CodeGed1.Value<>'') then
      CodeGed:=CODEGED1.Value
     else
      begin
       if (DPD_ARBOGED.ItemIndex=0) then CodeGed:='D';
       if (DPD_ARBOGED.ItemIndex=1) then CodeGed:='A';
      end;
   end;

 ActualiseBrancheEws;
end;

//------------------------------------------------------------------------------
//--- Nom : InitialiserClasseur
//--- Objet Met à jour la liste des classeurs dépendants de l'armoire choisie
//--- car on ne peut pas profiter de l'automatisme des tablettes hiérarchiques
//--- table YDATATYPETREES non commune, donc hiérarchie non visible depuis
//--- une appli connectée à une base dossier)
//------------------------------------------------------------------------------
procedure TFNewDocument.InitialiserClasseur;
var SQL : String;
    Q : TQuery;
begin
 SQL := 'SELECT YGD_CODEGED, YGD_LIBELLEGED FROM YGEDDICO WHERE'
      + ' EXISTS (SELECT YDT_CODEHDTLINK FROM ##DP##.YDATATYPETREES'
      + ' WHERE (YDT_CODEHDTLINK = "YYGEDNIV1GEDNIV2") AND (YDT_MCODE = "'+CODEGED1.Value+'")'
      + '  AND (##DP##.YDATATYPETREES.YDT_SCODE = YGEDDICO.YGD_CODEGED) )'
      + ' AND YGD_TYPEGED="CLS" ORDER BY YGD_TRIGED';

 CODEGED2.OnChange := Nil;
 CODEGED2.Value := '';
 CODEGED2.Values.Clear;
 CODEGED2.Items.Clear;

 Q := OpenSQL(SQL, True, -1, '', True);
 while (Not Q.Eof) do
  begin
   CODEGED2.Values.Add(Q.FindField('YGD_CODEGED').AsString);
   CODEGED2.Items.Add(Q.FindField('YGD_LIBELLEGED').AsString);
   Q.Next;
  end;
 Ferme(Q);

 CODEGED2.OnChange := CODEGED2Change;
end;

//------------------------------------------------------------------------------
//--- Nom : InitialiserIntercalaire
//------------------------------------------------------------------------------
procedure TFNewDocument.InitialiserIntercalaire;
var SQL : String;
    Q : TQuery;
begin
 SQL := 'SELECT YGD_CODEGED, YGD_LIBELLEGED FROM YGEDDICO WHERE'
      + ' EXISTS (SELECT YDT_CODEHDTLINK FROM ##DP##.YDATATYPETREES'
      + ' WHERE (YDT_CODEHDTLINK = "YYGEDNIV2GEDNIV3") AND (YDT_MCODE = "'+CODEGED2.Value+'")'
      + '  AND (##DP##.YDATATYPETREES.YDT_SCODE = YGEDDICO.YGD_CODEGED) )'
      + ' AND YGD_TYPEGED="ITL" ORDER BY YGD_TRIGED';

 CODEGED3.OnChange := Nil;
 CODEGED3.Value := '';
 CODEGED3.Values.Clear;
 CODEGED3.Items.Clear;

 Q := OpenSQL(SQL, True, -1, '', True);
 while (Not Q.Eof) do
  begin
   CODEGED3.Values.Add(Q.FindField('YGD_CODEGED').AsString);
   CODEGED3.Items.Add(Q.FindField('YGD_LIBELLEGED').AsString);
   Q.Next;
  end;
 Ferme(Q);

 CODEGED3.OnChange := CODEGED3Change;
end;

procedure TFNewDocument.BEfface1Click(Sender: TObject);
begin
  inherited;
  CODEGED1.Value := '';
 if (DPD_ARBOGED.ItemIndex=0) then CodeGed:='D'
 else if (DPD_ARBOGED.ItemIndex=1) then CodeGed:='A';

 // déclenche les onchange automatiquement
end;

procedure TFNewDocument.BEfface2Click(Sender: TObject);
begin
  inherited;
  CODEGED2.Value := '';
end;

procedure TFNewDocument.BEfface3Click(Sender: TObject);
begin
 inherited;
 Dossier.Text :='';
 NoDossier:='';
end;

procedure TFNewDocument.BEfface4Click(Sender: TObject);
begin
 inherited;
 YDO_NATDOC.Value := '';
end;

procedure TFNewDocument.BEfface5Click(Sender: TObject);
begin
 inherited;
 CODEGED3.Value := '';
end;

procedure TFNewDocument.BDocumentClick(Sender: TObject);
var
   sDoc_l : string;
begin
   // Extrait le document
   sDoc_l := GetParamSocDpSecur('SO_MDREPDOCUMENTS', TCbpPath.GetCegidDistriStd+'\BUREAU') + '\' + Fichier.text;
   V_GedFiles.Extract(sDoc_l, SFileGUID);
   if LanceAppliMaquette( sDoc_l, ExtractFileExt(sDoc_l) ) then
   begin
      // on sait pas ce qui a pu changer => suppression de l'ancien modèle
      SDocGUIDASupprimer := SDocGUID;

      // se considère comme en modification
      bChange := True;
      bDocChange_f := True;
   end;
end;


procedure TFNewDocument.DossierElipsisClick(Sender: TObject);
var St, codper : String;
begin
  // retourne NoDossier;GuidPer;Nom1
  St := AGLLanceFiche('YY','YYDOSSIER_SEL', '','',NoDossier);
  if St<>'' then
    begin
    NoDossier := READTOKENST(St);
    codper := READTOKENST(St);
    Dossier.Text := NoDossier+' '+GetNomDossier(NoDossier);
    // ou NoDossier+' '+READTOKENST(St);
    end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 23/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFNewDocument.FichierElipsisClick(Sender: TObject);
var
   ODModele_l : TOpenDialog;
begin
   //Open document a fusionner
   ODModele_l := TOpenDialog.Create(Application);
   try
   begin
      ODModele_l.Title := 'Sélectionner un fichier';
      ODModele_l.Filter := 'Tous les fichiers (*.*)|*.*';
      ODModele_l.InitialDir := GetParamSocDpSecur('SO_MDREPDOCUMENTS', TCbpPath.GetCegidDistriStd+'\BUREAU');
      if ODModele_l.execute then
         Fichier.Text := ODModele_l.FileName;
   end;
   finally
      ODModele_l.Free;
   end;
end;


procedure TFNewDocument.ActualiseTitre;
var titre: String;
begin
  titre := '';
  // Type d'enregistrement
  if TypeGed='COM' then
    titre := 'Commentaire'
  else if TypeGed='URL' then
    titre := 'URL'
  else if bModele then
    titre := 'Modèle'
  else
    titre := 'Document';

  // Nouvel enregistrement
  if SDocGUID = '' then
     if TypeGed <> 'URL' then
         titre := 'Nouveau ' + LowerCase(titre)
     else
         titre := 'Nouvelle url';

  // Libellé détaillé
  if YDO_LIBELLEDOC.Text<>'' then
    titre := titre + ' - ' + YDO_LIBELLEDOC.Text;
  Self.Caption := titre;
  UpdateCaption(Self);
end;


function TFNewDocument.ControleDatesValidite : Boolean;
begin
  Result := False;
  If Not IsValidDate(YDO_DATEDEB.Text) then exit;
  if Not IsValidDate(YDO_DATEFIN.Text) then exit;

  // Contrôle de cohérence sur des dates correctes
  if StrToDate(YDO_DATEDEB.Text) > StrToDate(YDO_DATEFIN.Text) then
    PGIInfo('La date de fin de validité ne peut être antérieure à la date de début de validité.')
  else
    Result := True;
end;


function TFNewDocument.SauveEnregDocument: Boolean;
// Sauvegarde d'un nouveau document
// True si on a passé les tests de validité
var
    NomFichier : String;
    bNewDocu, bError, bDocFiles : Boolean;
    TobDocFiles   : Tob;
    EwsId, Retour : String;
    EwsRegle      :string; // $$$ JP 11/04/06
    iNbError      : Integer;
{$IFDEF EWS}
    bEwsAssist    :boolean;
{$ENDIF}
begin
  Result := False;
  bNewDocu := False;

  // N'arrive jamais, le 000000 est par défaut
{  if NoDossier='' then
    begin
    PGIInfo('Pas de dossier choisi : impossible d''enregistrer.', Self.Caption);
    exit;
    end;}

  // Contrôles généraux
  if YDO_LIBELLEDOC.Text='' then
    begin
    PGIInfo('Entrez le libellé descriptif du document.', Self.Caption);
    PNewDoc.ActivePage := TSCaracteristiques;
    YDO_LIBELLEDOC.SetFocus;
    exit;
    end;

  if (YDO_ANNEE.Text<>'') and (Not IsNumeric(YDO_ANNEE.Text)) then
   begin
    PGIInfo('Zone numérique uniquement.', Self.Caption);
    PNewDoc.ActivePage := TSCaracteristiques;
    YDO_ANNEE.SetFocus;
    exit;
   end
  else
   if (StrToInt (YDO_ANNEE.Text)<1900) or (StrToInt (YDO_ANNEE.Text)>2099) then
    begin
     PGIInfo('Les valeurs de l''année doit être comprise entre 1900 et 2099.', Self.Caption);
     PNewDoc.ActivePage := TSCaracteristiques;
     YDO_ANNEE.SetFocus;
     exit;
    end;

  if (YDO_MOIS.Text<>'') then
   begin
    if (Not IsNumeric(YDO_MOIS.Text)) then
     begin
      PGIInfo('Zone numérique uniquement.', Self.Caption);
      PNewDoc.ActivePage := TSCaracteristiques;
      YDO_MOIS.SetFocus;
      exit;
     end
    else if (StrToInt (YDO_MOIS.Text)<1) or (StrToInt (YDO_MOIS.Text)>12) then
     begin
      PGIInfo('La valeur du mois doit être comprise entre 1 et 12.', Self.Caption);
      PNewDoc.ActivePage := TSCaracteristiques;
      YDO_MOIS.SetFocus;
      exit;
     end
    else if (Length (YDO_MOIS.Text)<2) then
     YDO_MOIS.Text:='0'+YDO_MOIS.Text;
   end;

  if Not ControleDatesValidite then exit;

  // En cas de modification du document WORD, ...
  if SDocGUIDASupprimer <> '' then
    begin
    // Dès qu'on est en phase de validation, on ne modifiera plus le document
    BVisualiser.Enabled := False;
    // Purge du modèle
    ExecuteSQL('DELETE FROM YMODELES WHERE YMO_DOCGUID= "' + SDocGUIDASupprimer+'"');
    SupprimeDocument(SDocGUIDASupprimer);
    // On est maintenant sur un nouveau document
    SDocGUIDASupprimer := '';
    SDocGUID := '';
    // le fichier à insérer est comme un nouveau fichier
    SFileGUID := '';
    Fichier.text := GetParamSocDpSecur('SO_MDREPDOCUMENTS', TCbpPath.GetCegidDistriStd+'\BUREAU') + '\' + Fichier.text;
    // sinon ne validera pas la création si aucune modif autre que le document
    // et si la clé calculée retombe sur le même n° de clé (dernier docid)
    TobDoc.SetAllModifie(True);
    TobDocGed.SetAllModifie(True);
    end;

  if (SDocGUID='') and (SFileGUID='') and (TypeGed<>'COM') then
  // Contrôles avant validation d'un nouvel enreg
  BEGIN
    // Tout type de fichier
    if (TypeFichier=CEXTAUTRE) then
     begin
      if Fichier.Text='' then
       begin
        if TypeGed = 'URL' then
         PgiInfo ('Aucune url à insérer', Self.Caption)
        else
         PGIInfo('Aucun fichier à insérer.', Self.Caption);
        PNewDoc.ActivePage := TSCaracteristiques;
        Fichier.SetFocus;
        exit;
       end
      else if (TypeGed <> 'URL') and Not FileExists(Fichier.Text) then
       begin
        PGIInfo('Le fichier '+Fichier.Text+' n''existe plus.', Self.Caption);
        PNewDoc.ActivePage := TSCaracteristiques;
        Fichier.SetFocus;
        exit;
       end;
     end

    //--- Document office
    else if (TypeFichier<>CEXTAUTRE) then
     begin
      NomFichier:=DonnerNomFichier;
      ForceDirectories(V_GedFiles.TempPath+'BUREAU\GED'); // FQ 11521

      if FileExists (V_GedFiles.TempPath+'BUREAU\GED\'+NomFichier) then
       begin
        if (PGIAsk ('Un fichier '+NomFichier+' existe déjà, voulez-vous l''écraser.',TitreHalley)=mrYes) then
         DeleteFile (PChar (V_GedFiles.TempPath+'BUREAU\GED\'+NomFichier))
        else
         exit;
       end;

      if (TypeFichier=CEXTDOC) then
       CreerDocumentDOC (V_GedFiles.TempPath+'BUREAU\GED\', NomFichier)
      else if (TypeFichier=CEXTXLS) then
       CreerDocumentXLS (V_GedFiles.TempPath+'BUREAU\GED\', NomFichier);
     end;
  END;

  // Nouvel enreg
  if SDocGUID = '' then
   begin
    // Document non encore référencé
    if (SFileGUID = '') and (TypeGed<>'COM') and (TypeGed<>'URL') then
     begin
      // Enregistre le fichier dans YFILES, YFILEPARTS...
      if (TypeFichier<>CEXTAUTRE) then
       SFileGUID := V_GedFiles.Import(V_GedFiles.TempPath+'BUREAU\GED\'+NomFichier)
      else
       SFileGUID := V_GedFiles.Import(Fichier.Text);

      if SFileGUID = '' then
       begin
        PGIInfo('Le fichier '+Fichier.Text+' ne peut pas être importé dans la GED.', TitreHalley);
        exit;
       end;

      if FileExists (V_GedFiles.TempPath+'BUREAU\GED\'+NomFichier) then DeleteFile (PChar (V_GedFiles.TempPath+'BUREAU\GED\'+NomFichier));
     end;
   end;

  // Récup des modifs
  TobDocGed.GetEcran(Self);
  TobDoc.GetEcran(Self);

    // Clé 2 : on remet à jour car a pu changer, même sur un enreg existant
    // #### pb, doit-on mettre à jour la référence documentaire YDO_DOCREF ?
    if Not bModele then
     begin
      TobDocGed.PutValue('DPD_NODOSSIER', NoDossier);
      TobDocGed.PutValue('DPD_CODEGED', CodeGed);
     end;

    // $$$ JP 19/12/06: doctype=DOC,COM ou URL; URL doit contenir l'url définie dans Fichier
    TobDoc.PutValue ('YDO_DOCTYPE', TypeGed);
    if TypeGed = 'URL' then
     TobDoc.PutValue ('YDO_URL', Trim (Fichier.Text));

    //--- Nouvel enregistrement
    if SDocGUID='' then
     begin
      // #### A REVOIR en transactionnel sur le modèle AfActivite :
      //    Result := Transactions (ValideLaLigneActivite, 5); et tests sur V_PGI.IoError
      // #### EN ATTENDANT correction rapide pour P.Lenormand chez IBM
      bError := True;

      // S'agit-il d'un document avec fichier
      bDocFiles := (SFileGUID<>'') and (TypeGed<>'COM') and (TypeGed<>'URL');

      //$$$ JP 06/12/05 - warning delphi
      if bDocFiles then TobDocFiles := Tob.Create('YDOCFILES', Nil, -1)
      else TobDocFiles := nil;

      for iNbError:=0 to 4 do
       begin
        if Not bError then break;

        // Clé 1 : clé unique
        SDocGUID := AGLGetGUID ();

        if bModele then TobDocGed.PutValue('YMO_DOCGUID', SDocGUID )
        else TobDocGed.PutValue('DPD_DOCGUID', SDocGUID);
        TobDoc.PutValue('YDO_DOCGUID',SDocGUID);

        // MD 31/05/06 : Ancienne clé mise à -2
        // pour éviter pb de maj des clés étrangères (comme dans ECRCOMPL)
        // qui prennent les enreg ydo_docid à 0 pour des enreg ged valides
        // lors de la conversion des integer en guid...
        TobDoc.PutValue('YDO_DOCID', -2);
        TobDocGed.PutValue('DPD_DOCID', -2);
        TobDoc.PutValue('YDO_DOCREF', NoDossier + ' ' + CodeGed + ' '+ SFileGUID + ' ' + SDocGUID );

        //--- Transaction
        Try
         BeginTrans;
         //--- YDOCFILES
         if bDocFiles then
          begin
           // #### pour l'instant, insère un seul fichier pour le document,
           // #### en lui attribuant le rôle "PRINCIPAL"
           TobDocFiles.InitValeurs;

           TobDocFiles.PutValue('YDF_DOCGUID', SDocGUID);
           TobDocFiles.PutValue('YDF_FILEGUID', SFileGUID);
           // MD 31/05/06
           TobDocFiles.PutValue('YDF_DOCID', -2);

           TobDocFiles.PutValue('YDF_FILEROLE', 'PAL');
           if Not TobDocFiles.InsertDB(Nil) then begin Rollback; Continue; end;
          end;

         //--- YDOCUMENTS
         if Not TobDoc.InsertDB(Nil) then begin Rollback; Continue; end;

         //--- DPDOCUMENT ou YMODELES
         if Not TobDocGed.InsertDB(Nil) then begin Rollback; Continue; end;

         CommitTrans;
         bError := False;
         // Code retour pour la fonction appelante
         bModifie := True;
         // On a inséré un vrai nouveau document inséré, avec fichier associé
         if bDocFiles then bNewDocu := True;

        except
         // L'erreur la plus fréquente sera un duplicate key row
         // d'où la boucle sur iNbError pour tenter une nouvelle clé
         RollBack;
         // PGIInfo('Erreur lors de la sauvegarde des modifications.', Self.Caption);
        end;
       end;

     // fin boucle sur iNbError
     if bDocFiles then TobDocFiles.Free;
    end

  //--- Modification d'un enregistrement
  else
    Try
      // Code retour pour la fonction appelante
      bModifie := True;
      TobDoc.InsertOrUpdateDB;
      TobDocGed.InsertOrUpdateDB;
    except
      On E:Exception do PGIInfo(E.Message, 'Erreur de sauvegarde des modifications.');
    end;



  // Si branche Ged abonnée à Ews, et client aussi, alors publication vers portail
{$IFDEF EWS}
  if bNewDocu and GetParamsocDpSecur('SO_NE_EWSACTIF', False) and IsDossierEws(Nodossier) then
  begin
    EwsId := CodeGedToEwsId(CodeGed);
    EwsRegle := CodeGedToEwsRegle (CodeGed); // $$$ JP 11/04/06
    if (EwsId<>'') or (EwsRegle <> '') then
    begin
         // $$$ JP 21/07/06: il faut publier avec l'ihm si nécessaire, mais si pas publier, on tente quand même de le publier automatiquement (sans ihm)
         // $$$ JP 24/07/06: non, il faut publier dans le mode choisi, c'est tout (réponse de TD du 22/07/06
         Retour     := '';
         bEwsAssist := GetParamsocDPSecur ('SO_NEMODEPUBLICATION', 'AUT') = 'ASS';
         EwsPublieUnDocument (SDocGUID, EwsID, EwsRegle, Retour, TRUE, bEwsAssist);
         if Retour <> '' then
            PGIInfo ('Erreur lors de la publication sur eWS:'#13#10' ' + Retour);
    end;
  end;
{$ENDIF EWS}

  SDocGUIDASupprimer := '';
  bChange := False;
  bDocChange_f := False;
  Result := True;
end;


procedure TFNewDocument.YDO_DATEDEBExit(Sender: TObject);
begin
  inherited;
//  If Not ControleDatesValidite then
//    begin
//    PNewDoc.ActivePage := TSCaracteristiques;
//    YDO_DATEDEB.SetFocus;
//    end;
end;

procedure TFNewDocument.YDO_DATEFINExit(Sender: TObject);
begin
  inherited;
//  If Not ControleDatesValidite then
//    begin
//    PNewDoc.ActivePage := TSCaracteristiques;
//    YDO_DATEFIN.SetFocus;
//    end;
end;

procedure TFNewDocument.ActualiseBrancheEws;
var
   EwsId    :string;
   EwsRegle :string;
begin
  if bDossierEws then
  begin
       EwsId    := CodeGedToEwsId (CodeGed);
       EwsRegle := CodeGedToEwsRegle (CodeGed);
  end
  else
  begin
       EwsId    := '';
       EwsRegle := '';
  end;

  if (EwsId<>'') or (EwsRegle<>'') then
  begin
       ImgEws.Visible        := True;
       LblBrancheEws.Visible := True;
{$IFDEF EWS}
       if EwsId <> '' then
           LblBrancheEws.Caption := 'Répertoire associé: ' + EwsRetourneLibelleNoeud(EwsId)
       else
           LblBrancheEws.Caption := 'Règle associée: ' + EwsRegle;
{$ENDIF}
  end
  else
  begin
    ImgEws.Visible        := False;
    LblBrancheEws.Visible := False;
    LblBrancheEws.Caption := '...';
  end;
end;

procedure TFNewDocument.DossierChange(Sender: TObject);
begin
  inherited;
  OnDocChange(Sender);
{$IFDEF EWS}
  bDossierEws := GetParamsocDpSecur('SO_NE_EWSACTIF', False) and IsDossierEws(NoDossier);
{$ENDIF}
  ActualiseBrancheEws;
end;


procedure TFNewDocument.BEffaceLib1Click(Sender: TObject);
begin
  inherited;
  YDO_LIBREGED1.Value := '';
end;

procedure TFNewDocument.BEffaceLib2Click(Sender: TObject);
begin
  inherited;
  YDO_LIBREGED2.Value := '';
end;

procedure TFNewDocument.BEffaceLib3Click(Sender: TObject);
begin
  inherited;
  YDO_LIBREGED3.Value := '';
end;

procedure TFNewDocument.BEffaceLib4Click(Sender: TObject);
begin
  inherited;
  YDO_LIBREGED4.Value := '';
end;

procedure TFNewDocument.BEffaceLib5Click(Sender: TObject);
begin
  inherited;
  YDO_LIBREGED5.Value := '';
end;

procedure TFNewDocument.DPD_ARBOGEDClick(Sender: TObject);
begin
 inherited;
 if not bDuringLoad then
  begin
   if (DPD_ARBOGED.ItemIndex=0) then
    begin
     CodeGed:='D';
     TobDocGed.PutValue ('DPD_ARBOGED','D')
    end
   else
    if (DPD_ARBOGED.ItemIndex=1) then
     begin
      CodeGed:='A';
      TobDocGed.PutValue ('DPD_ARBOGED','A');
     end;

   CodeGed1.Plus:=TobDocGed.GetValue ('DPD_ARBOGED');
   CodeGed2.Value := '';
   CodeGed2.Values.Clear;
   CodeGed2.Items.Clear;
   CodeGed3.Value := '';
   CodeGed3.Values.Clear;
   CodeGed3.Items.Clear;
  end;
end;

procedure TFNewDocument.BRemplacerClick(Sender: TObject);
var ODRemplacer    : TOpenDialog;
begin
 inherited;
 UnFichierRemplacer:='';
 if (PGIAsk ('Vous allez choisir un document qui viendra remplacer "'+Fichier.Text+'". Confirmez vous ?')=mrYes) then
  begin
   //--- Sélection du fichier de remplacement
   ODRemplacer := TOpenDialog.Create(Application);
   try
    begin
     ODRemplacer.Title      := 'Sélectionner un fichier';
     ODRemplacer.Filter     := 'Tous les fichiers (*.*)|*.*';
     ODRemplacer.InitialDir := GetParamSocDpSecur('SO_MDREPDOCUMENTS', TCbpPath.GetCegidDistriStd+'\BUREAU');
     if ODRemplacer.execute then
      begin
       UnFichierRemplacer := ODRemplacer.FileName;
       Fichier.Text:=ExtractFileName (UnFichierRemplacer);
      end;
    end;
   finally
    ODRemplacer.Free;
   end;
  end;
end;

end.
