unit UtomLiensOle;

interface

uses UTOM,
  Classes,
  SysUtils,
  HCtrls,
  HEnt1,
  Forms,
  {$IFDEF EAGLCLIENT}
  UTob,
  eFichList,
  MaineAGL,
  {$ELSE}
  DB,
  {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
  dbctrls,
  {$IFNDEF EAGLSERVER}
  {$IFNDEF ERADIO}
  FichList,
  Fe_Main,
  {$ENDIF !ERADIO}
  {$ENDIF !EAGLSERVER}
  {$ENDIF}
  {$IFNDEF EAGLSERVER}
  lookup,
  {$ENDIF !EAGLSERVER}
  M3FP,
  HMsgBox,
  UtilPGI,
  extctrls,
  Jpeg,
  Comctrls,
  EntGC;

type
  TKindOfMemo = (komNone, komMemo, komPhotoJpeg, komVignetteJpeg, komPhotoBmp, komData);

  TOM_LiensOle = class(TOM)
  private
    filtreQualif: boolean;
    filtreEmploi: boolean;
    rangeQualif: string;
    rangeEmploi: string;
    ObjDefaut: string; // JTR - Objet par défaut
      // JTR - Image libre stockées
    SelectByDblclick: boolean;
    CanChangeObj: boolean;
     // Fin JTR - Image libre stockées
    {$IFNDEF EAGLCLIENT}
    QualifiantBlob: ThdbvalCombobox;
    {$ELSE}
    QualifiantBlob: ThvalCombobox;
    {$ENDIF}
    {$IFDEF EAGLCLIENT}
    procedure OnExitMemo(Sender: TObject);
    {$ENDIF}
    {$IFNDEF EAGLSERVER}
    {$IFNDEF ERADIO}
    procedure FListeOnDblClick(Sender: TObject); // JTR - Image libre stockées
    {$ENDIF !ERADIO}
    {$ENDIF !EAGLSERVER}
    procedure QualifiantBlobOnChange(Sender: TObject);
  public
    procedure OnUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnNewRecord; override;
    procedure OnArgument(Arguments: string); override;
    {$IFDEF EAGLCLIENT}
    procedure chargeStream(s: TStringStream);
    {$ELSE}
    procedure chargeStream(s: TMemoryStream);
    {$ENDIF}

    function GetIDCourant: string;
    function GetKindOfMemo: TKindOfMemo;
    {$IFNDEF EAGLSERVER}
    procedure SetControlsState;
    {$ENDIF !EAGLSERVER}
  end;

implementation
{$IFDEF EAGLCLIENT}
procedure TOM_LiensOle.OnExitMemo(Sender: TObject);
begin
  {$IFDEF EAGLCLIENT}
  if getField('LO_OBJET') <> getControlText('LEMEMO') then
  begin
    if DS.State = dsBrowse then DS.Edit;
    SetField('LO_OBJET', getControlText('LEMEMO'));
  end;
  {$ELSE}
  TFFicheListe(ecran).ta.currentfille.modifie := true;
  {$ENDIF EAGLCLIENT}
end;
{$ENDIF}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{***********UNITE*************************************************
Auteur  ...... : JTR - Image libre stockées
Créé le ...... : 10/03/2004
Modifié le ... :
Description .. : Choix d'une image par double click sur une ligne
Mots clefs ... : FO;IMAGES;LIENSOLE;CLAVIERECRAN
*****************************************************************}
procedure TOM_LiensOle.FListeOnDblClick(Sender: TObject);
begin
  if not SelectByDblclick then exit;
  TFFicheListe(Ecran).Retour := GetField('LO_TABLEBLOB') + ';' +
    GetField('LO_IDENTIFIANT') + ';' +
    IntToStr(GetField('LO_RANGBLOB')) + ';' +
    GetField('LO_LIBELLE');
  TFFicheListe(Ecran).Close;
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}

{ Gestions des blobs :
  Arguments :  filtre des blobs
    LO_QUALIFIANTBLOB=xxx   xxx=MEM (mémo) ou PHO (photo)
    LO_EMPLOIBLOB=xxx   code d'emploi à filtrer
}
procedure TOM_LiensOle.OnArgument(Arguments: string);
var
  Critere: string;
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF EAGLSERVER}
  {$IFNDEF ERADIO}
  FF: TFFicheListe;
  {$ENDIF !ERADIO}
  {$ENDIF !EAGLSERVER}
  filterAND: string;
  {$ENDIF}
  stChamp, stVal: string;
  x: integer;
  {$IFDEF EAGLCLIENT}
  pMemo: TRichEdit;
  {$ENDIF}
  PhotoBtn: string; // JTR - Image libre stockées
begin
  inherited;
  filtreEmploi := false;
  filtreQualif := false;
  SetControlVisible('LEMEMO', true);
  SetControlVisible('LAPHOTO', false);
  ObjDefaut := ''; // JTR - Objet par défaut
  // JTR - Image libre stockées
  SelectByDblclick := false;
  PhotoBtn := '';
  CanChangeObj := True;
  {$IFNDEF EAGLSERVER}
  {$IFNDEF ERADIO}
  THGrid(GetControl('FLISTE')).OnDblClick := FListeOnDblClick;
  {$ENDIF !ERADIO}
  {$ENDIF !EAGLSERVER}
  // Fin JTR - Image libre stockées
  QualifiantBlob := THDBValComboBox(GetControl('LO_QUALIFIANTBLOB'));
  if Assigned(QualifiantBlob) then
    QualifiantBlob.OnChange := QualifiantBlobOnChange;

  {$IFDEF EAGLCLIENT}
  pMEMO := TRichEdit(GetControl('LEMEMO'));
  if pMEMO <> nil then pMemo.OnExit := onExitMemo;
  {$ELSE}
  // MD - 13/12/00
  if Ecran.ClassName = 'TFFicheListe' then
  begin
    {$IFNDEF EAGLSERVER}
    {$IFNDEF ERADIO}
    FF := TFFicheListe(ecran);
    FF.ta.filter := '';
    FF.ta.filtered := False;
    {$ENDIF !ERADIO}
    {$ENDIF !EAGLSERVER}
    filterAND := '';
  end
  else
    exit;
  {$ENDIF}

  {récup des arguments }
  repeat
    Critere := uppercase(Trim(ReadTokenSt(Arguments)));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        stChamp := copy(Critere, 1, x - 1);
        stVal := copy(Critere, x + 1, length(Critere));
        if copy(stChamp, 1, 3) = 'LO_' then
        begin
          // MD - 13/12/00
          {$IFNDEF EAGLCLIENT}
          if Ecran.ClassName = 'TFFicheListe' then
          begin
            {$IFNDEF EAGLSERVER}
            {$IFNDEF ERADIO}
            FF.ta.filter := FF.ta.filter + filterAND + stChamp + '=''' + stVal + '''';
            FF.ta.filtered := True;
            {$ENDIF !ERADIO}
            {$ENDIF !EAGLSERVER}
            filterAND := ' AND ';
          end;
          {$ENDIF}
          if stChamp = 'LO_QUALIFIANTBLOB' then
          begin
            {$IFNDEF EAGLCLIENT}
            filtreQualif := true;
            {$ENDIF EAGLCLIENT}
            rangeQualif := stVal;
          end;
          if stChamp = 'LO_EMPLOIBLOB' then
          begin
            {$IFNDEF EAGLCLIENT}
            filtreEmploi := true;
            {$ENDIF EAGLCLIENT}
            rangeEmploi := stVal;
          end;
        end
        else if stChamp = 'DEFAUT' then // JTR - Objet par défaut
          ObjDefaut := stVal
         // JTR - Image libre stockées
        else if (stChamp = 'SELECTBYDBCLICK') and (stVal = 'TRUE') then //Autorise le double click dans liste pour sélection
          SelectByDblClick := True
        else if stChamp = 'PHOTOBTN' then // Pour charger l'image en cours
          PhotoBtn := stVal
        else if (stChamp = 'FORCE') and (stVal = 'TRUE') then // Interdit de changer le type d'objet
          CanChangeObj := False;
        // Fin JTR - Image libre stockées
      end;
    end;
  until Critere = '';
end;

{$IFDEF EAGLCLIENT}
procedure TOM_LiensOle.chargeStream(s: TStringStream);
begin
  setField('LO_OBJET', s.datastring);
end;
{$ELSE}
procedure TOM_LiensOle.chargeStream(s: TMemoryStream);
begin
  TBlobField(DS.FindField('LO_OBJET')).loadFromStream(s);
end;
{$ENDIF}

procedure TOM_LiensOle.OnLoadRecord;
begin
  inherited;
  {$IFDEF EAGLCLIENT}
  // BDU - 25/05/07 - FQ : 14104
  // TFFicheListe(ecran).ta.currentfille.modifie := false;
  {$ENDIF}
  SetControlEnabled('LO_QUALIFIANTBLOB', ((DS.State = dsInsert) and (CanChangeObj))); // JTR - Image libre stockées
  SetControlEnabled('BREFRESH', DS.State <> dsInsert);

  if (Ecran.name = 'YYLIENSOLE') then
  begin
    {$IFNDEF EAGLSERVER}
    SetControlsState;
    {$ENDIF !EAGLSERVER}
    ObjDefaut := GetControlText('LO_QUALIFIANTBLOB'); // JTR - Objet par défaut
  end;
  setControlEnabled('LO_EMPLOIBLOB', (not filtreEmploi));
end;

procedure TOM_LiensOle.OnNewRecord;
begin
  inherited;
  SetField('LO_QUALIFIANTBLOB', 'MEM'); // JTR - Mémo par défaut (à la place de JPG)
  setControlEnabled('LO_QUALIFIANTBLOB', (not filtreQualif));
  if filtreQualif then SetField('LO_QUALIFIANTBLOB', rangeQualif);
  if filtreEmploi then SetField('LO_EMPLOIBLOB', rangeEmploi);
  if ObjDefaut <> '' then SetField('LO_QUALIFIANTBLOB', ObjDefaut); // JTR - Objet par défaut
  SetField('LO_PRIVE', '-');
  SetField('LO_DATEBLOB', date);
end;

procedure TOM_LiensOle.OnUpdateRecord;
var
  QQ: TQuery;
begin
  inherited;
  // MD - 13/12/00
  if (DS.State in [dsInsert]) and (GetField('LO_RANGBLOB') = 0) then
  begin
    QQ := OpenSQL('SELECT MAX(LO_RANGBLOB) AS MAXRANG FROM LIENSOLE WHERE LO_TABLEBLOB="' + GetField('LO_TABLEBLOB') + '" AND LO_IDENTIFIANT="' + GetField('LO_IDENTIFIANT') + '"', TRUE);
    if not QQ.EOF then
    begin
      SetField('LO_RANGBLOB', (QQ.Findfield('MAXRANG').AsInteger) + 1);
    end;
    Ferme(QQ);
  end;
end;

function TOM_LiensOle.GetIDCourant: string;
begin
  result := GetField('LO_IDENTIFIANT');
end;

function TOM_LiensOle.GetKindOfMemo: TKindOfMemo;
begin
  if QualifiantBlob.Value = 'MEM' then
    Result := komMemo
  else if QualifiantBlob.Value = 'PHJ' then
    Result := komPhotoJpeg
  else if QualifiantBlob.Value = 'VIJ' then
    Result := komVignetteJpeg
  else if QualifiantBlob.Value = 'PHO' then
    Result := komPhotoBmp
  else if QualifiantBlob.Value = 'DAT' then
    Result := komData
  else
    Result := komNone;
end;

{$IFNDEF EAGLSERVER}
procedure TOM_LiensOle.SetControlsState;
var
  CC: TImage;
  Kom: TKindOfMemo;
begin
  Kom := GetKindOfMemo;
  SetControlVisible('LEMEMO', Kom in [komMemo, komData]);
  SetControlVisible('LAPHOTO', Kom in [komPhotoJpeg, komVignetteJpeg, komPhotoBmp]);
  SetControlVisible('ZOOM', Kom in [komPhotoJpeg, komVignetteJpeg]);
  SetControlVisible('BIMAGE_INSERT', GetControlVisible('LAPHOTO'));
  if Kom in [komPhotoBmp, komVignetteJpeg, komPhotoJpeg] then
  begin
    CC := TImage(GetControl('LAPHOTO'));
    if (DS.State = dsInsert) then
      CC.Picture.Bitmap.ReleaseHandle
    else
    begin
      if kom = komPhotoBmp then
        LoadBitMapFromChamp(DS, 'LO_OBJET', CC)
      else
        LoadBitMapFromChamp(DS, 'LO_OBJET', CC, True, strToInt(GetControlText('ZOOM')));
    end;
  end
  else if GetKindOfMemo in [komMemo, KomData] then
  begin
    {$IFDEF EAGLCLIENT}
    with TFFicheListe(ecran) do
    begin
      SetControlText('LEMEMO', getField('LO_OBJET'));
      // BDU - 25/05/07 - FQ : 14104
      // Ta.Currentfille.Modifie := False;
    end;
    {$ENDIF}
  end;
end;
{$ENDIF !EAGLSERVER}

procedure TOM_LiensOle.QualifiantBlobOnChange(Sender: TObject);
begin
  {$IFNDEF EAGLSERVER}
  SetControlsState;
  {$ENDIF !EAGLSERVER}
end;

{$IFNDEF EAGLSERVER}
{$IFNDEF RADIO}
procedure AGLSetImage(parms: array of variant; nb: integer);
var
  I: TImage;
  Vide: Boolean;
  FileName: string;
  Q: TDataSet;
  F: TForm;
  {$IFDEF EAGLCLIENT}
  s: TStringStream;
  {$ELSE}
  s: TmemoryStream;
  {$ENDIF}
begin
  F := TForm(Longint(Parms[0]));
  I := TImage(F.FindComponent(string(Parms[1])));
  Vide := Boolean(Parms[2]);
  if i = nil then exit;
  Q := TFFicheListe(F).Ta;
  if vide then
  begin
    if Q.State in [dsInsert, dsEdit] then
    else
      Q.Edit;
    I.picture.graphic := nil;
  end
  else
  begin
    if GetFileName(tfOpenBMP, '*.BMP', FileName) then
    begin
      if Q.State in [dsInsert, dsEdit] then
      else
        Q.Edit;
      I.Picture.LoadFromFile(Filename);
      {$IFDEF EAGLCLIENT}
      s := TStringStream.Create('');
      {$ELSE}
      s := TMemoryStream.Create;
      {$ENDIF}
      I.picture.graphic.savetostream(s);
      TOM_LiensOle(TFFicheListe(F).OM).ChargeStream(s);
      S.Free;
    end;
  end;
end;

procedure AGLSetJpegImage(parms: array of variant; nb: integer);
var
  I: TImage;
  Vide: Boolean;
  FileName: string;
  Q: TDataSet;
  F: TForm;
  {$IFDEF EAGLCLIENT}
  s: TStringStream;
  {$ELSE}
  s: TmemoryStream;
  {$ENDIF}
  Scale: integer;
begin
  F := TForm(Longint(Parms[0]));
  I := TImage(F.FindComponent(string(Parms[1])));
  Vide := Boolean(Parms[2]);
  Scale := strToInt(Parms[3]);
  if i = nil then exit;
  Q := TFFicheListe(F).Ta;
  if vide then
  begin
    if Q.State in [dsInsert, dsEdit] then
    else
      Q.Edit;
    I.picture.graphic := nil;
  end
  else
  begin
    if GetFileName(tfOpenBMP, '*.JPG;*.JPEG', FileName) then
    begin
      if Q.State in [dsInsert, dsEdit] then
      else
        Q.Edit;
      I.Picture.LoadFromFile(Filename);
      SetJPEGOptions(I, Scale);
      {$IFDEF EAGLCLIENT}
      s := TStringStream.Create('');
      {$ELSE}
      s := TMemoryStream.Create;
      {$ENDIF}
      I.picture.graphic.savetostream(s);
      TOM_LiensOle(TFFicheListe(F).OM).ChargeStream(s);
      S.Free;
    end;
  end;
end;
{$ENDIF !RADIO}
{$ENDIF !EAGLSERVER}

procedure AglInitLienOle;
begin
  {$IFNDEF EAGLSERVER}
  RegisterAglProc('SetImage', TRUE, 2, AGLSetImage); //JD
  RegisterAglProc('SetJpegImage', TRUE, 2, AGLSetJpegImage);
  {$ENDIF !EAGLSERVER}
end;

initialization
  RegisterClasses([TOM_LiensOle]);
  {$IFNDEF EAGLSERVER}
  AglInitLienOle;
  {$ENDIF !EAGLSERVER}
end.

