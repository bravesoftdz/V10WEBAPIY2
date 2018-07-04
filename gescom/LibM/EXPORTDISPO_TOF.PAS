{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 09/08/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : EXPORTDISPO ()
Mots clefs ... : TOF;EXPORTDISPO
*****************************************************************}
unit EXPORTDISPO_TOF;

interface

uses UTOF, UTOB,
  {$IFDEF EAGLCLIENT}
  eMul, MaineAGL,
  {$ELSE}
  mul, Fe_Main, dbTables,
  {$ENDIF}
  HMsgBox, FileCtrl, sysutils, HEnt1, Controls, HStatus,
  StdCtrls, HCtrls, ComCtrls, forms, Classes, HQry, EntGC, Paramsoc, M3FP,
  EXPORTASCII_ENTETE_TOM, uToxFTP, UtilGC, UtilArticle, HDimension ;

type
  TOF_EXPORTDISPO = class(TOF)
  private
    stWhere, Separateur, ListeDesChamps, PathFichier: string;
    DateFormat, BufferAlpha, BufferNumerique: string;
    BFormatDynamique: Boolean;
    StGCORLIHOST, StGCORLIUSERNAME, StGCORLIPASSWORD, StGCORLIREPFTP: string;
    TOBChampASCII: TOB;
    FichierASCII: TextFile;
    procedure BatchExport;
    procedure OpenFichierAscii;
    procedure SetEXPORT;
    procedure SetAllEXPORT;
    function ChargeTOBChampASCII: boolean;
    procedure AjouteLignesDuQueryDansFichier(Q: TQuery);
    procedure ChargeCheminExportTRF;
    procedure ParamFTP;
    procedure ExportDuFichierParFTP;
  public
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
  end;

procedure LanceExportDispo(NatureExport: string);

implementation

procedure LanceExportDispo(NatureExport: string);
begin
  AGLLanceFiche('MBO', 'EXPORTDISPO', 'NATUREEXPORT=' + NatureExport, '', '');
end;

procedure TOF_EXPORTDISPO.OnLoad;
begin
  inherited;
  stWhere := RecupWhereCritere(TPageControl(TFMul(Ecran).Pages));
end;

procedure TOF_EXPORTDISPO.OnArgument(S: string);
var THLIB : THLabel;
    iCol,Nbr : integer ;
    FF : TFMUL;
    BVisible : Boolean;
begin
  inherited;
  StGCORLIHOST := GetParamSoc('SO_GCORLIHOST');
  StGCORLIUSERNAME := GetParamSoc('SO_GCORLIUSERNAME');
  StGCORLIPASSWORD := GetParamSoc('SO_GCORLIPASSWORD');
  StGCORLIREPFTP := GetParamSoc('SO_GCORLIREPFTP');

  FF:=TFMul(Ecran) ;

  // Paramétrage des libellés des familles, collection, stat. article et dimensions
  ChangeLibre2('TGA_COLLECTION',Ecran);
  for iCol:=1 to 3 do ChangeLibre2('TGA_FAMILLENIV'+InttoStr(iCol),Ecran);
  if (ctxMode in V_PGI.PGIContexte) and (GetPresentation=ART_ORLI) then
      begin
      for iCol:=4 to 8 do ChangeLibre2('TGA2_FAMILLENIV'+InttoStr(iCol),Ecran);
      for iCol:=1 to 2 do ChangeLibre2('TGA2_STATART'+InttoStr(iCol),Ecran);
      end ;

  // Mise en forme des libellés des dimensions    JD : 20/09/2002
  for iCol:=1 to MaxDimension do
      BEGIN
      THLIB:=THLabel(GetControl('DIMENSION'+IntToStr(iCol))) ;
      THLIB.Caption:=RechDom('GCCATEGORIEDIM','DI'+InttoStr(iCol),False);
      BVisible := Not ((Copy(THLIB.Caption,1,2)='.-') or (THLIB.Caption='??'));
      THLIB.Visible := BVisible; TControl(GetControl(THLIB.FocusControl.Name)).Visible := BVisible;
      END;

  // Paramètrage des libellés des onglets des zones et tables libres du dépôt
  if not VH_GC.GCMultiDepots then
     begin
     SetControlCaption('PTABLESLIBRESDEP','Tables Libres Etab.') ;
     SetControlCaption('PZONESLIBRESDEP','Zones Libres Etab.') ;
     end;

  // Paramétrage des libellés des tables libres articles et dépôts
  Nbr := 0;
  if (GCMAJChampLibre (FF, False, 'EDIT', 'GA_LIBREART', 10, '') = 0) then SetControlVisible('PTABLESLIBRES', False) ;
  if (GCMAJChampLibre (FF, False, 'EDIT', 'GDE_LIBREDEP', 10, '') = 0) then SetControlVisible('PTABLESLIBRESDEP', False) ;
  // Mise en forme des libellés des dates, booléans libres et montants libres
  if (GCMAJChampLibre (FF, False, 'EDIT', 'GA_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VAL', False) else inc(Nbr) ;
  if (GCMAJChampLibre (FF, False, 'EDIT', 'GA_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATE', False) else inc(Nbr) ;
  if (GCMAJChampLibre (FF, False, 'BOOL', 'GA_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOL', False) else inc(Nbr) ;
  if (Nbr = 0) then SetControlVisible('PZONESLIBRES', False) ;
  Nbr := 0;
  if (GCMAJChampLibre (FF, False, 'EDIT', 'GDE_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VALDEP', False) else inc(Nbr) ;
  if (GCMAJChampLibre (FF, False, 'EDIT', 'GDE_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATEDEP', False) else inc(Nbr) ;
  if (GCMAJChampLibre (FF, False, 'BOOL', 'GDE_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOLDEP', False) else inc(Nbr) ;
  if (Nbr = 0) then SetControlVisible('PZONESLIBRESDEP', False) ;

end;

procedure TOF_EXPORTDISPO.BatchExport;
var i: integer;
begin
  with TFMul(Ecran) do
  begin
    if (FListe.NbSelected = 0) and (not FListe.AllSelected) then
    begin
      PGIInfo('Veuillez sélectionner les stocks à exporter', Caption);
      exit;
    end;

    PathFichier := GetControlText('EMPLACEMENT');
    if (PathFichier = '') or not DirectoryExists(ExtractFileDir(PathFichier)) then
    begin
      PGIInfo('Vous devez sélectionner un emplacement de fichier.', Caption);
      exit;
    end;

    if not ChargeTOBChampASCII then
    begin
      PGIError('Impossible de charger le paramétrage d''export.', Caption);
      exit;
    end;

    SourisSablier;
    try
      if FListe.AllSelected then
      begin
        if PGIAsk('Voulez-vous exporter toutes les stocks ?', Caption) <> mrYes then exit;
        OpenFichierAscii;
        if Transactions(SetAllEXPORT, 2) <> oeOK then PGIError('Impossible d''exporter toutes les stocks', Caption);
        FListe.AllSelected := false;
      end
      else
      begin
        if PGIAsk('Voulez-vous exporter les stocks sélectionnées ?', Caption) <> mrYes then exit;
        OpenFichierAscii;
        InitMove(FListe.NbSelected, '');
        for i := 0 to FListe.NbSelected - 1 do
        begin
          FListe.GotoLeBookMark(i);
          {$IFDEF EAGLCLIENT}
          Q.TQ.Seek(FListe.Row - 1);
          {$ENDIF}
          if Transactions(SetEXPORT, 2) <> oeOK then
            PGIBox('Impossible d''exporter le stock de l''article ' + Q.FindField('GQ_ARTICLE').AsString +
              'et le dépôt ' + Q.FindField('GQ_DEPOT').AsString, Caption);
          MoveCur(False);
        end;
        FListe.ClearSelected;
        FiniMove;
      end;
      ChercheClick;
    finally
      CloseFile(FichierASCII);

      //Si EXPORTFTP est coché : envoie le fichier par FTP
      if GetCheckBoxState('EXPORTFTP') = cbChecked then ExportDuFichierParFTP;

      if TOBChampASCII <> nil then
      begin
        TOBChampASCII.Free;
        TOBChampASCII := nil;
      end;
      bSelectAll.Down := False;
      SourisNormale;
    end;
  end;
end;

procedure TOF_EXPORTDISPO.OpenFichierAscii;
begin
  AssignFile(FichierASCII, PathFichier);
  if (GetCheckBoxState('COMPLETER') = cbChecked) and FileExists(PathFichier) then
    Append(FichierASCII) else Rewrite(FichierASCII);
end;

procedure TOF_EXPORTDISPO.SetEXPORT;
var Article, Depot, DateCloture, Cloture: string;
  StWhereDispo, StSQL, StJointure: string;
  Q: TQuery;
begin
  StJointure := CreationJointure('GQ', ListeDesChamps);
  with TFMul(Ecran) do
  begin
    Article := Q.FindField('GQ_ARTICLE').AsString;
    Depot := Q.FindField('GQ_DEPOT').AsString;
    //DateCloture := USDateTime(Q.FindField('GQ_DATECLOTURE').AsDateTime);
    //Cloture := Q.FindField('GQ_CLOTURE').AsString;
    DateCloture := '01/01/1900';
    Cloture     := '-' ;
  end;
  StWhereDispo := 'WHERE GQ_ARTICLE="' + Article + '" AND GQ_DEPOT="' + Depot + '" AND ' +
    'GQ_DATECLOTURE="' + DateCloture + '" AND GQ_CLOTURE="' + Cloture + '"';
  StSQL := 'SELECT ' + ListeDesChamps + ' FROM DISPO ' +
    StJointure + StWhereDispo;
  Q := OpenSQL(StSQL, True);
  if not Q.Eof then AjouteLignesDuQueryDansFichier(Q);
  Ferme(Q);
end;

// Procédure appelée par une transaction pour exporter les dispos de tous les articles sélectionnés

procedure TOF_EXPORTDISPO.SetAllEXPORT;
var StJointure, StSQL: string;
  Q: TQuery;
begin
  StJointure := CreationJointure('GQ', ListeDesChamps);

  StSQL := 'SELECT ' + ListeDesChamps + ' FROM DISPO ' + StJointure + StWhere;
  Q := OpenSQL(StSQL, True);
  if not Q.Eof then AjouteLignesDuQueryDansFichier(Q);
  Ferme(Q);
end;

function TOF_EXPORTDISPO.ChargeTOBChampASCII: boolean;
var Q, QInfoSupp: TQuery;
  CodeExport, NatureExport: string;
begin
  Result := False;

  CodeExport := GetControlText('LISTEEXPORT');
  NatureExport := GetControlText('NATUREEXPORT');
  //Charge les champs du fichier ASCII
  TOBChampASCII := Tob.Create('Table EXPORTASCII', nil, -1);
  Q := OpenSQL('SELECT * FROM EXPORTASCII WHERE GEX_NATURE="' + NatureExport + '" AND GEX_CODEEX="' + CodeExport + '" ORDER BY GEX_POSITION', True);
  if not Q.Eof then TOBChampASCII.LoadDetailDB('EXPORTASCII', '', '', Q, False)
  else
  begin
    Ferme(Q);
    TOBChampASCII.Free;
    exit;
  end;
  Ferme(Q);

  ASCIIAjouteChampGaucheEtDate(TOBChampASCII);

  ListeDesChamps := ASCIIConstruireListeDeChamps(TOBChampASCII);

  QInfoSupp := OpenSQL('SELECT GAS_SEPARATEUR,GAS_DATEFORMAT,GAS_ALPHA,' +
  'GAS_NUMERIQUE, GAS_FORMATFIXE FROM EXPORTASCII_ENTETE WHERE GAS_NATURE="' +
  NatureExport + '" AND GAS_CODEEX="' + CodeExport + '"', True);

  Separateur := '';
  DateFormat := '';
  BufferAlpha := ' ';
  BufferNumerique := ' ';
  if not QInfoSupp.Eof then
  begin
    Separateur := QInfoSupp.FindField('GAS_SEPARATEUR').AsString;
    DateFormat := QInfoSupp.FindField('GAS_DATEFORMAT').AsString;
    BufferAlpha := QInfoSupp.FindField('GAS_ALPHA').AsString;
    BufferNumerique := QInfoSupp.FindField('GAS_NUMERIQUE').AsString;
    BFormatDynamique := QInfoSupp.FindField('GAS_FORMATFIXE').AsString = '-';
  end;
  Ferme(QInfoSupp);

  Result := True;
end;

procedure TOF_EXPORTDISPO.AjouteLignesDuQueryDansFichier(Q: TQuery);
var UneLigne: string;
begin
  Q.First;
  while not Q.Eof do
  begin
    UneLigne := ASCIIConstruireUneLigne(Q, TOBChampASCII, GetControlText('LISTEEXPORT'),
      Separateur, DateFormat, BufferAlpha, BufferNumerique, BFormatDynamique);
    WriteLn(FichierASCII, UneLigne);
    Q.Next;
  end;
end;

procedure TOF_EXPORTDISPO.ChargeCheminExportTRF;
var Q: TQuery;
begin
  Q := OpenSQL('SELECT GAS_FICHIEREXPORT FROM EXPORTASCII_ENTETE WHERE GAS_NATURE="' + GetControlText('NATUREEXPORT') + '" AND GAS_CODEEX="' +
    GetControlText('LISTEEXPORT') + '"', True);
  if not Q.EOF then SetControlText('EMPLACEMENT', Q.FindField('GAS_FICHIEREXPORT').AsString);
  Ferme(Q);
end;

procedure TOF_EXPORTDISPO.ParamFTP;
var LesParametres, UnParam, StParam, ValParam: string;
  x: integer;
begin
  LesParametres := AGLLanceFiche('MBO', 'ORLIFTP', '', '', '');
  while LesParametres <> '' do
  begin
    UnParam := READTOKENST(LesParametres);
    if UnParam <> '' then
    begin
      x := pos('=', UnParam);
      if x <> 0 then
      begin
        StParam := copy(UnParam, 1, x - 1);
        ValParam := copy(UnParam, x + 1, length(UnParam));
        if StParam = 'SO_GCORLIHOST' then StGCORLIHOST := ValParam
        else
          if StParam = 'SO_GCORLIUSERNAME' then StGCORLIUSERNAME := ValParam
        else
          if StParam = 'SO_GCORLIPASSWORD' then StGCORLIPASSWORD := ValParam
        else
          if StParam = 'SO_GCORLIREPFTP' then StGCORLIREPFTP := ValParam
          ;
      end;
    end;
  end;
end;

procedure TOF_EXPORTDISPO.ExportDuFichierParFTP;
var xpFtp: TOXFTP;
  PortFTP: integer;
  HostFileName: string;
begin
  if (StGCORLIHOST = '') or (StGCORLIUSERNAME = '') or (StGCORLIPASSWORD = '') then exit;
  PortFTP := 21;
  xpFtp := TOXFTP.Create(PortFtp);
  if xpFTP.Connect(StGCORLIHOST, StGCORLIUSERNAME, StGCORLIPASSWORD) then
  begin
    HostFileName := StGCORLIREPFTP + ExtractFileName(PathFichier);
    xpFTP.PutFileAscii(PathFichier, HostFileName);
    xpFTP.DisConnect;
  end;
  xpFTP.Free;
end;

///////////////////////////////////////////////////////////////////////////////

// procedure appellée par le bouton BOuvrir

procedure AGLBatchEXPORTDispo(Parms: array of Variant; Nb: Integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFmul) then TOTOF := TFMul(F).LaTOF else exit;
  if (TOTOF is TOF_EXPORTDISPO) then TOF_EXPORTDISPO(TOTOF).BatchExport;
end;

procedure AGLChargeCheminExportDispo(Parms: array of Variant; Nb: Integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFmul) then TOTOF := TFMul(F).LaTOF else exit;
  if (TOTOF is TOF_EXPORTDISPO) then TOF_EXPORTDISPO(TOTOF).ChargeCheminExportTRF;
end;

procedure AGLParamFTPDispo(Parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFmul) then TOTOF := TFMul(F).LaTOF else exit;
  if (TOTOF is TOF_EXPORTDISPO) then TOF_EXPORTDISPO(TOTOF).ParamFTP;
end;

initialization
  registerclasses([TOF_EXPORTDISPO]);
  RegisterAGLProc('BatchEXPORTDispo', True, 0, AGLBatchEXPORTDispo);
  RegisterAGLProc('ChargeCheminExportDispo', True, 0, AGLChargeCheminExportDispo);
  RegisterAglProc('ParamFTPDispo', True, 0, AGLParamFTPDispo);
end.
