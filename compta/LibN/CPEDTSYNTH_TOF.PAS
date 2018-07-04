{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 18/10/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPEDTSYNTH ()
Mots clefs ... : TOF;CPEDTSYNTH
*****************************************************************}
Unit CPEDTSYNTH_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     uTob,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     hdb,
     fe_main,
{$else}
     eMul,
     MaineAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     HTB97 ;

Type
  TEdtSynthAction = (aImport, aExport);

  TOF_CPEDTSYNTH = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    procedure Traitement ( pAction : TEdtSynthAction);
    procedure ExporteEtat (pstFichier  : string );
    procedure ImporteEtat (pstFichier  : string );
    procedure OnClickExport (Sender : TObject);
    procedure OnClickImport (Sender : TObject);
  end ;

procedure CPLanceFiche_GestionEtatSynthese;

Implementation

uses
  uYFILESTD
  , cbpPath
  , Lookup;

procedure CPLanceFiche_GestionEtatSynthese;
begin
  AglLanceFiche('CP', 'CPEDTSYNTH', '', '', '');
end;

procedure TOF_CPEDTSYNTH.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPEDTSYNTH.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPEDTSYNTH.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPEDTSYNTH.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPEDTSYNTH.OnArgument (S : String ) ;
begin
  Inherited ;
  TToolBarButton97(GetControl('BEXPORTMAQUETTE')).OnClick := OnClickExport;
  TToolBarButton97(GetControl('BIMPORTMAQUETTE')).OnClick := OnClickImport;
end ;

procedure TOF_CPEDTSYNTH.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPEDTSYNTH.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPEDTSYNTH.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPEDTSYNTH.Traitement ( pAction : TEdtSynthAction);
var
  {$IFDEF EAGLCLIENT}
  F : THGrid;
  {$ELSE}
  F : THDBGrid;
  {$ENDIF}
  n : integer;
  lstFichier : string;
begin
  {$IFDEF EAGLCLIENT}
  F := THGrid(TFMul(Ecran).FListe);
  {$ELSE}
  F := THDBGrid(TFMul(Ecran).FListe);
  {$ENDIF}

  {Aucune sélection, on sort}
  if (F.NbSelected = 0) and not F.AllSelected then begin
    PGIInfo (TraduireMemoire('Aucun élément n''est sélectionné.'));
    Exit;
  end;

  if GetFileName(tfSave ,'*.XML',lstFichier) then
  begin
    {$IFNDEF EAGLCLIENT}
    TFMul(Ecran).Q.First;
    if F.AllSelected then
      while not TFMul(Ecran).Q.EOF do
      begin
        if pAction = aExport then ExporteEtat (lstFichier);
        TFMul(Ecran).Q.Next;
      end
    else
    {$ENDIF}

    for n := 0 to F.nbSelected - 1 do begin
      F.GotoLeBookmark(n);
      {$IFDEF EAGLCLIENT}
      TFMul(Ecran).Q.TQ.Seek(F.Row - 1);
      {$ENDIF}
      if pAction = aExport then ExporteEtat (lstFichier);
    end;
  end;
  if pAction = aImport then TFMul(Ecran).BCherche.Click;
end;

procedure TOF_CPEDTSYNTH.ExporteEtat ( pstFichier : string );
var
    lstChemin : string;
    lRetour : integer;
    lstSQL : string;
    lT, lTFichier : TOB;
    i : integer;
    lFichier : TStringList;
    lstTitre : string;
begin
  { On efface le fichier sur le disque }
  lstChemin := AGL_YFILESTD_GET_PATH ('COMPTA', GetField('YFS_NOM'), 'ETATSYNTH',
      GetField('YFS_CRIT2'), GetField('YFS_CRIT3'), '', '', V_PGI.LanguePrinc, GetField('YFS_PREDEFINI'));
  DeleteFile (lstChemin);

  { On extrait le fichier sur le disque }
  lRetour := AGL_YFILESTD_EXTRACT( lstChemin, 'COMPTA',GetField('YFS_NOM'),
      'ETATSYNTH',GetField('YFS_CRIT2'), GetField('YFS_CRIT3'),'','',False,
      V_PGI.LanguePrinc,GetField('YFS_PREDEFINI'));
  if lRetour <> -1 then
  begin
    PGIInfo(AGL_YFILESTD_GET_ERR(lRetour) + #13#10 + lstChemin);
    exit;
  end;

  { Création de la TOB avec les éléments significatifs de l'état }
  lT := TOB.Create ('EDTSYNTH',nil,-1);
  lstSQL := 'SELECT * FROM YFILESTD WHERE YFS_CODEPRODUIT="COMPTA" AND '+
    'YFS_CRIT1="ETATSYNTH" AND YFS_CRIT2="'+GetField('YFS_CRIT2')+'" AND '+
    'YFS_CRIT3="'+GetField('YFS_CRIT3')+'" AND YFS_PREDEFINI="'+GetField('YFS_PREDEFINI')+'"';

  { Enregistrement de YFILESTD }
  lT.LoadDetailDBFromSQL('YFILESTD',lstSQL);

  { Enregistrement de la description de la maquette }
  lFichier := TStringList.Create;
  lFichier.LoadFromFile(lstChemin);
  lTFichier := TOB.Create('',lT.Detail[0],-1);
  for i := 0 to lFichier.Count - 1 do
  begin
    if i = 0 then lstTitre := lFichier[i];
    with TOB.Create('',lTFichier,-1) do
    begin
      AddChampSupValeur('LIGNE',lFichier[i]);
    end;
  end;
  lTFichier.AddChampSupValeur('TITRE',lstTitre);
  lFichier.Free;

  { Sauvegarde sous la forme d'un fichier XML }
  lT.SaveToXmlFile(pstFichier,True,False);
  lT.Free;
end;

procedure TOF_CPEDTSYNTH.OnClickExport(Sender: TObject);
begin
  Traitement (aExport);
end;

procedure TOF_CPEDTSYNTH.OnClickImport(Sender: TObject);
var lstFichier : string;
begin
  if GetFileName(tfOpen ,'*.XML',lstFichier) then
  begin
    ImporteEtat ( lstFichier);
  end;
end;

procedure TOF_CPEDTSYNTH.ImporteEtat(pstFichier: string);

  function _MaquettePresente (pstNom, pstCrit2, pstCrit3, pstPredefini : string) : boolean;
  begin
    Result := ExisteSQL ('SELECT * FROM YFILESTD WHERE YFS_NOM="'+pstNom+'" AND '+
        'YFS_CRIT1="ETATSYNTH" AND YFS_CODEPRODUIT="COMPTA" AND YFS_CRIT2="'+pstCrit2+
        '" AND YFS_CRIT3="'+pstCrit3+'" AND YFS_PREDEFINI="'+pstPredefini+'"'+
        ' YFS_LANGUE="' + V_PGI.LanguePrinc + '"');
  end;

  function _ExtraireFichierMaquette (pstNom : string; pT : TOB) : string;
  var i : integer;
      lSL :  TStringList;
  begin
    lSL := TStringList.Create;
    for i := 0 to pT.Detail.Count - 1 do
      lSL.Add(pT.GetString('LIGNE'));
    lSL.SaveToFile(pstNom);
    lSL.Free;
  end;

var
  lT : TOB;
  lstEntete : boolean;
  lstEncoding : string;
  i : integer;
  lstCrit2, lstCrit3, lstNom, lstPredefini : string;
  mrAsk : integer;
  lstMaquette : string;
  lRet : integer;
  lstTitre : string;
begin
  lT := TOB.Create('',nil,-1);
  try
    lT.LoadFromXMLFile(pstFichier,lstEntete,lstEncoding);
    for i:=0 to lT.Detail.Count - 1 do
    begin
      lstCrit2 := lT.Detail[i].GetString('YFS_CRIT2');
      lstCrit3 := lT.Detail[i].GetString('YFS_CRIT3');
      lstNom := lT.Detail[i].GetString('YFS_NOM');
      lstPredefini := lT.Detail[i].GetString('YFS_PREDEFINI');
      if _MaquettePresente (lstNom, lstCrit2, lstCrit3, lstPredefini) then
      begin
        mrAsk := PGIAskCancel(TraduireMemoire('La maquette '+lstNom+' existe déjà. Voulez-vous la remplacer ?'));
        case mrAsk of
          mrCancel : break;
          mrNo : continue;
        end;
      end;
      lstMaquette := _ExtraireFichierMaquette (TCBPPath.GetCegidUserTempPath+lstNom, lt.Detail[i].Detail[0]);
      lstTitre := lt.Detail[i].Detail[0].GetString('TITRE');
      lRet := AGL_YFILESTD_IMPORT(TCBPPath.GetCegidUserTempPath+lstNom, 'COMPTA',
            ExtractFileName(lstNom), ExtractFileExt(lstNom), 'ETATSYNTH', lstCrit2,
            lStCrit3, '', '', '-', '-', '-', '-', '-', V_PGI.LanguePrinc, lStPredefini, lstTitre, '000000');
      if lRet <> -1 then
      begin // Problème, on affiche un message d'erreur et on sort
        PGIInfo(AGL_YFILESTD_GET_ERR(lRet) + #13#10 + lstNom);
        break;
      end;
    end;
  finally
    lT.Free;
  end;
end;

Initialization
  registerclasses ( [ TOF_CPEDTSYNTH ] ) ;
end.
