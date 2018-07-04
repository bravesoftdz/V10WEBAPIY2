{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 24/04/2001
Modifié le ... : 03/05/2007
Description .. : Source TOM de la TABLE : STDMAQ (STDMAQ)
Suite ........ : - Suppression du niveau dossier inutilisé. Stockage des 
Suite ........ : numéros de maquette sur 3 caractères
Mots clefs ... : TOM;STDMAQ
*****************************************************************}
{***********MODIF*************************************************
- 27/11/2001 - CA - Enregistrement des maquettes cabinet dans le
                    répertoire DAT.
*****************************************************************}
Unit uTOMStdMaq;

Interface

Uses Windows
      , StdCtrls
      , Controls
      , Classes
      , db
      , forms
      , sysutils
      {$IFNDEF DBXPRESS} ,dbtables {$ELSE} ,uDbxDataSet {$ENDIF}
      , ComCtrls
      , HCtrls
      , HEnt1
      , HMsgBox
      , UTOM
//      , Fiche
{$IFDEF EAGLCLIENT}
      , eFichList
{$ELSE}
      , FichList
{$ENDIF}
      , UTob
      , HTB97
      , Ent1
      , uYFILESTD; // AGL_YFILESTD_GET_PATH

Type
  TOM_STDMAQ = Class (TOM)
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
      fTypeMaquette : string;
      procedure OnClickParamMaquette(Sender: TObject);
      procedure OnDblClickMaquette(Sender: TObject);
      function  GetMaquettePathName(bModele : boolean) : string;
      procedure BimprimeClick(Sender: TObject);
      function CopieMaquette: boolean;
      function SupprimeMaquette : boolean;
    end ;

Implementation

uses Gridsynth,Critedt, GalOutil,ImprimeMaquette;

procedure TOM_STDMAQ.OnNewRecord ;
begin
  Inherited ;
  SetControlVisible('ZMODELE',True);
  SetControlVisible('lZMODELE',True);
  SetControlEnabled ('STM_PREDEFINI',False);
  if ((ctxStandard in V_PGI.PGIContexte) and (not(EstSpecif('51502')))) then
  begin
    SetField('STM_PREDEFINI','STD');
    SetControlProperty('STM_NUMPLAN','MaxLength',3);
    SetControlProperty('STM_NUMPLAN','MinValue',21);
    SetControlProperty('STM_NUMPLAN','MaxValue',999);
  end else if ((ctxStandard in V_PGI.PGIContexte) and (EstSpecif('51502'))) then
  begin
    SetField('STM_PREDEFINI','CEG');
    SetControlProperty('STM_NUMPLAN','MaxLength',2);
    SetControlProperty('STM_NUMPLAN','MinValue',1);
    SetControlProperty('STM_NUMPLAN','MaxValue',20);
  end;
end ;

procedure TOM_STDMAQ.OnDeleteRecord ;
begin
  if not IsSuperviseur(TRUE) then
  begin
    LastErrorMsg := 'Vous n''êtes pas superviseur, la suppression est impossible';
    LastError := -1;
  end;
  if ((not (EstSpecif('51502')) and (StrToInt(GetField ('STM_NUMPLAN')) <= 20))) then
  begin
    LastErrorMsg := 'Maquette CEGID, la suppression est impossible';
    LastError := -1;
  end else if ((not (ctxStandard in V_PGI.PGIContexte)) and (GetField ('STM_PREDEFINI') = 'STD')) then
  begin
    LastErrorMsg := 'Maquette Standard, la suppression est impossible';
    LastError := -1;
  end;
  Inherited ;
  if LastError <> - 1 then
  begin
    if not SupprimeMaquette then
    begin
      LastErrorMsg := 'Suppression de la maquette en base impossible.';
      LastError := -1;
    end;
  end;
end ;

procedure TOM_STDMAQ.OnUpdateRecord ;
begin
  if DS.State in [dsInsert] then
  begin
    if ((not (EstSpecif('51502')) and (StrToInt(GetField ('STM_NUMPLAN')) <= 20))) then
    begin
      LastErrorMsg := 'Création impossible. Numéro réservé à CEGID.';
      LastError := -1;
      exit;
    end else if ((not (ctxStandard in V_PGI.PGIContexte)) and (GetField ('STM_PREDEFINI') = 'STD')) then
    begin
      LastErrorMsg := 'Création impossible. Numéro réservé aux standards.';
      LastError := -1;
      exit;
    end;
    if not CopieMaquette then
    begin
      LastErrorMsg := 'Copie de la maquette impossible.';
      LastError := -1;
    end;
  end;
  Inherited ;
  if LastError >=0 then
  begin
    SetControlProperty('STM_NUMPLAN','MaxLength',3);
    SetControlProperty('STM_NUMPLAN','MinValue',0);
    SetControlProperty('STM_NUMPLAN','MaxValue',999);
  end;
end ;

procedure TOM_STDMAQ.OnAfterUpdateRecord ;
begin
  Inherited ;
  SetControlVisible('ZMODELE',False);
  SetControlVisible('lZMODELE',False);
end ;

procedure TOM_STDMAQ.OnLoadRecord ;
begin
  Inherited ;
  if (not (DS.State in  [dsInsert])) then
  begin
    SetControlVisible('ZMODELE',False);
    SetControlVisible('lZMODELE',False);
    SetControlProperty('STM_NUMPLAN','MaxLength',3);
    SetControlProperty('STM_NUMPLAN','MinValue',0);
    SetControlProperty('STM_NUMPLAN','MaxValue',999);
  end;
end ;

procedure TOM_STDMAQ.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_STDMAQ.OnArgument ( S: String ) ;
begin
  Inherited ;
  ReadTokenSt (S);
  fTypeMaquette := ReadTokenSt(S);
  TToolbarButton97(GetControl('BMaquette')).Onclick := OnClickParamMaquette;
  TToolbarButton97(GetControl('FLISTE')).OnDblClick := OnDblClickMaquette;
  THEdit(GetControl('ZMODELE')).Plus := fTypeMaquette;
  TToolbarButton97(GetControl('BImprimer')).Onclick := BimprimeClick;
end ;

procedure TOM_STDMAQ.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_STDMAQ.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_STDMAQ.OnClickParamMaquette(Sender: TObject);
var Crit : TCritEdtPCL;
    CritEdtChaine : TCritEdtChaine;
    lStCheminMaquette : string;
    lStPredefini  : string;
    lStNomFichier : string;
begin
  if DS.State in [dsInsert] then
  begin
    PGIBox ('Veuillez sauvegarder votre maquette avant d''accéder au paramétrage.',Ecran.Caption);
    exit;
  end;
  if fTypeMaquette = 'CR' then Crit.Te := esCR;
  if fTypeMaquette = 'SIG' then Crit.Te := esSIG;
  if fTypeMaquette = 'BIL' then Crit.Te := esBIL;

  if GetField('STM_NUMPLAN') > 0 then
  begin
    Fillchar(CritEdtChaine,SizeOf(CritEdtChaine),#0) ;
    CritEdtChaine.Utiliser := False;

    // GCO - 15/09/2006 - FQ 18705 - Nouvelle gestion en base des maquettes
    // Suppression du fichier sur disque pour récupération systématique
    lStPredefini := GetColonneSQL('YFILESTD', 'YFS_PREDEFINI',
                    'YFS_CODEPRODUIT = "COMPTA" AND ' +
                    'YFS_CRIT1 = "ETATSYNTH" AND ' +
                    'YFS_CRIT2 = "' + FTypeMaquette + '" AND ' +
                    'YFS_CRIT3 = "' + IntToStr(GetField('STM_NUMPLAN')) + '"');

    lStNomFichier :=  GetColonneSQL('YFILESTD', 'YFS_NOM',
                    'YFS_CODEPRODUIT = "COMPTA" AND ' +
                    'YFS_CRIT1 = "ETATSYNTH" AND ' +
                    'YFS_CRIT2 = "' + FTypeMaquette + '" AND ' +
                    'YFS_CRIT3 = "' + IntToStr(GetField('STM_NUMPLAN')) + '" AND ' +
                    'YFS_PREDEFINI = "' + lStPredefini + '"');

    lStCheminMaquette := AGL_YFILESTD_GET_PATH('COMPTA', lStNomFichier, 'ETATSYNTH', FTypeMaquette, IntToStr(GetField('STM_NUMPLAN')), '', '', V_PGI.LanguePrinc, lStPredefini);
    DeleteFile(lStCheminMaquette);

    if AGL_YFILESTD_EXTRACT( lStCheminMaquette, 'COMPTA', lStNomFichier, 'ETATSYNTH', FTypeMaquette, IntToStr(GetField('STM_NUMPLAN')), '', '', False, V_PGI.LanguePrinc, lStPredefini) = -1 then
    begin
      LanceLiasse(lStCheminMaquette,FALSE,True,['N'],['#,##0.00'],['#,##0.00'],Crit,TRUE,CritEdtChaine);
    end
    else
      PGIBox ('Fichier maquette introuvable.',Ecran.Caption);

    // FIN GCO
  end;
end;

function TOM_STDMAQ.GetMaquettePathName(bModele : boolean) : string;
var
  lStPredefini, lStNomFichier, lStCheminMaquette : string;
  iNumMaq : integer;
begin
  if bModele then
    iNumMaq := StrToInt (GetControlText('ZMODELE'))
  else iNumMaq := StrToInt(GetField('STM_NUMPLAN'));
  lStPredefini := GetColonneSQL('YFILESTD', 'YFS_PREDEFINI',
                    'YFS_CODEPRODUIT = "COMPTA" AND ' +
                    'YFS_CRIT1 = "ETATSYNTH" AND ' +
                    'YFS_CRIT2 = "' + FTypeMaquette + '" AND ' +
                    'YFS_CRIT3 = "' + IntToStr(iNumMaq) + '"');

  lStNomFichier :=  GetColonneSQL('YFILESTD', 'YFS_NOM',
                    'YFS_CODEPRODUIT = "COMPTA" AND ' +
                    'YFS_CRIT1 = "ETATSYNTH" AND ' +
                    'YFS_CRIT2 = "' + FTypeMaquette + '" AND ' +
                    'YFS_CRIT3 = "' + IntToStr(iNumMaq) + '" AND ' +
                    'YFS_PREDEFINI = "' + lStPredefini + '"');

  lStCheminMaquette := AGL_YFILESTD_GET_PATH('COMPTA', lStNomFichier, 'ETATSYNTH', FTypeMaquette, IntToStr(iNumMaq), '', '', V_PGI.LanguePrinc, lStPredefini);
  DeleteFile(lStCheminMaquette);
  if AGL_YFILESTD_EXTRACT( lStCheminMaquette, 'COMPTA', lStNomFichier, 'ETATSYNTH', FTypeMaquette, IntToStr(iNumMaq), '', '', False, V_PGI.LanguePrinc, lStPredefini) = -1 then
    Result := lStCheminMaquette
  else Result := '';
end;

function TOM_STDMAQ.CopieMaquette : boolean;
var
  lStPredefini, lStNomFichier, lStCheminMaquette : string;
  lStNouveauFichier : string;
  iNumModele : integer;
  lRet : integer;
begin
  { Extraction du modèle sur le disque }
  iNumModele := StrToInt (GetControlText('ZMODELE'));
  lStPredefini := GetColonneSQL('YFILESTD', 'YFS_PREDEFINI',
                    'YFS_CODEPRODUIT = "COMPTA" AND ' +
                    'YFS_CRIT1 = "ETATSYNTH" AND ' +
                    'YFS_CRIT2 = "' + FTypeMaquette + '" AND ' +
                    'YFS_CRIT3 = "' + IntToStr(iNumModele) + '"');

  lStNomFichier :=  GetColonneSQL('YFILESTD', 'YFS_NOM',
                    'YFS_CODEPRODUIT = "COMPTA" AND ' +
                    'YFS_CRIT1 = "ETATSYNTH" AND ' +
                    'YFS_CRIT2 = "' + FTypeMaquette + '" AND ' +
                    'YFS_CRIT3 = "' + IntToStr(iNumModele) + '" AND ' +
                    'YFS_PREDEFINI = "' + lStPredefini + '"');
  lStCheminMaquette := AGL_YFILESTD_GET_PATH('COMPTA', lStNomFichier, 'ETATSYNTH', FTypeMaquette, IntToStr(iNumModele), '', '', V_PGI.LanguePrinc, lStPredefini);
  DeleteFile(lStCheminMaquette);
  if AGL_YFILESTD_EXTRACT( lStCheminMaquette, 'COMPTA', lStNomFichier, 'ETATSYNTH', FTypeMaquette, IntToStr(iNumModele), '', '', False, V_PGI.LanguePrinc, lStPredefini) = -1 then
  begin
    { Constitution du nom du fichier maquette }
    lStNouveauFichier := FTypeMaquette + Format('%.03d', [StrToInt(GetField('STM_NUMPLAN'))]) + '.TXT';
    { Recopie du fichier maquette depuis le modéle }
    CopyFile(PChar(lStCheminMaquette), PChar(ExtractFilePath(lStCheminMaquette)+'\'+lStNouveauFichier),False);
    { Intégration du fichier en base }
    if StrToInt(GetField('STM_NUMPLAN'))>20 then lStPredefini:='STD' else lStPredefini:='CEG';
    lRet := AGL_YFILESTD_IMPORT(ExtractFilePath(lStCheminMaquette)+'\'+lStNouveauFichier, 'COMPTA', lStNouveauFichier, ExtractFileExt(lStNouveauFichier),
      'ETATSYNTH', FTypeMaquette, IntToStr(StrToInt(GetField('STM_NUMPLAN'))), '', '', '-', '-', '-', '-', '-', V_PGI.LanguePrinc, lStPredefini, GetField('STM_LIBELLE'), '000000');
    Result := (lRet=-1);
  end else Result := False;
end;

procedure TOM_STDMAQ.OnDblClickMaquette(Sender: TObject);
begin
  OnClickParamMaquette(nil);
end;

procedure TOM_STDMAQ.BimprimeClick(Sender: TObject);
var
Chemin    : string;
begin
  inherited;
  Chemin := GetMaquettePathName (False);
  ControlTextToPrinter (Chemin);
end;

function TOM_STDMAQ.SupprimeMaquette: boolean;
var
  lStFichier, lStPredefini : string;
  lStFileGUID : string;
  Q : TQuery;
  bErreur : boolean;
begin
  bErreur := False;
  lStFichier := FTypeMaquette + Format('%.03d', [StrToInt(GetField('STM_NUMPLAN'))]) + '.TXT';
  if (StrToInt(GetField('STM_NUMPLAN')) > 20) then lStPredefini := 'STD' else lStPredefini :='CEG';
  Q := OpenSql ('SELECT YFS_FILEGUID FROM YFILESTD WHERE YFS_CODEPRODUIT="COMPTA" AND YFS_NOM="' + Trim (lStFichier) + '" AND YFS_LANGUE="FRA" AND YFS_PREDEFINI="' + Trim (lSTPredefini) + '" AND YFS_CRIT1="ETATSYNTH" AND YFS_NODOSSIER="000000"', True) ;
  if not Q.Eof and (Q.FindField ('YFS_FILEGUID').AsString <> '' ) then
  begin
    lStFileGUID := Q.FindField ('YFS_FILEGUID').AsString;
    Ferme (Q);
    BeginTrans ();
    try
      ExecuteSQL ('DELETE FROM NFILES WHERE NFI_FILEGUID="'+lStFileGUID+'"');
      ExecuteSQL ('DELETE FROM NFILEPARTS WHERE NFS_FILEGUID="'+lStFileGUID+'"');
      bErreur := (ExecuteSQL ('DELETE FROM YFILESTD WHERE YFS_FILEGUID="'+lStFileGUID+'"') <> 1);
    finally
      if bErreur then ROLLBACK else COMMITTRANS ;
    end;
  end else Ferme (Q);
  Result := not bErreur;
end;

Initialization
  registerclasses ( [ TOM_STDMAQ ] ) ;
end.

