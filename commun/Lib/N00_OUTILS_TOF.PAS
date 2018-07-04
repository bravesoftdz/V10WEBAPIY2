{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 17/02/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : N00_OUTILS ()
Mots clefs ... : TOF;N00_OUTILS
*****************************************************************}
unit N00_OUTILS_TOF;

interface

uses StdCtrls,
  Controls,
  Classes,
{$IFNDEF EAGLCLIENT}
  db,
  dbtables,
  mul,
{$ELSE}
  eMul,
{$ENDIF}
  uTob,
  uYFILESTD,
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HTB97,
  BobGestion,

{$IFNDEF SANS_NLP}
  NLP_VAR,
  NLP_API,
{$ENDIF}
  HMsgBox,
  UTOF;

type
  TOF_N00_OUTILS = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
  private
    procedure SuppressionBOB(sender: tObject);
    procedure ImportationBOB(sender: tObject);
    procedure ActivationBOB(sender: tObject);
    procedure SuppressionFEB_ORPHELIN(sender: tObject);
    procedure SuppressionFEB(sender: tObject);
    procedure SuppressionParametrage(sender: tObject);
    procedure BValidSuppressionParametrage(sender: tObject);
  end;

implementation

procedure TOF_N00_OUTILS.OnNew;
begin
  inherited;
end;

procedure TOF_N00_OUTILS.OnDelete;
begin
  inherited;
end;

procedure TOF_N00_OUTILS.OnUpdate;
begin
  inherited;
end;

procedure TOF_N00_OUTILS.OnLoad;
begin
  inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PASCAL BASSET
Créé le ...... : 08/03/2006
Modifié le ... :   /  /
Description .. : on passe en argument, sinon on est obligé d'emporter
Suite ........ : V8LIAPGI, etc
Mots clefs ... :
*****************************************************************}

procedure TOF_N00_OUTILS.OnArgument(S: string);
var sSOCREF_PLUS_UN, LEPRODUIT, LEMILLESIME: string;
begin
  inherited;
  sSOCREF_PLUS_UN := '0' + IntToStr(V_PGI.NumVersionSoc + 1);
{$IFNDEF SANS_NLP}
  LEPRODUIT := V_NLP.Nom_Bob_Produit;
  LEMILLESIME := V_NLP.MillesimeDiff;
{$ELSE}
  LEPRODUIT := ReadTokenSt(s);
  LEMILLESIME := ReadTokenSt(s);
{$ENDIF}




  TToolBarButton97(GetControl('BValider')).OnClick := BValidSuppressionParametrage;
  /////////////////////////////
  // ONGLET BOB
  /////////////////////////////
  if assigned(TToolBarButton97(Getcontrol('BSUPBOB')))
    then TToolBarButton97(Getcontrol('BSUPBOB')).OnClick := SuppressionBOB
  else PGIINFO('TOF_N00_OUTILS:BSUPBOB non trouvé');

  if assigned(TToolBarButton97(Getcontrol('BACTBOB')))
    then TToolBarButton97(Getcontrol('BACTBOB')).OnClick := ActivationBOB
  else PGIINFO('TOF_N00_OUTILS:BACTBOB non trouvé');

  if assigned(TToolBarButton97(Getcontrol('BIMPBOB')))
    then TToolBarButton97(Getcontrol('BIMPBOB')).OnClick := ImportationBOB
  else PGIINFO('TOF_N00_OUTILS:BIMPBOB non trouvé');

  if not assigned(TMemo(Getcontrol('BOBMEMO')))
    then PGIINFO('TOF_N00_OUTILS:BOBMEMO non trouvé');

  //Pour Test mais faudra alimenter avec le code produit
  if assigned(THEdit(Getcontrol('THPRODUIT_SUP')))
    then THEdit(Getcontrol('THPRODUIT_SUP')).Text := LEPRODUIT
  else PGIINFO('TOF_N00_OUTILS:THPRODUIT_SUP non trouvé');
  if assigned(THEdit(Getcontrol('THPRODUIT_SUP')))
    then THEdit(Getcontrol('THPRODUIT_SUP')).Text := LEPRODUIT
  else PGIINFO('TOF_N00_OUTILS:THPRODUIT_SUP non trouvé');
  if assigned(THEdit(Getcontrol('THPRODUIT_SUP')))
    then THEdit(Getcontrol('THPRODUIT_SUP')).Text := LEPRODUIT
  else PGIINFO('TOF_N00_OUTILS:THPRODUIT_SUP non trouvé');
  if assigned(THEdit(Getcontrol('THPRODUIT_IMP')))
    then THEdit(Getcontrol('THPRODUIT_IMP')).Text := LEPRODUIT
  else PGIINFO('TOF_N00_OUTILS:THPRODUIT_IMP non trouvé');
  if assigned(THEdit(Getcontrol('THPRODUIT_IMP')))
    then THEdit(Getcontrol('THPRODUIT_IMP')).Text := LEPRODUIT
  else PGIINFO('TOF_N00_OUTILS:THPRODUIT_IMP non trouvé');


  //Pour Test mais faudra alimenter avec le code produit
  if assigned(THEdit(Getcontrol('THSOCREF_SUP')))
    then THEdit(Getcontrol('THSOCREF_SUP')).Text := sSOCREF_PLUS_UN
  else PGIINFO('TOF_N00_OUTILS:THSOCREF_SUP non trouvé');

  //Pour Test mais faudra alimenter avec le code produit
  if assigned(THEdit(Getcontrol('THPRODUIT_ACT')))
    then THEdit(Getcontrol('THPRODUIT_ACT')).Text := THEdit(Getcontrol('THPRODUIT_SUP')).Text
  else PGIINFO('TOF_N00_OUTILS:THPRODUIT_ACT non trouvé');
  //Pour Test mais faudra alimenter avec le code produit
  if assigned(THEdit(Getcontrol('THSOCREF_ACT')))
    then THEdit(Getcontrol('THSOCREF_ACT')).Text := THEdit(Getcontrol('THSOCREF_SUP')).Text
  else PGIINFO('TOF_N00_OUTILS:THSOCREF_ACT non trouvé');

  /////////////////////////////
  // ONGLET FICHIER EN BASE
  /////////////////////////////
  if not assigned(TMemo(Getcontrol('BOBFEB')))
    then PGIINFO('TOF_N00_OUTILS:BOBFEB non trouvé');
  if assigned(TToolBarButton97(Getcontrol('BSUPFEBORPH')))
    then TToolBarButton97(Getcontrol('BSUPFEBORPH')).OnClick := SuppressionFEB_ORPHELIN
  else PGIINFO('TOF_N00_OUTILS:BSUPFEBORPH non trouvé');
//  if assigned(THEdit(Getcontrol('THPRODUIT_PARAMORPH')))
//    then THEdit(Getcontrol('THPRODUIT_PARAMORPH')).Text := 'AAAA'
//    else PGIINFO('TOF_N00_OUTILS:THPRODUIT_PARAMORPH non trouvé');
  if assigned(TToolBarButton97(Getcontrol('BSUPFEB')))
    then TToolBarButton97(Getcontrol('BSUPFEB')).OnClick := SuppressionFEB
  else PGIINFO('TOF_N00_OUTILS:BSUPFEB non trouvé');
  if assigned(THEdit(Getcontrol('THPRODUIT_PARAM')))
    then THEdit(Getcontrol('THPRODUIT_PARAM')).Text := 'AAAA'
  else PGIINFO('TOF_N00_OUTILS:THPRODUIT_PARAM non trouvé');
  if assigned(THEdit(Getcontrol('THCRIT1_PARAM')))
    then THEdit(Getcontrol('THCRIT1_PARAM')).Text := 'AAAA'
  else PGIINFO('TOF_N00_OUTILS:THCRIT1_PARAM non trouvé');
  if assigned(THEdit(Getcontrol('THCRIT2_PARAM')))
    then THEdit(Getcontrol('THCRIT2_PARAM')).Text := 'AAAA'
  else PGIINFO('TOF_N00_OUTILS:THCRIT2_PARAM non trouvé');

  /////////////////////////////
  // ONGLET PARAMETRAGE
  /////////////////////////////
  if assigned(TToolBarButton97(Getcontrol('BPARAM')))
    then TToolBarButton97(Getcontrol('BPARAM')).OnClick := SuppressionParametrage
  else PGIINFO('TOF_N00_OUTILS:BPARAM non trouvé');


  if Assigned(THEdit(Getcontrol('THPRODUIT_PARAM'))) then
{$IFDEF SANS_NLP}
    THEdit(Getcontrol('THPRODUIT_PARAM')).Text := LEPRODUIT
{$ELSE}
    THEdit(Getcontrol('THPRODUIT_PARAM')).Text := V_NLP.Nom_Produit
{$ENDIF SANS_NLP}
  else
    PGIINFO('TOF_N00_OUTILS:THPRODUIT_PARAM non trouvé');



  if assigned(THEdit(Getcontrol('THCRIT1_PARAM')))
    then THEdit(Getcontrol('THCRIT1_PARAM')).Text := LEMILLESIME
  else PGIINFO('TOF_N00_OUTILS:THCRIT1_PARAM non trouvé');
  if assigned(THEdit(Getcontrol('THCRIT2_PARAM')))
    then THEdit(Getcontrol('THCRIT2_PARAM')).Text := 'REQ'
  else PGIINFO('TOF_N00_OUTILS:THCRIT2_PARAM non trouvé');
  if assigned(THEdit(Getcontrol('THFIC_PARAM')))
    then THEdit(Getcontrol('THFIC_PARAM')).Text := LEPRODUIT + LEMILLESIME + 'PARAM.TXT'
  else PGIINFO('TOF_N00_OUTILS:THFIC_PARAM non trouvé');
  if not assigned(THGrid(Getcontrol('HG_PARAM'))) then PGIINFO('TOF_N00_OUTILS:HG_PARAM non trouvé');
  if not assigned(THEdit(Getcontrol('THREQUETE')))
    then PGIINFO('TOF_N00_OUTILS:THREQUETE non trouvé');


end;

procedure TOF_N00_OUTILS.OnClose;
begin
  inherited;
end;

procedure TOF_N00_OUTILS.OnDisplay();
begin
  inherited;
end;

procedure TOF_N00_OUTILS.OnCancel();
begin
  inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PASCAL BASSET
Créé le ...... : 17/02/2006
Modifié le ... :   /  /
Description .. : Suppression des enregistrements dans la table YMYBOBS
Suite ........ : pour un produit et pour ue SOCREF
Mots clefs ... :
*****************************************************************}

procedure TOF_N00_OUTILS.SuppressionBOB(sender: tObject);
var st: string;
  sSQL: string;
  sNom: string;
begin
  st := BOB_GET_CHEMIN + '\' + THEdit(Getcontrol('THPRODUIT_SUP')).Text + '\'
    + THEdit(Getcontrol('THPRODUIT_SUP')).Text + THEdit(Getcontrol('THSOCREF_SUP')).text + '%';
  sNom := THEdit(Getcontrol('THPRODUIT_SUP')).Text + THEdit(Getcontrol('THSOCREF_SUP')).text + '%';
  sSQL := 'DELETE FROM YMYBOBS WHERE YB_BOBNAME LIKE "' + sNom + '%"';
  if pgiask('Confirmez l''exécution de la requete suivante ' + #13#10 + sSQL) = mryes then
  begin
    ExecuteSQL(sSQL);
    TMemo(Getcontrol('BOBMEMO')).Lines.Add('Supression des bobs de type : ' + st);
    TMemo(Getcontrol('BOBMEMO')).Lines.Add(' ');
    TMemo(Getcontrol('BOBMEMO')).Lines.Add('--> Installer KIT ou cliquer sur Activation des BOBs');
    if Copy(sNom, 1, 4) = 'BICP' then
    begin
      sSQL := 'DELETE FROM YMYBOBS WHERE YB_BOBNAME LIKE "COMM%"';
      if pgiask('Confirmez l''exécution de la requete suivante ' + #13#10 + sSQL) = mryes then
      begin
        ExecuteSQL(sSQL);
        TMemo(Getcontrol('BOBMEMO')).Lines.Add('Supression des bobs de type : ' + st);
        TMemo(Getcontrol('BOBMEMO')).Lines.Add(' ');
        TMemo(Getcontrol('BOBMEMO')).Lines.Add('--> Installer KIT ou cliquer sur Activation des BOBs');
      end;
      sSQL := 'DELETE FROM YMYBOBS WHERE YB_BOBNAME LIKE "PLAQ%"';
      if pgiask('Confirmez l''exécution de la requete suivante ' + #13#10 + sSQL) = mryes then
      begin
        ExecuteSQL(sSQL);
        TMemo(Getcontrol('BOBMEMO')).Lines.Add('Supression des bobs de type : ' + st);
        TMemo(Getcontrol('BOBMEMO')).Lines.Add(' ');
        TMemo(Getcontrol('BOBMEMO')).Lines.Add('--> Installer KIT ou cliquer sur Activation des BOBs');
      end;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PASCAL BASSET
Créé le ...... : 17/02/2006
Modifié le ... :   /  /
Description .. : Renomme les bobs _BOB en BOB
Mots clefs ... :
*****************************************************************}

procedure TOF_N00_OUTILS.ActivationBOB(sender: tObject);
var
  sFileBOB, Chemin: string;
  SearchRec: TSearchRec;
  ret: integer;
  Oldname, NewName: string;
  bOK: boolean;
begin
  // LE NOM DES BOB SE COMPOSE DE
  // - Code Produit   XXXX
  // - Num version base 9999
  // - type de BOB (F:fiche,M:Menu,D:data);
  // - Num version BOB 999
  // - extension .BOB
  // - exemple CCS50582F0001.BOB
  TMemo(Getcontrol('BOBMEMO')).Lines.Add('Activation des BOB :');
  Chemin := BOB_GET_CHEMIN + '\' + THEdit(Getcontrol('THPRODUIT_ACT')).Text + '\'; //EX : C:\PGI00\BOB\2035\
  ret := FindFirst(Chemin + '_' + THEdit(Getcontrol('THPRODUIT_ACT')).Text + '*.BOZ', faAnyFile, SearchRec);
  if ret <> 0 then TMemo(Getcontrol('BOBMEMO')).Lines.Add('Aucun BOB renommé avec les critères _'
      + THEdit(Getcontrol('THPRODUIT_ACT')).Text + THEdit(Getcontrol('THSOCREF_ACT')).Text);
  bOK := FALSE;

  while ret = 0 do
  begin
    sFileBOB := SearchRec.Name; //RECUPERE NOM DU BOB  {_25610900A001.BOB}
    if ((Copy(sFileBOB, 2, 4) = THEdit(Getcontrol('THPRODUIT_ACT')).Text)
      and (Copy(sFileBOB, 6, 4) = THEdit(Getcontrol('THSOCREF_ACT')).Text)) then
    begin
      OldName := sFileBOB;
      NewName := OldName;
      System.Delete(NewName, 1, 1);
      RenameFile(Chemin + OldName, Chemin + NewName);
      TMemo(Getcontrol('BOBMEMO')).Lines.Add(OldName + ' renommé en ' + NewName);
      bOK := TRUE;
    end;
    ret := FindNext(SearchRec);
  end;
  Sysutils.FindClose(SearchRec);
  if bOK then
  begin
    TMemo(Getcontrol('BOBMEMO')).Lines.Add(' ');
    TMemo(Getcontrol('BOBMEMO')).Lines.Add('--> Cliquer sur Importation des BOBs');
  end;
end;

procedure TOF_N00_OUTILS.SuppressionFEB_ORPHELIN(sender: tObject);
begin
  TMemo(Getcontrol('BOBFEB')).Lines.Add('Suppressions des orphelins :');
  PGIINFO('SuppressionFEB_ORPHELIN dans N00_OUTILS_TOF');
end;

procedure TOF_N00_OUTILS.SuppressionFEB(sender: tObject);
var
  sSQL: string;
begin
  TMemo(Getcontrol('BOBFEB')).Lines.Add('Suppressions des fichier en base :');
  sSQL := THEdit(Getcontrol('THPRODUIT_SUPFEB')).Text
    + THEdit(Getcontrol('THCRIT1_SUPFEB')).Text
    + THEdit(Getcontrol('THCRIT2_SUPFEB')).Text;
  PGIINFO('SuppressionFEB dans N00_OUTILS_TOF');
end;

procedure TOF_N00_OUTILS.SuppressionParametrage(sender: tObject);
var
  SFileGED: string;
  iCodeGED: integer;
  sProduit: string;
  c1, c2: string;
  sFichier: string;
  TH: THGRID;
  TOB_PARAM: TOB;
begin
//  PGIINFO('DEV04:');
  sFichier := THEdit(Getcontrol('THFIC_PARAM')).Text;
  sProduit := THEdit(Getcontrol('THPRODUIT_PARAM')).Text;
  c1 := THEdit(Getcontrol('THCRIT1_PARAM')).Text;
  c2 := THEdit(Getcontrol('THCRIT2_PARAM')).Text;
  iCodeGED := AGL_YFILESTD_EXTRACT(sFileGED, sProduit, sFichier, c1, c2);
  if iCodeGED <> -1 then
  begin
    PGIINFO(AGL_YFILESTD_GET_ERR(iCodeGED) + #13#10 + sFichier + ' TOF_N00_OUTILSSuppressionParametrage');
    exit;
  end;
  TH := THGrid(Getcontrol('HG_PARAM'));
  TOB_PARAM := TOB.Create('', nil, -1);
  TOB_PARAM.LoadDetailFile(sFileGed);
  TOB_PARAM.PutGridDetail(TH, False, True, '');
  TH.ColWidths[0] := round(TH.Width / 2);
  TH.ColWidths[1] := round(TH.Width / 2);
  TOB_PARAM.free;
  DeleteFile(sFileGED);
  TToolBarButton97(GetControl('BValider')).Visible := TRUE;
  PGIINFO('Cliquer sur Valider pour supprimer le paramétrage');


end;

procedure TOF_N00_OUTILS.BValidSuppressionParametrage(sender: tObject);
var
  TH: THGRID;
  i: integer;
begin
  TH := THGrid(Getcontrol('HG_PARAM'));
  TH.AllSelected := TRUE;
  if pgiask('Lancer la suppression du paramétrage') = mrYes then
  begin
    if pgiask('Confirmez la suppression du paramétrage') = mrYes then
    begin
      for i := 1 to TH.RowCount - 1 do
      begin
        //PGIINFO('exécution de la requete : '+TH.cells[0,i]);
        try
          THEdit(Getcontrol('THREQUETE')).Text := TH.cells[0, i];
          THEdit(Getcontrol('THREQUETE')).Refresh;
          executeSQL(TH.cells[0, i]);
        except
          on E: exception do
          begin
            PGIINFO(TH.cells[0, i] + #13#10 + E.message, 'PROBELEME REQUETE');
          end;
        end;
      end;
      PGIINFO('Suppression du paramétrage terminée, relancer l''application');
    end;
  end;
  TToolBarButton97(GetControl('BValider')).Visible := FALSE;
end;

procedure TOF_N00_OUTILS.ImportationBOB(sender: tObject);
var
  sNomBob: string;
  s: string;
begin
  sNomBob := THEdit(Getcontrol('THPRODUIT_IMP')).Text;
  s := 'Confirmez l''importation des BOB de type ' + sNomBob; ;
  if pgiask(s) = mryes then
  begin
    BOB_IMPORT_PCL(sNomBob, TRUE); //Importation des bobs COMMUN
    BOB_RENAME_BOB(sNomBob, TRUE); //Renommage des bobs COMMUN pour éviter qu'ils s'importent tout le temps
  end;
end;

initialization
  registerclasses([TOF_N00_OUTILS]);
end.

