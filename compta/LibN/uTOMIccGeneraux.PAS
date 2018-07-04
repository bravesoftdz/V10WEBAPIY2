{***********UNITE*************************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 23/02/2001
Modifié le ... : 28/08/2001
Description .. : Source TOM de la TABLE : ICCGENERAUX
Suite ........ : (ICCGENERAUX)
Mots clefs ... : ICCGENERAUX
*****************************************************************}
unit uTOMIccGeneraux;

interface

uses Graphics,
  StdCtrls,
  Controls,
  Classes,
{$IFDEF EAGLCLIENT}
  MaineAGL, // AGLLanceFiche
  UtileAGL,
  eFiche,
  eFichList,
{$ELSE}
  DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  FE_Main, // AGLLanceFiche
  DbCtrls,
  Fiche,
  FichList,
  HDB,
{$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTob,
  UTom,
  Grids,
  Windows,
  Menus,
  Ent1,
  ParamSoc,
  Htb97,
  HPanel,
  IccGlobale,
  uTOFIccParametre,  // CPLanceFiche_ICCPARAMETRE
  uTOFIccTauxCompte, // CPLanceFiche_FicheICCTAUX
  Messages;

procedure CPLanceFiche_FicheICCGENERAUX( vStLeCompte : string = '');
procedure CPLanceFiche_CreateICCGENERAUX;

type
  TOM_ICCGENERAUX = class(TOM)

    ZPersonne : THEdit;

    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure OnCancelRecord; override;
    procedure OnClickBSupprimer(Sender: TObject);
    procedure OnClickBAjouter(Sender: TObject);
    procedure OnClickBCalculer(Sender: TObject);
    procedure OnClickBRecuperation(Sender: TObject);
    procedure OnClickBMiseAJour(Sender: TObject);
    procedure OnClickBTaux(Sender: TObject);

    procedure OnKeyPressZPersonne(Sender: TObject; var Key: Char);

    { Evénement de la grille }
    procedure OnEnterGridEcriture(Sender: TObject);
    procedure OnCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure OnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure OnKeyPressGridEcriture(Sender: TObject; var Key: Char);
    procedure DessineCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
    procedure PostDrawCell(ACol, ARow: LongInt; Canvas: TCanvas; AState: TGridDrawState);
    procedure OnRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure OnElipsisClickGridEcriture(Sender: TObject);
    { Fin des événements de la grille }

    function IsRowValid(TheRow: integer): Boolean;
    procedure AutoEditDataSource;
    procedure BlocageComposant(Sender: TObject);

    procedure UpdateTobAndGrille(Sender: TObject);

    procedure OnPopupImporter(Sender: TObject);

    procedure OnKeyPressICGGeneral(Sender: TObject; var Key: Char);
    procedure OnExitICGGeneral(Sender: TObject);

    procedure OnEnterSoldeAnterieur ( Sender : TObject );
    procedure OnKeyPressSoldeAnterieur(Sender: TObject; var Key: Char);
    procedure OnExitSoldeAnterieur(Sender: TObject);

    procedure CreationEnregistrement(Indice, Ligne: Integer; LaDate: TDateTime);
    procedure OnKeyDownForm(Sender: TObject; var Key: Word; Shift: TShiftState);
    function GetGridSens(ACol, ARow: Integer): Integer;
    procedure LoadDataFromEcriture(Mode: Integer);

    procedure OnStateChange(Sender: TObject);
    procedure CalculDesSoldes;
    procedure ParametrageTypeCompte;
    procedure ParametrageGrille;

    procedure OnElipsisClickZPersonne(Sender : TObject);

  private
    TobMere : Tob;
    GridEcriture: THGrid;

    BValider, BCalculer, BImporter, BTaux, BSupprimer, BAjouter, BInsert, BDelete: TToolBarButton97;
    {$IFDEF EAGLCLIENT}
    CBGroupe, CBLimitation, CBDirigeant: TCheckBox;
    {$ELSE}
    CBGroupe, CBLimitation, CBDirigeant: TDBCheckBox;
    {$ENDIF}
    SoldeAnterieur, SoldeEnDate1, SoldeEnDate2, SoldeTotal: THNumEdit;

    TheKeyPressed: Word;

    FCapital : Boolean;
    FDate1, FDate2: TDateTime;

    procedure EnregistrementEcriture;
    procedure ChargeTobEcriture(vTobEcr: Tob);
    procedure OnAfterFormShowEcran;

  end;

const

  LargeurColonne = 105; // Largeur des colonnes pour les montants

  ColLigne = 0;
  ColLibelle = 1;
  ColDate = 2;
  ColDebit = 3;
  ColCredit = 4;
  ColSolde = 5;

implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPTypeCons,
  {$ENDIF MODENT1}
  uLibWindows,
  LookUp,       // LookUpList
  uLibExercice;

procedure CPLanceFiche_FicheICCGENERAUX( vStLeCompte : string = '');
begin
  if GetParamSocSecur('SO_ICCCOMPTECAPITAL','') = '' then
  begin
    PgiInfo('Vous devez renseigner le compte de capital avant d''utiliser ICC.', 'Paramètre Société');
    Exit;
  end;
  AGLLanceFiche('CP', 'ICCFICHEGENERAUX', '', vStLeCompte, 'ACTION=MODIFICATION');
end;

procedure CPLanceFiche_CreateICCGENERAUX;
begin
  if GetParamSocSecur('SO_ICCCOMPTECAPITAL','') = '' then
  begin
    PgiInfo('Vous devez renseigner le compte de capital avant d''utiliser ICC.', 'Paramètre Société');
    Exit;
  end;
  AGLLanceFiche('CP', 'ICCFICHEGENERAUX', '', '', 'ACTION=CREATION;');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/09/2001
Modifié le ... : 10/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnArgument(S: string);
var
  StArg: string;
  SArgument: string;
begin
  inherited;
  //AglNettoieListesPlus('CPICCGENERAUX', '', nil, False, 'ICG_SOLDEDEBEX;ICG_SOLDEFINEX');

  Ecran.HelpContext := 999999111;
  SetControlVisible('HELPBTN', True);
  Ecran.OnKeyDown := OnKeyDownForm;

  Ecran.Caption := '';
  UpDateCaption(Ecran);
  
{$IFDEF EAGLCLIENT}
//(TFFiche(Ecran)).QFiche. := OnStateChange;
{$ELSE}
  (TFFiche(Ecran)).SFiche.OnStateChange := OnStateChange;
{$ENDIF}

  ParametrageGrille;

  SoldeAnterieur := THNumEdit(GetControl('SOLDEANTERIEUR',True));
  SoldeAnterieur.Masks.NegativeMask := '#,##0.00 C'; // GC - 22/11/2001
  SoldeAnterieur.Masks.PositiveMask := '#,##0.00 D'; // GC - 22/11/2001
  SoldeAnterieur.Masks.ZeroMask := '#,##0.00 C';
  SoldeAnterieur.OnKeyPress := OnKeyPressSoldeAnterieur;
  SoldeAnterieur.OnEnter := OnEnterSoldeAnterieur;
  SoldeAnterieur.OnExit := OnExitSoldeAnterieur;

  SoldeEnDate1 := THNumEdit(GetControl('ZSOLDEENDATE1',True));
  SoldeEnDate1.Masks.NegativeMask := '#,##0.00 C'; // GC - 22/11/2001
  SoldeEnDate1.Masks.PositiveMask := '#,##0.00 D'; // GC - 22/11/2001
  SoldeEnDate1.Masks.ZeroMask := '#,##0.00 C';

  SoldeEnDate2 := THNumEdit(GetControl('ZSOLDEENDATE2',True));
  SoldeEnDate2.Masks.NegativeMask := '#,##0.00 C'; // GC - 22/11/2001
  SoldeEnDate2.Masks.PositiveMask := '#,##0.00 D'; // GC - 22/11/2001
  SoldeEnDate2.Masks.ZeroMask := '#,##0.00 C';

  SoldeTotal := THNumEdit(GetControl('ZTOTAL',True));
  SoldeTotal.Masks.NegativeMask := '#,##0.00 C'; // GC - 22/11/2001
  SoldeTotal.Masks.PositiveMask := '#,##0.00 D'; // GC - 22/11/2001
  SoldeTotal.Masks.ZeroMask := '#,##0.00 C';

  BValider  := TToolBarButton97(GetControl('BValider',True));
  BCalculer := TToolBarButton97(GetControl('BCalculer',True));
  BImporter := TToolBarButton97(GetControl('BImporter',True));
  BTaux     := TToolBarButton97(GetControl('BTaux',True));
  BAjouter  := TToolBarButton97(GetControl('BAjouter',True));
  BSupprimer:= TToolBarButton97(GetControl('BSupprimer',True));
  BInsert   := TToolBarButton97(GetControl('BInsert',True));
  BDelete   := TToolBarButton97(GetControl('BDelete',True));

  BCalculer.OnClick := OnClickBCalculer;
  BTaux.OnClick := OnClickBTaux;
  BAjouter.OnClick := OnClickBAjouter;
  BSupprimer.OnClick := OnClickBSupprimer;

  // GCO - 13/03/2006
  ZPersonne := THEdit(GetControl('ZPERSONNE', True));
  ZPersonne.OnKeyPress     := OnKeyPressZPersonne;
  ZPersonne.OnElipsisClick := OnElipsisClickZPersonne;
  // FIN GCO

  SetControlVisible('ZLPERSONNE',(CtxPcl in V_Pgi.PgiContexte));
  SetControlVisible('ZPERSONNE',(CtxPcl in V_Pgi.PgiContexte));
  // GCO - 18/01/2004 - FQ 10865
  SetControlVisible('BDELETE', False);
  SetControlProperty('ICG_GENERAL', 'MAXLENGTH', VH^.Cpta[fbGene].Lg);

  {$IFDEF EAGLCLIENT}
  CBGroupe := TCheckBox(GetControl('ICG_GROUPE',True));
  CBLimitation := TCheckBox(GetControl('ICG_LIMITATION',True));
  CBDirigeant := TCheckBox(GetControl('ICG_DIRIGEANT',True));
  {$ELSE}
  CBGroupe := TDBCheckBox(GetControl('ICG_GROUPE',True));
  CBLimitation := TDBCheckBox(GetControl('ICG_LIMITATION',True));
  CBDirigeant := TDBCheckBox(GetControl('ICG_DIRIGEANT',True));
  {$ENDIF}

  CBGroupe.OnClick := BlocageComposant;
  CBLimitation.OnClick := BlocageComposant;
  CBDirigeant.OnClick := BlocageComposant;

  TPopupMenu(GetControl('POPUPMENU',True)).Items[0].OnClick := OnClickBAjouter;
  TPopupMenu(GetControl('POPUPMENU',True)).Items[1].OnClick := OnClickBSupprimer;

  { Activation ou Désactivation des menus }
  TPopupMenu(GetControl('POPUPIMPORTER',True)).OnPopup := OnPopupImporter;

  THEdit(GetControl('ICG_GENERAL',True)).OnExit := OnExitICGGeneral;
  THEdit(GetControl('ICG_GENERAL',True)).OnKeyPress := OnKeyPressICGGeneral;
{$IFNDEF EAGLCLIENT}
  THDBEdit(GetControl('ICG_TXANNUEL',True)).DisplayFormat := '#,##0.00 %';
{$ELSE}

{$ENDIF}

  { Lecture des paramètres du OnArgument }
  StArg := S;
  SArgument := ReadTokenSt(StArg); // ACTION

  // Détermination des Dates sur lesquelles, on doit charger les écritures
  // Lecture de paramètre qui indique la date de début
  SArgument := ReadTokenST(StArg);
  if SArgument = '' then
    FDate1 := IIF( CtxPCL in V_PGI.PgiContexte, VH^.CPExoRef.Deb, VH^.Encours.Deb )
  else
    FDate1 := StrToDate(SArgument);

  // Lecture de paramètre qui indique la date de fin
  SArgument := ReadTokenST(StArg);
  if SArgument = '' then
    FDate2 := IIF( CtxPCL in V_PGI.PgiContexte, VH^.CPExoRef.Fin, VH^.Encours.Fin )
  else
    FDate2 := StrToDate(SArgument);

  { Calcul du taux maximum légal }
  SetControlText('ZTXMAX',StrfMontant(CalculTauxMaxLegal( FDate1, FDate2 ), 15, V_PGI.OkDecV, '', True) + ' %');

  TobMere := Tob.Create('Tob Mere', nil, -1);

  if TFFiche(Ecran).FLequel = GetParamSocSecur('SO_ICCCOMPTECAPITAL','') then
  begin
    FCapital := True;
    ParametrageTypeCompte;
  end;

  TFFiche(Ecran).OnAfterFormShow := OnAfterFormShowEcran;
  // GCO - 13/01/2005 - FQ 14956
  THLabel(GetControl('DATEFINEXERCICE', True)).Caption := 'au ' + DateToStr(FDate2);

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/09/2001
Modifié le ... : 10/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnLoadRecord;
var Q: TQuery;
    lStCodePersonne : string;
begin
  inherited;
  CExoRefOuvert( True );

  // Affichage de la personne associé au compte
  lStCodePersonne := GetField('ICG_GUIDPER');
  if lStCodePersonne <> '' then
  begin
    Q := nil;
    try
      Q := OpenSQL('SELECT ANN_GUIDPER, ANN_NOM1,ANN_NOM2,ANN_CVA FROM ANNUAIRE WHERE ANN_GUIDPER = "' + lStCodePersonne + '"', True);
      SetControlText('ZPERSONNE',IIF(not Q.Eof, Trim(Trim(Q.FindField('ANN_CVA').AsString) + ' ' +
        Trim(Q.FindField('ANN_NOM2').AsString) + ' ' +
        Trim(Q.FindField('ANN_NOM1').AsString)), ''));
    finally
      if Assigned(Q) then Ferme(Q);
    end;
  end
  else
    SetControlText('ZPERSONNE', '');

  // Chargement du solde antérieur
  SoldeAnterieur.Value := GetField('ICG_SOLDEDEBEX');

  // Chargement des écritures du compte ICC
  Q := nil;
  try
    Q := OpenSQL('SELECT ICE_GENERAL, ICE_LIGNE, ICE_LIBELLE, ICE_DATEVALEUR, ICE_DEBIT, ICE_CREDIT FROM ICCECRITURE WHERE ' +
      '(ICE_GENERAL = "' + GetField('ICG_GENERAL') + '") AND ' +
      '(ICE_DATEVALEUR >= "' + USDateTime(FDate1) + '") AND ' +
      '(ICE_DATEVALEUR <= "' + USDateTime(FDate2) + '") ORDER BY ICE_DATEVALEUR', True);

    TobMere.ClearDetail;
    TobMere.LoadDetailDB('ICCECRITURE', '', '', Q, True, False);
    UpdateTobAndGrille(nil);

  finally
    if Assigned(Q) then Ferme(Q);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/09/2001
Modifié le ... : 10/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.ParametrageGrille;
begin
  GridEcriture := THGrid(GetControl('GRIDECRITURE',True));
  GridEcriture.PostDrawCell := PostDrawCell;
  GridEcriture.OnRowExit := OnRowExit;
  GridEcriture.OnCellEnter := OnCellEnter;
  GridEcriture.OnKeyPress := OnKeyPressGridEcriture;
  GridEcriture.OnCellExit := OnCellExit;
  GridEcriture.GetCellCanvas := DessineCell;
  GridEcriture.OnEnter := OnEnterGridEcriture;
  GridEcriture.OnElipsisClick := OnElipsisClickGridEcriture;

  GridEcriture.ColCount := ColSolde + 1;
  GridEcriture.FixedCols := 1;
  GridEcriture.FixedRows := 1;

  GridEcriture.Cells[ColLigne, 0] := TraduireMemoire('N°');
  GridEcriture.ColAligns[ColLigne] := TaCenter;
  GridEcriture.ColWidths[ColLigne] := 40;

  GridEcriture.Cells[ColLibelle, 0] := TraduireMemoire('Libellé');
  GridEcriture.ColAligns[ColLibelle] := TaLeftJustify;
  GridEcriture.ColWidths[ColLibelle] := 247;

  GridEcriture.Cells[ColDate, 0] := TraduireMemoire('Date');
  GridEcriture.ColTypes[ColDate] := 'D';
  GridEcriture.ColFormats[Coldate] := ShortDateFormat;
  GridEcriture.ColAligns[ColDate] := TaCenter;
  GridEcriture.ColWidths[ColDate] := 75;

  GridEcriture.Cells[ColDebit, 0] := TraduireMemoire('Débit');
  GridEcriture.ColTypes[ColDebit] := 'F';
  GridEcriture.ColFormats[ColDebit] := '#,##0.00';
  GridEcriture.ColAligns[ColDebit] := TaRightJustify;
  GridEcriture.ColWidths[ColDebit] := LargeurColonne;

  GridEcriture.Cells[ColCredit, 0] := TraduireMemoire('Crédit');
  GridEcriture.ColTypes[ColCredit] := 'F';
  GridEcriture.ColFormats[ColCredit] := '#,##0.00';
  GridEcriture.ColAligns[ColCredit] := TaRightJustify;
  GridEcriture.ColWidths[ColCredit] := LargeurColonne;

  GridEcriture.Cells[ColSolde, 0] := TraduireMemoire('Solde');
  GridEcriture.ColAligns[ColSolde] := TaRightJustify;
  GridEcriture.ColWidths[ColSolde] := LargeurColonne;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/09/2001
Modifié le ... : 10/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.ParametrageTypeCompte;
begin
  // Réorganise l'interface de l' ecran si on ouvre le Compte de CAPITAL
  SetControlVisible('TICG_TXANNUEL',False);
  SetControlVisible('ICG_TXANNUEL',False);

  SetControlVisible('TICG_TXMAX',False);
  SetControlVisible('ZTXMAX',False);
  SetControlVisible('DATEFINEXERCICE', False);

  CBGroupe.Visible := False;
  CBLimitation.Visible := False;
  CBDirigeant.Visible := False;

  BTaux.Visible := False;
  BInsert.Visible := False;
  BDelete.Visible := False;
  BCalculer.Visible := False;

  SetControlVisible('TICG_DATEEFFET',False);
  SetControlVisible('ICG_DATEEFFET',False);

  SetControlVisible('ZLPERSONNE',False);
  SetControlVisible('ZPERSONNE',False);

  // Remonte les composants
  THLabel(GetControl('TICG_SOLDEDEBEX')).Top := THLabel(GetControl('TICG_TXANNUEL')).Top;
  {$IFDEF EAGLCLIENT}
  SoldeAnterieur.Top := THEdit(GetControl('ICG_TXANNUEL')).Top;
  {$ELSE}
  SoldeAnterieur.Top := THDBEdit(GetControl('ICG_TXANNUEL')).Top;
  {$ENDIF}

  // Solde au __/__/____ numéro 1
  THLabel(GetControl('LSOLDEENDATE1')).Top := CBGroupe.Top;
  SoldeEnDate1.Top := CBGroupe.Top;

  // Solde au __/__/____ numéro 2
  THLabel(GetControl('LSOLDEENDATE2')).Top := CBGroupe.Top;
  SoldeEnDate2.Top := CBGroupe.Top;

  THPanel(GetControl('PTOP')).Height := 90;

  // Recherche du compte du compte de capital
  if not ExisteSQL('SELECT ICG_GENERAL FROM ICCGENERAUX WHERE ICG_GENERAL = "' + GetParamSocSecur('SO_ICCCOMPTECAPITAL','') + '"') then
  begin // Création automatique du compte de capital
    try
      BeginTrans;
      ExecuteSql('INSERT INTO ICCGENERAUX(ICG_GENERAL,ICG_TXANNUEL,ICG_GROUPE,ICG_LIMITATION,ICG_DIRIGEANT,ICG_DATEEFFET,ICG_SOUMISIS,ICG_GUIDPER,ICG_SOLDEDEBEX,ICG_SOLDEFINEX) ' +
        'Values ("' + GetParamSocSecur('SO_ICCCOMPTECAPITAL','') + '",0,"-","-","X","' + UsDateTime(iDate1900) + '","-","",0,0)');
      CommitTrans;
    except
      RollBack;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/09/2001
Modifié le ... : 11/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.BlocageComposant(Sender: TObject);
var
  bEnable : Boolean;
begin
  CBLimitation.Enabled := (not CBGroupe.Checked);
  CBGroupe.Enabled := (not CbLimitation.Checked);

  bEnable := not CBGroupe.Checked;
  SetControlEnabled('ZPERSONNE',bEnable);
  if (not bEnable) then
    SetControlProperty('ZPERSONNE','COLOR',clBtnFace)
  else
    SetControlProperty('ZPERSONNE','COLOR',clWindow);

  if (not CBLimitation.Checked) then SetField('ICG_DIRIGEANT', '-');

  if CtxPCL in V_PGi.PgiContexte then
  begin
    CBDirigeant.Enabled := CBLimitation.Checked and Icc_Data.SoumisIs;
    if not Icc_Data.SoumisIs then
    begin
      if (GetField('ICG_DIRIGEANT') = 'X') then
      // On décoche le Champ ICG_DIRIGEANT car la société n'est plus soumise à l'IS
      begin
        DS.Edit;
        SetField('ICG_DIRIGEANT', '-');
      end;
    end;
  end
  else
  begin
    CBDirigeant.Enabled := CBLimitation.Checked;
  end;

  bEnable := (CBDirigeant.Checked);
  SetControlEnabled('ICG_DATEEFFET',bEnable);
  if bEnable then
    SetControlProperty('ICG_DATEEFFET','COLOR',clWindow)
  else
    SetControlProperty('ICG_DATEEFFET','COLOR',clBtnFace);

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/09/2001
Modifié le ... : 10/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnKeyPressZPersonne(Sender: TObject; var Key: Char);
begin
  if Key = #8 then
  begin
    SetControlText('ZPERSONNE','');
    SetField('ICG_GUIDPER', '');
  end
  else
    Key := #0;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/09/2001
Modifié le ... : 10/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnChangeField(F: TField);
var
  SQL: TQuery;
begin
  inherited;
  // Recherche le libelle du compte dans la table des généraux de la compta
  if (F.FieldName = 'ICG_GENERAL') then
  begin
    Ecran.Caption := IIF(FCapital, 'Compte de capital : ', 'Compte à rémunérer : ');

    SQL := nil;
    try
      SQL := OpenSQL('SELECT G_GENERAL, G_LIBELLE FROM GENERAUX WHERE G_GENERAL = "' + GetField('ICG_GENERAL') + '"', True);
      if not SQL.EOF then
      begin
        Ecran.Caption := Ecran.Caption + GetField('ICG_GENERAL') + ' - ' + SQL.FindField('G_LIBELLE').AsString;
        SetControlText('ZLIBELLE',SQL.FindField('G_LIBELLE').AsString);
      end;

    finally
      if Assigned(SQL) then Ferme(SQL);
      UpDateCaption(Ecran);
    end;
  end;
end;

procedure TOM_ICCGENERAUX.OnExitICGGeneral(Sender: TObject);
begin
  if Trim(GetControlText('ICG_GENERAL')) <> '' then
  begin
    if Length(GetControlText('ICG_GENERAL')) > VH^.Cpta[FbGene].Lg then
      SetField('ICG_GENERAL', Copy(GetField('ICG_GENERAL'), 1, 6))
    else
      SetField('ICG_GENERAL', BourreLaDonc(GetField('ICG_GENERAL'), FbGene));
  end;
end;

{==============================================================================}
{================== AJOUT, MISE A JOUR, SUPPRESSION DU COMPTE =================}
{==============================================================================}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/09/2001
Modifié le ... : 19/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnNewRecord;
begin
  inherited;
  Ecran.Caption := 'Compte à rémunérer : ';
  UpDateCaption(Ecran);

  SetControlText('ZLIBELLE', '');
  SetField('ICG_DATEEFFET', FDate1);
  SetField('ICG_LIMITATION', '-');
  SetField('ICG_DIRIGEANT', '-');
  SetField('ICG_GROUPE', '-');
  {$IFDEF EAGLCLIENT}
  Ecran.ActiveControl := THEdit(GetControl('ICG_GENERAL'));
  {$ELSE}
  Ecran.ActiveControl := THDBEdit(GetControl('ICG_GENERAL'));
  {$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/09/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnDeleteRecord;
begin
  if not BlocageMonoPoste(True) then
    Exit;
  try
    // Efface les écritures associées au compte supprimé
    try
      BeginTrans;
      ExecuteSQL('DELETE FROM ICCECRITURE WHERE ICE_GENERAL = "' + GetField('ICG_GENERAL') + '"');
      CommitTrans;
    except
      RollBack;
    end;
  finally
    DeblocageMonoPoste(True);
  end;
  inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/09/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnStateChange(Sender: TObject);
begin
  inherited;
  BImporter.Enabled := (Ds.State = DsBrowse);
  BCalculer.Enabled := BImporter.Enabled;
end;

procedure TOM_ICCGENERAUX.OnCancelRecord;
begin
  inherited;
end;

{==============================================================================}
{========================= GESTION DE LA TOB DES ECRITURES ====================}
{==============================================================================}
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/10/2001
Modifié le ... :   /  /
Description .. : Evénement de l' entrée en grille
Mots clefs ... : OnEnter
*****************************************************************}
procedure TOM_ICCGENERAUX.OnEnterGridEcriture(Sender: TObject);
begin

  if TobMere.Detail.Count = 0 then
  begin
    CreationEnregistrement(-1, 1, FDate1);
    UpdateTobAndGrille(nil);

    //GridEcriture.SetFocus;
    //GridEcriture.EditorMode := True;
  end;

  //  Ecran.ActiveControl := THGrid(GetControl('GridEcriture'));

  //  GridEcriture.Col := ColLibelle;
  //  GridEcriture.SetFocus;
   // GridEcriture.EditorMode := True;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/09/2001
Modifié le ... : 19/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.CreationEnregistrement(Indice, Ligne: Integer; LaDate: TDateTime);
var
  TobFille: TOB;
begin
  AutoEditDataSource;
  TobFille := Tob.Create('ICCECRITURE', TobMere, Indice);
  TobFille.PutValue('ICE_LIGNE', Ligne);
  TobFille.PutValue('ICE_DEBIT', 0);
  TobFille.PutValue('ICE_CREDIT', 0);
  TobFille.PutValue('ICE_DATEVALEUR', LaDate);

  GridEcriture.SetFocus;
  GridEcriture.Col := ColLibelle;
  GridEcriture.EditorMode := True;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/09/2001
Modifié le ... : 19/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnClickBSupprimer(Sender: TObject);
var
  i: integer;
  NumLigne: Integer;
begin
  if (TobMere.Detail.Count = 0) or
    ((TobMere.Detail.Count = 1) and (not IsRowValid(GridEcriture.Row))) then
    Exit;

  GridEcriture.EditorMode := False;
  // Effacement par rapport au champ ICE_LIGNE de l'enregistement pointé
  Numligne := StrToInt(GridEcriture.Cells[0, GridEcriture.Row]);
  TobMere.Detail[NumLigne - 1].Free;

  for i := NumLigne to TobMere.Detail.Count do
    TobMere.Detail[i - 1].PutValue('ICE_LIGNE', TobMere.Detail[i - 1].GetValue('ICE_LIGNE') - 1);

  if TobMere.Detail.Count = 0 then
    CreationEnregistrement(-1, 1, FDate1);

  UpdateTobAndGrille(nil);
  AutoEditDataSource;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/09/2001
Modifié le ... : 19/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnClickBAjouter(Sender: TObject);
var
  i: integer;
begin
  if TobMere.Detail.Count = 0 then
    CreationEnregistrement(-1, 1, FDate1)
  else
    CreationEnregistrement(GridEcriture.Row - 1, GridECriture.Row - 1, StrToDate(GridEcriture.Cells[2, GridEcriture.Row]));

  for i := TobMere.Detail[GridEcriture.Row - 1].GetValue('ICE_LIGNE') to TobMere.Detail.Count - 1 do
    TobMere.Detail[i].PutValue('ICE_LIGNE', TobMere.Detail[i].GetValue('ICE_LIGNE') + 1);

  UpdateTobAndGrille(nil);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/09/2001
Modifié le ... : 19/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.CalculDesSoldes;
var
  Q: TQuery;
  Solde, TotalDebit, TotalCredit : Double;
  i: integer;
  lSoldeComptable : Double;
begin
  TotalDebit := 0;
  TotalCredit := 0;

  Solde := GetField('icg_soldedebex');

  Q := nil;
  try
    Q := OpenSQL('Select Sum(ice_debit - ice_credit) as solde from iccecriture where (ice_general = "' + GetField('icg_general') + '") and ' +
      '(ice_datevaleur < "' + USDateTime(FDate1) + '")', True);

    if not Q.EOF then
      Solde := Solde + Q.FindField('Solde').AsFloat;
  finally
    if Assigned(Q) then Ferme(Q);
  end;

  // Affichage du solde en datedu
  SetControlText('LSOLDEENDATE1','Solde au ' + DateToStr(FDate1));
  SoldeEnDate1.Value := Solde;

  // Affichage du solde antérieur ( soit Fdate1 -1 )
  SetControlText('TICG_SOLDEDEBEX', 'Solde au ' + DateToStr(FDate1 - 1));

  // Calcul du solde progressif dans la grille des ecritures
  for i := 0 to TobMere.Detail.Count - 1 do
  begin
    TotalDebit := TotalDebit + TobMere.Detail[i].GetValue('ICE_DEBIT');
    TotalCredit := TotalCredit + TobMere.Detail[i].GetValue('ICE_CREDIT');
    GridEcriture.Cells[ColSolde, i + 1] := AfficheDBCR(Solde + TotalDebit - TotalCredit);
    GridEcriture.Cells[ColLigne, i + 1] := IntToStr(I+1);
  end;

  SetControlText('ZTotalDebit',StrfMontant(TotalDebit, 15, V_PGI.OkDecV, '', True));
  SetControlText('ZTotalCredit',StrfMontant(TotalCredit, 15, V_PGI.OkDecV, '', True));

  SoldeTotal.Value := Solde + TotalDebit - TotalCredit;
  SetControlText('LSOLDEENDATE2','Solde au ' + DateToStr(FDate2));
  SoldeEnDate2.Text := GetControlText('ZTotal');

  SetControlText('LSoldeDifferent','Le solde est différent de celui du compte général');
  lSoldeComptable := GetSoldeComptable(GetField('ICG_GENERAL'), FDate1, FDate2);

  // GCO - 07/09/2006 - FQ 18354
  // Arrondi a deux chiffres pour comparer les soldes.
  Icc_Data.SoldeDifferent := Arrondi(SoldeTotal.Value,2) <> Arrondi(lSoldeComptable,2);
  SetControlVisible('LSoldeDifferent', Icc_Data.SoldeDifferent);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/09/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.UpdateTobAndGrille(Sender: TObject);
var
  i, j: integer;
begin
  // Inutile apparement du fait du GRidEcriture.RowCount qui est après
  //GridEcriture.Row := GridEcriture.FixedRows;
  //GridEcriture.Col := GridEcriture.FixedCols;
  //GridEcriture.VidePile(False);

  if TobMere.Detail.Count = 0 then
    GridEcriture.RowCount := 2
  else
    GridEcriture.RowCount := TobMere.Detail.Count + 1;

  for i := 1 to TobMere.Detail.Count + 1 do
    for j := 0 to GridEcriture.Colcount - 1 do
      GridEcriture.Cells[j, i] := '';

  //TobMere.PutGridDetail(GridEcriture, False, False, 'ICE_LIGNE; ICE_LIBELLE; ICE_DATEVALEUR; ICE_DEBIT; ICE_CREDIT');
  TobMere.PutGridDetail(GridEcriture, False, False, ''';ICE_LIBELLE; ICE_DATEVALEUR; ICE_DEBIT; ICE_CREDIT');
  CalculDesSoldes;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/09/2001
Modifié le ... : 06/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnKeyPressGridEcriture(Sender: TObject; var Key: Char);
begin
  // Simuler une modif dans la table.....
  if not (Key in [#9, #13]) then
    AutoEditDataSource;

  case GridEcriture.Col of

    ColDate : if key = #32 then
    begin
      Key := #0;
      V_PGI.ParamDateproc(GridEcriture);
    end;

    // Contrôle des montants
    ColDebit: if TobMere.Detail[GridEcriture.Row - 1].GetValue('ICE_CREDIT') > 0.001 then Key := #0;

    ColCredit: if TobMere.Detail[GridEcriture.Row - 1].GetValue('ICE_DEBIT') > 0.001 then Key := #0;
  end;
end;

procedure TOM_ICCGENERAUX.AutoEditDataSource;
begin
  if DS.State = DsBrowse then
    DS.Edit;
  SetField('ICG_TXANNUEL', GetField('ICG_TXANNUEL'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 06/09/2001
Modifié le ... : 06/09/2001
Description .. : Gestion du grisage des colonnes de la grille
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.PostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var
  TheRect: TRect;
begin
  if ((GridEcriture.Cells[ColDebit, ARow] = '0.00') and
    (GridEcriture.Cells[ColCredit, ARow] = '0.00')) then
    Exit;

  if (ACol = ColDebit) or (ACol = ColCredit) then
  begin
    if GridEcriture.Cells[ACol, ARow] = '0.00' then
    begin
      TheRect := GridEcriture.CellRect(ACol, ARow);
      Canvas.TextRect(TheRect, TheRect.Left, TheRect.Top, '');
      //Canvas.Brush.Color := ClRed;
      //Canvas.FrameRect(TheRect); // Dessine une bordure autour de la cellule avec la couleur du Brush
      Canvas.Brush.Color := GridEcriture.FixedColor;
      Canvas.Brush.Style := bsBDiagonal;
      Canvas.Pen.Color := GridEcriture.FixedColor;
      Canvas.Pen.Mode := pmCopy;
      Canvas.Pen.Style := psClear;
      Canvas.Pen.Width := 1;
      Canvas.Rectangle(TheRect); // Afficher les rayures
    end;
  end;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/09/2001
Modifié le ... : 06/09/2001
Description .. :
Mots clefs ... :                            ColSolde
*****************************************************************}
procedure TOM_ICCGENERAUX.DessineCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
  if ARow = 0 then
    Exit;
  case ACol of
    //ColLigne: Canvas.Font.Style := [fsBold];
    ColSolde: Canvas.Font.Color := IIF(Pos('C', GridEcriture.Cells[ACol, Arow]) > 0, ClGreen, ClRed);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JLD
Créé le ...... : 06/09/2001
Modifié le ... : 06/09/2001
Description .. : Détermine le sens de déplacement dans la grille
Mots clefs ... :
*****************************************************************}
function TOM_ICCGENERAUX.GetGridSens(ACol, ARow: Integer): Integer;
begin
  // Sens de déplacement dans le Grid
  if (GridEcriture.Row = ARow) then
  begin
    if (GridEcriture.Col > ACol) then
      Result := 1
    else
      Result := -1;
  end
  else if (GridEcriture.Row > ARow) then
    Result := 1
  else
    Result := -1;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/09/2001
Modifié le ... : 06/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
  Sens: Integer;
begin
  // Recherche le sens de fonctionnement dans la grille
  Sens := GetGridSens(ACol, ARow);

  ACol := GridEcriture.Col;
  ARow := GridEcriture.Row;
  GridEcriture.ElipsisButton := (ACol = ColDate);

  (*
  if (ACol = ColDebit) and (Valeur(GridEcriture.Cells[ColCredit,ARow]) <> 0) then
  begin
    ACol := ColCredit;
    Cancel := True;
  end;

  if (ACol = ColCredit) and (Valeur(GridEcriture.Cells[ColDebit,ARow]) <> 0) then
  begin
    ACol := ColSolde;
    Cancel := True;
  end;*)

  if (ACol = ColSolde) then
  begin
    if (ARow <> GridEcriture.RowCount - 1) then
    begin
      ARow := ARow + 1;
      ACol := ColLibelle;
      Cancel := True;
    end
    else if (Sens = 1) and (isRowValid(ARow)) and (ARow = GridEcriture.RowCount - 1) then
    begin
      // On est sur la dernière ligne de la grille donc ==> création
      CreationEnregistrement(-1, GridEcriture.Row + 1, StrToDate(GridEcriture.Cells[2, GridEcriture.Row]));
      UpdateTobAndGrille(nil);
      GridEcriture.Row := GridEcriture.Row + 1;
      GridEcriture.EditorMode := True;
    end
    else
    begin
      ACol := ColLibelle;
      Cancel := True;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 05/09/2001
Modifié le ... : 05/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
  DateSaisie, DateSuivante, DatePrecedente: TDateTime;
begin
  case ACol of
    ColLibelle: TobMere.Detail[ARow - 1].PutValue('ICE_LIBELLE', GridEcriture.Cells[ACol, ARow]);

    ColDate:
      begin
        if Trim(GridEcriture.Cells[ACol, ARow]) = '' then
          GridEcriture.Cells[ACol, ARow] := DateToStr(Date)
        else
        begin
          // Verifie la date saisie
          if IsValidDate(GridEcriture.Cells[ACol, ARow]) then
          begin
            DateSaisie := StrToDate(GridEcriture.Cells[ACol, ARow]);

            // La date est elle comprise entre DateDebExo et DateFinExo
            if not ((DateSaisie >= FDate1) and (DateSaisie <= FDate2)) then
            begin
              PgiInfo('La date doit être comprise entre le ' + DateToStr(FDate1) + ' et le ' + DateToStr(FDate2) + '.', 'Saisie de la date');
              Cancel := True;
              Exit;
            end;

            // La date est elle comprise entre DateLignePrecedente et DateLigneSuivante
            if (ARow = 1) then
              DatePrecedente := FDate1
            else
              DatePrecedente := StrToDate(GridEcriture.Cells[ACol, ARow - 1]);

            if (ARow = TobMere.Detail.Count) then
              DateSuivante := FDate2
            else
              DateSuivante := StrToDate(GridEcriture.Cells[ACol, ARow + 1]);

            if not ((DateSaisie >= DatePrecedente) and (DateSaisie <= DateSuivante)) then
            begin
              PgiInfo('La date doit être comprise entre le ' + DateToStr(DatePrecedente) + ' et le ' + DateToStr(DateSuivante), 'Saisie de la date');
              Cancel := True;
              Exit;
            end;

            // La date est ok, on l' enregistre dans la TOB
            TobMere.Detail[ARow - 1].PutValue('ICE_DATEVALEUR', DateSaisie);
          end
          else
          begin
            PgiInfo('Le format de la date est incorrect', 'Saisie de la date');
            Cancel := True;
          end;
        end;
      end;

    // Colonne Debit et Credit
    ColDebit, ColCredit:
      begin
        if Trim(GridEcriture.Cells[ACol, ARow]) = '' then
        begin
          GridEcriture.Cells[ACol, ARow] := '0';
          Cancel := True;
          Exit;
        end;

        if not IsNumeric(GridEcriture.Cells[ACol, ARow]) then
        begin
          Cancel := True;
          Exit;
        end;

        if ACol = ColDebit then
          TobMere.Detail[ARow - 1].PutValue('ICE_DEBIT', Valeur(GridEcriture.Cells[ACol, ARow]))
        else
          TobMere.Detail[ARow - 1].PutValue('ICE_CREDIT', Valeur(GridEcriture.Cells[ACol, ARow]));

        GridEcriture.Cells[ACol, ARow] := StrfMontant(Valeur(GridEcriture.Cells[ACol, ARow]), 20, 2, '', True);
        CalculdesSoldes;
      end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/09/2001
Modifié le ... : 06/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  Cancel := not IsRowValid(Ou);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/09/2001
Modifié le ... : 06/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_ICCGENERAUX.IsRowValid(TheRow: Integer): Boolean;
begin
  Result := (Valeur(GridEcriture.Cells[ColDebit, TheRow]) <> 0) or (Valeur(GridEcriture.Cells[ColCredit, TheRow]) <> 0);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/09/2001
Modifié le ... : 06/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnElipsisClickGridEcriture(Sender: TObject);
begin
  V_PGI.ParamDateproc(GridEcriture);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/09/2001
Modifié le ... : 20/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnKeyDownForm(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  TheKeyPressed := Key;

  case Key of
    VK_F6: if BCalculer.Enabled then
        BCalculer.Click;

    VK_F7: if GridEcriture.Row > 1 then
           begin
             AutoEditDataSource;
             GridEcriture.Cells[GridEcriture.Col, GridEcriture.Row] := GridEcriture.Cells[GridEcriture.Col, GridEcriture.Row - 1];
           end;

    VK_F10: BValider.Click;

    VK_F11: TPopUpMenu(GetControl('POPUPMENU')).Popup(Mouse.CursorPos.x, Mouse.CursorPos.y);

    VK_F12: if THGrid(GetControl('GRIDECRITURE')).Focused then
              SetFocusControl('ICG_TXANNUEL')
            else
              SetFocusControl('GRIDECRITURE');

    VK_INSERT : //if THGrid(GetControl('GRIDECRITURE')).Focused then
                  OnClickBAjouter(nil);
                //else
                //  BInsert.Click;

    VK_DELETE : if Shift = [ssCtrl] then OnClickBSupprimer(nil);

    VK_DOWN:
      begin
        if (GridEcriture.Row = GridEcriture.RowCount - 1) and isRowValid(GridEcriture.Row) then
        begin
          CreationEnregistrement(-1, GridEcriture.Row + 1, StrToDate(GridEcriture.Cells[2, GridEcriture.Row]));
          UpdateTobAndGrille(nil);
          GridEcriture.Row := GridEcriture.Row + 1;
          GridEcriture.EditorMode := True;
        end;
      end;

    VK_ESCAPE:
      begin
        GridEcriture.CacheEdit;
        Key := 0;
      end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/09/2001
Modifié le ... : 06/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnClickBTaux(Sender: TObject);
begin
  CPLanceFiche_FicheICCTAUX(GetField('ICG_GENERAL'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/09/2001
Modifié le ... : 06/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnClickBCalculer(Sender: TObject);
begin
  CPLanceFiche_ICCPARAMETRE(DateToStr(FDate1) + ';' + DateToStr(FDate2) + ';' + GetField('ICG_GENERAL'));
end;

{******************************************************************************}
                        { IMPORTATION DES ECRITURES }
{******************************************************************************}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/09/2001
Modifié le ... : 20/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnClickBRecuperation(Sender: TObject);
begin
  LoadDataFromEcriture(0);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/09/2001
Modifié le ... : 20/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnClickBMiseAJour(Sender: TObject);
begin
  LoadDataFromEcriture(1);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/09/2001
Modifié le ... : 20/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.LoadDataFromEcriture(Mode: Integer);
begin
  if not BlocageMonoPoste(True) then
    Exit;

  try
    case Mode of
      0: if PgiAsk('Confirmez vous la récupération complète des écritures de la comptabilité  ?', Ecran.Caption) = MrYes then
        begin
          BeginTrans;
          if TraitementEcriture(GetField('ICG_GENERAL'), FDate1, FDate2, False) then
          begin
            CommitTrans;
            PgiInfo('Le traitement s''est correctement terminé. Aucune anomalie n''a été détectée.', 'Récupération des nouvelles écritures');
          end
          else
          begin
            RollBack;
            //PgiInfo(E.Message, 'Erreur pendant la récupération complète des écritures de la comptabilité.');
          end;
        end;

      1: if PgiAsk('Confirmez vous la mise à jour du compte avec les nouvelles écritures de la comptabilité ?', 'Importation des écritures') = MrYes then
        begin
          BeginTrans;
          if TraitementEcriture(GetField('ICG_GENERAL'), FDate1, FDate2, True) then
          begin
            CommitTrans;
            PgiInfo('Le traitement s''est correctement terminé. Aucune anomalie n''a été détectée.', 'Récupération des nouvelles écritures');
          end
          else
            RollBack;
        end;
    else
      ;
    end;

    OnLoadRecord;
  finally
    DeblocageMonoposte(True);
  end;
end;

procedure TOM_ICCGENERAUX.OnKeyPressICGGeneral(Sender: TObject; var Key: Char);
begin
  Key := TestKeyPressed(ftInteger, Key);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 05/09/2001
Modifié le ... : 05/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnClose;
begin
  inherited;
  if TobMere <> nil then
  begin
    TobMere.ClearDetail;
    TobMere.Free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/09/2001
Modifié le ... : 19/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnKeyPressSoldeAnterieur(Sender: TObject; var Key: Char);
begin
  if Ds.State = dsBrowse then
    Ds.Edit;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/11/2001
Modifié le ... :   /  /
Description .. : Resout le probleme de perte du signe dans le THNumEdit avec les
Suite ........ : Négatif, positif et nul
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnEnterSoldeAnterieur(Sender: TObject);
begin
  SoldeAnterieur.Text := FormatFloat('#,##0.00',SoldeAnterieur.Value);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/09/2001
Modifié le ... : 10/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnExitSoldeAnterieur(Sender: TObject);
begin
  if GetField('icg_soldedebex') <> SoldeAnterieur.Value then
  begin
    AutoEditDataSource;
    SetField('icg_soldedebex', SoldeAnterieur.Value);
    UpdateTobAndGrille(nil);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/09/2001
Modifié le ... : 19/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_ICCGENERAUX.OnUpdateRecord;
var
  i, Acol, ARow: Integer;
  BCancel: boolean;
begin
  inherited;

  { Vérifie la saisie du numero de compte }
  if (not Presence('GENERAUX', 'G_GENERAL', GetField('ICG_GENERAL'))) then
  begin
    PgiInfo('Le compte général sélectionné n''existe pas.','Vérification du compte saisi');
    SetFocusControl('ICG_GENERAL');
    Abort;
  end;

  if TobMere.Detail.Count > 0 then // Si une ecriture existe
  begin
    ACol := GridECriture.Col;
    ARow := GridEcriture.Row;
    BCancel := False;
    GridEcriture.OnCellExit(nil, ACol, ARow, BCancel);

    { Efface la ligne vide sauf si 1 seule ligne vide dans la grille }
    i := 0;
    while i <> TobMere.Detail.Count do
    begin
      if (TobMere.Detail[i].GetValue('ICE_DEBIT') < 0.001) and
        (TobMere.Detail[i].GetValue('ICE_CREDIT') < 0.001) then
        TobMere.Detail[i].Free
      else
        Inc(i);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/11/2007
Modifié le ... :   /  /    
Description .. : FQ 21836
Mots clefs ... : 
*****************************************************************}
procedure TOM_ICCGENERAUX.OnAfterUpdateRecord;
begin
  inherited;
  if not BlocageMonoposte(True) then Abort;

  try
    EnregistrementEcriture;
  finally
    Ecran.ActiveControl := TTabSheet(GetControl('P_GENERAL'));
    DeblocageMonoPoste(True);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/10/2001
Modifié le ... :   /  /
Description .. : Enregistre les écritures dans la table écritures
Suite ........ : et met à jour le nyméro de ligne
Mots clefs ... : INSERTDB,  TOB
*****************************************************************}
procedure TOM_ICCGENERAUX.EnregistrementEcriture;
var
  TobEcriture: Tob;
begin
  // Enregistrement des écritures
  TobEcriture := Tob.Create('Tob Ecr', nil, -1);
  ChargeTobEcriture(TobEcriture);
  TobEcriture.SetAllModifie(True);
//TobEcriture.SaveToFile('C:\TobEcriture.txt',False,True,True);

  try
    BeginTrans;

    ExecuteSQL('Delete from ICCECRITURE where (ICE_GENERAL = "' + GetField('ICG_GENERAL') + '") and ' +
      '(ICE_DATEVALEUR >= "' + USDateTime(FDate1) + '")');

    TobEcriture.InsertDB(nil, False);

    CommitTrans;
  except
    RollBack;
    MessageAlerte('Erreur en création des écritures');
  end;
  UpDateTobAndGrille(nil);

  if TobEcriture <> nil then
    TobEcriture.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/10/2001
Modifié le ... :   /  /
Description .. : Duplique Tobmere et ajoute les ecritures dont la date > date2
Suite ........ : et renumérote les numéros de ligne des écritures
Mots clefs ... : TOB; DUPLIQUER
*****************************************************************}
procedure TOM_ICCGENERAUX.ChargeTobEcriture(vTobEcr: Tob);
var
  CleUnique, i: Integer;
  SQL: TQuery;
begin
  vTobEcr.Dupliquer(TobMere, True, True);

  { Ajout des lignes d' écritures dont la date est supérieure à Date2 }
  SQL := nil;
  try
    SQL := OpenSql('Select ICE_GENERAL, ICE_LIGNE, ICE_LIBELLE, ICE_DATEVALEUR, ICE_DEBIT, ICE_CREDIT from ICCECRITURE where ' +
      '(ICE_GENERAL = "' + GetField('ICG_GENERAL') + '") and ' +
      '(ICE_DATEVALEUR > "' + USDateTime(FDate2) + '") order by ICE_DATEVALEUR', True);

    if not SQL.EOF then
      vTobEcr.LoadDetailDB('ICCECRITURE', '', '', SQL, True, False);

  finally
    if Assigned(SQL) then
      Ferme(SQL);
  end;

  // Récupère la clé de départ pour la valeur de ICE_LIGNE
  CleUnique := GetNumLigneFromEcriture( GetField('ICG_GENERAL') , FDate1 );

  vTobEcr.Detail.Sort('ICE_DATEVALEUR');

  for i := 0 to vTobEcr.Detail.Count - 1 do
  begin
    {  1 - Met à jour le numero de ligne de toutes les ecritures }
    vTobEcr.Detail[i].PutValue('ICE_LIGNE', CleUnique + 1);
    { 2 - Enregistre le numéro de compte dans les ecritures }
    vTobEcr.Detail[i].PutValue('ICE_GENERAL', GetField('ICG_GENERAL'));

    Inc(CleUnique);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/10/2001
Modifié le ... :   /  /
Description .. : Se produit lors de l' apparation du popup du popupmenu d' importation
Mots clefs ... :
*****************************************************************}
procedure TOM_ICCGENERAUX.OnPopupImporter(Sender: TObject);
begin
  TPopupMenu(GetControl('POPUPIMPORTER')).Items[0].OnClick := OnClickBRecuperation;
  TPopupMenu(GetControl('POPUPIMPORTER')).Items[1].OnClick := OnClickBMiseAJour;
  TPopupMenu(GetControl('POPUPIMPORTER')).Items[1].Enabled := ExisteNouvelleEcriture(GetField('ICG_GENERAL'), FDate1, FDate2);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/10/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_ICCGENERAUX.OnAfterFormShowEcran;
begin
  // GCO - 15/10/2004 - FQ 13926
  TToolBarButton97(GetControl('BFIRST', True)).Visible := False;
  TToolBarButton97(GetControl('BPREV', True)).Visible  := False;
  TToolBarButton97(GetControl('BNEXT', True)).Visible  := False;
  TToolBarButton97(GetControl('BLAST', True)).Visible  := False;
  // FIN GCO
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_ICCGENERAUX.OnElipsisClickZPersonne( Sender : TObject);
begin
  LookUpList(ZPersonne, 'Code de la personne', 'ANNUAIRE', 'ANN_GUIDPER',
             'ANN_APNOM', '', '', True, 0);

  SetField('ICG_GUIDPER', ZPersonne.Text);

  if ZPersonne.Text <> '' then
    ZPersonne.Text := GetColonneSQL('ANNUAIRE', 'ANN_APNOM', 'ANN_GUIDPER = "' + ZPersonne.Text + '"');
end;

////////////////////////////////////////////////////////////////////////////////

initialization
  registerclasses([TOM_ICCGENERAUX]);
end.

