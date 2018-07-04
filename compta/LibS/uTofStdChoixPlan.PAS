{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 08/03/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : STDCHOIXPLAN ()
Mots clefs ... : TOF;STDCHOIXPLAN
*****************************************************************}
unit uTofStdChoixPlan;

interface

uses StdCtrls,
  Controls,
  Classes,
{$IFDEF EAGLCLIENT}
  uLanceProcess, // LanceProcessServer
{$ELSE}
  db,
  dbtables,
{$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  Windows,
  Grids,
  Graphics,
  HCtrls,
  HEnt1,
  HMsgBox,
  Vierge,
  UTOF,
  uTOB,
  HTB97,
  HSysMenu,
  HStatus,
  Ent1,
  ParamSoc,
  galOutil,
  uLibStdCpta,
  uBOB,
  LookUp;

type
  TOF_STDCHOIXPLAN = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  private
    FDosType: TOB;
    FGrille: THGrid;
    FHSystemMenu  : THSystemMenu;
    procedure PostDrawCell(ACol, ARow: LongInt; Canvas: TCanvas; AState: TGridDrawState);
    procedure OnListeDblClick(Sender: TObject);
    procedure SuppressionStandard;
    procedure OnExportClick ( Sender : TObject );
    procedure OnImportClick ( Sender : TObject );
    procedure BobAddTableStandard ( TStd : TOB; stSQL : string );    
  end;

procedure ChargeListeDossierType(T: TOB);
var
  NumPlanCompte: integer;

implementation

const
  COL_NUMPLAN = 1;
  COL_LIBELLE = 2;
  COL_ABREGE = 3;
  COL_STANDARD = 5;

procedure ChargeListeDossierType(T: TOB);
begin
  T.ClearDetail;
  T.LoadDetailDB('STDCPTA', '', 'STC_NUMPLAN', nil, False);
end;

procedure TOF_STDCHOIXPLAN.OnNew;
begin
  inherited;
end;

procedure TOF_STDCHOIXPLAN.OnDelete;
var NoStd: integer;
begin
  inherited;
  if not IsSuperviseur(TRUE) then
  begin
    PGIInfo('Vous n''êtes pas un utilisateur privilégié', '');
    exit;
  end;
  NoStd := StrToInt(FGrille.Cells[COL_NUMPLAN, FGrille.Row]);
  if not EstSpecif('51502') then
    if NoStd < 21 then
    begin
      PGIInfo('Dossier type CEGID, suppression impossible', Ecran.Caption);
      exit;
    end;
  if Transactions(SuppressionStandard, 1) <> oeOK then
  begin
    MessageAlerte('Suppression du dossier type impossible.');
  end
  else
  begin
    PGIInfo('Suppression du dossier type terminée.', Ecran.Caption);
    // Mise à jour de la liste des dossiers type
    ChargeListeDossierType(FDosType);
    FDosType.PutGridDetail(THGrid(GetControl('FLISTE')), False, False,
      'STC_NUMPLAN;STC_LIBELLE;STC_ABREGE;STC_PREDEFINI;;',True);
  end;
end;

procedure TOF_STDCHOIXPLAN.OnUpdate;
var NoChoixStd : integer;
    lTTraitementFiltre : TTraitementFiltre;
{$IFDEF EAGLCLIENT}
    lTobParam  : Tob;
    lTobResult : Tob;
{$ENDIF}
begin
  // Chargement du dossier type
  NoChoixStd := StrToInt(FGrille.Cells[COL_NUMPLAN, FGrille.Row]);
  // Si pas de changement de plan para rapport au précédent , on ne recharge pas.
  if NumPlanCompte = NoChoixStd then Ecran.Close;
  if (PGIAsk('Voulez-vous recharger le dossier type ' + IntToStr(NoChoixStd) + ' ? ', Ecran.Caption) <> mrYes) then
    Exit;
  NumPlanCompte := NoChoixStd;

  if Blocage(['nrBatch', 'nrCloture'], True, 'nrCloture') then
    Exit;

{$IFDEF EAGLCLIENT}
//  lTobParam := InitTobParamProcessServer;
  //lTobParam.SaveToFile('C:\TOB.TXT', False, True, False);
  lTobParam := TOB.Create ('',nil,-1);
  lTobResult := LanceProcessServer('cgiStdCpta', 'CHARGESTD', IntToStr(NumPlanCompte), lTobParam, True ) ;
  if lTobResult.GetValue('RESULT') = 'PASOKCHARGE' then
{$ELSE}
  if not ChargeToutLeStandard( NumPlanCompte, True, True ) then
{$ENDIF}
  begin
    // Erreur durant la mise à jour depuis les standards
    PGIBox('Chargement du dossier type impossible. Veuillez recommencer.', 'Chargement d''un dossier type');
    Bloqueur('nrCloture', False);
    Exit;
  end
  else
  begin
    // GCO - 27/05/2004
    // Conversion des Filtres ancien format en Filtre XML
    lTTraitementFiltre := TTraitementFiltre.Create;
    lTTraitementFiltre.MAJFiltreVersionXML;
    lTTraitementFiltre.Free;
    // FIN - GCO - 27/05/2004
    
    SetParamSoc('SO_NUMPLANREF', NumPlanCompte);
    RechargeParamSoc;
    Bloqueur('nrCloture', False);
    RajouteCaptionDossier('Dossier type : ' + IntToStr(NumPlanCompte));
    PGIInfo ('Chargement du dossier type terminé.',Ecran.Caption);
    ECRAN.Close;
  end;


{$IFDEF EAGLCLIENT}
  // Libération des Tob du ProcessServer
  FreeAndNil( lTobParam  );
  FreeAndNil( lTobResult );
{$ENDIF}
  inherited;
end;

procedure TOF_STDCHOIXPLAN.OnLoad;
begin
  inherited;
  ChargeListeDossierType(FDosType);
  FDosType.PutGridDetail(THGrid(GetControl('FLISTE')), False, False,
    'STC_NUMPLAN;STC_LIBELLE;STC_ABREGE;STC_PREDEFINI;;');
end;

procedure TOF_STDCHOIXPLAN.OnArgument(S: string);
begin
  inherited;
  FHSystemMenu := THSystemMenu (TFVierge(ECRAN).HMTrad);
  FDosType := TOB.Create('', nil, -1);
  FGrille := THGrid(GetControl('FLISTE'));
  FGrille.OnDblclick := OnListeDblClick;
  FGrille.ColTypes[COL_NUMPLAN] := 'R';
  FGrille.ColLengths[COL_STANDARD - 1] := -1;
  FGrille.ColWidths[COL_STANDARD - 1] := -1;
  FGrille.PostDrawCell := PostDrawCell;
  TFVierge(Ecran).FormResize := True;
  FHSystemMenu.ResizeGridColumns ( FGrille );
  TToolbarButton97(GetControl('BEXPORT')).OnClick := OnExportClick;
  TToolbarButton97(GetControl('BIMPORT')).OnClick := OnImportClick;
end;

procedure TOF_STDCHOIXPLAN.OnClose;
begin
  FDosType.Free;
  inherited;
end;

{procedure TOF_STDCHOIXPLAN.OnDeleteClick(Sender: TObject);
begin
  // Suppression d'un dossier type
end;}

procedure TOF_STDCHOIXPLAN.OnListeDblClick(Sender: TObject);
begin
  TToolbarButton97(GetControl('BVALIDER')).Click;
end;

procedure TOF_STDCHOIXPLAN.PostDrawCell(ACol, ARow: Integer; Canvas: TCanvas;
  AState: TGridDrawState);
var
  Rect: TRect;
  St: string;
  bCegid: boolean;
  Couleur: TColor;
begin
  if (ARow = 0) or (ACol <> COL_STANDARD) then
    exit;
  Rect := FGrille.CellRect(ACol, ARow);
  bCegid := FGrille.Cells[COL_STANDARD - 1, ARow] = 'CEG';
  if ARow = FGrille.Row then Couleur := clWhite else
  begin
    if bCegid then
      Couleur := clMaroon
    else
      Couleur := clNavy;
  end;
  if bCegid then
    St := TraduireMemoire('Cegid')
  else
    St := TraduireMemoire('Cabinet');
  FGrille.Canvas.Font.Color := Couleur;
  FGrille.Canvas.TextOut((Rect.Left + Rect.Right) div 2 -
    FGrille.Canvas.TextWidth(St) div 2,
    (Rect.Top + Rect.Bottom) div 2 - FGrille.Canvas.TextHeight(St) div 2, St);
end;

procedure TOF_STDCHOIXPLAN.SuppressionStandard;
var
  NoStd: integer;
begin
  NoStd := StrToInt(FGrille.Cells[COL_NUMPLAN, FGrille.Row]);
  SupprimeDossierType(NoStd);
end;

procedure TOF_STDCHOIXPLAN.OnExportClick(Sender: TObject);
var stFichier : string;
    TStd : TOB;
    NumStd : integer;
begin
  NumStd := StrToInt(FGrille.Cells[COL_NUMPLAN, FGrille.Row]);

  { Exportation autorisée uniquement si code spécif 51502 ( Cegid ) }
  if not EstSpecif('51502') then
  if NumStd < 21 then
  begin
    PGIInfo('Dossier type CEGID, exportation impossible', Ecran.Caption);
    exit;
  end;

  if GetFileName(tfSave ,'*.BOB',stFichier) then
  begin
    Application.ProcessMessages;
    NumStd := StrToInt(FGrille.Cells[COL_NUMPLAN, FGrille.Row]);
    TStd := TOB.Create('',nil,-1);
    try
      TStd.AddChampSupValeur('BOBNAME',stFichier);
      TStd.AddChampSupValeur('BOBVERSION',NumStd);
      TStd.AddChampSupValeur('BOBNUMSOCREF',V_PGI.NumVersionBase);
      TStd.AddChampSupValeur('BOBDATEGEN',date);
      BobAddTableStandard ( TStd, 'SELECT * FROM GENERAUXREF WHERE GER_NUMPLAN='+IntToStr(NumStd));
      BobAddTableStandard ( TStd, 'SELECT * FROM GUIDEREF WHERE GDR_NUMPLAN='+IntToStr(NumStd));
      BobAddTableStandard ( TStd, 'SELECT * FROM JALREF WHERE JR_NUMPLAN='+IntToStr(NumStd));
      BobAddTableStandard ( TStd, 'SELECT * FROM TIERSREF WHERE TRR_NUMPLAN='+IntToStr(NumStd));
      BobAddTableStandard ( TStd, 'SELECT * FROM ECRGUIREF WHERE EGR_NUMPLAN='+IntToStr(NumStd));
      BobAddTableStandard ( TStd, 'SELECT * FROM ANAGUIREF WHERE AGR_NUMPLAN='+IntToStr(NumStd));
      BobAddTableStandard ( TStd, 'SELECT * FROM CORRESPREF WHERE CRF_NUMPLAN='+IntToStr(NumStd));
      BobAddTableStandard ( TStd, 'SELECT * FROM RUPTUREREF WHERE RUR_NUMPLAN='+IntToStr(NumStd));
      BobAddTableStandard ( TStd, 'SELECT * FROM REFAUTOREF WHERE RAF_NUMPLAN='+IntToStr(NumStd));
      BobAddTableStandard ( TStd, 'SELECT * FROM AXEREF WHERE XRE_NUMPLAN='+IntToStr(NumStd));
      BobAddTableStandard ( TStd, 'SELECT * FROM SECTIONREF WHERE SRE_NUMPLAN='+IntToStr(NumStd));
      BobAddTableStandard ( TStd, 'SELECT * FROM CHOIXCODREF WHERE CCR_NUMPLAN='+IntToStr(NumStd));
      BobAddTableStandard ( TStd, 'SELECT * FROM VENTILREF WHERE VR_NUMPLAN='+IntToStr(NumStd));
      BobAddTableStandard ( TStd, 'SELECT * FROM NATCPTEREF WHERE NTR_NUMPLAN='+IntToStr(NumStd));
      BobAddTableStandard ( TStd, 'SELECT * FROM FILTRESREF WHERE FIR_NUMPLAN='+IntToStr(NumStd));
      BobAddTableStandard ( TStd, 'SELECT * FROM LISTEREF WHERE LIR_NUMPLAN='+IntToStr(NumStd));
      BobAddTableStandard ( TStd, 'SELECT * FROM MODEPAIEREF WHERE MPR_NUMPLAN='+IntToStr(NumStd));
      BobAddTableStandard ( TStd, 'SELECT * FROM MODEREGLREF WHERE MRR_NUMPLAN='+IntToStr(NumStd));
      BobAddTableStandard ( TStd, 'SELECT * FROM PARSOCREF WHERE PRR_NUMPLAN='+IntToStr(NumStd));
      BobAddTableStandard ( TStd, 'SELECT * FROM TXCPTTVAREF WHERE TVR_NUMPLAN='+IntToStr(NumStd));
      BobAddTableStandard ( TStd, 'SELECT * FROM CMASQUECRITREF WHERE CMF_NUMPLAN='+IntToStr(NumStd));
      BobAddTableStandard ( TStd, 'SELECT * FROM CMASQUESAISIEREF WHERE CMR_NUMPLAN='+IntToStr(NumStd));                  
      BobAddTableStandard ( TStd, 'SELECT * FROM STDCPTA WHERE STC_NUMPLAN='+IntToStr(NumStd));
      if AglExportBob ( stFichier, False,False, TStd,True) then PGIInfo ('Export du dossier type '+IntToStr(NumStd)+' terminé.', ECRAN.Caption)
      else PGIBox ( 'Erreur lors de l''exportation du dossier type', ECRAN.Caption);
    finally
      TStd.Free;
    end;
  end;
end;

procedure TOF_STDCHOIXPLAN.OnImportClick(Sender: TObject);
var stFichier : string;
    T : TOB;
    NumPlan : integer;
    bImport : boolean;
begin
  bImport := False;
  if GetFileName(tfOpen ,'*.BOB',stFichier) then
  begin
    Application.ProcessMessages;
    T := TOB.Create('Maman',Nil,-1);
    try
      if AglReadContentBob ( stFichier , T, True ) then
      begin
        if (T.GetValue('BOBCREATEUR')<>NomHalley) then PGIBOX('Import impossible : fichier non standard !')
        else
        begin
          NumPlan := T.GetValue('BOBVERSION');
          if not EstSpecif('51502') then
            if NumPlan < 21 then
            begin
              PGIInfo('Dossier type CEGID, import impossible');
              T.Free;
              exit;
            end;
          if ExisteSQL ('SELECT * FROM STDCPTA WHERE STC_NUMPLAN='+IntToStr(NumPlan)) then
          begin
            if PGIAsk('Le dossier type '+IntToStr(NumPlan)+' existe déjà. Voulez-vous l''écraser ?')=mrYes then bImport := true;
          end else bImport := True;
          if bImport then
          begin
            SourisSablier;
            if AglImportBob ( stFichier ) then
            begin
              SourisNormale;
              PGIInfo ('Import du dossier type '+IntToStr(NumPlan)+' terminé.', ECRAN.Caption);
              // Mise à jour de la liste des dossiers type
              ChargeListeDossierType(FDosType);
              FDosType.PutGridDetail(THGrid(GetControl('FLISTE')), False, False,
                    'STC_NUMPLAN;STC_LIBELLE;STC_ABREGE;STC_PREDEFINI;;');
            end
            else
            begin
              SourisNormale;
              PGIBox ( 'Erreur lors de l''importation du dossier type', ECRAN.Caption);
            end;
          end;
        end;
      end;
    finally
      T.Free
    end;
  end;
end;

procedure TOF_STDCHOIXPLAN.BobAddTableStandard(TStd: TOB; stSQL: string);
begin
  with TOB.Create('',TStd,-1) do
  begin
    AddChampSupValeur('_OBJECTTYPE', toSQL);
    AddChampSupValeur('_OBJECTNAME', stSQL);
    AddChampSupValeur('_OBJECTWITHDATA', False );
    AddChampSupValeur('_OBJECTDOMAINE','C');
  end;
end;

initialization
  registerclasses([TOF_STDCHOIXPLAN]);
  NumPlanCompte := 0;  // A l'entrée sur l'appli, on initialise le standard à 0 ==> pas de standard chargé
end.

