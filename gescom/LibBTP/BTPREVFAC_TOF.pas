{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 25/01/2017
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPREVFAC ()
Mots clefs ... : TOF;BTPREVFAC
*****************************************************************}
Unit BTPREVFAC_TOF ;

Interface

Uses StdCtrls, 
     Controls,
     Classes, 
     db,
     {$IFNDEF DBXPRESS}
     dbtables,
     {$ELSE}
     uDbxDataSet,
     {$ENDIF}
     fe_main,
     mul,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HPanel,
     HTB97,
     HEnt1,
     HMsgBox,
     uTOB,
     Paramsoc,
     LookUp,
     UtilsGrille,
     UtilsEtat,
     Vierge,
     HSysMenu,
     HRichEdt,
     HRichOLE,
     Graphics,
     Grids,
     Types,
     UTOF,
     Windows,
     Messages;


Type
  TOF_BTPREVFAC = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

private
    //Variable pour définition de l'écran de saisie....
    client        : THEdit;
    //
    Affaire       : THEdit;
    Affaire0      : THEdit;
    Affaire1      : THEdit;
    Affaire2      : THEdit;
    Affaire3      : THEdit;
    Avenant       : THEdit;    
    //
    AFFAdr1       : THEdit;
    AFFAdr2       : THEdit;
    AFFAdr3       : THEdit;
    AFFCP         : THEdit;
    AFFVille      : THEdit;
    //
    DateDeb       : THEdit;
    DateFin       : THEdit;
    //
    SaisieDate    : THEdit;
    //
    Libclient     : THLabel;
    LibAffaire    : THLabel;
    //
    Grp_AdresseInt: TGroupBox;
    //
    GrilleSaisie  : THGrid;
    //
    TMemoRemarque : TMemo;
    //
    BtnNewLine    : TToolbarButton97;
    BtnValide     : TToolbarButton97;
    BtnImprime    : TToolbarButton97;
    BtnDelete     : TToolbarButton97;
    //
    TotMtMarche   : THNumEdit;
    TotMtPlCharge : THNumEdit;
    TotMtPrevu    : THNumEdit;
    TotMtFact     : THNumEdit;
    TotMtResteAPrevoir : THNumEdit;
    TotMtEcart    : THNumEdit;

    THPanel2      : THPanel;

    //Variable nécessaire pour la gestion de l'état
    OptionEdition : TOptionEdition;
    fEtat         : THValComboBox;

    //
    TheType       : String;
    TheNature     : String;
    TheTitre      : String;
    TheModele     : String;
    //
    OkPlanCharge  : Boolean;
    //
    Dbl_TotMtMarche   : Double;
    Dbl_TotMtPLCharge : Double;
    Dbl_TotMtPrevu    : Double;
    Dbl_TotMtFact     : Double;
    Dbl_TotMtEcart    : Double;
    Dbl_TotResteAPrevoir : Double;
    //
    Sens          : Integer;
    //
    GestGridSaisie: TGestionGS;
    //
    TOBPrevFact   : TOB;
    TobAffaire    : TOB;
    TobEdition    : TOB;
    //
    FF            : String;
    //
    procedure AddchampsSuppTobPrevFact(TobLPrevFact: TOB);
    Procedure AffichageBtnMemo;
    procedure AffichageMemoRemarque(Arow: Integer);
    procedure AffichageTobDansGrille;
    //
    Procedure CalculMtFacture(TOBLPrevFact : TOB);
    Procedure CalculMtPlCharge(TOBLPrevFact : TOB);
    procedure CalculTotMtFacture;
    procedure CalculTotMtMarche;
    procedure CalculTotPlcharge;
    procedure ChargeAdresse;
    procedure ChargeEnteteWithTobAffaire;
    procedure ChargePiedWithTotal;
    procedure ChargeTobAffaire;
    procedure ChargeTobEdition;
    procedure ChargeTOBPrevFact;
    procedure ChargeZoneSupToTOBPrevFact;
    procedure ClickBtnRemarque(Sender: Tobject);
    function  ControleDateSaisie(var mm, yy: Word; Arow : Integer): Boolean;
    procedure Controlechamp(Champ, Valeur: String);
    procedure CreateEdition;
    procedure CreateEnregTobPrevFact(OkFirst : Boolean = False);
    Procedure CreateObject;
    procedure CreateTOB;

    procedure DefinieGrille;
    procedure DestroyTOB;

    procedure FermeMemoRemarque(Arow : Integer);

    procedure GetObjects;
    procedure GSDblClick(Sender: Tobject);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer;  var Cancel: Boolean);
    procedure GSGetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
    procedure GSKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GSPostDrawCell(ACol, ARow: Integer; Canvas: TCanvas;  AState: TGridDrawState);

    procedure InitGrille;
    procedure InitTobEdition(TobLEdit : TOB);

    procedure OnClickDelete(Sender: TObject);
    procedure OnClickImprime(Sender: TObject);
    procedure OnClickNewLine(Sender: TObject);
    procedure OnClickValider(Sender: Tobject);
    procedure OnExitSaisieDate(sender: TObject);

    procedure RAZZoneEcran;
    Function  RecupfirstLine(TobLPrevFact : TOB): String;
    procedure RemarqueKeydown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure RenumeroteLigneTob(Mois : String);

    procedure SetScreenEvents;

    procedure ZoneSuivanteOuOk(var ACol, ARow: Integer; var Cancel: boolean);




end ;

  const
    SG_MOIS       = 1;
    SG_MTPLCHARGE = 2;
    SG_MTPREVU    = 3;
    SG_MTFACT     = 4;
    SG_MTECART    = 5;
    SG_REMARQUE   = 6;
    SG_VERROU     = 7;

Implementation

Uses  BTPUtil,
      AffaireUtil,
      DateUtils,
      EntGC,
      SAISUTIL,
      TntGrids;

procedure TOF_BTPREVFAC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTPREVFAC.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPREVFAC.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTPREVFAC.OnLoad ;
Var Arow, Acol : Integer;
    Cancel     : Boolean;
begin
  Inherited ;
  //Gestion des objects liés à la fiche
  CreateObject;

  ChargeTobAffaire;
  ChargeEnteteWithTobAffaire;
  //
  ChargeTOBPrevFact;
  //
  ChargeZoneSupToTOBPrevFact;
  //
  CalculTotMtMarche;
  CalculTotPlcharge;
  CalculTotMtFacture;
  //
  DBL_TotMtEcart := Dbl_TotMtPrevu - Dbl_TotMtFact;
  Dbl_TotResteAPrevoir := DBL_TotMtMarche - Dbl_TotMtPLCharge;
  //
  ChargePiedWithTotal;
  //
  AffichageTobDansGrille;
  //
  TFVierge(Ecran).HMTrad.ResizeGridColumns (GrilleSaisie);
  //
  Arow := GrilleSaisie.Rowcount -1;
  GrilleSaisie.Row := Arow;
  ACol := SG_MOIS;
  GSCellEnter(GrilleSaisie, Acol, ARow, cancel);
  
end ;

procedure TOF_BTPREVFAC.OnArgument (S : String ) ;
var Critere : string;
    Champ   : string;
    Valeur  : string;
    i       : Integer;
    x       : Integer;
begin
  Inherited ;

  OkPlanCharge := (VH_GC.SeriaPlanCharge);

  //gestion de l'écran
  GetObjects;
  //
  CreateTOB;

  //gestion de la grille
  InitGrille;

  //gestion de l'écran
  RAZZoneEcran;

  //Récupération des paramètres de lancements.....
  Critere := uppercase(Trim(ReadTokenSt(S)));
  while Critere <> '' do
  begin
     x := pos('=', Critere);
     if x <> 0 then
        begin
        Champ  := copy(Critere, 1, x - 1);
        Valeur := copy(Critere, x + 1, length(Critere));
        end
     else
        Champ  := Critere;
     ControleChamp(Champ, Valeur);
     Critere:= uppercase(Trim(ReadTokenSt(S)));
  end;

  FF:='#';
  if V_PGI.OkDecV>0 then
  begin
    FF:='# ##0.';
    for i:=1 to V_PGI.OkDecP-1 do
    begin
      FF:=FF+'0';
    end;
    FF:=FF+'0';
  end;

  SetScreenEvents;

  //
  DefinieGrille;
  //
  With GestGridSaisie do
  begin
    DessineGrille;
    GS.Colformats[SG_MTPLCHARGE] := FF;
    GS.Colformats[SG_MTPREVU]    := FF;
    GS.Colformats[SG_MTFACT]     := FF;
    GS.Colformats[SG_MTECART]    := FF;
  end;

  //Gestion de l'édition
  CreateEdition;


end ;

procedure TOF_BTPREVFAC.OnClose ;
begin
  Inherited ;

  DestroyTOB;

  FreeAndNil(TMemoRemarque);

end ;

procedure TOF_BTPREVFAC.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPREVFAC.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTPREVFAC.CreateTOB;
begin

  TOBPrevFact := TOB.Create('PREVISIONNEL', Nil, -1);
  TobAffaire  := TOB.Create('AFFAIRE',   Nil, -1);

end;

Procedure TOF_BTPREVFAC.CreateObject;
Begin

//Création du bloc-note associé au Ttoolwindows...
  TMemoRemarque             := TMemo.Create(Ecran);
  TMemoRemarque.Parent      := Ecran;
  TMemoRemarque.WordWrap    := True;
  TMemoRemarque.ScrollBars  := ssBoth ;
  TMemoRemarque.Enabled     := True;
  TMemoRemarque.ReadOnly    := False;
  TMemoRemarque.OnKeyDown   := RemarqueKeyDown;
  TMemoRemarque.Visible     := False;

end;

procedure TOF_BTPREVFAC.DestroyTOB;
begin

  FreeAndNil(TobPrevFact);
  FreeAndNil(TobAffaire);

  FreeAndNil(OptionEdition);
  FreeAndNil(GestGridSaisie);

end;

Procedure TOF_BTPREVFAC.Controlechamp(Champ, Valeur : String);
begin

  if Champ = 'CODEAFFAIRE' then Affaire.text := Valeur;

end;

Procedure TOF_BTPREVFAC.ChargeTobAffaire;
Var StSQL : string;
    QQ    : TQuery;
begin

  Try
    STSQL := 'SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="' + Affaire.text + '" AND AFF_ETATAFFAIRE="ACP"';
    QQ := OpenSQL(StSQL, False, -1, '', True);
    if not QQ.eof then
    begin
      TobAffaire.SelectDB('AFFAIRE', QQ, True);
    end;
  finally
    Ferme(QQ);
  end;

end;

Procedure TOF_BTPREVFAC.ChargeEnteteWithTobAffaire;
begin

  client.Text := TobAffaire.GetValue('AFF_TIERS');
  Libclient.Caption := RecupLibelleTiers(Client.text);

  ChargeAffaire(Affaire, Affaire1, Affaire2,Affaire3,Avenant, client, Affaire0,False,taModif,TobAffaire, nil, False, False, False);
  LibAffaire.caption := TobAffaire.GetValue('AFF_LIBELLE');

  ChargeAdresse;

  DateDeb.Text := TobAffaire.GetValue('AFF_DATEDEBUT');
  DateFin.Text := TobAffaire.GetValue('AFF_DATEFIN');

end;

Procedure TOF_BTPREVFAC.ChargeAdresse;
Var StSQL : string;
    QQ    : TQuery;
Begin

  Try
    STSQL := 'SELECT ADR_LIBELLE, ADR_LIBELLE2, ADR_ADRESSE1, ADR_ADRESSE2, ADR_ADRESSE3, ADR_CODEPOSTAL, ADR_VILLE';
    StSQL := StSQL + ' FROM ADRESSES WHERE ADR_REFCODE="' + Affaire.text + '" AND ADR_TYPEADRESSE="INT" AND ADR_NADRESSE=1';
    QQ := OpenSQL(StSQL, False, -1, '', True);
    if not QQ.eof then
    begin
      Grp_AdresseInt.Caption := QQ.FindField('ADR_LIBELLE').AsString;
      //
      AFFAdr1.Text   := QQ.FindField('ADR_ADRESSE1').AsString;
      AFFAdr2.Text   := QQ.FindField('ADR_ADRESSE2').AsString;
      AFFAdr3.Text   := QQ.FindField('ADR_ADRESSE3').AsString;
      AFFCP.Text     := QQ.FindField('ADR_CODEPOSTAL').AsString;
      AFFVille.Text  := QQ.FindField('ADR_VILLE').AsString;
    end;
  finally
    Ferme(QQ);
  end;

end;

procedure TOF_BTPREVFAC.ChargeTOBPrevFact;
Var StSQL : string;
    QQ    : TQuery;
Begin

  Try
    STSQL := 'SELECT * FROM BTPREVFAC WHERE BPF_AFFAIRE="' + Affaire.text + '" ORDER BY BPF_MOIS';
    QQ := OpenSQL(StSQL, False, -1, '', True);
    if not QQ.eof then
    begin
      TobPrevFact.LoadDetailDB('BTPREVFAC','','',QQ, True, True);
    end;
  finally
    Ferme(QQ);
  end;

end;

procedure TOF_BTPREVFAC.ChargeZoneSupToTOBPrevFact;
Var yy            : Word;
    mm            : Word;
    ind           : Integer;
    MtPrevu       : Double;
    TOBLPrevFact  : TOB;
begin

  if TobPrevFact = Nil then Exit;

  //Chargement du mois et de l'annnée
  if TobPrevFact.detail.count = 0 then
    CreateEnregTobPrevFact(True)
  else
  begin
    for Ind := 0 To tobPrevFact.detail.Count - 1 do
    begin
      TOBLPrevFact := TobPrevFact.detail[Ind];
      //
      AddchampsSuppTobPrevFact(TOBLPrevFact);
      //
      yy := StrToInt(Copy(TOBLPrevFact.getValue('BPF_MOIS'),0,4));
      mm := StrToInt(Copy(TOBLPrevFact.getValue('BPF_MOIS'),5,2));
      //
      SaisieDate.Text := IntToStr(mm) + '-' + IntToStr(yy);
      //
      MtPrevu         := TOBLPrevFact.getValue('BPF_MTPREVU');
      Dbl_TotMtPrevu  := Dbl_TotMtPrevu + MtPrevu;
      //
      CalculMtPlCharge(TOBLPrevFact);
      //
      CalculMtFacture(TOBLPrevFact);
      //
      RecupfirstLine(TOBLPrevFact);   
      //
    end;
  end;

end;

procedure TOF_BTPREVFAC.CreateEnregTobPrevFact(OkFirst : Boolean = False);
Var ToblPrevFact  : TOB;
    yy            : Word;
    mm            : Word;
    dd            : Word;
    Mois          : string;
    DateSaisie    : TDateTime;
begin

  ToblPrevFact := Tob.Create('BTPREVFAC', TobPrevFact, -1);

  AddchampsSuppTobPrevFact(ToblPrevFact);

  TOBLPrevFact.PutValue('BPF_AFFAIRE', Affaire.text);
  //
  TOBLPrevFact.PutValue('MTPLCHARGE',  0.00);
  TOBLPrevFact.PutValue('BPF_MTPREVU', 0.00);
  //
  if OkFirst then
  Begin
    DateSaisie := StrToDate(Datedeb.text);
    Decodedate(DateSaisie, yy, mm, dd);
  end
  else
  begin
    yy    := StrToInt(Copy(TobPrevFact.detail[GrilleSaisie.RowCount - 2].GetValue('BPF_MOIS'),0,4));
    mm    := StrToInt(Copy(TobPrevFact.detail[GrilleSaisie.RowCount - 2].GetValue('BPF_MOIS'),5,2));
    //
    DateSaisie := IncMonth(EncodeDate(yy, mm, 01), 1);
    Decodedate(DateSaisie, yy, mm, dd);
  end;

  SaisieDate.Text     := IntToStr(mm) + '-' + IntToStr(yy);

  DecodeDate(DateSaisie,YY, MM, DD);

  Mois := Format('%.4d%.2d', [yy,mm]);
  ToblPrevFact.PutValue('BPF_MOIS', Mois);

  CalculMtPlCharge(TOBLPrevFact);

  CalculMtFacture(TOBLPrevFact);

end;

procedure TOF_BTPREVFAC.AddchampsSuppTobPrevFact(TobLPrevFact : TOB);
begin

  TobLPrevFact.AddChampSupValeur('FLINEREMARQUE', '');

  TobLPrevFact.AddChampSupValeur('MTPLCHARGE',    '0.00');
  TobLPrevFact.AddChampSupValeur('MTFACT',        '0.00');
  TobLPrevFact.AddChampSupValeur('MTECART',       '0.00');

  TobLPrevFact.AddChampSupValeur('NOLIGNE',       '0');

  TobLPrevFact.AddChampSupValeur('OKVALIDE', 'O');

end;

Procedure TOF_BTPREVFAC.CalculMtPlCharge(TOBLPrevFact : TOB);
Var Ratio           : Double;  //NbHeurePC/NbHeureAff
    NbheurePC       : Double;
    NbHeureAFF      : Double;  //GP_TPSUNITAIRE
    MTDevisACC      : Double;  //somme des devis acceptés
    MtPlCharge      : Double;
    StSQl           : string;
    QQ              : TQuery;
    Date1           : TDateTime;
    Date2           : TDateTime;
    mm, yy          : Word;
begin

  MtPlCharge  := 0;
  NbheureAff  := 0;
  NbheurePC   := 0;
  MtDevisAcc  := 0;
  Ratio       := 0;

  if not OkPlanCharge then
  Begin
    ToblPrevFact.PutValue('MTPLCHARGE', StrF00(MtPlCharge   , V_PGI.OkDecP));
    Exit;
  end;

  //On extrait le mois et l'année de la zone de grille.....
  mm := StrToInt(copy(SaisieDate.Text,0,2));
  yy := StrToInt(Copy(saisieDate.text,4,4));

  Date1 := EncodeDate(yy, mm, 1);
  Date2 := EndOfAMonth(yy, mm);

  //On calcul d'abord le ratio du nombre d'heure
  try
    StSQl := 'SELECT SUM(BAT_NBHRS) AS HEUREPC FROM BAFFECTCHANT ';
    StSQl := StSQl + 'WHERE BAT_AFFAIRE = "'  + Affaire.Text        + '" ';
    StSQl := StSQl + '  AND BAT_DATE >="'     + USDATETIME(Date1)   + '" ';
    StSQl := StSQl + '  AND BAT_DATE <="'     + USDATETIME(Date2)   + '" ';

    QQ := OpenSQL(StSQl, False, -1, '', False);

    if not QQ.Eof then NbheurePC := QQ.findField('HEUREPC').AsFloat;
  finally
    Ferme(QQ);
  end;

  Date1 := StrToDate(DateDeb.text);
  Date2 := StrToDate(Datefin.text);

  //On calcul ensuite le nombre d'heures totales de l'affaire, le montant total des devis accepté sur l'affaire...
  try
    StSQl := 'SELECT SUM(GP_TOTALHEURE) AS HEUREAFF, SUM(GP_TOTALHTDEV) AS MONTANTHT FROM PIECE ';
    StSQL := StSQL + 'LEFT JOIN AFFAIRE AS AA on GP_AFFAIRE = AA.AFF_AFFAIRE ';
    STSQL := StSQL + 'LEFT JOIN AFFAIRE AS AD on GP_AFFAIREDEVIS = AD.AFF_AFFAIRE ';
    StSQl := StSQl + 'WHERE GP_AFFAIRE = "'   + Affaire.Text + '" ';
    StSQl := StSQl + '  AND GP_NATUREPIECEG = "DBT" ';
    StSQl := StSQl + '  AND GP_DATEPIECE >="' + USDATETIME(Date1)   + '" ';
    StSQl := StSQl + '  AND GP_DATEPIECE <="' + USDATETIME(Date2)   + '" ';
    StSQl := StSQl + '  AND AD.AFF_ETATAFFAIRE ="ACP" ';

    QQ := OpenSQL(StSQl, False, -1, '', False);

    if not QQ.Eof then
    Begin
      NbheureAff := QQ.findField('HEUREAFF').AsFloat;
      MTDevisACC := QQ.findField('MONTANTHT').AsFloat;
    end;
  finally
    Ferme(QQ);
  end;

  //Calcul du Ratio du nombre d'heure
  if NbheureAff <> 0 then Ratio := NbheurePC/NbheureAff;

  MtPlCharge := Ratio * MtDevisAcc;
  MtPlCharge := Arrondi(MtPlCharge, 2);

  ToblPrevFact.PutValue('MTPLCHARGE', FloatTostr(MtPlCharge));

  Dbl_TotMtPlcharge  := Dbl_TotMtPlCharge + MtPlCharge;

end;

Procedure TOF_BTPREVFAC.CalculMtFacture(TOBLPrevFact: TOB);
Var StSql           : string;
    QQ              : Tquery;
    Date1           : Tdatetime;
    Date2           : TDatetime;
    MtFacture       : Double;
    MtPrevu         : Double;
    MtEcart         : Double;
    mm, yy          : Word;
begin

  MtFacture := 0;

  MtPrevu := Valeur(ToblPrevFact.GetValue('BPF_MTPREVU'));

  //On extrait le mois et l'année de la zone de grille.....
  mm := StrToInt(copy(SaisieDate.Text,0,2));
  yy := StrToInt(Copy(saisieDate.text,4,4));

  Date1 := EncodeDate(yy, mm, 1);
  Date2 := EndOfAMonth(yy, mm);

  try
    StSQl := 'SELECT SUM(GP_TOTALHTDEV) AS MTFACT FROM PIECE ';
    StSQL := StSQL + 'LEFT JOIN AFFAIRE AS AA on GP_AFFAIRE = AA.AFF_AFFAIRE ';
    STSQL := StSQL + 'LEFT JOIN AFFAIRE AS AD on GP_AFFAIREDEVIS = AD.AFF_AFFAIRE ';
    StSQl := StSQl + 'WHERE GP_AFFAIRE = "'   + Affaire.Text + '" ';
    StSQl := StSQl + '  AND (GP_NATUREPIECEG = "FBT" OR GP_NATUREPIECEG = "ABT") '; //GP_NATUREPIECEG = "FBT" ';
    StSQl := StSQl + '  AND GP_DATEPIECE >="' + USDATETIME(Date1)   + '" ';
    StSQl := StSQl + '  AND GP_DATEPIECE <="' + USDATETIME(Date2)   + '" ';
    StSQl := StSQl + '  AND AD.AFF_ETATAFFAIRE = "ACP" ';

    QQ := OpenSQL(StSQl, False, -1, '', False);

    if not QQ.Eof then
    Begin
      MtFacture := QQ.findField('MTFACT').AsFloat;
    end;
  finally
    Ferme(QQ);
  end;

  MtFacture := Arrondi(MtFacture, 2);
  MtPrevu   := Arrondi(MtPrevu, 2);

  MtEcart   := MtPrevu - MtFacture;
  MtEcart   := Arrondi(MtEcart, 2);

  ToblPrevFact.PutValue('MTFACT',  MtFacture);
  ToblPrevFact.PutValue('MTECART', MtEcart);

end;

Function TOF_BTPREVFAC.RecupfirstLine(TobLPrevFact : TOB) : String;
var ind      : Integer;
    Contenu  : TStrings;
begin

  Result := '';

  if TobLPrevFact = nil then exit;

  Contenu := TStringList.Create;

  Contenu.Text := TOBLPrevFact.GetValue('BPF_REMARQUE');

  if Contenu = nil then Exit;

  For ind := 0 To Contenu.Count - 1 do
  begin
    Result := Contenu[Ind];
    if result <> '' then break;
  end;

  ToblPrevFact.PutValue('FLINEREMARQUE', Result);

end;


Procedure TOF_BTPREVFAC.CalculTotMtMarche;
Var StSql           : string;
    QQ              : Tquery;
begin

  try
    StSQl := 'SELECT SUM(GP_TOTALHTDEV) AS MTMARCHE FROM PIECE ';
    StSQL := StSQL + 'LEFT JOIN AFFAIRE AS AA on GP_AFFAIRE = AA.AFF_AFFAIRE ';
    STSQL := StSQL + 'LEFT JOIN AFFAIRE AS AD on GP_AFFAIREDEVIS = AD.AFF_AFFAIRE ';
    StSQl := StSQl + 'WHERE GP_AFFAIRE = "'   + Affaire.Text + '" ';
    StSQl := StSQl + '  AND GP_NATUREPIECEG = "DBT" ';
    StSQl := StSQl + '  AND AD.AFF_ETATAFFAIRE = "ACP" ';

    QQ := OpenSQL(StSQl, False, -1, '', False);

    if not QQ.Eof then
    Begin
      Dbl_TotMtMarche := QQ.findField('MTMARCHE').AsFloat;
    end;
  finally
    Ferme(QQ);
  end;

end;

Procedure TOF_BTPREVFAC.CalculTotPlcharge;
Var StSql           : string;
    QQ              : Tquery;
    NbheurePC       : Double;
    NbheureAff      : Double;
    MTDevisACC      : Double;
    Ratio           : Double;
begin

  if not OkPlanCharge then Exit;

  NbheureAff  := 0;
  MtDevisAcc  := 0;
  Ratio       := 0;
  NbheurePC   := 0;

  //On calcul d'abord le ratio du nombre d'heure
  try
    StSQl := 'SELECT SUM(BAT_NBHRS) AS HEUREPC FROM BAFFECTCHANT ';
    StSQl := StSQl + 'WHERE BAT_AFFAIRE = "'  + Affaire.Text        + '" ';

    QQ := OpenSQL(StSQl, False, -1, '', False);

    if not QQ.Eof then NbheurePC := QQ.findField('HEUREPC').AsFloat;
  finally
    Ferme(QQ);
  end;

  //On calcul ensuite le nombre d'heures totales de l'affaire, le montant total des devis accepté sur l'affaire...
  try
    StSQl := 'SELECT SUM(GP_TOTALHEURE) AS HEUREAFF, SUM(GP_TOTALHTDEV) AS MONTANTHT FROM PIECE ';
    StSQL := StSQL + 'LEFT JOIN AFFAIRE AS AA on GP_AFFAIRE = AA.AFF_AFFAIRE ';
    STSQL := StSQL + 'LEFT JOIN AFFAIRE AS AD on GP_AFFAIREDEVIS = AD.AFF_AFFAIRE ';
    StSQl := StSQl + 'WHERE GP_AFFAIRE = "'   + Affaire.Text + '" ';
    StSQl := StSQl + '  AND GP_NATUREPIECEG = "DBT" ';
    StSQl := StSQl + '  AND AD.AFF_ETATAFFAIRE ="ACP" ';

    QQ := OpenSQL(StSQl, False, -1, '', False);

    if not QQ.Eof then
    Begin
      NbheureAff := QQ.findField('HEUREAFF').AsFloat;
      MTDevisACC := QQ.findField('MONTANTHT').AsFloat;
    end;
  finally
    Ferme(QQ);
  end;

  //Calcul du Ratio du nombre d'heure
  if NbheureAff <> 0 then Ratio := NbheurePC/NbheureAff;

  Dbl_TotMtPLCharge := MtDevisAcc * Ratio;

end;

Procedure TOF_BTPREVFAC.CalculTotMtFacture;
Var StSql           : string;
    QQ              : Tquery;
begin

  Dbl_TotMtFact := 0;

  try
    StSQl := 'SELECT SUM(GP_TOTALHTDEV) AS MTFACT FROM PIECE ';
    StSQL := StSQL + 'LEFT JOIN AFFAIRE AS AA on GP_AFFAIRE = AA.AFF_AFFAIRE ';
    STSQL := StSQL + 'LEFT JOIN AFFAIRE AS AD on GP_AFFAIREDEVIS = AD.AFF_AFFAIRE ';
    StSQl := StSQl + 'WHERE GP_AFFAIRE = "'   + Affaire.Text + '" ';
    StSQl := StSQl + '  AND (GP_NATUREPIECEG = "FBT" OR GP_NATUREPIECEG = "ABT")';
    StSQl := StSQl + '  AND AD.AFF_ETATAFFAIRE = "ACP" ';

    QQ := OpenSQL(StSQl, False, -1, '', False);

    if not QQ.Eof then
    Begin
      Dbl_TotMtFact := QQ.findField('MTFACT').AsFloat;
    end;
  finally
    Ferme(QQ);
  end;

end;

Procedure TOF_BTPREVFAC.ChargePiedWithTotal;
begin

  TotMtMarche.text    := StrF00(Dbl_TotMtMarche,    V_PGI.OkDecP);
  TotMtPlCharge.Text  := StrF00(Dbl_TotMtPlCharge,  V_PGI.OkDecP);
  TotMtPrevu.text     := StrF00(Dbl_TotMtPrevu ,    V_PGI.OkDecP);
  TotMtFact.text      := StrF00(Dbl_TotMtFact  ,    V_PGI.OkDecP);
  TotMtResteAPrevoir.text := StrF00(Dbl_TotResteAPrevoir,    V_PGI.OkDecP);
  TotMtEcart.text     := StrF00(Dbl_TotMtEcart ,    V_PGI.OkDecP);

end;

Procedure TOF_BTPREVFAC.AffichageTobDansGrille;
Var ARow    : Integer;
    Acol    : Integer;
    cancel  : Boolean;
begin

  with GestGridSaisie do
  begin
    Cancel  := False;

    TOBPrevFact.Detail.sort('BPF_MOIS');
    //
    ChargementGrille;
    //
    //Arow := GS.RowCount-1;
    //ACol := SG_MOIS;
    //
    //GSCellEnter(GS, Acol, ARow, cancel);
  end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : F.Vautrain
Créé le ...... : 25/01/2017
Modifié le ... :   /  /
Description .. : Gestion de l'écran et des zones de saisie
Mots clefs ... :
*****************************************************************}
procedure TOF_BTPREVFAC.GetObjects;
begin

  THPanel2    := THPanel(GetControl('THPANEL2'));

  client      := THEdit(GetControl('CLIENT'));
  //
  Affaire     := THEdit(GetControl('AFFAIRE'));
  Affaire0    := THEdit(GetControl('AFFAIRE0'));
  Affaire1    := THEdit(GetControl('AFFAIRE1'));
  Affaire2    := THEdit(GetControl('AFFAIRE2'));
  Affaire3    := THEdit(GetControl('AFFAIRE3'));
  Avenant     := THEdit(GetControl('AVENANT'));
  //
  Libclient   := THLabel(GetControl('LIBCLIENT'));
  LibAffaire  := THLabel(GetControl('LIBAFFAIRE'));
  //
  Grp_AdresseInt := TGroupBox((GetControl('GRP_ADRESSEINT')));
  //
  AFFAdr1     := THEdit(GetControl('AFFADR1'));
  AFFAdr2     := THEdit(GetControl('AFFADR2'));
  AFFAdr3     := THEdit(GetControl('AFFADR3'));
  AFFCP       := THEdit(GetControl('AFFCP'));
  AffVille    := THEdit(GetControl('AFFVILLE'));
  //
  SaisieDate  := THEdit(GetControl('SAISIEDATE'));
  //
  DateDeb     := THEdit(GetControl('DATEDEB'));
  DateFin     := THEdit(GetControl('DATEFIN'));
  //
  GrilleSaisie:= THGrid(GetControl('GRILLESAISIE'));
  //
  TotMtMarche     := THNumEdit(GetControl('TOTMTMARCHE'));
  TotMtPlCharge   := THNumEdit(GetControl('TOTMTPREV'));
  TotMtPrevu      := THNumEdit(GetControl('TOTMTPREVU'));
  TotMtFact       := THNumEdit(GetControl('TOTMTFACT'));
  TotMtResteAPrevoir   := THNumEdit(GetControl('RESTAPREVOIR'));
  TotMtEcart      := THNumEdit(GetControl('TOTECART'));
  TotMtResteAPrevoir := THNumEdit(GetControl('TOTMTRESTEAPREVOIR'));
  //
  BtnNewLine      := TToolbarButton97(GetControl('BTNEWLINE'));
  BtnValide       := TToolbarButton97(GetControl('BValider'));
  BtnImprime      := TToolbarButton97(GetControl('BImprimer'));
  BtnDelete       := TToolbarButton97(GetControl('BDelete'));

end;

procedure TOF_BTPREVFAC.SetScreenEvents;
begin

  BtnNewLine.OnClick          := OnClickNewLine;
  BtnValide.OnClick           := OnClickValider;
  BtnImprime.OnClick          := OnClickImprime;
  BtnDelete.OnClick           := OnClickDelete;
  SaisieDate.OnExit           := OnExitSaisieDate;

  GrilleSaisie.OnCellEnter    := GSCellEnter;
  GrilleSaisie.OnCellExit     := GSCellExit;
  GrilleSaisie.OnDblClick     := GSDblClick;
  GrilleSaisie.OnKeyDown      := GSKeyDown;

  GrilleSaisie.GetCellCanvas  := GSGetCellCanvas;
  GrilleSaisie.PostDrawCell   := GSPostDrawCell;
  GrilleSaisie.OnElipsisClick := ClickBtnRemarque;

end;

procedure TOF_BTPREVFAC.RAZZoneEcran;
begin

  Sens := 1;

  client.text   := '';
  //
  Affaire.Text  := '';
  Affaire0.Text := '';
  Affaire1.Text := '';
  Affaire2.Text := '';
  Affaire3.Text := '';
  Avenant.Text  := '';
  //
  Libclient.Caption   := '...';
  LibAffaire.Caption  := '...';
  //
  Grp_AdresseInt.Caption  := 'Adresse Intervention';
  //
  AFFAdr1.text        := '';
  AFFAdr2.text        := '';
  AFFAdr3.text        := '';
  AFFCP.text          := '';
  AffVille.text       := '';
  //
  DateDeb.text        := DateToStr(idate1900);
  DateFin.text        := DateToStr(idate2099);
  //
  TOBPrevFact.ClearDetail;
  //
  TotMtMarche.Text    := '0.00';
  TotMtPlCharge.Text  := '0.00';
  TotMtPrevu.Text     := '0.00';
  TotMtFact.Text      := '0.00';
  TotMtResteAPrevoir.Text  := '0.00';
  TotMTEcart.Text     := '0.00';
  //
end;


Procedure TOF_BTPREVFAC.OnExitSaisieDate(sender : TObject);
Var ToblPrevFact  : TOB;
    ARow,ACol     : Integer;
    yy            : Word;
    mm            : Word;
    MtEcart       : Double;
    MtPlCharge    : Double;
    //MtPrevu       : Double;
    Mois          : String;
    ZoneSAV       : string;
    Cancel        : Boolean;
begin

  Arow  := GrilleSaisie.Row;

  if ARow < GrilleSaisie.FixedRows then Exit;

  ToblPrevFact := GestGridSaisie.GetTOBGrille;

  If not ControleDateSaisie(mm, yy, Arow) then
  begin
    SaisieDate.SetFocus;
    Exit;
  end;

  ZoneSAV := TobLPrevFact.GetValue('BPF_MOIS');

  SaisieDate.Visible := False;

  MtPlCharge    := Valeur(ToblPrevFact.GetValue('MTPLCHARGE'));
  MtEcart       := Valeur(ToblPrevFact.GetValue('MTECART'));
  //
  Dbl_TotMtPlCharge := Dbl_TotMtPlcharge  - MtPlcharge;
  Dbl_TotMtEcart    := Dbl_TotMtEcart     - MtEcart;
  //
  CalculMtPlCharge(ToblPrevFact);
  CalculMtFacture(ToblPrevFact);
  //
  ChargePiedWithTotal;
  //
  Mois := Format('%.4d%.2d', [yy,mm]);
  //
  ToblPrevFact.PutValue('BPF_MOIS', mois);
  GrilleSaisie.Cells[SG_MOIS, ARow] := mois;

  if ZoneSAV <> Mois then RenumeroteLigneTob(Mois);
  //
  GrilleSaisie.SetFocus;
  //
  GrilleSaisie.Col := SG_MTPREVU;
  //
	GrilleSaisie.CacheEdit;
  GrilleSaisie.SynEnabled := false;
  GSCellEnter(self, ACol, Arow, cancel);
  GrilleSaisie.SynEnabled := true;
	GrilleSaisie.ShowEditor;

  //SendMessage(GrilleSaisie.Handle, WM_KeyDown, VK_TAB, 0);
  //
end;

function TOF_BTPREVFAC.ControleDateSaisie(var mm, yy : Word; Arow : Integer) : Boolean;
Var ToblPrevFac : Tob;
    Y, M, J     : Word;
    TempoDt     : Tdate;
begin

  Result := False;

  if (IsNumeric(copy(SaisieDate.Text,0,2))) And (Trim(copy(SaisieDate.Text,0,2)) <> '') then
    mm := StrToInt(copy(SaisieDate.Text,0,2))
  else
    Exit;

  if (IsNumeric(copy(SaisieDate.Text,4,4))) And (Trim(copy(SaisieDate.Text,4,4)) <> '') then
    yy    := StrToInt(Copy(SaisieDate.text,4,4))
  else
   Exit;

  if mm > 12   then mm := MonthOf(Now);
  if (yy > 2099) Or (yy < 1900) Then yy := YearOf(Now);

  //on vérifie si dans la table cette date n'existe pas déjà...
  ToblPrevFac := TOBPrevFact.findfirst(['BPF_MOIS'],[Format('%.4d%.2d', [yy,mm])],False);
  if (ToblPrevFac <> nil) And (ToblPrevFac.GetIndex <> ARow -1) then
  Begin
    PGIError('Date déjà présente dans la grille', 'Saisie Prévisionnel');
    Exit;
  end;

  //On vérifie que la date saisie ne soit pas inférieur à la date de début chantier
  //en effet il est peu probable que l'on fasse un prévisionnel à postériori du début du chantier.
  DecodeDate(StrToDate(DateDeb.Text), Y, M, J);
  TempoDt := EncodeDate(yy,mm,01);

  if (tempoDt < EncodeDate(y, M, 01)) then
  begin
    PGIError('La date saisie ne peut-être inférieure à la date de début du chantier', 'Saisie Prévisionnel');
    Exit;
  end;

  Result := True;

end;

Procedure TOF_BTPREVFAC.AffichageBtnMemo;
begin

  GrilleSaisie.ElipsisButton := True;

  //AffichageMemoRemarque(GestGridSaisie.GS.Row);

end;

procedure TOF_BTPREVFAC.ClickBtnRemarque(Sender : Tobject);
Var TOBLPrevFact  : TOB;
    Arow          : Integer;
begin

  With GestGridSaisie do
  begin
    Arow := GS.Row;
    TOBLPrevFact := GetTobGrille;
    If TobLPrevFact = nil then Exit;

    if TMemoRemarque.visible then
      FermeMemoRemarque(Arow)
    else
      AffichageMemoRemarque(Arow);
  end;

end;

procedure TOF_BTPREVFAC.AffichageMemoRemarque(Arow : Integer);
Var ARect : Trect;
begin

  with GestGridSaisie do
  begin
    ARect := GS.CellRect(SG_REMARQUE, ARow);

    //Affichage du bloc-note associé au Ttoolwindows...
    TMemoRemarque.Top         := Arect.Top + THPanel2.Top;
    TMemoRemarque.Left        := Arect.left;
    TMemoRemarque.Height      := GS.DefaultRowHeight * 15;
    TMemoRemarque.Width       := GS.ColWidths[SG_REMARQUE] - 18;
    //
    TMemoRemarque.Visible     := True;
    //
    TMemoRemarque.Text        := TOBPrevFact.detail[ARow - 1].GetValue('BPF_REMARQUE');
    //
    TMemoRemarque.SetFocus;
  end;

end;

Procedure TOF_BTPREVFAC.RemarqueKeydown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin

  if Key=VK_ESCAPE then FermeMemoRemarque(GestGridSaisie.GS.row);

end;

procedure TOF_BTPREVFAC.FermeMemoRemarque(Arow : Integer);
Var TOBLPrevFact  : TOB;
    ZoneGrid      : String;
    Contenu       : TStrings;
    ind           : Integer;
begin

  TOBLPrevFact := TOBPrevFact.detail[Arow -1];

  if ToblPrevFact = nil then Exit;

  if TMemoRemarque <> nil then
  begin
    TobLPrevFact.PutValue('BPF_REMARQUE', TMemoRemarque.text);
    TMemoRemarque.visible := False;
  end
  else
  begin
    ZoneGrid := GrilleSaisie.Cells[SG_REMARQUE, Arow];
    if TobLPrevFact.GetValue('BPF_REMARQUE') = '' then
      TobLPrevFact.PutValue('BPF_REMARQUE', ZoneGrid)
    else
    begin
      if ToblPrevFact.GetValue('FLINEREMARQUE') <> ZoneGrid then
      begin
        //on charge zone grid en 1ère ligne du text... et on remplace FLineRemarque...
        Contenu := TStringList.Create;
        Contenu.Text := TOBLPrevFact.GetValue('BPF_REMARQUE');
        if Contenu = nil then Exit;
        For ind := 0 To Contenu.Count - 1 do
        begin
          if Contenu[Ind] <> '' then
          begin
            Contenu[Ind] := ZoneGrid;
            break;
          end;
        end;
        TobLPrevFact.PutValue('BPF_REMARQUE', Contenu.Text)
      end;
    end;
  end;

  //On récupère les X premiers caractères....
  RecupfirstLine(TobLPrevFact);

  //11/04/2017
  GestGridSaisie.ChargementGrille;
  //TobLPrevFact.PutLigneGrid(GrilleSaisie,Arow,False,False,'FLINEREMARQUE');

end;


Procedure TOF_BTPREVFAC.OnClickNewLine(Sender : TObject);
Var Arow, Acol : Integer;
    Cancel     : Boolean;
begin

  CreateEnregTobPrevFact;
  //
  ChargePiedWithTotal;
  //
  AffichageTobDansGrille;
  //
  Arow := GrilleSaisie.Rowcount -1;
  GrilleSaisie.Row := Arow;
  ACol := SG_MOIS;
  GSCellEnter(GrilleSaisie, Acol, ARow, cancel);

end;

Procedure TOF_BTPREVFAC.OnClickValider(Sender : Tobject);
begin
  BEGINTRANS;
  TRY
    ExecuteSql('DELETE BTPREVFAC WHERE BPF_AFFAIRE = "' + Affaire.Text + '" ');
    TOBPrevFact.SetAllModifie(true);
    TOBPrevFact.InsertOrUpdateDB;
    COMMITTRANS;
  excepT
    // ICIC MESSAGE DERREUR A LA CON
    ROLLBACK;
    close;
  end;
end;

procedure TOF_BTPREVFAC.OnClickDelete(Sender : TObject);
Var TOBLPrevFact  : TOB;
    MtPrevu       : Double;
    ARow          : Integer;
    ACol          : Integer;
    Cancel        : Boolean;
    MtFact        : Double;
    MtEcart       : double;
begin

  ARow := GrilleSaisie.Row;
  Acol := SG_MOIS;
  cancel := false;

  TobLPrevFact := GestGridSaisie.GetTOBGrille;

  If TOBLPrevFact = nil then Exit;

  If TOBLprevFact.GetBoolean('BPF_VERROU') then
    PGIInfo('Suppression impossible cette ligne est verrouillée', 'Suppression')
  else
  Begin

    If PGIAsk('Confirmez-vous la suppression de cette ligne ?', 'Suppression')=mrNo then Exit;

    MtPrevu := Valeur(TOBLPrevFact.GetValue('BPF_MTPREVU'));
    //MtFact  := Valeur(TOBLPrevFact.GetValue('MTFACT'));
    //MtEcart := MtFact - MtPrevu;
    //ToblPrevFact.PutValue('MTECART',      MtEcart);
    //
    //On Reclacul le total prévu....
    Dbl_TotMtPrevu  := MtPrevu - Dbl_TotMtPrevu;
    Dbl_TotMtEcart  := Dbl_TotMtPrevu - Dbl_TotMtFact;
    //
    ChargePiedWithTotal;
    //
    GrilleSaisie.SynEnabled := false;
    GrilleSaisie.DeleteRow(GrilleSaisie.row);
    GSCellEnter (Self,Acol,ARow,cancel);
    GrilleSaisie.SynEnabled := true;
    //
    TobLPrevFact.Free;
  end;

end;

procedure TOF_BTPREVFAC.OnClickImprime(Sender: TObject);
begin

  TobEdition  := TOB.Create('EDITION',   Nil, -1);

  ChargeTobEdition;
  //
  if OptionEdition.LanceImpression('', TobEdition) < 0 then V_PGI.IoError:=oeUnknown;

  FreeAndNil(TobEdition);

end;



{***********A.G.L.***********************************************
Auteur  ...... : F.Vautrain
Créé le ...... : 25/01/2017
Modifié le ... :   /  /    
Description .. : Gestion de la grille de saisie
Mots clefs ... :
*****************************************************************}
procedure TOF_BTPREVFAC.InitGrille;
begin

  //Une recherche de la grille au niveau de la table des liste serait bien venu !!!
  GestGridSaisie          := TGestionGS.Create;

  With GestGridSaisie do
  Begin
    Ecran         := TFVierge(Ecran);
    GS            := GrilleSaisie;
    TOBG          := TOBPrevFact;
    NomListe      := '';
    ColNamesGS    := '';
    AlignementGS  := '';
    LibColNameGS  := '';
    titreGS       := '';
    LargeurGS     := '';
  end;

end;

procedure TOF_BTPREVFAC.DefinieGrille;
begin

  with GestGridSaisie do
  begin
    titreGS       := 'Saisie du Prévisionnel de Facturation';
    // Définition de la liste de saisie pour la grille Détail
    ColNamesGS    := 'BPF_MOIS;MTPLCHARGE;BPF_MTPREVU;MTFACT;MTECART;FLINEREMARQUE;BPF_VERROU';
    AlignementGS  := 'G.0  ---;D/' + IntToStr(V_PGI.OkDecP) + '  -X-;D/' + IntToStr(V_PGI.OkDecP) + '  -X-;D/' + IntToStr(V_PGI.OkDecP) + '  -X-;D/' + IntToStr(V_PGI.OkDecP) + '  -X-;G.0  ---;C.0  ---';
    LibColNameGS  := 'Mois/Année;Prévisionnel Plan Charge;Montant Prévu;Montant Facturé;Ecart;Remarque;V.';
    if OkPlanCharge then
      LargeurGS   := '25;35;30;30;30;245;5'
    else
      LargeurGS   := '25;-1;30;30;30;250;5';
    ColEditableGS := 'X;-;X;-;-;X;-';
  end;

end;

procedure TOF_BTPREVFAC.GSCellEnter (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
Var Arect   : TRect;
    TheText : String;
begin

  if Cancel then exit;

  if ARow <> GrilleSaisie.Row then Arow := GrilleSaisie.Row;

  ZoneSuivanteOuOk(Acol, ARow, Cancel);

  If ACol = SG_MOIS then
  begin
    TheText             := GrilleSaisie.CellValues[ACol, Arow];
    ARect               := GrilleSaisie.CellRect(SG_MOIS, ARow);
    SaisieDate.Top      := Arect.Top;
    SaisieDate.Left     := Arect.left;
    SaisieDate.Height   := GrilleSaisie.DefaultRowHeight;
    SaisieDate.Width    := GrilleSaisie.ColWidths[SG_MOIS];
    SaisieDate.text     := Copy(theText,5,2) + '-' + Copy(theText,0,4);
    Saisiedate.visible  := True;
    SaisieDate.SetFocus;
  end
  else if ACol = SG_MTPREVU  then
  Begin
  end
  else if Acol = SG_REMARQUE then
    AffichageBtnMemo
  else if Acol = SG_VERROU   then
  begin
    //On doit afficher une petite image de cadenas en fonction si X ou - (Fermé/Ouvert)
  end;                         

end;

procedure TOF_BTPREVFAC.GSCellExit (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
Var Zone    : string;
    ZoneSAV : string;
    TobLPrevFact : TOB;
    MtPrevu : Double;
    MtEcart : Double;
    MtFact  : Double;
    Contenu : TStringList;
    ind     : Integer;
begin

  ZoneSAV := '';
  Zone := '';

  with GestGridSaisie do
  Begin
    TobLPrevFact := TobPrevFact.detail[Arow - 1];
    If TobLPrevFact = nil then Exit;
    //
    if Acol = SG_MOIS           then
    begin
      ZoneSAV := TobLPrevFact.GetValue('BPF_MOIS');
      Zone    := GS.Cells[SG_MOIS, Arow];
      if ZoneSAV <> Zone then TobLPrevFact.PutValue('OKVALIDE', 'N');
      TobLPrevFact.PutValue('BPF_MOIS', zone);
    end
    else if Acol = SG_MTPREVU  then
    Begin
      ZoneSAV         := TobLPrevFact.GetValue('BPF_MTPREVU');

      MtPrevu         := TOBLPrevFact.getValue('BPF_MTPREVU');
      //
      Dbl_TotMtPrevu  := Dbl_TotMtPrevu - MtPrevu;
      //
      MtPrevu         := Valeur(GS.cells[SG_MTPREVU,Arow]);
      MtFact          := Valeur(GS.cells[SG_MTFACT, Arow]);
      //
      MtEcart         := MtFact - MtPrevu;
      GS.cells[SG_MTECART, Arow] := StrF00(MtEcart ,    V_PGI.OkDecP);
      //
      Dbl_TotMtPrevu  := Dbl_TotMtPrevu + MtPrevu;
      Dbl_TotMtEcart  := Dbl_TotMtPrevu - Dbl_TotMtFact;
      //
      Zone            := FloatToStr(MtPrevu);
      //
      ChargePiedWithTotal;
      //
      if ZoneSAV <> Zone then TobLPrevFact.PutValue('OKVALIDE', 'N');
      //
      ToblPrevFact.PutValue('BPF_MTPREVU',  MtPrevu);
      ToblPrevFact.PutValue('MTECART',      MtEcart);
      //

    end
    else if Acol = SG_REMARQUE  then
    Begin
      ZoneSAV := GrilleSaisie.Cells[SG_REMARQUE, Arow];
      if TobLPrevFact.GetValue('BPF_REMARQUE') = '' then
        TobLPrevFact.PutValue('BPF_REMARQUE', ZoneSAV)
      else
      begin
        if ToblPrevFact.GetValue('FLINEREMARQUE') <> ZoneSAV then
        begin
          //on charge FLineRemarque... et on remplace zone grid en 1ère ligne du text...
          Contenu      := TStringList.Create;
          Contenu.Text := TOBLPrevFact.GetValue('BPF_REMARQUE');
          if Contenu = nil then Exit;
          For ind := 0 To Contenu.Count - 1 do
          begin
            if Contenu[Ind] <> '' then
            begin
              Contenu[Ind] := ZoneSAV;
              break;
            end;
          end;
          TobLPrevFact.PutValue('BPF_REMARQUE', Contenu.Text);
          TobLPrevFact.PutValue('OKVALIDE', 'N');
        end;
      end;
    end;
  end;

end;

Procedure TOF_BTPREVFAC.GSKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin

  case Key of
    VK_RETURN : if (Shift = []) then
    begin
    	Key := 0;
      SendMessage(THedit(Sender).Handle, WM_KeyDown, VK_TAB, 0);
    end;
  end;

end;

Procedure TOF_BTPREVFAC.RenumeroteLigneTob(Mois : String);
Var ind           : Integer;
    ToblPrevFact  : TOB;
begin

  with GestGridSaisie do
  begin
    TOBPrevFact.Detail.sort('BPF_MOIS');

    For ind := 0 to TOBPrevFact.Detail.count -1 do
    begin
      TOBPrevFact.detail[Ind].PutValue('NOLIGNE', Ind+1);
    end;

    ChargementGrille;

    //on recherche la ligne en cours de saisie
    ToblPrevFact := TobPrevFact.FindFirst(['BPF_MOIS'], [Mois], False);
    if TobLPrevFact <> nil then
    begin
       GestGridSaisie.GS.Row := TobLPrevFact.GetValue('NOLIGNE');
    end;
  end;

end;

procedure TOF_BTPREVFAC.GSDblClick(Sender : Tobject);
Var Arow : Integer;
begin

  With GestGridSaisie do
  begin
    Arow := GS.Row;
    //Si la ligne est verrouillée
    if GS.Cells[SG_VERROU, Arow]= 'X' then
    begin
      TOBPrevFact.Detail[Arow -1].PutValue('BPF_VERROU', '-');
      GS.Cells[SG_VERROU, Arow]    := '-'
    end
    else
    begin
      TOBPrevFact.Detail[Arow -1].PutValue('BPF_VERROU', 'X');
      GS.Cells[SG_VERROU, Arow]    := 'X';
    end;
    //
    GS.InvalidateRow(Arow);
  end;

end;

procedure TOF_BTPREVFAC.ZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: boolean);
var LastCol : integer;
    LastRow : Integer;
    FirstCol: Integer;
    FirstRow: Integer;
    //
    OldEna  : Boolean;
    ChgSens : Boolean;
begin

  With GestGridSaisie do
  begin
    GS.ElipsisButton := False;
    //
    OldEna := GS.SynEnabled;
    GS.SynEnabled := False;
    //
    LastCol := GS.ColCount-1;
    LastRow := GS.RowCount-1;
    FirstCol:= SG_MOIS;
    FirstRow:= 1;
    //
    ChgSens := False;
    //
    if Acol > GS.col Then ChgSens := True;
    //
    ACol := GS.Col;
    ARow := GS.row;
    //
    if GS.Cells[SG_VERROU, Arow]= '-' then
      GS.ColEditables[SG_MTPREVU]  := True
    else
      GS.ColEditables[SG_MTPREVU]  := False;
    //
    while not ZoneAccessible(ACol, ARow) do
    begin
      //Déplacement de Gauche à Droite
      if not ChgSens then
      begin
        if Acol > LastCol then
        begin
          Acol := FirstCol;
          If Arow < LastRow then
            inc(Arow)
          else
            //La meilleure option serait de créer une nouvelle ligne (!!!!)
            Arow := FirstRow;
        end
        else
        begin
          Inc(Acol)
        end;
      end
      //Déplacement de Droite à Gauche
      else
      Begin
        if Acol < FirstCol then
        begin
          Acol := LastCol;
          if Arow > FirstRow then
            Dec(Arow)
          else
            Arow := FirstRow;
        end
        else
        begin
          Dec(Acol)
        end;
      end;
    end;
    if Acol > Lastcol then
    begin
      Acol := FirstCol;
      If Arow < LastRow then
        inc(Arow)
      else
        //La meilleure option serait de créer une nouvelle ligne (!!!!)
        Arow := FirstRow;
    end;
    //
    GS.col := Acol;
    GS.Row := Arow;
    //
    GS.SynEnabled := OldEna;
    //
  end;

end;

procedure TOF_BTPREVFAC.GSPostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var Arect     : Trect;
    //Decalage  : integer;
    TheText   : string;
    //Variables gestion mois
    DateSaisie  : TDateTime;
    yy, mm : Word;
    TobLPrevFact : TOB;
begin


  With GestGridSaisie do
  Begin
    if GS.RowHeights[ARow] <= 0 then Exit;

    TobLPrevFact := GetTOBGrille;
    If TobLPrevFact = nil then Exit;

    if (ACol < GS.FixedCols) or (Arow < GS.Fixedrows) then Exit;

    ARect := GS.CellRect(ACol, ARow);


    GS.Canvas.Pen.Style   := psSolid;
    GS.Canvas.Brush.Style := BsSolid;

    //Decalage := 0; //Si l'on veux effectuer un décalage au niveau de l'affichage... là 2 caractère sur la gauche.... canvas.TextWidth('W') * 2;

    TheText := GS.Cells[ACol, Arow];
    //
    if (Acol = SG_MOIS) then
    begin
      mm    := StrToInt(Copy(TheText,5,2));
      yy    := StrToInt(Copy(TheText,0,4));
      //
      DateSaisie := EncodeDate(yy, mm, 01);
      //
      TheText := FormatDateTime('mmmm', DateSaisie);
      TheText := TheText + ' ' + IntToStr(yy);
      Canvas.FillRect(ARect);
      GS.Canvas.TextOut (Arect.left+1,Arect.Top+2, Thetext);
    end;
  end;

end;

procedure TOF_BTPREVFAC.GSGetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
Var MtEcart : Double;
begin

  with GestGridSaisie do
  Begin
    if GS.RowHeights[ARow] <= 0 then Exit;
    //
    if (ACol < GS.FixedCols) or (Arow < GS.Fixedrows) then Exit;
    //
    // gestion colonne Montant Prévu
    if GS.ColEditables[Acol] = False then
    begin
      if (ACol = SG_MTPREVU) then
      begin
        if GS.Cells[SG_VERROU, Arow] = '-' then
        begin
          //GS.Canvas.Brush.Color := ClWhite;
          GS.Canvas.Font.Color  := ClBlack;
        end
        else
        begin
          //GS.Canvas.Brush.Color := $f8dfdf;
          GS.Canvas.Font.Color  := ClRed;
        end;
      end
      else if (Acol = SG_MTECART) then
      begin
        GS.Canvas.Brush.Color := $f8dfdf;
        MtEcart := Valeur(GS.Cells[SG_MTECART, Arow]);
        if MtEcart < 0 then
          GS.Canvas.Font.Color := clred
        else
          GS.Canvas.Font.Color := clGreen;
      end
      else
      Begin
        GS.Canvas.Font.Style := Canvas.Font.Style + [fsBold];
        GS.Canvas.Brush.Color := $f8dfdf;
      end;
    end
    Else
    begin
      if (ACol = SG_MTPREVU) then
      begin
        if GS.Cells[SG_VERROU, Arow] = '-' then
        begin
          //GS.Canvas.Brush.Color := ClWhite;
          //GS.Canvas.Font.Color  := ClBlack;
        end
        else
        begin
          //GS.Canvas.Brush.Color := $f8dfdf;
          //GS.Canvas.Font.Color  := ClRed;
        end;
      end
    end;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : F.Vautrain
Créé le ...... : 25/01/2017
Modifié le ... :   /  /
Description .. : Gerstion de l'édition
Mots clefs ... :
*****************************************************************}
procedure TOF_BTPREVFAC.CreateEdition;
begin

  TheType       := 'E';
  TheNature     := 'BPF';
  TheTitre      := 'Prévisionnel Facturation';
  TheModele     := 'PFA';
  //
  OptionEdition := TOptionEdition.Create(TheType,TheNature,TheModele, TheTitre, '', True, True, True, False, False, nil, fEtat);
  //
  OptionEdition.Apercu    := True;
  OptionEdition.DeuxPages := False;
  OptionEdition.Spages    := nil;

end;

procedure TOF_BTPREVFAC.InitTobEdition(ToblEdit : TOB);
begin

  ToblEdit.AddChampSupValeur('CLIENT',        '');
  ToblEdit.AddChampSupValeur('AFFAIRE',       '');
  ToblEdit.AddChampSupValeur('AFFAIRE0',      '');
  ToblEdit.AddChampSupValeur('AFFAIRE1',      '');
  ToblEdit.AddChampSupValeur('AFFAIRE2',      '');
  ToblEdit.AddChampSupValeur('AFFAIRE3',      '');
  ToblEdit.AddChampSupValeur('AVENANT',       '');
  ToblEdit.AddChampSupValeur('LIBCLIENT',     '');
  ToblEdit.AddChampSupValeur('LIBAFFAIRE',    '');
  ToblEdit.AddChampSupValeur('AFFADR1',       '');
  ToblEdit.AddChampSupValeur('AFFADR2',       '');
  ToblEdit.AddChampSupValeur('AFFADR3',       '');
  ToblEdit.AddChampSupValeur('AFFCP',         '');
  ToblEdit.AddChampSupValeur('AFFVILLE',      '');
  //
  ToblEdit.AddChampSupValeur('DATEDEB',       DateToStr(idate1900));
  ToblEdit.AddChampSupValeur('DATEFIN',       DateToStr(idate2099));
  //
  ToblEdit.AddChampSupValeur('MTPLCHARGE',    '0.00');
  ToblEdit.AddChampSupValeur('LINEREMARQUE',  '');
  ToblEdit.AddChampSupValeur('REMARQUE',      '');
  //
  ToblEdit.AddChampSupValeur('TOTMTMARCHE',   '0.00');
  ToblEdit.AddChampSupValeur('TOTMTPLCHARGE', '0.00');
  ToblEdit.AddChampSupValeur('TOTMTFACT',     '0.00');
  ToblEdit.AddChampSupValeur('TOTMTPREVU',    '0.00');
  ToblEdit.AddChampSupValeur('TOTMTECART',    '0.00');
  ToblEdit.AddChampSupValeur('TOTMTRESTAPREVOIR',  '0.00');
  //
end;

procedure TOF_BTPREVFAC.ChargeTobEdition;
Var Ind           : Integer;
    ToblPrevFact  : TOB;
    ToblEdit      : TOB;
    yy,mm         : Word;
    mmmm          : string;
    DateSaisie    : TDateTime;
begin

  //chargement des lignes...
  For Ind := 0 to Tobprevfact.Detail.count - 1 do
  begin
    ToblPrevFact := tobPrevfact.Detail[Ind];
    if ToblPrevFact.GetBoolean('DELETE') Then Continue;
    ToblEdit := TOB.Create('Ligne_Edition', TobEdition, -1);
    ToblEdit.Dupliquer(ToblPrevFact, true, True);
    //
    InitTobEdition(ToblEdit);
    //
    //chargement de l'entête
    ToblEdit.PutValue('CLIENT',        client.Text);
    ToblEdit.PutValue('AFFAIRE',       Affaire.Text);
    ToblEdit.PutValue('AFFAIRE0',      Affaire0.Text);
    ToblEdit.PutValue('AFFAIRE1',      Affaire1.Text);
    ToblEdit.PutValue('AFFAIRE2',      Affaire2.Text);
    ToblEdit.PutValue('AFFAIRE3',      Affaire3.Text);
    ToblEdit.PutValue('AVENANT',       Avenant.Text);
    ToblEdit.PutValue('LIBCLIENT',     Libclient.Caption);
    ToblEdit.PutValue('LIBAFFAIRE',    LibAffaire.caption);
    ToblEdit.PutValue('AFFADR1',       AFFAdr1.Text);
    ToblEdit.PutValue('AFFADR2',       AFFAdr2.Text);
    ToblEdit.PutValue('AFFADR3',       AFFAdr3.Text);
    ToblEdit.PutValue('AFFCP',         AFFCP.Text);
    ToblEdit.PutValue('AFFVILLE',      AFFVille.Text);
    //
    ToblEdit.PutValue('DATEDEB',       DateDeb.Text);
    ToblEdit.PutValue('DATEFIN',       DateFin.Text);
    //
    TOBLEdit.PutValue('OKVALIDE',      TOBLPrevFact.GetValue('OKVALIDE'));
    //
    mm    := StrToInt(Copy(ToblEdit.GetValue('BPF_MOIS'),5,2));
    yy    := StrToInt(Copy(ToblEdit.GetValue('BPF_MOIS'),0,4));
    //
    DateSaisie := EncodeDate(yy, mm, 01);
    //
    mmmm := FormatDateTime('mmmm', DateSaisie);
    //
    ToblEdit.AddChampSupValeur('ANNEE', yy);
    ToblEdit.AddChampSupValeur('MOIS',  mmmm);
    //
    //Chargement du pieds de lignes...
    ToblEdit.PutValue('TOTMTMARCHE',   TotMtMarche.text);
    ToblEdit.PutValue('TOTMTPLCHARGE', TotMtPlCharge.Text);
    ToblEdit.PutValue('TOTMTFACT',     TotMtFact.text);
    ToblEdit.PutValue('TOTMTPREVU',    TotMtPrevu.text);
    ToblEdit.PutValue('TOTMTECART',    TotMtEcart.text);
    ToblEdit.PutValue('TOTMTRESTAPREVOIR',  TotMtResteAPrevoir.text);
    //
  end;

end;

Initialization
  registerclasses ( [ TOF_BTPREVFAC ] ) ;
end.

