{***********UNITE*************************************************
Auteur  ...... : Compta
Créé le ...... : 01/01/1900
Modifié le ... : 12/04/2005
Description .. :
Suite ........ : 02/09/2004 : FQ 14389 Mise en place des boutons plats
Suite ........ : pour le look V6
Suite ........ :
Suite ........ : 12/04/2005 : Modification gestion interface : utilisation des
Suite ........ : possibilité AGL
Mots clefs ... :
*****************************************************************}
unit Echeance;
                    
interface

uses
  {$IFNDEF ERADIO}
  {$IFNDEF EAGLSERVER}
    WinTypes,WinProcs,Messages,Classes,Graphics,Controls,
    Forms,Dialogs,StdCtrls,ExtCtrls,Tabs,Buttons,ComCtrls, Mask,UtilPGI,
    Grids, Menus, HSysMenu, HTB97, FichComm,
  {$ENDIF EAGLSERVER}
  {$ENDIF !ERADIO}

  SysUtils,
  Ent1,
  HEnt1,
  hmsgbox,
  Hctrls,
  UTOB,
  {$IFDEF EAGLCLIENT}
  {$ELSE}
    {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
    DB,
  {$ENDIF}
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  ParamSoc, TntGrids, TntStdCtrls
  ;

{$IFNDEF ERADIO}
{$IFNDEF EAGLSERVER}
  Function  SaisirEcheance ( var Mode : T_ModeRegl ) : boolean ;
{$ENDIF EAGLSERVER}
{$ENDIF !ERADIO}
function CreerModeRegle(var Mode: T_ModeRegl): string;
procedure CalculModeRegle(var Mode: T_ModeRegl; Multi: boolean; TobModeregl : TOB=nil);
//function ProchaineDate(DD: TDateTime; SEP, ARR: String3): TDateTime;
//function NextEche(DD: TDateTime; Separe: String3): TDateTime;
//function EcheArrondie(DRef: TDateTime; ArrondirAu: String3; JP1, JP2: integer): TDateTime;

{$IFNDEF ERADIO}
{$IFNDEF EAGLSERVER}
type
  TFEcheance = class(TForm)
    ButtonPanel: TPanel;
    Panel4: TPanel;
    Label18: TLabel;
    FSaisie: THNumEdit;
    Label19: TLabel;
    FReste: THNumEdit;
    FAPayer: THNumEdit;
    Label17: TLabel;
    FModeRegle: THValComboBox;
    Label1: TLabel;
    Label2: TLabel;
    FModeInit: THValComboBox;
    FListe: THGrid;
    MsgBox: THMsgBox;
    POPS: TPopupMenu;
    BZoom: TToolbarButton97;
    HMTrad: THSystemMenu;
    ISigneEuro: TImage;
    BValide: TToolbarButton97;
    ToolbarButton971: TToolbarButton97;
    BAide: TToolbarButton97;
    BSolde: TToolbarButton97;
    BRepart: TToolbarButton97;
    FVerouiller: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure BSoldeClick(Sender: TObject);
    procedure FModeRegleChange(Sender: TObject);
    procedure BValideClick(Sender: TObject);
    procedure FListeCellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BRepartClick(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure FListeRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure POPSPopup(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BZoomClick(Sender: TObject);
    procedure FListeExit(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FVerouillerClick(Sender: TObject);
  private
    FMode, FOldMode: T_ModeRegl;
    FDebut, MvtLettre: Boolean;
    FTOBModePaie, FTobModeRegle : TOB ;
    procedure ChargeLimiteModePaie;
    procedure VerifLimiteModepaie(ARow: Integer);
    procedure AfficheEcheance;
    function RecupEche: boolean;
    procedure AfficheTotal;
    function  GetLimiteMP( vStCodeMP : String ) : Single ;
    function  GetModeRempMP( vStCodeMP : String ) : String ;
    procedure SetModePaie( vStCodeMP : String ; vRow : Integer ) ;

    procedure Bloque(Bloque : boolean);
  public
  end;
{$ENDIF EAGLSERVER}
{$ENDIF !ERADIO}

implementation

Uses
  {$IFDEF MODENT1}
  CPProcGen,
  {$ENDIF MODENT1}
  SaisUtil,
  UlibEcriture ,
  {$IFNDEF ERADIO}
  {$IFNDEF EAGLSERVER}
    Filtre,
  {$ENDIF EAGLSERVER}
  {$ENDIF !ERADIO}
  SaisComm
  ;

{$IFNDEF ERADIO}
{$IFNDEF EAGLSERVER}
{$R *.DFM}
{$ENDIF EAGLSERVER}
{$ENDIF !ERADIO}

{$IFNDEF ERADIO}
{$IFNDEF EAGLSERVER}
function SaisirEcheance(var Mode: T_ModeRegl): boolean;
var
  FEcheance: TFEcheance;
begin
  FEcheance := TFEcheance.Create(Application);
  try
    FEcheance.FMode := Mode;
    Result := (FEcheance.ShowModal = mrOK);
    Mode := FEcheance.FMode;
  finally
    Fecheance.Free;
  end;
  Screen.Cursor := SyncrDefault;
end;
{$ENDIF EAGLSERVER}
{$ENDIF !ERADIO}


procedure DecalDepartEche(var D: TDateTime; var Mode: T_ModeRegl);
begin
  D := D + Mode.PlusJour;
end;


procedure SepareEche(var D: TDateTime; var Mode: T_ModeRegl);
begin
  D := NextEche(D, Mode.SeparePar);
end;

procedure ArrondirEche(i: integer; var Mode: T_ModeRegl);
var
  D: TDateTime;
begin
  D := Mode.TabEche[i].DateEche;
  D := EcheArrondie(D, Mode.ArrondirAu, Mode.JourPaiement1, Mode.JourPaiement2);
  Mode.TabEche[i].DateEche := D;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 03/03/2003
Modifié le ... :   /  /
Description .. : - 03/03/2003 - msg si Apartir est vide ( paramètrage
Suite ........ : incorrecte )
Mots clefs ... :
*****************************************************************}
function DateDepartEche(var Mode: T_ModeRegl): TDateTime;
var
  D: TDateTime;
begin
  // GP REGL
  if trim(Mode.APartir) = '' then
  begin
    Result := iDate1900;
    PGIError('Paramètrage du mode de règlement du tiers incorrect', 'Attention');
    exit;
  end; // if
  if Mode.APartir = 'BL' then D := Mode.DateBL else
    if Mode.APartir = 'FIN' then D := FinDeMois(Mode.DateFact) else
    if Mode.APartir = 'DEB' then D := DebutDeMois(Mode.DateFact) else
    if Mode.APartir = 'ECR' then D := Mode.DateFact else
    if Mode.APartir = 'FAC' then
  begin
    D := Mode.DateFactExt;
    if D = IDate1900 then D := Mode.DateFact;
  end;
  if Mode.APartir = 'FAF' then
  begin
    D := Mode.DateFactExt;
    if D = IDate1900 then D := Mode.DateFact;
    D := FinDeMois(D);
  end;
  DecalDepartEche(D, Mode);
  Result := D;
end;

procedure InitEcheance(var FEche: T_Eche; DD: TDateTime);
begin
  with FEche do
  begin
    EtatLettrage := 'AL';
    CodeLettre := '';
    LettrageDev := '-';
    DatePaquetMax := DD;
    DatePaquetMin := DD;
    Couverture := 0;
    CouvertureDev := 0; //CouvertureEuro:=0 ;
  end;
  {Reste géré par Saisie.pas}
end;

procedure CalculEcheance(var Mode: T_ModeRegl);
var
  D: TDateTime;
  i: integer;
  StE: string;
  TotP, TotD, XP, XD: double;
begin
  if Mode.ModeGuide then
  begin
    D := Mode.DateFact;
    StE := Mode.EcartJours;
  end else D := DateDepartEche(Mode);
  TotP := 0;
  TotD := 0;
  for i := 1 to Mode.NbEche do
  begin
    if Mode.TabEche[i].PourC <> 0 then
    begin
      if i < Mode.NbEche then
      begin
        XP := Arrondi(Mode.TotalAPayerP * Mode.TabEche[i].Pourc / 100.0, V_PGI.OkDecV);
        XD := Arrondi(Mode.TotalAPayerD * Mode.TabEche[i].Pourc / 100.0, Mode.Decimale);
        TotP := Arrondi(TotP + XP, V_PGI.OkDecV);
        TotD := Arrondi(TotD + XD, Mode.Decimale);
      end else
      begin
        XP := Arrondi(Mode.TotalAPayerP - TotP, V_PGI.OkDecV);
        XD := Arrondi(Mode.TotalAPayerD - TotD, Mode.Decimale);
      end;
      Mode.TabEche[i].MontantP := XP;
      Mode.TabEche[i].MontantD := XD;
    end;
    if not Mode.ModeGuide then
    begin
      Mode.TabEche[i].DateEche := D;
      ArrondirEche(i, Mode);
      SepareEche(D, Mode);
    end else
    begin
      Mode.TabEche[i].DateEche := D + ReadTokenI(StE);
    end;
    if ((Mode.TabEche[i].EtatLettrage <> 'TL') and (Mode.TabEche[i].EtatLettrage <> 'PL')) then InitEcheance(Mode.TabEche[i], Mode.DateFact);
  end;
end;

{============================================================================}
procedure TrouveModeFinal(var Mode: T_ModeRegl);
var
  TotMin: Double;
  ModeMin: String3;
  TabTrouve: array[1..100] of String3;
  Nb, i: Integer;
  okok: Boolean;
  Q: TQuery;
begin
  Q := OpenSQL('SELECT * FROM MODEREGL ORDER BY MR_MODEREGLE', TRUE);
  Mode.ModeFinal := Mode.ModeInitial;
  okok := TRUE;
  Fillchar(TabTrouve, Sizeof(TabTrouve), #0);
  Nb := 1;
  TabTrouve[Nb] := Mode.ModeFinal;
  if Q.Locate('MR_MODEREGLE', Mode.ModeFinal, []) and (Mode.ModeFinal <> '') then // fiche 12590 SBO
  begin
    repeat
      ModeMin := Q.FindField('MR_REMPLACEMIN').AsString;
      TotMin := Q.FindField('MR_MONTANTMIN').AsFloat;
      if Abs(Mode.TotalAPayerP) < TotMin then
      begin
        for i := 1 to nb do if TabTrouve[i] = ModeMin then Okok := TRUE;
        if not OKok then
        begin
          if not Q.Locate('MR_MODEREGLE', ModeMin, [])
            then
            begin
            okok := True;
            ModeMin := Mode.ModeFinal;
          end
          else
          begin
            Mode.ModeFinal := ModeMin;
            Inc(nb);
            TabTrouve[nb] := ModeMin;
          end;
        end;
      end else
      begin
        ModeMin := Mode.ModeFinal;
        okok := TRUE;
      end;
    until okok;
  end else
  begin
    Q.First;
    ModeMin := Q.FindField('MR_MODEREGLE').AsString;
    Mode.ModeInitial := ModeMin;
  end;
  if ModeMin = '' then ModeMin := Mode.ModeInitial;
  Mode.ModeFinal := ModeMin;
  Ferme(Q);
end;

procedure TrouveModeFinalTob(var Mode: T_ModeRegl; TobModeregl : TOB);
var
  TotMin: Double;
  ModeMin: String3;
  TabTrouve: array[1..100] of String3;
  Nb, i: Integer;
  okok: Boolean;
  T1  : TOB;
begin
  Mode.ModeFinal := Mode.ModeInitial;
  okok := TRUE;
  Fillchar(TabTrouve, Sizeof(TabTrouve), #0);
  Nb := 1;
  TabTrouve[Nb] := Mode.ModeFinal;
  T1 := TobModeregl.FindFirst(['MR_MODEREGLE'], [Mode.ModeFinal], FALSE);
  if (T1 <> nil) and  (Mode.ModeFinal <> '') then
  begin
    repeat
      ModeMin := T1.GetValue('MR_REMPLACEMIN');
      TotMin := T1.Getdouble('MR_MONTANTMIN');
      if Abs(Mode.TotalAPayerP) < TotMin then
      begin
        for i := 1 to nb do if TabTrouve[i] = ModeMin then Okok := TRUE;
        if not OKok then
        begin
          T1 := TobModeregl.FindFirst(['MR_MODEREGLE'], [ModeMin], FALSE);
          if T1 = nil then
          begin
            okok := True;
            ModeMin := Mode.ModeFinal;
          end
          else
          begin
            Mode.ModeFinal := ModeMin;
            Inc(nb);
            TabTrouve[nb] := ModeMin;
          end;
        end;
      end else
      begin
        ModeMin := Mode.ModeFinal;
        okok := TRUE;
      end;
    until okok;
  end else
  begin
    ModeMin := TobModeregl.detail[0].Getvalue ('MR_MODEREGLE');
    Mode.ModeInitial := ModeMin;
  end;
  if ModeMin = '' then ModeMin := Mode.ModeInitial;
  Mode.ModeFinal := ModeMin;
end;

{$IFDEF ME}
procedure ChargeModeRegle(var Mode: T_ModeRegl; Multi: boolean);
var
  i: integer;
  Q: TQuery;
  MRDef, MPDef: string;
begin
  Q := OpenSQL('Select * from MODEREGL Where MR_MODEREGLE="' + Mode.ModeFinal + '"', TRUE,-1,'',true);
  {Première vérif sur le nb echéances, si <=0 passer au mode règlement par défaut}
  if not Q.EOF then
  begin
    Mode.NbEche := Q.FindField('MR_NOMBREECHEANCE').AsInteger;
    if Mode.NbEche <= 0 then
    begin
      MRDef := GetParamSocSecur('SO_GCMODEREGLEDEFAUT', False);
      if MRDef <> '' then
      begin
        Ferme(Q);
        Mode.ModeFinal := MRDef;
        Q := OpenSQL('Select * from MODEREGL Where MR_MODEREGLE="' + Mode.ModeFinal + '"', TRUE,-1,'',true);
      end;
    end;
  end;
  if not Q.EOF then
  begin
    Mode.NbEche := Q.FindField('MR_NOMBREECHEANCE').AsInteger;
    Mode.APartir := Q.FindField('MR_APARTIRDE').AsString;
    Mode.ArrondirAu := Q.FindField('MR_ARRONDIJOUR').AsString;
    Mode.SeparePar := Q.FindField('MR_SEPAREPAR').AsString;
    Mode.PlusJour := Q.FindField('MR_PLUSJOUR').AsInteger;
    Mode.ModeGuide := (Q.FindField('MR_MODEGUIDE').AsString = 'X');
    Mode.EcartJours := Q.FindField('MR_ECARTJOURS').AsString;
    FillChar(Mode.TabEche, Sizeof(Mode.TabEche), #0);
    if Mode.NbEche <= 0 then
    begin
      Mode.NbEche := 1;
      Mode.TabEche[1].PourC := 100.0;
      if GetInfoAcceptation.Count > 0 then
      begin
        MPDef := GetInfoAcceptation[0];
        MPDef := ReadTokenSt(MPDef);
        Mode.TabEche[1].ModePaie := MPDef;
      end;
    end;
    if Multi then
    begin
      for i := 1 to Mode.NbEche do
      begin
        Mode.TabEche[i].ModePaie := Q.FindField('MR_MP' + IntToStr(i)).AsString;
        Mode.TabEche[i].PourC := Q.FindField('MR_TAUX' + IntToStr(i)).AsFloat;
      end;
    end else
    begin
      Mode.NbEche := 1;
      Mode.TabEche[1].ModePaie := Q.FindField('MR_MP1').AsString;
      Mode.TabEche[1].PourC := 100.0;
    end;
  end;
  Ferme(Q);
end;
{$ENDIF}

procedure ChargeModeRegle(var Mode: T_ModeRegl; Multi: boolean; TobModeregl : TOB=nil; PourQui : string='');
var
  i: integer;
  MRDef, MPDef: string;
  T1 : TOB;
  Destinataire : string;
begin
  T1 := TobModeregl.FindFirst(['MR_MODEREGLE'], [Mode.ModeFinal], FALSE);
  {Première vérif sur le nb echéances, si <=0 passer au mode règlement par défaut}
  if T1 <> nil then
  begin
    Mode.NbEche := T1.GetValue('MR_NOMBREECHEANCE');
    if Mode.NbEche <= 0 then
    begin
      MRDef := GetParamSocSecur('SO_GCMODEREGLEDEFAUT', False);
      if MRDef <> '' then
      begin
        Mode.ModeFinal := MRDef;
        T1 := TobModeregl.FindFirst(['MR_MODEREGLE'], [Mode.ModeFinal], FALSE);
      end;
    end;
  end;
  if T1 <> nil then
  begin
    Mode.NbEche := T1.GetValue('MR_NOMBREECHEANCE');
    Mode.APartir := T1.GetString('MR_APARTIRDE');
    Mode.ArrondirAu := T1.GetString('MR_ARRONDIJOUR');
    Mode.SeparePar := T1.GetString('MR_SEPAREPAR');
    Mode.PlusJour := T1.GetValue('MR_PLUSJOUR');
    Mode.ModeGuide := (T1.GetString('MR_MODEGUIDE') = 'X');
    Mode.EcartJours := T1.GetString('MR_ECARTJOURS');
    FillChar(Mode.TabEche, Sizeof(Mode.TabEche), #0);
    if Mode.NbEche <= 0 then
    begin
      Mode.NbEche := 1;
      Mode.TabEche[1].PourC := 100.0;
      Mode.TabEche[1].PourQui := PourQui;
      if GetInfoAcceptation.Count > 0 then
      begin
        MPDef := GetInfoAcceptation[0];
        MPDef := ReadTokenSt(MPDef);
        Mode.TabEche[1].ModePaie := MPDef;
      end;
    end;
    if Multi then
    begin
      for i := 1 to Mode.NbEche do
      begin
        Mode.TabEche[i].ModePaie := T1.GetString('MR_MP' + IntToStr(i));
        Mode.TabEche[i].PourC := T1.Getdouble('MR_TAUX' + IntToStr(i));
      	Mode.TabEche[i].PourQui := PourQui;
  			if PourQui = '' then Mode.TabEche[i].PourQuilib := GetParamSocSecur('SO_LIBELLE','Notre Société');
      end;
    end else
    begin
      Mode.NbEche := 1;
      Mode.TabEche[1].ModePaie := T1.GetString('MR_MP1');
      Mode.TabEche[1].PourC := 100.0;
      Mode.TabEche[1].PourQui := PourQui;
      if PourQui = '' then Mode.TabEche[1].PourQuilib := GetParamSocSecur('SO_LIBELLE','Notre Société');
    end;
  end;
end;


procedure CalculModeRegle(var Mode: T_ModeRegl; Multi: boolean; TobModeregl : TOB=nil);
var
TobM    : TOB;
Q       : TQuery;
begin
  if LienS1S3 then Multi := False;

  TobM :=TOB.Create('TOBM',Nil,-1) ;
  Try
    if TobModeregl = nil then
    begin
       Q := OpenSQL('SELECT * FROM MODEREGL ORDER BY MR_MODEREGLE', TRUE,-1,'',true);
       TobM.LoadDetailDB('MODEREGL', '', '', Q, TRUE, FALSE);
       Ferme (Q);
    end
    else
    begin
      TobM.Dupliquer(TobModeregl, TRUE, TRUE, TRUE);
    end;

    TrouveModeFinalTob(Mode, TobM);
    ChargeModeRegle(Mode, Multi, TobM);
    CalculEcheance(Mode);
  Finally
    FreeAndNil (TobM);    // modif gm le 21/12/08 pbm fuite mémoire truovée  en prépa facture SIC
  End;
end;

function CreerModeRegle(var Mode: T_ModeRegl): string;
var
  TM, Q: TQuery;
  i: integer;
  ValSt, Ind, StEcart: string;
begin
  Q := OpenSQL('Select MAX(MR_MODEREGLE) from MODEREGL Where MR_MODEREGLE>="001" AND MR_MODEREGLE<="999"', True,-1,'',true);
  if not Q.EOF then
  begin
    ValSt := Q.Fields[0].AsString;
    if ValSt <> '' then
    begin
      i := StrToInt(ValSt);
      ValSt := IntToStr(i + 1);
    end else ValSt := '001';
    while Length(ValSt) < 3 do ValSt := '0' + ValSt;
  end else ValSt := '001';
  Ferme(Q);
  TM := OpenSQL('SELECT * FROM MODEREGL WHERE MR_MODEREGLE="Trr"', False);
  TM.Insert;
  InitNew(TM);
  TM.FindField('MR_MODEREGLE').AsString := ValSt;
  TM.FindField('MR_LIBELLE').AsString := ValSt;
  TM.FindField('MR_ABREGE').AsString := ValSt;
  TM.FindField('MR_APARTIRDE').AsString := 'ECR';
  TM.FindField('MR_PLUSJOUR').AsInteger := 0;
  TM.FindField('MR_ARRONDIJOUR').AsString := 'PAS';
  TM.FindField('MR_NOMBREECHEANCE').AsInteger := Mode.NbEche;
  TM.FindField('MR_SEPAREPAR').AsString := '1M';
  TM.FindField('MR_MONTANTMIN').AsFloat := 0;
  TM.FindField('MR_REMPLACEMIN').AsString := ValSt;
  TM.FindField('MR_REPARTECHE').AsString := '';
  StEcart := '';
  for i := 1 to Mode.NbEche do
  begin
    Ind := IntToStr(i);
    TM.FindField('MR_MP' + Ind).AsString := Mode.TabEche[i].ModePaie;
    TM.FindField('MR_TAUX' + Ind).AsFloat := Mode.TabEche[i].Pourc;
    StEcart := StEcart + IntToStr(Round(Mode.TabEche[i].DateEche - Mode.DateFact)) + ';';
  end;
  TM.FindField('MR_ECARTJOURS').AsString := StEcart;
  TM.FindField('MR_MODEGUIDE').AsString := 'X';
  TM.Post;
  Ferme(TM);
  AvertirTable('ttModeRegle');
  Result := ValSt;
end;

{$IFNDEF ERADIO}
{$IFNDEF EAGLSERVER}
procedure TFEcheance.FormShow(Sender: TObject);
var i: integer;
Q    : TQuery;
begin

  // Chargement des limites de valeurs / mode de paiement
  FTobModePaie := TOB.Create( 'V_MODEPAIE', nil, -1 ) ;
  ChargeLimiteModePaie;

  // Paramétrage de la liste
  FListe.ColFormats[1] := 'CB=TTMODEPAIE||' ;

  // Montant de l'échéance
  FListe.ColTypes[3]  := 'R' ;         // Montant
  FListe.ColFormats[3] := '#,#00.00' ;  // Montant

  // Libellé du bénéficiaire
  FListe.ColEditables [2] := false;  // libellé

  // Date d'échéance
  FListe.ColTypes[4]   := 'D';
  FListe.ColFormats[4] := ShortDateFormat;
  FListe.ColAligns[4]  := taCenter;

  MvtLettre := False;
  FOldMode := FMode;

  ChangeMask(FApayer, FMode.Decimale, FMode.Symbole);
  ChangeMask(FReste, FMode.Decimale, FMode.Symbole);
  ChangeMask(FSaisie, FMode.Decimale, FMode.Symbole);

  if not assigned(FTobModeRegle) then
  begin
    Q := OpenSQL('SELECT * FROM MODEREGL ORDER BY MR_MODEREGLE', TRUE,-1,'',true);
    try
      FTobModeRegle :=TOB.Create('TOBM',Nil,-1) ;
      if not Q.Eof then
        FTobModeRegle.LoadDetailDB('MODEREGL', '', '', Q, TRUE, FALSE);
    finally
      Ferme (Q);
    end;
  end;

  if FMode.Action <> taCreat then
    begin
    FDebut := TRUE;
    FModeInit.Value := FMode.ModeInitial;
    FModeRegle.Value := FMode.ModeFinal;
    if FMode.Action = TaModif then
      for i := 1 to FMode.NbEche do
        if FMode.TabEche[i].CodeLettre <> '' then MvtLettre := True;
  end else
  begin
    FDebut := FALSE;
    FModeInit.Value := FMode.ModeInitial;
    TrouveModeFinal(FMode);
    FModeRegle.Value := FMode.ModeFinal;
  end;
  FDebut := FALSE;

{$IFDEF GCGC}
 {$IFNDEF CHR}
  FVerouiller.Visible := True;
  FVerouiller.Checked := FMode.Bloque;
  FVerouiller.Enabled := (FMode.Action <> taConsult);
  Bloque(FMode.Bloque);
 {$ELSE CHR}
  FMode.Bloque := False;
  FVerouiller.Checked := False;
 {$ENDIF !CHR}
{$ELSE GCGC}
  FVerouiller.Visible := False;
  FVerouiller.Checked := False;
{$ENDIF GCGC}

  if FMode.Action = taConsult then
    Bloque(True);

  if MvtLettre then
  begin
    FModeRegle.Enabled := False;
    BRepart.Enabled := False;
  end;
end;

procedure TFEcheance.BSoldeClick(Sender: TObject);
begin
  if FMode.Action = taConsult then Exit;
  if GoRowSelect in FListe.Options then Exit;
  FListe.Cells[3, FListe.Row] := StrfMontant(Valeur(FListe.Cells[3, FListe.Row]) + FReste.Value, 15, FMode.Decimale, '', TRUE);
  AfficheTotal;
end;

procedure TFEcheance.FModeRegleChange(Sender: TObject);
begin
  if MvtLettre then Exit;
  if not FDebut then
  begin
    FMode.ModeFinal := FModeRegle.Value;
    ChargeModeRegle(FMode, True, FTobModeRegle);
    CalculEcheance(FMode);
  end;
  AfficheEcheance;
end;

procedure TFEcheance.AfficheEcheance;
var i: integer;
begin
  for i := 1 to 12 do
    begin
    FListe.Cells[0, i] := IntToStr(i);
    FListe.Cells[1, i] := '';
    FListe.Cells[2, i] := '';
    FListe.Cells[3, i] := '';
    FListe.Cells[4, i] := '';
    if i <= FMode.NbEche then
      begin
      // Maj mode de paiement
      SetModePaie( FMode.TabEche[i].ModePaie, i ) ;
      // Pour Qui
      FListe.Cells[2, i]      := FMode.TabEche[i].PourQuiLib;
      // Maj montant
      FListe.Cells[3, i]      := StrfMontant(FMode.TabEche[i].MontantD, 15, FMode.Decimale, '', TRUE);
      // Maj date
      FListe.Cells[4, i]      := DateToStr(FMode.TabEche[i].DateEche);
      end;
    end;
  AfficheTotal;
end;

procedure TFEcheance.AfficheTotal;
var
  i: integer;
  X, Y: Double;
begin
  FAPayer.Value := FMode.TotalAPayerD;
  X := 0;
  for i := 1 to 12 do
    begin
    Y := Valeur(FListe.Cells[3, i]);
    if Y = 0
      then FListe.Cells[3, i] := ''
      else FListe.Cells[3, i] := StrfMontant(Y, 15, FMode.Decimale, '', TRUE);
    X := X + Y;
    end;
  FReste.Value := FAPayer.Value - X;
  FSaisie.Value := X;
end;

procedure TFEcheance.BValideClick(Sender: TObject);
begin
  {JP 30/06/05 : FQ 16087 : la correction de la FQ 14822 n'est pas suffisante : avant de lancer la sortie
                 de grille, on force un changement de contrôle}
  NextPrevControl(TForm(Self));

  //  BBI pour corriger le fait que si on saisit une date et qu'on valide sans sortir
  //      de la cellule la saisie est perdue.
  FListeExit(nil);          //FQ 14822 SG6 : 02/11/2004
  if ((FMode.Action = taConsult) or (RecupEche)) then ModalResult := mrOk;
end;

function TFEcheance.RecupEche: boolean;
var
  i, nb, Last: Integer;
  totD, TotP, XX: double;
  DD: TDateTime;
begin
  AfficheTotal;
  nb := 0;
  TotD := 0;
  TotP := 0;
  Result := Arrondi(FReste.Value, FMode.Decimale) = 0;
  Last := 0;
  if not Result then
  begin
    MsgBox.Execute(0, '', '');
    exit;
  end;
  {Contrôles de cohérence}
  for i := 1 to MaxEche do
    if ((FListe.Cells[1, i] <> '') and (FListe.Cells[3, i] <> '') and (FListe.Cells[4, i] <> '')) then
    begin
    (*
      if ((not (GetParamSocSecur('SO_MONTANTNEGATIF',False))) and (Valeur(FListe.Cells[2, i]) < 0)) then
      begin
        Result := False;
        MsgBox.Execute(2, '', '');
        Exit;
      end;
    *)
      DD := StrToDate(FListe.Cells[4, i]);
      if not NbJoursOK(FMode.DateFact, DD) then
      begin
        Result := False;
        MsgBox.Execute(3, '', '');
        Exit;
      end;
    end;
  for i := 1 to MaxEche do
    if ((FListe.Cells[1, i] <> '') and (FListe.Cells[3, i] <> '') and (FListe.Cells[4, i] <> '')) then
    begin
      Inc(nb);
      FMode.TabEche[nb].DateEche := StrToDate( FListe.Cells[4, i] );
      FMode.TabEche[nb].ModePaie := FListe.CellValues[1, i] ;
      XX := Valeur(FListe.Cells[3, i]);
      if GetParamSocSecur('SO_TENUEEURO',True) then
      begin
        {Compta tenue en Euro}
        if FMode.CodeDevise = V_PGI.DevisePivot then
        begin
          {Saisie en Euro nécessairement}
          FMode.TabEche[nb].MontantP := XX;
          FMode.TabEche[nb].MontantD := XX;
        end else
        begin
          {Saisie en devise}
          FMode.TabEche[nb].MontantP := DeviseToEuro(XX, FMode.TauxDevise, FMode.Quotite);
          FMode.TabEche[nb].MontantD := XX;
        end;
      end else
      begin
        {Saisie en franc ou devise}
        FMode.TabEche[nb].MontantP := DeviseToPivot(XX, FMode.TauxDevise, FMode.Quotite);
        FMode.TabEche[nb].MontantD := XX;
      end;
      TotP := TotP + FMode.TabEche[nb].MontantP;
      TotD := TotD + FMode.TabEche[nb].MontantD;
      Last := nb;
      if ((FMode.TabEche[Nb].EtatLettrage <> 'TL') and (FMode.TabEche[Nb].EtatLettrage <> 'PL')) then InitEcheance(FMode.TabEche[Nb], FMode.DateFact);
    end;
  FMode.NbEche := nb;
  if Last > 0 then
  begin
    FMode.TabEche[Last].MontantP := FMode.TabEche[Last].MontantP + (FMode.TotalApayerP - TotP);
    FMode.TabEche[Last].MontantD := FMode.TabEche[Last].MontantD + (FMode.TotalApayerD - TotD);
  end;
  if FMode.TotalAPayerP <> 0 then for i := 1 to MaxEche do FMode.TabEche[i].Pourc := Arrondi(100.0 * FMode.TabEche[i].MontantP / FMode.TotalAPayerP, ADecimP);
end;

procedure TFEcheance.VerifLimiteModepaie(ARow: Integer);
var ModeRemp : string;
    S        : Single;
    rep      : Integer ;
begin
  if FMode.Action = taConsult then Exit;

  if FListe.CellValues[1, Arow] = '' then Exit ;

  ModeRemp := GetModeRempMP( FListe.CellValues[1, Arow] ) ;
  S        := GetLimiteMP(   FListe.CellValues[1, Arow] ) ;

  // Dépasse-t-on la limite fixé pour le mode de paiement
  if ( S > 0 ) and ( S < Valeur(FListe.Cells[3, ARow]) ) and ( ModeRemp <> '' ) then
    begin
    rep := MsgBox.execute(1, '', '');
    case Rep of
      // On remplace le mode de paiement
      mrYes : SetModePaie( ModeRemp, ARow ) ;
      // On stagne au max pour le mode de paiement
      mrNo  : FListe.Cells[3, ARow] := StrfMontant(S, 15, FMode.Decimale, '', TRUE);
      end;

    end;

end;

procedure TFEcheance.FListeCellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
begin

  if FMode.Action = taConsult then Exit;
  if ((FListe.Cells[1, ARow] = '') and (FListe.Cells[3, ARow] = '') and (FListe.Cells[4, ARow] = '')) then exit;

  // MAJ mode paiement => plus nécessaire
  if ( FListe.Cells[1, ARow] = '' ) and ( FTobModePaie.Detail.count > 0 ) then
    SetModePaie( FTobModePaie.Detail[0].GetValue('MP_MODEPAIE'), ARow ) ;

  // MAJ montant
  if FListe.Cells[3, ARow] = '' then
    FListe.Cells[3, ARow] := StrfMontant(FReste.Value, 15, FMode.Decimale, '', TRUE);

  // MAJ date
  if FListe.Cells[4, ARow] = '' then
    FListe.Cells[4, ARow] := DateToStr(FMode.DateFact);

  VerifLimiteModePaie(ARow);
  AfficheTotal;

end;


procedure TFEcheance.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if FMode.Action <> taConsult then CanClose := RecupEche else CanClose := True;
end;

procedure TFEcheance.BRepartClick(Sender: TObject);
var
  i, nb, Last: integer;
begin
  if MvtLettre then Exit;
  if FMode.Action = taConsult then Exit;
  nb := 0;
  Last := 1;
  for i := 1 to 12 do if FListe.Cells[1, i] <> '' then inc(nb);
  if nb <= 0 then exit;
  for i := 1 to 12 do if FListe.Cells[1, i] <> '' then
    begin
    FListe.Cells[3, i] := StrfMontant(FAPayer.Value / nb, 15, FMode.Decimale, '', TRUE);
    Last := i;
    end;
  AfficheTotal;
  FListe.Row := Last;
  BSoldeClick(nil);
end;

procedure TFEcheance.BAbandonClick(Sender: TObject);
begin
  if FMode.Action <> taCreat then
  begin
    FMode := FOldMode;
    AfficheEcheance;
  end;
end;

procedure TFEcheance.FListeRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
begin
  // Gestion droit modif sur la ligne
  if ( FMode.TabEche[FListe.Row].ReadOnly and (FMode.Action <> taCreat) )
     or ( FMode.Action = taConsult )
    // ne peut pas modifer :
    then FListe.Options := FListe.Options - [goEditing] + [goRowSelect]
    // peut modifer :
    else FListe.Options := FListe.Options + [goEditing] - [goRowSelect] ;
end;

procedure TFEcheance.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_ESCAPE: if Shift = [] then
      begin
        Key := 0;
        BAbandonClick(nil);
      end;
    VK_F6: if Shift = [] then
      begin
        Key := 0;
        BSoldeClick(nil);
      end;
    VK_F10: if Shift = [] then
      begin
        Key := 0;
        NextControl(Self);
        BValideClick(nil);
      end;
    {AR}82: if Shift = [ssAlt] then
      begin
        Key := 0;
        BRepartClick(nil);
      end;
    {JP 02/08/05 : FQ 16071 : Mise en place du ctrl + Suppr}
    VK_DELETE : if Shift = [ssCtrl] then begin
                  Key := 0;
                  FListe.DeleteRow(FListe.Row);
                end;
  end;
end;

procedure TFEcheance.POPSPopup(Sender: TObject);
begin
  InitPopup(Self);
end;

procedure TFEcheance.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  PurgePopup(POPS);
 // if FTobModePaie.Detail.count > 0 then
 //   FTobModePaie.Detail.Clear ;
  FreeAndNil( FTobModePaie ) ;
end;

procedure TFEcheance.BZoomClick(Sender: TObject);
begin
  FicheRegle_AGL(FModeRegle.Value, False, taConsult);
end;

procedure TFEcheance.FListeExit(Sender: TObject);
var
  Acol, Arow: longint;
  Cancel: boolean;
begin
  if FMode.Action = taConsult then Exit;
  ACol := FListe.Col;
  ARow := FListe.Row;
  Cancel := False;
  FListeCellExit(nil, ACol, ARow, Cancel);
end;

procedure TFEcheance.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

procedure TFEcheance.ChargeLimiteModePaie;
begin
  FTobModePaie.ClearDetail ;
  FTobModePaie.LoadDetailFromSQL( 'SELECT MP_MODEPAIE, MP_CONDITION, MP_MONTANTMAX, MP_REMPLACEMAX FROM MODEPAIE ORDER BY MP_MODEPAIE' ) ;
end;

function TFEcheance.GetLimiteMP( vStCodeMP : String ) : Single;
var lTobMP : Tob ;
begin

  result := 0.0 ;
  if vStCodeMP = '' then Exit ;

  lTobMP := FTobModePaie.FindFirst( ['MP_MODEPAIE'], [ vStCodeMP ], False ) ;
  if lTobMP = nil then Exit ;

  if lTobMP.GetValue('MP_CONDITION') = 'X'
      then result := lTobMP.GetValue('MP_MONTANTMAX') ;

  if GetParamSocSecur('SO_TENUEEURO',True)
      then result := EuroToDevise(  result, FMode.TauxDevise, FMode.Quotite, FMode.Decimale )
      else result := PivotToDevise( result, FMode.TauxDevise, FMode.Quotite, FMode.Decimale ) ;
end;

function TFEcheance.GetModeRempMP(vStCodeMP: String): String;
var lTobMP : Tob ;
begin
  result := '' ;
  if vStCodeMP = '' then Exit ;

  lTobMP := FTobModePaie.FindFirst( ['MP_MODEPAIE'], [ vStCodeMP ], False ) ;
  if lTobMP = nil then Exit ;

  if lTobMP.GetValue('MP_CONDITION') = 'X'
      then result := lTobMP.GetValue('MP_REMPLACEMAX') ;
end;

procedure TFEcheance.SetModePaie( vStCodeMP : String; vRow : Integer ) ;
begin
  if vStCodeMP = '' then
    begin
    FListe.CellValues [1, vRow] := '' ;
    FListe.Cells      [1, vRow] := '' ;
    end
  else
    begin
    // Maj code de l'item sélectionné dans la combo
    FListe.CellValues [1, vRow] := vStCodeMP;
    // Maj texte affiché
    FListe.Cells      [1, vRow] := RechDom( 'TTMODEPAIE', vStCodeMP, False ) ;
    end ;
end;

procedure TFEcheance.Bloque(Bloque : boolean);

  procedure GereAcces(Active : boolean);
  var Cpt : integer;
  begin
    for Cpt := 1 to 3 do
      FListe.ColEditables[Cpt] := Active;
    if Active then
      FListe.Options := FListe.Options + [goEditing] - [goRowSelect]
      else
      FListe.Options := FListe.Options + [goEditing] - [goRowSelect];
  end;


begin
{$IFDEF GCGC}
 {$IFNDEF CHR}
  FMode.Bloque := Bloque;
 {$ELSE CHR}
  FMode.Bloque := False;
 {$ENDIF !CHR}
{$ENDIF GCGC}
  FModeRegle.Enabled := (not Bloque);
  if FMode.Action = taConsult then
    GereAcces(False)
    else
    GereAcces(not FVerouiller.Checked);
end;

procedure TFEcheance.FVerouillerClick(Sender: TObject);
begin
  Bloque(FVerouiller.Checked);
end;

{$ENDIF EAGLSERVER}
{$ENDIF !ERADIO}


end.

