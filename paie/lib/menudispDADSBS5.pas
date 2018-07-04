{***********UNITE*************************************************
Auteur  ...... : VG
Créé le ...... : 23/09/2004
Modifié le ... :   /  /
Description .. : Menu de la paie
Mots clefs ... : PGDADSB
*****************************************************************}

unit menudispDADSBS5;

interface
uses HEnt1,
  Forms,
  sysutils,
  HMsgBox,
  Classes,
  Controls,
  HPanel,
  UIUtil,
//     Windows,
  Hdebug,
{$IFDEF EAGLCLIENT}
  utileagl,
  MenuOLX,
  MaineAgl,
  eTablette,
{$ELSE}
  uedtcomp,
  Tablette,
  FE_Main,
  MenuOLG,
{$ENDIF}
  AGLInit,
//     PgiExec,
  EntPgi,
//     ParamSoc,
  Hctrls,
  Ent1,
  ImgList,
  AglInitPlus,
  UtilPGI
  ;

procedure InitApplication;
type
  TFMenuDisp = class(TDatamodule)
  end;

var
  FMenuDisp: TFMenuDisp;

implementation

uses PgOutils2
  ;

{$R *.DFM}
// Procedure qui indique que le dossier en mode nomade . Uniquement PCL !!!

procedure Dispatch(Num: Integer; PRien: THPanel; var retourforce, sortiehalley: boolean);
begin
{$IFNDEF VG}
{ PLUS DE SERIA donc pas de setseria et plus de version de démo
  if ((Num >= 41220) and (V_PGI.VersionDemo = TRUE)) then
  begin
    PGIBox('Version non sérialisée', TitreHalley);
    exit;
  end;
}
{$ENDIF}
  V_PGI.VersionDemo := FALSE;
  case Num of
    10: ;
    11: ;
    12: ;
{    begin // avant connexion
        if (V_PGI.ModePcl = '1') then
          FMenuG.SetSeria('00280011', ['05051060'], ['TD Bilatéral']) //PCL
        else
          FMenuG.SetSeria('05970012', ['05051060'], ['TD Bilatéral']); // Entreprise
      end;
}
    13: ;
    16: ;
    100: ; // action A Faire quand on passe par le lanceur


    41220: begin
        AglLanceFiche('PAY', 'ETABLISSEMENT', '', '', '');
        MajEtabCompl; // Analyse et controle des champs à NULL
      end;

    {------------------------------------------------------------
        Gestion des envois des concernant le social DADS-B
    -------------------------------------------------------------}
    41817: AglLanceFiche('PAY', 'MUL_ENVOISOCIAL', '', '', 'DADSB');
    41818: AglLanceFiche('PAY', 'MUL_CONSULTENVSOC', '', '', 'DADSB');
    41819: AglLanceFiche('PAY', 'MUL_ENVOISOCIAL', '', '', 'DADSBP');

    {------------------------------------------------------------
                          GESTION DES DADS-B
    ------------------------------------------------------------}
    42563: AglLanceFiche('PAY', 'MUL_DADS2TIERS', '', '', '');
    42567: AglLanceFiche('PAY', 'MUL_DADS2', '', '', 'S');
    42568: AglLanceFiche('PAY', 'MUL_DADS2_HONOR', '', '', 'S');
    42571: AglLanceFiche('PAY', 'DADSB_EDITION', '', '', '');
    42576: AglLanceFiche('PAY', 'DADS2_PREP', '', '', '');
  end;
end;

procedure DispatchTT(Num: Integer; Action: TActionFiche; Lequel, TT: string; Range: string);
begin
  case Num of
    1: ;
  end;
end;

procedure ZoomEdt(Qui: integer; Quoi: string);
begin
  Screen.Cursor := crDefault;
end;

procedure ZoomEdtEtat(Quoi: string);
var
  i, Qui: integer;
begin
  i := Pos(';', Quoi);
  Qui := StrToInt(Copy(Quoi, 1, i - 1));
  Quoi := Copy(Quoi, i + 1, Length(Quoi) - i);
  ZoomEdt(Qui, Quoi);
end;


procedure AfterChangeModule(NumModule: Integer);
begin
  FMenuG.RemoveGroup(-41100, TRUE);
  FMenuG.RemoveGroup(-41300, TRUE);
  FMenuG.RemoveGroup(-41400, TRUE);
  FMenuG.RemoveGroup(-41600, TRUE);
  FMenuG.RemoveGroup(-41700, TRUE);

  FMenuG.RemoveItem(41210);
  FMenuG.RemoveItem(41230);
  FMenuG.RemoveItem(41240);
  FMenuG.RemoveItem(41250);
  FMenuG.RemoveItem(41260);
  FMenuG.RemoveItem(41270);
  FMenuG.RemoveItem(-41280);
  FMenuG.RemoveItem(41290);
  FMenuG.RemoveItem(41802);
  FMenuG.RemoveItem(-41810);
  FMenuG.RemoveItem(-41820);
  FMenuG.RemoveItem(-41840);
  FMenuG.RemoveItem(-41860);
  FMenuG.RemoveItem(-41870);

  FMenuG.RemoveGroup(-42100, TRUE);
  FMenuG.RemoveGroup(-42200, TRUE);
  FMenuG.RemoveGroup(-42300, TRUE);
  FMenuG.RemoveGroup(-42400, TRUE);
  FMenuG.RemoveGroup(-42450, TRUE);
  FMenuG.RemoveGroup(-42500, TRUE);
  FMenuG.RemoveItem(42562);
  FMenuG.RemoveGroup(-42600, TRUE);
  FMenuG.RemoveGroup(-42700, TRUE);
  FMenuG.RemoveGroup(-42800, TRUE);
  FMenuG.RemoveGroup(-42900, TRUE);
  FMenuG.RenameItem (42, 'TD Bilatéral');
end;

procedure AfterProtec(sAcces: string); //Sérialisation des modules
begin
// cas où on coche "version non sérialisée' dans
// l'écran de séria. On force la non sérialisation
// PH il n'y a plus de séria donc plus de version de démo
{
  if V_PGI.NoProtec then
    V_PGI.VersionDemo := TRUE
  else
    V_PGI.VersionDemo := False;
  if (sAcces[1] = '-') then
    V_PGI.VersionDemo := TRUE;
}
  V_PGI.VersionDemo := FALSE;
  FMenuG.SetModules([42, 41], [42, 41]);
end;

procedure AGLAfterChangeModule(Parms: array of variant; nb: integer);
begin
  AfterChangeModule(Integer(Parms[0]));
end;

function ReadParam(var St: string): string;
var
  i: Integer;
begin
  i := Pos(';', St);
  if i <= 0 then
    i := Length(St) + 1;
  result := Copy(St, 1, i - 1);
  Delete(St, 1, i);
end;

function FunctPGDADSB(sf, sp: string): variant;
var
  BisTer, NomRue, Num, periode, TypeFonc: string;
begin
  result := '';
  TypeFonc := ReadParam(sp);

  if (TypeFonc = 'PER') then
  begin
    periode := (ReadParam(sp));
    if (periode <> '') then
      result := copy(periode, 1, 2) + '/' + copy(periode, 3, 2) + '    ' +
        copy(periode, 5, 2) + '/' + copy(periode, 7, 2);
  end
  else
    if (TypeFonc = 'ADR') then
    begin
      Num := ReadParam(sp);
      BisTer := ReadParam(sp);
      NomRue := ReadParam(sp);
      BisTer := RechDom('PGADRESSECOMPL', BisTer, False);
      result := Num + ' ' + BisTer + ' ' + NomRue;
    end;
end;

function CalcOLEEtat(sf, sp: string): variant;
var
  Val, Table: string;
begin
  if sf = 'RECHDOM' then
  begin
    Val := Trim(ReadParam(sp));
    Table := Trim(ReadParam(sp));
    if (Table <> '') and (Val <> '') then
      Result := RechDom(Table, val, False)
    else
      result := '';
    exit;
  end
  else
    if Sf = 'DADSB' then
      result := FunctPGDADSB(sf, sp);
end;

procedure InitApplication;
begin
  ProcCalcEdt := CalcOLEEtat;
  FMenuG.OnDispatch := Dispatch;
  FMenuG.OnChargeMag := ChargeMagHalley;
{$IFNDEF EAGLCLIENT}
  FMenuG.OnMajAvant := nil;
  FMenuG.OnMajApres := nil;
{$ENDIF}
{
  if (V_PGI.ModePcl = '1') then
    FMenuG.SetSeria('00280011', ['05051060'], ['TD Bilatéral']) //PCL
  else
    FMenuG.SetSeria('05970012', ['05051060'], ['TD Bilatéral']); // Entreprise
}
  FMenuG.OnAfterProtec := AfterProtec;
  FMenuG.SetModules([42, 41], [42, 41]);
  FMenuG.OnChangeModule := AfterChangeModule;
  V_PGI.DispatchTT := DispatchTT;
end;

{ TFMenuDisp }

procedure InitLaVariablePGI;
begin
  Apalatys := 'CEGID';
  HalSocIni := 'cegidpgi.ini';
  Copyright := '© Copyright ' + Apalatys;
  V_PGI.OutLook := TRUE;
  TitreHalley := 'TD Bilatéral';
{$IFNDEF CPS3}
  V_PGI.LaSerie := S5;
  NomHalley := 'CPTDBS5';
  Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CCS5.HLP';
{$ELSE}
  V_PGI.LaSerie := S3;
  NomHalley := 'CPTDBS3';
  Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CCS5.HLP';
{$ENDIF}
  V_PGI.OfficeMsg := FALSE;
  V_PGI.ToolsBarRight := TRUE;
  ChargeXuelib;
  if V_PGI.ModePcl = '1' then
     begin
{$IFNDEF EAGLCLIENT}
     V_PGI.ToolBarreUp := False;
     V_PGI.PortailWeb := 'http://experts.cegid.fr/home.asp';
{$ENDIF}
     V_PGI.CompterLesimpressions:= true;
     V_PGI.PRocCompteImpression:= CompteImpression;
     end
  else
    V_PGI.PortailWeb := 'http://utilisateurs.cegid.fr';
  V_PGI.OfficeMsg := TRUE;
  V_PGI.VersionDemo := True;
  V_PGI.MenuCourant := 0;
  V_PGI.VersionReseau := False;
  V_PGI.NumVersion := '8.00';
  V_PGI.NumBuild := '000.000';
  V_PGI.NumVersionBase := 832;
  V_PGI.DateVersion := EncodeDate(2007, 03, 29);
  V_PGI.ImpMatrix := True;
  V_PGI.OKOuvert := FALSE;
  V_PGI.Halley := TRUE;
  V_PGI.MenuCourant := 0;
  V_PGI.DebugSQL := False;
  V_PGI.Debug := False;
  V_PGI.QRPdf := True;
  V_PGI.NumMenuPop := 27;
  V_PGI.CegidBureau := True;
  V_PGI.CegidApalatys := FALSE;
// pour autoriser l'acces au dp comme dossier de travail,au multi dossier
  V_PGI.StandardSurDP := true;
// pour autoriser l'acces au dp comme base de reference pour tous les dossiers ==> Table DP+STD
  V_PGI.MajPredefini := True;
// Confidentialité
  V_PGI.PGIContexte := [CtxPaie];
  V_PGI.RAZForme := TRUE;
// Blocage des mises à jour de structure par l'appli. Elle doit etre faite par PgiMajVEr fourni dans le kit socref
  V_PGI.BlockMAJStruct := TRUE;
  v_pgi.sav := False;
end;

initialization
  ProcChargeV_PGI := InitLaVariablePGI;
end.

