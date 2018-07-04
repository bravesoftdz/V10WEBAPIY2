{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 11/03/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : STDENREGPLAN ()
Mots clefs ... : TOF;STDENREGPLAN
*****************************************************************}
unit uTofStdEnregPlan;

interface

uses StdCtrls,
  Controls,
  Classes,
{$IFDEF EAGLCLIENT}
  uLanceProcess, // LanceProcessServer
  MainEagl,      // AGLLanceFiche
{$ELSE}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  FE_Main,
{$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  Spin,
  ParamSoc,
  UTOB,
  Ent1;

type
  TOF_STDENREGPLAN = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  private
    FNumStd: TSpinEdit;
    FNumPlanCompte : integer;
    FListeStd : TOB;
    procedure OnChangeNumStd (Sender : TObject);
  end;

procedure CPLanceFiche_EnregistrementSurStandard( vInNumPlan : integer );

implementation

uses uLibStdCpta;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/05/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure CPLanceFiche_EnregistrementSurStandard( vInNumPlan : integer );
begin
  AGLLanceFiche ('CP','STDENREGPLAN','','', IntToStr(vInNumPlan) );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... :   /  /
Modifié le ... : 03/05/2005
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_STDENREGPLAN.OnArgument(S: string);
var Q: TQuery;
begin
  inherited;

  FNumPlanCompte := StrToInt(ReadTokenSt(S)); 

  FNumStd := TSpinEdit (GetControl ('NUMSTD'));
  if EstSpecif('51502') then
    FNumStd.MinValue := 1;

  if FNumPlanCompte <> 0 then
  begin
    if (FNumStd.MinValue = 21) and (FNumPlanCompte < 21) then
    begin
      Q := OpenSQL('SELECT MAX(STC_NUMPLAN) AS MAXNO FROM STDCPTA', True);
      if not Q.EOF then
        FNumStd.value := Q.Fields[0].AsInteger + 1;
    end
    else
    begin
      Q := OpenSql('SELECT STC_LIBELLE, STC_ABREGE FROM STDCPTA WHERE STC_NUMPLAN=' +
                   IntToStr(FNumPlanCompte), TRUE);
      if not Q.EOF then
      begin
        SetControlText('LIBELLESTD', Q.FindField('STC_LIBELLE').asstring);
        SetControlText('ABREGESTD', Q.FindField('STC_ABREGE').asstring);
      end;
      FNumStd.Value := FNumPlanCompte;
    end;
    Ferme(Q);
  end;
  FListeStd := TOB.Create ('', nil, -1);
  FListeStd.LoadDetailDB('STDCPTA','','',nil,False);
  FNumStd.OnChange := OnChangeNumStd;

  THLabel(GetControl ('LCOMMENTAIRE1', True)).Visible := False;
  THLabel(GetControl ('LCOMMENTAIRE2', True)).Visible := False;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_STDENREGPLAN.OnLoad;
begin
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_STDENREGPLAN.OnNew;
begin
  inherited;
end;

procedure TOF_STDENREGPLAN.OnDelete;
begin
  inherited;
end;

procedure TOF_STDENREGPLAN.OnUpdate;
{$IFDEF EAGLCLIENT}
var lTobParam  : Tob;
    lTobResult : Tob;
{$ENDIF}
begin
  inherited;
  if (not EstSpecif('51502')) and (FNumStd.Value < 21 ) then
  begin
    PGIBox('Le numéro du dossier type doit être supérieur à 21.#10#13Les numéros inférieurs à 21 sont réservés à Cegid.', ECRAN.Caption);
    exit;
  end;
  if GetControlText('LIBELLESTD') = '' then
  begin
    PGIBox('Libellé incorrect', ECRAN.Caption);
    exit;
  end;
  if ExisteSQL('SELECT * FROM STDCPTA WHERE STC_NUMPLAN=' +
    IntToStr(FNumStd.Value)) then
    if (PGIAsk('Le dossier type existe déjà, voulez-vous le remplacer ?', Ecran.Caption)
      <> mrYes) then
      exit;
  if not ExisteSQL('SELECT * FROM MODEREGL') then
  begin
    PGIBox ('Vous devez renseigner au moins un mode de réglement.',ECRAN.Caption);
    exit;
  end;
  // Inutile : fait systématiquement dans l'enregistrement
  // ExecuteSQL('UPDATE  STDCPTA SET STC_LIBELLE="'+ GetControlText('LIBELLESTD')+'" where STC_NUMPLAN='+ IntTostr(NUMSTD.Value)) ;

{$IFDEF EAGLCLIENT}
  lTobParam := TOB.Create('ParamCloture', nil, -1) ;
  // Paramètres pour l'appel du process server :
(*  lTobParam.AddChampSupValeur('USERLOGIN' , V_PGI.UserLogin ) ;
  lTobParam.AddChampSupValeur('INIFILE' , HalSocIni ) ;
  lTobParam.AddChampSupValeur('PASSWORD' , V_PGI.Password ) ;
  lTobParam.AddChampSupValeur('DOMAINNAME' , '' ) ;
  lTobParam.AddChampSupValeur('DATEENTREE' , V_PGI.DateEntree ) ;
  lTobParam.AddChampSupValeur('DOSSIER' , V_PGI.CurrentAlias ) ;
*)
  // Libellé + Abrege du Standard à sauvegarder...
  lTobParam.AddChampSupValeur('LIBELLESTD', GetControlText('LIBELLESTD'));
  lTobParam.AddChampSupValeur('ABREGESTD', GetControlText('ABREGESTD'));

  lTobResult := LanceProcessServer('cgiStdCpta', 'ENREGSTD', IntToStr(FNumStd.Value), lTobParam, True ) ;
  if lTobResult.GetValue('RESULT') = 'PASOKENREG' then
{$ELSE}
  if not EnregistreToutLeStandard( FNumStd.Value, GetControlText('LIBELLESTD'), GetControlText('ABREGESTD'), True ) then
{$ENDIF}
  begin
    MessageAlerte('Enregistrement du dossier type impossible.');
  end
  else
  begin
    PGIInfo('Enregistrement du dossier type terminé.', Ecran.Caption);
    FNumPlanCompte := FNumStd.Value;
  end;

{$IFDEF EAGLCLIENT}
  // Libération des Tob du ProcessServer
  FreeAndNil( lTobParam  );
  FreeAndNil( lTobResult );
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_STDENREGPLAN.OnChangeNumStd(Sender: TObject);
var T : TOB;
begin
  T := FListeStd.FindFirst(['STC_NUMPLAN'],[FNumStd.Value],False);
  if T <> nil then
  begin
    SetControlText('LIBELLESTD',T.GetValue('STC_LIBELLE'));
    SetControlText('ABREGESTD',T.GetValue('STC_ABREGE'));
  end
  else
  begin
    SetControlText('LIBELLESTD','');
    SetControlText('ABREGESTD','');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_STDENREGPLAN.OnClose;
begin
  FListeStd.Free;
  inherited;
end;

initialization
  registerclasses([TOF_STDENREGPLAN]);
end.

