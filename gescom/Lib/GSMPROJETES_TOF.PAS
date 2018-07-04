{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 14/05/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : GSMPROJETES ()
Mots clefs ... : TOF;GSMPROJETES
*****************************************************************}
Unit GSMPROJETES_TOF ;

Interface

{$IFDEF STK}

Uses
  StdCtrls,
  Controls,
  Classes,
{$IFNDEF EAGLCLIENT}
  db,
  dbtables,
  Fe_Main,
  Mul,
{$ELSE}
  MainEAgl,
  eMul,
{$ENDIF}
  HTB97,
  Menus,
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  wTof,
  Dispo,
  wMnu
  ;

Type
  TOF_GSMPROJETES = Class (tWTOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    PmFlux: TPopupMenuFlux;

    { Loupe }
    procedure MnLpArticle_OnClick(Sender: TObject);
    procedure MnLpDispo_OnClick(Sender: TObject);
    procedure MnLpRefOrigine_OnClick(Sender: TObject);

    { Get }
    function GetCleGQ: tCleGQ;
    function GetTobGSM: Boolean;

    { Set }
    procedure SetPhysique;

    { Flux }
    procedure MNFlux_OnClick(Sender: TObject);
  end ;

{$ENDIF}

Implementation

{$IFDEF STK}

uses
  wCommuns,
  UtilArticle,
  EntGP,
  uTob,
  StkMouvement
  ;

procedure TOF_GSMPROJETES.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_GSMPROJETES.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_GSMPROJETES.OnUpdate ;
begin
  Inherited ;
  { Tob projetée }
  GetTobGSM;

  SetPhysique;
end ;

procedure TOF_GSMPROJETES.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_GSMPROJETES.OnArgument (S : String ) ;
begin
  FTableName := 'STKMOUVEMENT';

  Inherited ;

  { Unité de flux }
  PmFlux := TPopupMenuFlux.Create(Ecran, MNFlux_OnClick, S, GetArgumentValue(S, 'FLUX'));

  { Gestion des colonnes }
  InitParamsForColumns('QPREVUE;PHYSIQUE;QRESTE;LIBUNITE', PmFlux);

  { Loupe }
  if Assigned(GetControl('MNLPARTICLE')) then
    TMenuItem(GetControl('MNLPARTICLE')).OnClick := MnLpArticle_OnClick;
  if Assigned(GetControl('MNLPDISPO')) then
    TMenuItem(GetControl('MNLPDISPO')).OnClick := MnLPDispo_OnClick;
  if Assigned(GetControl('MNLPREFORIGINE')) then
    TMenuItem(GetControl('MNLPREFORIGINE')).OnClick := MnLPRefOrigine_OnClick;

  { TobTemp }
  VH_Gp.TobTemp := TOB.Create('GSM', nil, -1);

  { bAgrandir }
  if Assigned(GetControl('BAGRANDIR')) then
    ttoolbarButton97(GetControl('BAGRANDIR')).Click;
end;

procedure TOF_GSMPROJETES.OnClose ;
begin
  Inherited ;

  PmFlux.Free;
  VH_Gp.TobTemp.free;
end ;

procedure TOF_GSMPROJETES.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_GSMPROJETES.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_GSMPROJETES.MnLpArticle_OnClick(Sender: TObject);
  function GetRange: string;
  begin
    Result := 'GA_ARTICLE=' + GetString('GSM_ARTICLE');
  end;
begin
	wCallGA(GetRange);
  RefreshDB;
end;

procedure TOF_GSMPROJETES.MnLpDispo_OnClick(Sender: TObject);
begin
	CallMulGQ(GetCleGQ, PmFlux.Flux);
  RefreshDB;
end;

function TOF_GSMPROJETES.GetCleGQ: tCleGQ;
begin
  Result.Article := GetString('GSM_ARTICLE');
  Result.Depot := GetString('GSM_DEPOT');
  Result.Cloture := False;
  Result.DateCloture := iDate1900;
end;

procedure TOF_GSMPROJETES.MNFlux_OnClick(Sender: TObject);
begin
  PmFlux.Flux := StringReplace(TPopupMenu(Sender).Name, 'MN', '', [rfIgnoreCase]);
  SetColsVisible;
  SetPhysique;
  VH_GP.TobTemp.AddChampSupValeur('Flux', PmFlux.Flux);
end;


function TOF_GSMPROJETES.GetTobGSM: Boolean;
var
  i       : integer;
  sql     : string;
  Physique: Double;

  function GetWhereGQ: string;
  begin
    Result := 'GQ_ARTICLE="' + wGetArticleFromCodeArticle(GetControlText('GSM_CODEARTICLE')) + '"'
            + ' AND GQ_DEPOT="' + GetControlText('GSM_DEPOT') + '"'
            + ' AND GQ_CLOTURE="' + wFalse + '"'
            + ' AND GQ_DATECLOTURE="' + UsDateTime(iDate1900) +'"'
  end;
begin
  Result := false;

  sql := 'SELECT *'
       + ' FROM GSMPROJETES'
       + ' WHERE GSM_ARTICLE="' + wGetArticleFromCodeArticle(GetControlText('GSM_CODEARTICLE')) + '"'
       + ' AND GSM_DEPOT="' + GetControlText('GSM_DEPOT') + '"'
       + ' ORDER BY GSM_DATEPREVUE, GSM_QUALIFMVT, GSM_IDENTIFIANT'
       ;
  if wLoadTobFromSql('GSM', Sql, VH_GP.TobTemp) then
  begin
    Result := true;
    VH_GP.TobTemp.AddChampSupValeur('Flux', PmFlux.Flux);
    for i := 0 to VH_GP.TobTemp.Detail.Count -1 do
    begin
      if i = 0 then
      begin
        Physique := wGetSqlFieldValue('GQ_PHYSIQUE', 'DISPO', GetWhereGQ);
        VH_GP.TobTemp.Detail[i].AddChampSupValeur('QPROJETESTO', Physique                                                                                             + VH_GP.TobTemp.Detail[i].G('QRESTESTO'));
        VH_GP.TobTemp.Detail[i].AddChampSupValeur('QPROJETEVTE', wDivise(Physique, VH_GP.TobTemp.Detail[i].G('QUOTITEVTE') * VH_GP.TobTemp.Detail[i].G('QUOTITESTO')) + VH_GP.TobTemp.Detail[i].G('QRESTEVTE'));
        VH_GP.TobTemp.Detail[i].AddChampSupValeur('QPROJETEACH', wDivise(Physique, VH_GP.TobTemp.Detail[i].G('QUOTITEACH') * VH_GP.TobTemp.Detail[i].G('QUOTITESTO')) + VH_GP.TobTemp.Detail[i].G('QRESTEACH'));
        VH_GP.TobTemp.Detail[i].AddChampSupValeur('QPROJETEPRO', wDivise(Physique, VH_GP.TobTemp.Detail[i].G('QUOTITEPRO') * VH_GP.TobTemp.Detail[i].G('QUOTITESTO')) + VH_GP.TobTemp.Detail[i].G('QRESTEPRO'));
        VH_GP.TobTemp.Detail[i].AddChampSupValeur('QPROJETECON', wDivise(Physique, VH_GP.TobTemp.Detail[i].G('QUOTITECON') * VH_GP.TobTemp.Detail[i].G('QUOTITESTO')) + VH_GP.TobTemp.Detail[i].G('QRESTECON'));
      end
      else
      begin
        VH_GP.TobTemp.Detail[i].AddChampSupValeur('QPROJETESTO', VH_GP.TobTemp.Detail[i-1].G('QPROJETESTO') + VH_GP.TobTemp.Detail[i].G('QRESTESTO'));
        VH_GP.TobTemp.Detail[i].AddChampSupValeur('QPROJETEVTE', VH_GP.TobTemp.Detail[i-1].G('QPROJETEVTE') + VH_GP.TobTemp.Detail[i].G('QRESTEVTE'));
        VH_GP.TobTemp.Detail[i].AddChampSupValeur('QPROJETEACH', VH_GP.TobTemp.Detail[i-1].G('QPROJETEACH') + VH_GP.TobTemp.Detail[i].G('QRESTEACH'));
        VH_GP.TobTemp.Detail[i].AddChampSupValeur('QPROJETEPRO', VH_GP.TobTemp.Detail[i-1].G('QPROJETEPRO') + VH_GP.TobTemp.Detail[i].G('QRESTEPRO'));
        VH_GP.TobTemp.Detail[i].AddChampSupValeur('QPROJETECON', VH_GP.TobTemp.Detail[i-1].G('QPROJETECON') + VH_GP.TobTemp.Detail[i].G('QRESTECON'));
      end
    end;
  end;
end;

procedure TOF_GSMPROJETES.SetPhysique;
var
  Physique: Double;
begin
  Physique := 0;

  if Assigned(VH_GP.TobTemp) then
  begin
    Physique := VH_GP.TobTemp.Detail[0].G('QPROJETE' + PmFlux.Flux) - VH_GP.TobTemp.Detail[0].G('QRESTE' + PmFlux.Flux)
  end;

  SetControlProperty('PHYSIQUE', 'Value', FloatToStr(Physique));
end;

procedure TOF_GSMPROJETES.MnLpRefOrigine_OnClick(Sender: TObject);
begin
  CallFicRefOrigine(GetString('GSM_REFORIGINE'));
  RefreshDB;
end;

Initialization
  registerclasses ( [ TOF_GSMPROJETES ] ) ;

{$ENDIF}  
end.
