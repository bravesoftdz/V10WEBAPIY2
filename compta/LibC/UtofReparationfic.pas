{***********UNITE*************************************************
Auteur  ...... : M. ENTRESSANGLE
Créé le ...... : 13/09/2007
Modifié le ... :   /  /    
Description .. : Ecran de lancement de calcul des contreparties 
Suite ........ : automatiques
Mots clefs ... : 
*****************************************************************}

unit UtofReparationfic;

interface

uses
  SysUtils, UTob, UTOF, HTB97, HEnt1, HCtrls, HMsgBox, Controls,
  Classes, Ent1,
{$IFDEF EAGLCLIENT}
  MaineAGL,
  uWA,
{$ELSE}
  FE_Main,
  CalCulContr ,
{$ENDIF}
  RecupUtil;

type
  TOF_REPARATIONFIC = Class (TOF)
    procedure OnLoad ; override ;
    procedure Onclick(Sender: TObject);
    procedure OnChangeCtr(Sender: TObject);
    procedure EXERCICEChange(Sender: TObject);
    function  InitTobParam : TOB ;
    procedure CtrEnabled ( OKCRT : Boolean);
  end;

procedure CCLanceFiche_ReparationFic(pszArg : String);

implementation

procedure CCLanceFiche_ReparationFic(pszArg : String);
begin
  AGLLanceFiche('CP','REPARATIONFIC','','',pszArg);
end;


procedure TOF_REPARATIONFIC.OnLoad;
var
Jrl : string;
Q   : TQuery;
begin
  TToolBarButton97(GetControl ('BValider')).Onclick := Onclick;
  THValComboBox(GetControl('CEXO',True) ).OnChange  := EXERCICEChange;
  SetControlEnabled ('PANEL1', FALSE);
  THCheckBox(GetControl ('BCTR1')).Onclick := OnChangeCtr;
  CtrEnabled(FALSE);
  Q := OpenSql ('SELECT J_JOURNAL FROM JOURNAL WHERE J_NATUREJAL<>"ANO" AND J_NATUREJAL<>"CLO"', TRUE);
  if not Q.EOF then
  Jrl := Q.FindField ('J_JOURNAL').asstring;
  ferme (Q);
  SetControlText('E_JOURNAL', Jrl);
  Q := OpenSql ('SELECT MAX(J_JOURNAL) jrl FROM JOURNAL WHERE J_NATUREJAL<>"ANO" AND J_NATUREJAL<>"CLO"', TRUE);
  if not Q.EOF then
  Jrl := Q.FindField ('Jrl').asstring ;
  ferme (Q);
  SetControlText('E_JOURNAL1', Jrl);

end;

procedure TOF_REPARATIONFIC.OnChangeCtr(Sender: TObject);
begin
  if THCheckBox(GetControl ('BCTR1')).Checked then
  begin
    SetControlEnabled ('PANEL1', TRUE);
    CtrEnabled (TRUE);
  end
  else
  begin
    SetControlEnabled ('PANEL1', FALSE);
    CtrEnabled (FALSE);
  end;
end;

procedure TOF_REPARATIONFIC.Onclick(Sender: TObject);
var 
TB              : TOB;
{$IFDEF EAGLCLIENT}
lTobResult      : TOB;
{$ENDIF}
begin
if not VH^.EnSerie then
  if PGiAsk('Confirmez-vous le traitement ? ') <> mrYes then Exit ;

TB := InitTobParam;
{$IFDEF EAGLCLIENT}
      with cWA.create do
      begin
          lTobResult  := Request('CtrCpta.CALCTREPARTIE','', TB,'','');
         free ;
      end ;
      if lTobResult = nil then
      begin
        PgiBox('Erreur, Server CtrCpta introuvable',' Calcul des contreparties') ;
      end
      else
      if lTobResult.GetValue ('ERROR') <> '' then
           PgiBox(lTobResult.GetValue ('ERROR'), 'Calcul des contreparties')
      else PgiInfo ('Réparation terminée');
{$ELSE}
      LanceTraitementContreparties (TB);
{$ENDIF}

TB.Free;

end;

function TOF_REPARATIONFIC.InitTobParam : TOB ;
var
LT, L1   : TOB;
Date1,Date2 : string;
begin
    LT := TOB.Create('$PARAM', nil, -1) ;
    LT.AddChampSupValeur('USERLOGIN' , V_PGI.UserLogin ) ;
    LT.AddChampSupValeur('INIFILE'   , HalSocIni ) ;
    LT.AddChampSupValeur('PASSWORD'  , V_PGI.Password ) ;
    LT.AddChampSupValeur('DOMAINNAME', '' ) ;
    LT.AddChampSupValeur('DATEENTREE', V_PGI.DateEntree ) ;
    LT.AddChampSupValeur('DOSSIER'   , V_PGI.NoDossier ) ;
    LT.AddChampSupValeur('APPLICATION', 'CgiCtrCpta') ;

    LT.AddChampSupValeur('BaseCommune', EstBaseCommune);

    L1 := TOB.Create('Trans', LT, -1) ;
    L1.AddChampSupValeur('E_EXERCICE'           , THValComboBox(GetControl('CEXO',True) ).value) ;
    Date1 := GetControlText('E_DATECOMPTABLE');
    Date2 := GetControlText('E_DATECOMPTABLE1');
    L1.AddChampSupValeur('E_DATECOMPTABLE'      , UsDateTime(StrToDate(Date1))) ;
    L1.AddChampSupValeur('E_DATECOMPTABLE_'     , UsDateTime(StrToDate(Date2))) ;
    L1.AddChampSupValeur('E_JOURNAL_'           , GetControlText('E_JOURNAL1')) ;
    L1.AddChampSupValeur('E_JOURNAL'            , GetControlText('E_JOURNAL')) ;
    L1.AddChampSupValeur('E_NATUREPIECE'        , GetControlText('E_NATUREPIECE')) ;

    Result := LT;
end;

procedure TOF_REPARATIONFIC.EXERCICEChange(Sender: TObject);
begin
  inherited;
     ExoToDates(THValComboBox(GetControl('CEXO',True) ).value, THEdit(GetControl ('E_DATECOMPTABLE')), THEdit(GetControl ('E_DATECOMPTABLE1')));
end;

procedure TOF_REPARATIONFIC.CtrEnabled ( OKCRT : Boolean);
begin
    SetControlEnabled ('E_DATECOMPTABLE' , OKCRT);
    SetControlEnabled ('E_DATECOMPTABLE1', OKCRT);
    SetControlEnabled ('CEXO', OKCRT);
    SetControlEnabled ('E_JOURNAL1'      , OKCRT) ;
    SetControlEnabled ('E_JOURNAL'       , OKCRT) ;
    SetControlEnabled ('E_NATUREPIECE'   , OKCRT) ;
    SetControlEnabled ('BValider'        , OKCRT) ;
end;

Initialization
RegisterClasses([TOF_REPARATIONFIC]);

end.
