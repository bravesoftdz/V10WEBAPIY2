unit UtilPlannifchantier;

interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,
{$ELSE}
     MainEagl,
{$ENDIF}
     forms,
     menus,
     AglInit,
     saisUtil,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     UTOB,
     UTOF,
     UPlannifchUtil,
     FactComm,
     FactUtil,
     FactCalc,
     FactGrp,
     FactGrpBTP,
     paramSoc,
     EntGC,
     FactOuvrage,
     FactVariante,
     ExtCtrls,
     Graphics,
     FactTob,
     FactAdresse
     ;

type    
  TtypeTraitPlannif = (TtcDemPrix,ttcChantier);
  TOF_BTTYPEPLANNIF = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
    	TOBR : TOB;
  end ;

  TOF_BTMODIFPHASE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
    DateInit : TDateTime;
    DateDebut : THedit;
    procedure DatedebutExit(Sender: Tobject);
  end ;

function ChoixPlannifdevis (Typetrait : TtypeTraitPlannif) : integer;
function ChoixPlannifCtrEtude : integer;
function AppelModifPhase (TOBpassage : TOB) : boolean;


implementation

function AppelModifPhase (TOBpassage : TOB) : boolean;
begin
	result := true;
  TheTob := TOBPassage;
  AGLLanceFiche('BTP','BTMODIFPHASE','','','ACTION=MODIFICATION') ;
  if TheTOB = nil then result := false;
  TheTob := nil;
end;


function ChoixPlannifdevis (Typetrait : TtypeTraitPlannif) : integer;
var TOBparam : TOB;
begin
	result := -1;
	TOBParam := TOB.Create ('RESULTAT',nil,-1);
  TOBPARAM.addChampSupValeur ('CONDITION',0,false);

  if TypeTrait = TtcDemPrix then
    TOBPARAM.addChampSupValeur ('TYPETRAIT','D',false)
  else
    TOBPARAM.addChampSupValeur ('TYPETRAIT','C',false);

  TheTOB := TOBPARAM;
  AGLLanceFiche('BTP','BTTYPEPLANNIF','','','DEVIS') ;
  if (TheTOB <> nil) and (TheTOB.GetValue('CONDITION')<>0 ) then
  begin
  	result := TheTOB.GetValue('CONDITION');
  end;
  TheTOB := nil;
  TOBPARAM.free;
end;

function ChoixPlannifCtrEtude : integer;
var TOBparam : TOB;
begin
	result := -1;
	TOBParam := TOB.Create ('RESULTAT',nil,-1);
  TOBPARAM.addChampSupValeur ('CONDITION',0,false);
  TOBPARAM.addChampSupValeur ('TYPETRAIT','C',false);
  TheTOB := TOBPARAM;
  AGLLanceFiche('BTP','BTTYPEPLANNIF','','','CTRETUDE') ;
  if (TheTOB <> nil) and (TheTOB.GetValue('CONDITION')<>0 ) then
  begin
  	result := TheTOB.GetValue('CONDITION');
  end;
  TheTOB := nil;
  TOBPARAM.free;
end;

(* BTTYPEPLANNIF *)

procedure TOF_BTTYPEPLANNIF.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTTYPEPLANNIF.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTTYPEPLANNIF.OnUpdate ;
begin
  Inherited ;
  if TCheckBox(GetControl('CGLOBAL')).checked then TOBR.putvalue('CONDITION',1);
  if TCheckBox(GetControl('CPHASE')).checked then TOBR.putvalue('CONDITION',TOBR.Getvalue('CONDITION')+2);
  if TCheckBox(GetControl('COUVRAGE')).checked then TOBR.putvalue('CONDITION',TOBR.Getvalue('CONDITION')+4);
  TheTob := TOBR;
end ;

procedure TOF_BTTYPEPLANNIF.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTTYPEPLANNIF.OnArgument (S : String ) ;
begin
  Inherited ;
  if S = 'CTRETUDE' then
  begin
  	TCheckBox(GetControl('CGLOBAL')).visible := false;
  	TCheckBox(GetControl('CGLOBAL')).checked := false;
  	TCheckBox(GetControl('CPHASE')).checked := true;
  end;
  TOBR := LaTOB;
  if TOBR.GetString('TYPETRAIT')='D' then TCheckBox(GetControl('CGLOBAL')).Caption := 'Devis globalisé';
end ;

procedure TOF_BTTYPEPLANNIF.OnClose ;
begin
  Inherited ;
end ;

(* BTMODIFPHASE *)

procedure TOF_BTMODIFPHASE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTMODIFPHASE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTMODIFPHASE.OnUpdate ;
begin
  Inherited ;
  if LaTob.GetValue('DATEDEBUT') < DateInit then
  begin
    PGIBox (TraduireMemoire('La date de début est antérieure à celle du démarrage chantier'),ecran.caption);
    TForm(Ecran).ModalResult:=0;
    exit;
  end;
  if LaTob.getValue('DATEDEBUT_') < LaTOB.GetValue('DATEDEBUT') then
  begin
    PGIBox (TraduireMemoire('La date de fin des travaux est antérieure à la date de début'),ecran.caption);
    TForm(Ecran).ModalResult:=0;
    exit;
  end;
  TheTob := laTob;
end ;

procedure TOF_BTMODIFPHASE.OnLoad ;
begin
  Inherited ;
  laTOB.PutEcran(ecran);
  DateInit := LaTOB.GetValue('DATEINIT');
end ;

procedure TOF_BTMODIFPHASE.OnArgument (S : String ) ;
begin
  Inherited ;
  DateDebut := THedit(Getcontrol('DATEDEBUT'));
  DateDebut.OnExit := DatedebutExit;
end ;

procedure TOF_BTMODIFPHASE.DatedebutExit (Sender : Tobject);
begin
  if StrToDate(THedit(Getcontrol('DATEDEBUT_')).Text) < StrToDate(THedit(Getcontrol('DATEDEBUT')).Text) then
  	THedit(Getcontrol('DATEDEBUT_')).Text := DateDebut.Text;
end;

procedure TOF_BTMODIFPHASE.OnClose ;
begin
  Inherited ;
end ;



Initialization
  registerclasses ( [ TOF_BTTYPEPLANNIF, TOF_BTMODIFPHASE] ) ;

end.
