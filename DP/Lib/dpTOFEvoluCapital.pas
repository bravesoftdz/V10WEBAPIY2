unit dpTOFEvoluCapital;
// TOF de la fiche DP EVOLUCAPITAL

interface
uses  StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
      HCtrls, HEnt1, HMsgBox, UTOF, UTob, HDB,
{$IFDEF EAGLCLIENT}
      MaineAGL, eMul,
{$ELSE}
      FE_Main, Mul,
{$ENDIF}
      HTB97,
      AGLInit;

Type
  TOF_EVOLUCAPITAL = Class (TOF)
    procedure OnArgument(Arguments : String ) ; override ;
  private
    EvovaleurNomin    : String; // utile pour la Tom TOM_DPMVTCAP
    procedure FListe_OnDblClick(Sender: TObject);
    procedure BINSERT_OnClick(Sender: TObject);
    procedure BSELECT_OnClick(Sender: TObject);
    procedure BDELETE_OnClick(Sender: TObject);
  end;


///////////// IMPLEMENTATION ////////////
implementation

uses dpOutils, DpJurOutilsEve
     {$IFDEF VER150}
     ,Variants
     {$ENDIF}
     ;

procedure TOF_EVOLUCAPITAL.OnArgument(Arguments : String ) ;
var action : String;
begin
  Inherited; // traite ACTION=
{$IFDEF EAGLCLIENT}
  THGrid(GetControl('FLISTE')).OnDblClick := FListe_OnDblClick;
{$ELSE}
  THDBGrid(GetControl('FLISTE')).OnDblClick := FListe_OnDblClick;
{$ENDIF}
  TToolbarButton97(GetControl('BINSERT')).OnClick := BINSERT_OnClick;
  TToolbarButton97(GetControl('BSELECT')).OnClick := BSELECT_OnClick;
  TToolbarButton97(GetControl('BDELETE')).OnClick := BDELETE_OnClick;

  if Arguments='' then exit;
  action := ReadTokenSt(Arguments);
  EvovaleurNomin := ReadTokenSt(Arguments);
end;


procedure TOF_EVOLUCAPITAL.FListe_OnDblClick(Sender: TObject);
var St: String;
begin
  if VarIsNull(GetField('DPM_GUIDPER')) then exit;
  // pour se balader sur les enreg, il faut indiquer le mul d'origine,
  // sinon malgré Range large, on reste bloqué sur l'enreg passé dans Lequel,
  // et on ne peut pas faire suivant/précédent dans la fiche ...
{$IFDEF EAGLCLIENT}
  TheMulQ:=TFMul(Ecran).Q.TQ;
{$ELSE}
  TheMulQ:=TFMul(Ecran).Q;
{$ENDIF}
  AGLLanceFiche('DP','MOUVCAPITAL', GetField('DPM_GUIDPER'), GetField('DPM_GUIDPER') + ';' + IntToStr(GetField('DPM_NOORDRE')),
    'ACTION=MODIFICATION;+;'+EvovaleurNomin) ;
  AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;


procedure TOF_EVOLUCAPITAL.BINSERT_OnClick(Sender: TObject);
begin
  AGLLanceFiche('DP','MOUVCAPITAL', GetControlText('DPM_GUIDPER'), GetControlText('DPM_GUIDPER'),
    'ACTION=CREATION;'+GetControlText('DPM_SENS')+';'+EvovaleurNomin) ;
  AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;


procedure TOF_EVOLUCAPITAL.BSELECT_OnClick(Sender: TObject);
begin
  FListe_OnDblClick(Sender);
end;


procedure TOF_EVOLUCAPITAL.BDELETE_OnClick(Sender: TObject);
begin
  SupprimeListeEnreg(TFMul(Ecran), 'DPMVTCAP');
  AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;


Initialization
  registerclasses([TOF_EVOLUCAPITAL]) ;
end.
