unit TofTLMvtRupt;

interface

uses Classes, StdCtrls,
{$IFDEF EAGLCLIENT}
        MaineAGL,
{$ELSE}
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
     FE_Main,
     QRS1,
{$ENDIF}
     comctrls,UTof, HCtrls,
     Ent1, Graphics, HTB97, spin,
     SysUtils;

function CPLanceFiche_MvtRuptureTL(vStRange, vStLequel, vStArgs : string) : String ;

type
  TOF_TLMVTRUPT = class(TOF)
  private
    ArgRetour: string;
    procedure E00OnClick(Sender : TObject);
    procedure TDeOnChange(Sender : TObject);
    procedure TaOnChange(Sender : TObject);

  public
    procedure OnArgument(Arguments : string); override ;
    procedure OnUpdate; override ;
    procedure OnClose; override ;
  end ;

implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  HEnt1, LicUtil, Vierge ;

// ===============================================================================

function CPLanceFiche_MvtRuptureTL(vStRange, vStLequel, vStArgs : string) : String ;
begin
  result := AglLanceFiche('CP','TLMVTRUPT',vStRange, vStLequel, vStArgs) ;
end ;

// ===============================================================================

procedure TOF_TLMVTRUPT.OnUpdate;
Var St   : string;
    StDe : string;
    StA  : string;
    i    : integer;
    CB   : TCheckBox;
    TC   : tComponent ;
begin

  NextPrevControl( Ecran ) ;

  for i := 0 to 3 do
    begin
    TC := Ecran.FindComponent( 'E0'+IntToStr(i) ) ;
    CB := TCheckBox(TC) ;
    if CB.enabled=True then
      if CB.checked then
        if GetControlText('DETABLEE0'+IntToStr(i))='' then
          begin
          StDe := StDe + '*;' ;
          StA  := StA  + '*;' ;
          end
        else
          begin
          StDe := StDe + GetControlText('DETABLEE0'+IntToStr(i)) + ';';
          StA  := StA  + GetControlText('ATABLEE0'+IntToStr(i))  + ';';
          end
      else
        begin
        StDe := StDe + '-;' ;
        StA  := StA  + '-;' ;
        end
    else
      begin
      StDe := StDe + '#;';
      StA  := StA  + '#;';
      end
    end;

  St := StDe + '|' + StA;

  ArgRetour := St ;
//  TFVierge(Ecran).retour := St ;

END;

procedure TOF_TLMVTRUPT.OnArgument(Arguments : string);
Var i : integer;
    CB : TCheckBox;
    Tde,Ta : THEdit;
    LibA : TLabel;
    TE,TE1,TC,LA : tComponent ;
    LesLib : HTStringList;
    St,St1,St2,St3,St4,St5 : string;
BEGIN

  inherited ;
  Ecran.HelpContext := 7490100 ;
  ArgRetour := Arguments;
  St2       := Arguments;
  St3       := ReadTokenPipe(St2,'|');

  for i:= 0 to 3 do
    begin
    SetControlEnabled( 'DETABLEE0'+IntToStr(i), FALSE);
    SetControlEnabled( 'ATABLEE0'+IntToStr(i),  FALSE);
    SetControlEnabled( 'TLibE0'+IntToStr(i),    FALSE) ;
    SetControlEnabled( 'TDeE0'+IntToStr(i),     FALSE) ;
    SetControlEnabled( 'TaE0'+IntToStr(i),      FALSE) ;
    end;

  LesLib:=HTStringList.Create ;
  GetLibelleTableLibre('E',LesLib);

  For i:=0 To 3 Do
    begin
    TC  := Ecran.FindComponent('E0'+IntToStr(i)) ;
    TE  := Ecran.FindComponent('DETABLEE0'+IntToStr(i)) ;
    TE1 := Ecran.FindComponent('ATABLEE0'+IntToStr(i)) ;
    LA  := Ecran.FindComponent('TaE0'+IntToStr(i)) ;
    if TC<>NIL Then
      begin
      CB    :=TCheckBox(TC) ;
      TDe   :=THEdit(TE) ;
      Ta    :=THEdit(TE1) ;
      LibA  :=TLabel(LA) ;
      St    :=LesLib[i];
      St1   :=ReadTokenSt(St) ;

      CB.enabled := St='X' ;
      SetControlCaption('TlibE0'+IntToStr(i),St1);
      CB.OnClick := E00OnClick;

      St4 := ReadTokenSt(St2);
      St5 := ReadTokenSt(St3);

      if (St5<>'-') and (St5<>'#') and (St5<>'') then
        if (St5='*') then
          begin
          CB.checked   := True;
          TDe.enabled := True;
          end
        else
          begin
          CB.checked   := True;
          TDe.enabled  := True;
          Ta.enabled   := True;
          LibA.enabled := True;
          TDe.text     := St5;
          Ta.text      := St4;
          end;

//      TDe.OnChange:=TDeOnChange;
//      Ta.OnChange:=TaOnChange;
        Tde.OnExit := TDeOnChange;
        Ta.OnExit  := TaOnChange;
      end ;
    end ;

  LesLib.Free ;

end;

Procedure TOF_TLMVTRUPT.TDeOnChange(Sender : TObject);
var Tde : THEdit ;
    i   : Integer ;
begin

  Tde := THEdit(Sender) ;
  If Tde=NIL Then Exit ;

  i:=Tde.tag ;
  if GetControlText('ATABLEE0'+IntToStr(i)) = '' then
    SetControlText( 'ATABLEE0'+IntToStr(i), GetControlText('DETABLEE0'+IntToStr(i) ) ) ;

{  if GetControlText('DETABLEE0'+IntToStr(i)) = '' then
    begin
    SetControlEnabled( 'TaE0'+IntToStr(i),     False) ;
    SetControlEnabled( 'ATABLEE0'+IntToStr(i), False);
    end
  else
    begin
    SetControlEnabled( 'TaE0'+IntToStr(i),     GetControlText( 'E0' + IntToStr(i) )='X' ) ;
    SetControlEnabled( 'ATABLEE0'+IntToStr(i), GetControlText( 'E0' + IntToStr(i) )='X' );
    end;
}
end;

Procedure TOF_TLMVTRUPT.TaOnChange(Sender : TObject);
var Ta : THEdit ;
    i  : Integer ;
begin

  Ta := THEdit(Sender) ;
  if Ta=NIL Then Exit ;

  i:=Ta.tag ;
  if GetControlText( 'DETABLEE0' + IntToStr(i) ) > Ta.text then
    SetControlText(  'DETABLEE0' + IntToStr(i), GetControlText('ATABLEE0' + IntToStr(i) ) );

end;

Procedure TOF_TLMVTRUPT.E00OnClick(Sender : TObject);
var CB : TCheckBox ;
    i : Integer ;
begin

  CB := TCheckBox(Sender) ;
  if CB=NIL Then Exit ;
  i:=CB.tag ;
  if CB.checked=False then
    begin
    SetControlText('DETABLEE0'+IntToStr(i),'');
    SetControlText('ATABLEE0'+IntToStr(i),'');
    end;

  SetControlEnabled( 'DETABLEE0' + IntToStr(i), GetControlText('E0'+IntToStr(i)) = 'X' );
  SetControlEnabled( 'TLibE0'    + IntToStr(i), GetControlText('E0'+IntToStr(i)) = 'X' ) ;
  SetControlEnabled( 'TDeE0'     + IntToStr(i), GetControlText('E0'+IntToStr(i)) = 'X' ) ;
  SetControlEnabled( 'TaT0'      + IntToStr(i), GetControlText('E0'+IntToStr(i)) = 'X' ) ;
  SetControlEnabled( 'ATABLEE0'  + IntToStr(i), GetControlText('E0'+IntToStr(i)) = 'X' );

End;

procedure TOF_TLMVTRUPT.OnClose;
begin
  inherited;
  TFVierge(Ecran).retour:=ArgRetour;
end;

initialization
RegisterClasses([TOF_TLMVTRUPT]) ;

end.
