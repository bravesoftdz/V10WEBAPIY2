
unit UTofPG_GestionMotPasse;

interface
uses  StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}DBGrids,
{$ENDIF}
      Grids,HCtrls,HEnt1,HMsgBox,HSysMenu,UTOF,UTOB,Vierge,AGLInit;

Type
     TOF_PG_GestionMotPasse = Class (TOF)
       private
       LeDroit :   Boolean ;
       TypAct  :   String ;
       procedure ValidClick(Sender: TObject);
       procedure SortieClick(Sender: TObject);
       procedure ControleSaisie () ;
       public
       procedure OnClose ; override ;
       procedure OnArgument(Arguments : String ) ; override ;
     END ;

implementation

uses ParamSoc,LicUtil,EntPaie;

procedure TOF_PG_GestionMotPasse.OnArgument(Arguments: String);
var BTNVAL, BTNSOR : TToolbarButton97 ;
begin
inherited ;
  TypAct := Trim (Arguments) ;
  If TypAct = 'C' then   // Phase de controle
  begin
    SetControlVisible ('MOTDEPASSE', TRUE) ;
    SetControlVisible ('TMOTDEPASSE', TRUE) ;
    SetControlVisible ('NEWMOT', TRUE) ;
    SetControlVisible ('TNEWMOT', TRUE) ;
    SetControlVisible ('CNOUVMOT', TRUE) ;
    SetControlVisible ('TCNOUVMOT', TRUE) ;
  end
  else
  begin // Phase de création ou de modification
    SetControlVisible ('MOTDEPASSE', FALSE ) ;
    SetControlVisible ('TMOTDEPASSE', FALSE ) ;
    SetControlVisible ('NEWMOT', TRUE ) ;
    SetControlVisible ('TNEWMOT', TRUE ) ;
    SetControlVisible ('CNOUVMOT', TRUE ) ;
    SetControlVisible ('TCNOUVMOT',TRUE ) ;
  end;

  BTNVAL := TToolbarButton97 (GetControl ('BValider')) ;
  If BTNVAL <> NIL then BTNVAL.OnClick := ValidClick ;
  BTNSOR := TToolbarButton97 (GetControl ('BFerme')) ;
  If BTNSOR <> NIL then BTNSOR.OnClick := SortieClick ;

end;

procedure TOF_PG_GestionMotPasse.OnClose;
begin
  SortieClick (NIL);
end;
{
}
procedure TOF_PG_GestionMotPasse.SortieClick(Sender: TObject);
begin
  ControleSaisie ();
  if LeDroit AND (TypAct <> 'C') then // Saisie
  begin
    SetParamSoc ('SO_PGTPCL', CryptageSt (GetControlText ('NEWMOT'))) ;
    VH_Paie.PGMotDePasse := CryptageSt (GetControlText ('NEWMOT')) ;
  end;
  if LeDroit then TFVierge(Ecran).Retour := 'N'
    else TFVierge(Ecran).Retour := 'O' ;
end;

procedure TOF_PG_GestionMotPasse.ValidClick(Sender: TObject);
begin
  SortieClick (Sender);
end;

Procedure TOF_PG_GestionMotPasse.ControleSaisie ( ) ;
begin
  if TypAct = 'C' then // Controle
  begin
  if GetControlText ('NEWMOT') = DeCryptageSt (VH_Paie.PGMotDePasse) then
    LeDroit := TRUE
    else LeDroit := FALSE ;
  end
  else
  begin
  if GetControlText ('NEWMOT') = GetControlText ('CNOUVMOT') then
    LeDroit := TRUE
    else LeDroit := FALSE ;
  end;
end;

Initialization
registerclasses([TOF_PG_GestionMotPasse]);
end.
