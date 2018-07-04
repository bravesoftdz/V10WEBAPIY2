Unit UTOFGCcptadiff ;

Interface

Uses  StdCtrls, Controls, Classes,  forms, sysutils, ComCtrls,grids ,
      graphics,windows,vierge,Messages,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ENDIF}
      HCtrls, HEnt1, HMsgBox, UTOF,FactComm ,Utob, AGLInit,EntGC,M3FP,CptaDiff ;

Type
  TOF_GCCPTADIFF = Class (TOF)
  private
    TobParam : Tob ;
    procedure InitLaTob ;
    procedure PurgeEcritures;
  protected
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnUpdate                 ; override ;
  end ;


Implementation

procedure TOF_GCCPTADIFF.OnArgument (S : String ) ;
begin
inherited;
InitLaTob;
end;

procedure TOF_GCCPTADIFF.OnLoad ;
var QQ : Tquery;
begin
inherited;
QQ:=OpenSql('SELECT MIN(GCD_DATEPIECE) from COMPTADIFFEREE WHERE GCD_COMPTABILISE="-"',true);
if not QQ.EOF then TobParam.PutValue('DATECPTA',QQ.Fields[0].asDateTime);
Ferme(QQ) ;

TobParam.PutEcran(Ecran);
end;

procedure TOF_GCCPTADIFF.OnUpdate ;
Var Ret : integer;
    StNats : String ;
begin
inherited;
TobParam.GetEcran(Ecran);
StNats:=TOBParam.GetValue('GCD_NATURECOMPTA') ; if Copy(StNats,1,2)='<<' then TOBParam.PutValue('GCD_NATURECOMPTA','') ;
if PGIAsk('Confirmez-vous le traitement ?',TFVierge(ECran).caption)=mrNo then Exit; ;
Ret:=PassageComptaDifferee ( TOBParam  ) ;
Case Ret of
  0 : begin
      // Purge des écritures
      if TobParam.GetValue('PURGE')='X' then PurgeEcritures ;
      PGIInfo('Traitement terminé',TFVierge(ECran).caption);
      end;
  1 : PGIInfo('Aucune pièce à traiter',TFVierge(ECran).caption);
  2 : PGIBox('Le traitement a échoué',TFVierge(ECran).caption);
  end;

end;

procedure TOF_GCCPTADIFF.PurgeEcritures;
var stWhere : string;
    nbJour : integer ;
begin
NbJour :=TobParam.GetValue('JOURSPURGE');
stWhere:='';
If nBJour<>0 then  stWhere:=' AND GCD_DATEPIECE<="'+USDateTime(Date-nbJour)+'"' ;
ExecuteSQL('DELETE FROM COMPTADIFFEREE WHERE GCD_COMPTABILISE="X" '+stWhere);
end;

procedure TOF_GCCPTADIFF.OnClose;
begin
inherited;
TobParam.free;
end;

procedure TOF_GCCPTADIFF.InitLaTob;
begin
TobParam:=TOB.Create('_Param Compta Diff',Nil,-1) ;
TobParam.addchampsup('DATECPTA',false)           ; TobParam.PutValue('DATECPTA',iDate1900);
TobParam.addchampsup('DATECPTA_',false)           ; TobParam.PutValue('DATECPTA_',iDate2099);
TobParam.addchampsup('CUMULDATE',false)         ; TobParam.PutValue('CUMULDATE','JOU');
TobParam.addchampsup('GCD_NATURECOMPTA',false)  ; TobParam.PutValue('GCD_NATURECOMPTA','');
TobParam.addchampsup('DEJAPASSE',false)         ; TobParam.PutValue('DEJAPASSE','-');
TobParam.addchampsup('DETAILENTREPRISE',false)  ; TobParam.PutValue('DETAILENTREPRISE','-');
TobParam.addchampsup('DETAILPARTICULIER',false) ; TobParam.PutValue('DETAILPARTICULIER','-');
TobParam.addchampsup('COMPTECENTRAL',false)     ; TobParam.PutValue('COMPTECENTRAL','X');
TobParam.addchampsup('DISTINGUERMPREGLE',false) ; TobParam.PutValue('DISTINGUERMPREGLE','X');
TobParam.addchampsup('PURGE',false)             ; TobParam.PutValue('PURGE','X');
TobParam.addchampsup('JOURSPURGE',false)        ; TobParam.PutValue('JOURSPURGE',0);
TobParam.addchampsup('CENTRALISERREGLEBQE',false) ; TobParam.PutValue('CENTRALISERREGLEBQE','-');
end;

Initialization
  registerclasses ( [TOF_GCCPTADIFF] ) ;
end.
