{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 23/04/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTCPTPIECE_S1 ()
Mots clefs ... : TOF;BTCPTPIECE_S1
*****************************************************************}
Unit BTCPTPIECE_S1_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOB,
     UTOF ;

Type
  TOF_BTCPTPIECE_S1 = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    TOBSouches : TOB;
    procedure ChargeLesSOuches;
    procedure EcritInfos;

    function ExistDocument(Nature : string;Numero : integer) : boolean;
    procedure PositionneEcran;
    procedure PositionneCompteurs;
    function verifCompteur : boolean;

  end ;

Implementation

procedure TOF_BTCPTPIECE_S1.ChargeLesSOuches;
var QQ : Tquery;
begin
  QQ := OpenSql ('SELECT * FROM SOUCHE WHERE SH_TYPE="GES" AND SH_SOUCHE IN ("ABT","FBT","DBT")',True);
  TOBSOuches.LoadDetailDB ('SOUCHE','','',QQ,false);
  ferme (QQ);
end;

function TOF_BTCPTPIECE_S1.ExistDocument(Nature: string; Numero: integer): boolean;
var QQ : TQuery;
begin
  QQ := OpenSql ('SELECT GP_NUMERO FROM PIECE WHERE GP_NATUREPIECEG="'+Nature+'" AND GP_NUMERO='+IntToStr(Numero),True);
  result := not QQ.eof;
  ferme (QQ);
end;

procedure TOF_BTCPTPIECE_S1.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTCPTPIECE_S1.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTCPTPIECE_S1.OnUpdate ;
begin
  Inherited ;
  if not verifCompteur then
  begin
    TForm(Ecran).ModalResult:=0;
    Exit;
  end;
  PositionneCompteurs;
  EcritInfos;
end ;

procedure TOF_BTCPTPIECE_S1.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTCPTPIECE_S1.OnArgument (S : String ) ;
begin
  Inherited ;
  TOBSouches := TOB.Create ('LES SOUCHES',nil,-1);
  ChargeLesSouches;
  PositionneEcran;
end ;

procedure TOF_BTCPTPIECE_S1.OnClose ;
begin
  Inherited ;
  TOBSouches.free;
end ;

procedure TOF_BTCPTPIECE_S1.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTCPTPIECE_S1.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTCPTPIECE_S1.PositionneEcran;
var Indice : integer;
    UneTOB : TOB;
begin
  for Indice := 0 To TOBSouches.detail.count -1 do
  begin
    UneTOB := TOBSouches.detail[Indice];
    if UneTOB.GetValue('SH_SOUCHE')='ABT' then
    begin
      THSpinEdit(GetControl ('CPTAVOIR')).Value := UneTOB.GetValue('SH_NUMDEPART');
    end else if (Pos(UneTOB.GetValue('SH_SOUCHE'),'FBT;FBP')>0) then
    begin
      THSpinEdit(GetControl ('CPTFACTURE')).Value := UneTOB.GetValue('SH_NUMDEPART');
    end else
    begin
      THSpinEdit(GetControl ('CPTDEVIS')).Value := UneTOB.GetValue('SH_NUMDEPART');
    end;
  end;
end;

procedure TOF_BTCPTPIECE_S1.PositionneCompteurs;
var Indice : integer;
    UneTOB : TOB;
begin
  for Indice := 0 To TOBSouches.detail.count -1 do
  begin
    UneTOB := TOBSouches.detail[Indice];
    if UneTOB.GetValue('SH_SOUCHE')='ABT' then
    begin
      UneTOB.PutValue('SH_NUMDEPART',THSpinEdit(GetControl ('CPTAVOIR')).Value);
    end else if (Pos(UneTOB.GetValue('SH_SOUCHE'),'FBT;FBP')>0) then
    begin
      UneTOB.PutValue('SH_NUMDEPART',THSpinEdit(GetControl ('CPTFACTURE')).Value);
    end else
    begin
      UneTOB.PutValue('SH_NUMDEPART',THSpinEdit(GetControl ('CPTDEVIS')).Value);
    end;
  end;
end;

function TOF_BTCPTPIECE_S1.verifCompteur: boolean;
begin
  result := false;
  if ExistDocument('DBT',THSpinEdit(GetControl ('CPTDEVIS')).Value) then
  begin
    PgiBox('Le Numéro de devis existe déjà');
    SetFocusControl ('CPTDEVIS');
    exit;
  end;
  if ExistDocument('FBT',THSpinEdit(GetControl ('CPTFACTURE')).Value) then
  begin
    PgiBox('Le Numéro de facture existe déjà');
    SetFocusControl ('CPTFACTURE');
    exit;
  end;
  if ExistDocument('ABT',THSpinEdit(GetControl ('CPTAVOIR')).Value) then
  begin
    PgiBox('Le Numéro d''avoir existe déjà');
    SetFocusControl ('CPTAVOIR');
    exit;
  end;
  result := true;
end;

procedure TOF_BTCPTPIECE_S1.EcritInfos;
begin
  TOBSouches.UpdateDB (true);
end;

Initialization
  registerclasses ( [ TOF_BTCPTPIECE_S1 ] ) ;
end.

