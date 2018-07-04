{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 05/01/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : OPERATIONS_TV et OPERATIONS_TVF()
Mots clefs ... : TOF;OPERATIONS_TV;OPERATIONS_TVF
*****************************************************************}
Unit UtofOPERATIONS_TV ;

Interface

Uses StdCtrls, 
     Classes,
     UTobView,
{$IFNDEF EAGLCLIENT}
     Fe_Main,
{$else}
     MainEAGL,
{$ENDIF}
     forms,
     sysutils, 
     HEnt1,
     UTOF,
{$ifdef AFFAIRE}
    UTOFAFTRADUCCHAMPLIBRE,
{$endif}
    UtilSelection,ParamSoc,
     EntGC,
     FactUtil,
     UtilPGI,uEntCommun
     ;

Function RTLanceFiche_OPERATIONS_TV(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
{$ifdef AFFAIRE}
                //PL le 18/05/07 pour gérer les champs libres si paramétrés
     TOF_OPERATIONS_TV = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
     TOF_OPERATIONS_TV = Class (TOF)
{$endif}
     private
         TobViewer1: TTobViewer;
         procedure TVOnDblClickCell(Sender: TObject ) ;
     public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation
uses facture;

Function RTLanceFiche_OPERATIONS_TV(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_OPERATIONS_TV.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_OPERATIONS_TV.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_OPERATIONS_TV.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_OPERATIONS_TV.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_OPERATIONS_TV.OnArgument (S : String ) ;
var F : TForm;
    Memo : TMemo;
begin
{$ifdef AFFAIRE}
  fSansRestrictEtab := true; // PL le 12/09/07 : on ne veut pas appliquer la restriction sur les établissements dans le cas du marketing
{$endif}
  Inherited ;

  if S <> 'GRF' then
    begin
    F := TForm (Ecran);
    if GetParamSocSecur('SO_RTGESTINFOS002',False) = True then
        MulCreerPagesCL(F,'NOMFIC=RTOPERATIONS');
    if GetParamSocSecur ('SO_RTLIENOPERATIONPIECE',False) then
      begin
      SetControlVisible('PPIECE',True) ;
//      SetControlText ('GP_VENTEACHAT','VEN');
      memo := TMemo(GetControl('FSQL'));
      memo.lines[2] := 'RTINFOS002.*,PIECE.*,RTINFOS00D.* FROM OPERATIONS ';
      Memo.lines.Insert(4,'LEFT JOIN PIECE ON ROP_OPERATION = GP_OPERATION AND GP_VENTEACHAT = "VEN"');
      Memo.lines.Insert(5,'LEFT JOIN RTINFOS00D ON RDD_CLEDATA=CAST(GP_NBTRANSMIS AS varchar(7)) ');
      if GetParamSocSecur('SO_RTGESTINFOS00D',True) = True then
        begin
        F:=TForm (Ecran);
        MulCreerPagesCL(F,'NOMFIC=PIECES');
        end;
      end;
    end;
  TobViewer1:=TTobViewer(getcontrol('TV'));
  TobViewer1.OnDblClick:= TVOnDblClickCell ;
  SetControlText ('XX_WHERE','ROP_OPERATION<>"MODELES D''ACTIONS"');

end ;

procedure TOF_OPERATIONS_TV.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_OPERATIONS_TV.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_OPERATIONS_TV.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_OPERATIONS_TV.TVOnDblClickCell(Sender: TObject );
var CleDoc: R_CleDoc;
begin
with TTobViewer(sender) do
  begin
    if (copy(ColName[CurrentCol],1,2) = 'GP') or (copy(ColName[CurrentCol],1,3) = 'RDD') then
      begin
      DecodeRefPiece (
          AsString[ColIndex('GP_NATUREPIECEG'), CurrentRow] + ';' + AsString[ColIndex('GP_SOUCHE'), CurrentRow] + ';' +
               IntToStr(AsInteger[ColIndex('GP_NUMERO'), CurrentRow]) +
          ';' + IntToStr(AsInteger[ColIndex('GP_INDICEG'), CurrentRow]) + ';',
          CleDoc);
      CleDoc.DatePiece:=AsDateTime[ColIndex('GP_DATEPIECE'), CurrentRow];
      SaisiePiece (CleDoc, taConsult);
      end
    else
      if TForm (Ecran).Name <> 'RTOPERATIONS_TVF' then V_PGI.DispatchTT (23,taConsult ,AsString[ColIndex('ROP_OPERATION'), CurrentRow], '','');
  end;
end;

Initialization
  registerclasses ( [ TOF_OPERATIONS_TV ] ) ;
end.
