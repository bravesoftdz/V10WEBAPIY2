{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 03/10/2011
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTMULPIECEINTERV ()
Mots clefs ... : TOF;BTMULPIECEINTERV
*****************************************************************}
Unit BTMULPIECEINTERV_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     forms,
     Sysutils,
     ComCtrls,
{$IFDEF EAGLCLIENT}
    MaineAGL,emul,
{$ELSE}
    {$IFNDEF ERADIO}
    Fe_Main,
    {$ENDIF !ERADIO}
    {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
    HDB,
    Mul,
{$ENDIF}
		 Grids,
     uTob,
     HCtrls,
     HEnt1,
     HMsgBox,
     uEntCommun,
     UCotraitance,
     HTB97,
     UTOF,
     uTOFComm,
     Ent1,
     AglInit;
//

Type
  TOF_BTMULPIECEINTERV = Class (tTOFComm)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    TOBParam  : TOB;
    TOBPiece  : TOB;
    fCledoc   : R_cledoc;
    Binsert   : TToolbarButton97;
    procedure GSDblclick (Sender : TObject);
    procedure OnInsert(Sender: Tobject);
  end ;

Implementation

procedure TOF_BTMULPIECEINTERV.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTMULPIECEINTERV.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTMULPIECEINTERV.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTMULPIECEINTERV.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTMULPIECEINTERV.OnArgument (S : String ) ;
var Refpiece : string;
begin
  fMulDeTraitement := true;
  Inherited ;

  BInsert := TToolbarBUtton97(getControl('BINSERT'));
  Binsert.OnClick := OnInsert;

  TOBPARAM := LaTOB;
  TOBPiece := TOB(LaTOB.Data);

  Refpiece := EncodeCleDoc (TOBPiece);

  DecodeCleDoc(Refpiece,fcledoc);

  SetControlText('BPI_NATUREPIECEG',fCledoc.NaturePiece);
  SetControlText('BPI_SOUCHE',fCledoc.Souche);
  SetControlText('BPI_NUMERO',IntToStr(fCledoc.NumeroPiece));
  SetControlText('BPI_INDICEG',IntToStr(fCledoc.Indice));
  SetControlText('AFFAIRE',TOBPiece.GetValue('GP_AFFAIRE'));
  THDBEdit(GetControl('fliste')).OnDblClick := GSDblclick;
end ;

procedure TOF_BTMULPIECEINTERV.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTMULPIECEINTERV.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTMULPIECEINTERV.OnCancel () ;
begin
  Inherited ;
end ;


procedure TOF_BTMULPIECEINTERV.GSDblclick(Sender: TObject);
var Tiers,ModePaie: string;
begin
  Tiers:=ThDbGrid(GetControl('Fliste')).datasource.dataset.FindField('BPI_TIERSFOU').AsString;
  ModePaie:=ThDbGrid(GetControl('Fliste')).datasource.dataset.FindField('BPI_TYPEPAIE').AsString;
  TOBParam.putValue('SOUSTRAIT',Tiers);
  TOBParam.putValue('MODEPAIE',ModePaie);
  TFMUL(Ecran).Close;
end;

Procedure TOF_BTMULPIECEINTERV.OnInsert(Sender: Tobject);
Var Params : string;
begin

   params := 'BPI_NATUREPIECEG='+ GetControlText('BPI_NATUREPIECEG')  + ';' +
             'BPI_SOUCHE='      + GetControlText('BPI_SOUCHE')        + ';' +
             'BPI_NUMERO='      + GetControlText('BPI_NUMERO')        + ';' +
             'BPI_INDICEG='     + GetControlText('BPI_INDICEG')       + ';' +
             'BPI_TIERSFOU='    + GetControlText('BPI_TIERSFOU')      + ';' +
             'AFFAIRE='         + GetControlText('AFFAIRE');

   AGLLanceFiche('BTP','BTPIECEINTERV','','','ACTION=CREATION;' + Params);
   RefreshDB;

end;

Initialization
  registerclasses ( [ TOF_BTMULPIECEINTERV ] ) ;
end.
