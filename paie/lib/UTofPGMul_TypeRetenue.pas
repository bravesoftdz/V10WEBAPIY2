{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 01/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULTYPERETENUE ()
Mots clefs ... : TOF;PGMULTYPERETENUE
*****************************************************************}
Unit UTofPGMul_TypeRetenue ;

Interface

Uses StdCtrls,Controls,Classes,                     
{$IFNDEF EAGLCLIENT}
     DBGrids, HDB,Mul,Fiche,db,FE_Main,
{$ELSE}
     UtileAGL,MaineAgl,emul,HCTrls,
{$ENDIF}
     HEnt1,UTOF,Hqry,HTB97;

Type
  TOF_PGMULTYPERETENUE = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    procedure GrilleDblClick (Sedner : Tobject);
    procedure CreerTypeRetenue (Sender : TObject);
  end ;

Implementation

procedure TOF_PGMULTYPERETENUE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGMULTYPERETENUE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGMULTYPERETENUE.OnArgument (S : String ) ;
var {$IFNDEF EAGLCLIENT}
    Liste:THDBGrid;
    {$ELSE}
    Liste:THGrid;
    {$ENDIF}
    BOuv:TToolBarButton97;
begin
  Inherited ;
        {$IFNDEF EAGLCLIENT}
        Liste:=THDBGrid(GetControl('FLISTE'));
        {$ELSE}
        Liste:=THGrid(GetControl('FLISTE'));
        {$ENDIF}
        if Liste <> NIL then Liste.OnDblClick := GrilleDblClick;
        BOuv := TToolbarButton97 (getcontrol ('BInsert'));
        if Bouv <> NIL then  BOuv.OnClick := CreerTypeRetenue;
        SetControlProperty('PAVANCE','TabVisible',False);
        SetControlProperty('PCOMPLEMENT','TabVisible',False);
        If Not V_PGI.Debug then SetControlVisible('BINSERT',False);
end ;

procedure TOF_PGMULTYPERETENUE.GrilleDblClick (Sedner : Tobject);
var Q_Mul:THQuery ;
    St : String;
    {$IFNDEF EAGLCLIENT}
    Liste:THDBGrid;
    {$ELSE}
    Liste:THGrid;
    {$ENDIF}
begin
        {$IFNDEF EAGLCLIENT}
        Liste:=THDBGrid(GetControl('FLISTE'));
        {$ELSE}
        Liste:=THGrid(GetControl('FLISTE'));
        {$ENDIF}
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(Liste.Row-1) ;
        {$ENDIF}
        Q_Mul:=THQuery(Ecran.FindComponent('Q'));
        St := Q_Mul.FindField ('PTR_PREDEFINI').AsString + ';' + Q_Mul.FindField ('PTR_NODOSSIER').AsString +
        ';' + Q_Mul.FindField ('PTR_RETENUESAL').AsString;
        AGLLanceFiche('PAY','TYPERETENUE','',St,'ACTION=MODIFCATION');
        TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

procedure TOF_PGMULTYPERETENUE.CreerTypeRetenue (Sender : TObject);
begin
        AGLLanceFiche('PAY','TYPERETENUE','','','ACTION=CREATION');
        TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

Initialization
  registerclasses ( [ TOF_PGMULTYPERETENUE ] ) ;
end.

