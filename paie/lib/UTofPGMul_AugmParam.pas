{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 12/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULPARAMAUGM ()
Mots clefs ... : TOF;PGMULPARAMAUGM
*****************************************************************}
Unit UTofPGMul_AugmParam ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     hdb,
     FE_Main,
{$else}
     eMul,
     uTob,
     MainEAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     HQry,
     UTOF ;

Type
  TOF_PGMULPARAMAUGM = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    procedure GrilleDblClick(Sender : TObject);
    procedure CreerCritere(Sender : TObject);
  end ;

Implementation

procedure TOF_PGMULPARAMAUGM.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGMULPARAMAUGM.OnArgument (S : String ) ;
var    {$IFNDEF EAGLCLIENT}
        Liste : THDBGrid ;
        {$ELSE}
        Liste : THGrid ;
        {$ENDIF}
        BOuv:TToolBarButton97;
begin
  Inherited ;
        {$IFNDEF EAGLCLIENT}
        Liste := THDBGrid(GetControl('FListe'));
        {$ELSE}
        Liste := THGrid(GetControl('FListe'));
        {$ENDIF}
        If Liste <> Nil Then Liste.OnDblClick := GrilleDblClick ;
        BOuv := TToolbarButton97 (getcontrol ('BInsert'));
        if Bouv <> NIL then  BOuv.OnClick := CreerCritere;
end ;

procedure TOF_PGMULPARAMAUGM.GrilleDblClick (Sender : TObject) ;
var St : String ;
    Q_Mul : THQuery ;
begin
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(TFmul(Ecran).FListe.Row-1) ;
        {$ENDIF}
        Q_Mul := THQuery(Ecran.FindComponent('Q')) ;
        St := Q_Mul.FindField('PAP_CRITEREAUGM').AsString;
        AGLLanceFiche('PAY','AUGMPARAM','',St,'');
end ;

procedure TOF_PGMULPARAMAUGM.CreerCritere (Sender : TObject);
begin
        AGLLanceFiche('PAY','AUGMPARAM','','','ACTION=CREATION');
        TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

Initialization
  registerclasses ( [ TOF_PGMULPARAMAUGM ] ) ; 
end.

