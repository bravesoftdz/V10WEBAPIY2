{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 29/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULCALCFRACTION ()
Mots clefs ... : TOF;PGMULCALCFRACTION
*****************************************************************}
Unit UTofPGMul_calcFraction ;

Interface

Uses Classes,
{$IFNDEF EAGLCLIENT}
     db,FE_Main,Mul,  HDB,
{$ELSE}
     MainEAGL,eMul,
{$ENDIF}
     HCtrls,HEnt1,UTOF,HTB97 ;

Type
  TOF_PGMULCALCFRACTION = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    Arg : String;
    procedure CreerFraction (Sender : TObject);
    procedure GrilleDblClick (Sender : TObject);
  end ;

Implementation


procedure TOF_PGMULCALCFRACTION.OnArgument (S : String ) ;
var BOuv : TToolBarButton97;
    {$IFNDEF EAGLCLIENT}
    Liste : THDBGrid;
    {$ELSE}
    Liste : THGrid;
    {$ENDIF}
begin
  Inherited ;
        Arg := ReadTokenPipe(S,';');
        BOuv  :=  TToolbarButton97 (getcontrol ('BInsert'));
        if Bouv  <>  NIL then  BOuv.OnClick  :=  CreerFraction;
        {$IFNDEF EAGLCLIENT}
        Liste := THDBGrid(GetControl('FLISTE'));
        {$ELSE}
        Liste := THGrid(GetControl('FLISTE'));
        {$ENDIF}
        if Liste  <>  NIL then Liste.OnDblClick  :=  GrilleDblClick;
        SetControlVisible('PCF_TYPECALCFRAC',False);
        SetControlVisible('TPCF_TYPECALCFRAC',False);
        SetControlProperty('PAVANCE','TabVisible',False);
        SetControlProperty('PCOMPLEMENT','TabVisible',False);
        If Not V_PGI.Debug then SetControlVisible('BINSERT',False);
end ;

procedure TOF_PGMULCALCFRACTION.CreerFraction (Sender : TObject);
begin
        AGLLanceFiche('PAY','CALCFRACTION','','',Arg+';ACTION=CREATION');
        TFMul(Ecran).BCherche.Click;
end;

procedure TOF_PGMULCALCFRACTION.GrilleDblClick (Sender : TObject);
var St : String;
begin
        St := TFmul(Ecran).Q.FindField('PCF_CALCFRACTION').AsString;
        AGLLanceFiche('PAY','CALCFRACTION','',St,Arg+';ACTION=MODIFICATION');
        TFMul(Ecran).BCherche.Click;
end;

Initialization
  registerclasses ( [ TOF_PGMULCALCFRACTION ] ) ;
end.

