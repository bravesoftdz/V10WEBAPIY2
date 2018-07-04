{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 15/04/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULCOMPTEURSDIF ()
Mots clefs ... : TOF;PGMULCOMPTEURSDIF
*****************************************************************}
Unit UtofPGMul_ParamFormAbs ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
     FE_Main,
     HDB,
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
     HQry,
     HMsgBox,
     HTB97,
     UTOF ; 

Type
  TOF_PGMULPARAMFORMABS = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    procedure GrilleDblClick (Sender : TObject) ;
    procedure Nouveau(Sender : TObject);
  end ;

Implementation


procedure TOF_PGMULPARAMFORMABS.GrilleDblClick (Sender : TObject) ;
var Code,Retour : String;
    Q_Mul : THQuery ;
begin
     {$IFDEF EAGLCLIENT}
     TFmul(Ecran).Q.TQ.Seek(TFmul(Ecran).FListe.Row-1) ;  //PT1
     {$ENDIF}
     Q_Mul := THQuery(Ecran.FindComponent('Q')) ;
     Code := Q_Mul.FindField('PPF_PARAMFORMPAIE').AsString;
     Retour := AGLLanceFiche('PAY','PARAMFORMABS','',Code,'');
     If Retour <> '' then TFMul(Ecran).BChercheClick(TFMul(Ecran).BCHerche);
end;

procedure TOF_PGMULPARAMFORMABS.Nouveau(Sender : TObject);
Var Retour : String;
begin
     Retour := AGLLanceFiche('PAY','PARAMFORMABS','','','ACTION=CREATION');
     If Retour <> '' then TFMul(Ecran).BChercheClick(TFMul(Ecran).BCHerche);
end;

procedure TOF_PGMULPARAMFORMABS.OnArgument (S : String ) ;
var    {$IFNDEF EAGLCLIENT}
        Liste : THDBGrid ;
        {$ELSE}
        Liste : THGrid ;
        {$ENDIF}
        BNew : TToolBarButton97;
begin
  Inherited ;
            {$IFNDEF EAGLCLIENT}
        Liste := THDBGrid(GetControl('FListe'));
        {$ELSE}
        Liste := THGrid(GetControl('FListe'));
        {$ENDIF}
        If Liste <> Nil Then Liste.OnDblClick := GrilleDblClick ;
        BNew := TToolBarButton97(Getcontrol('BInsert'));
       If BNew <> Nil then BNew.OnClick := Nouveau;
end ;

Initialization
  registerclasses ( [ TOF_PGMULPARAMFORMABS ] ) ;
end.

