{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 12/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULPARAMAUGM ()
Mots clefs ... : TOF;PGMULPARAMAUGM
*****************************************************************}
Unit UTofPGMul_AugmExclus;

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
  TOF_PGMULAUGMEXCLUS = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    Q_Mul : THQuery ;
    procedure GrilleDblClick(Sender : TObject);
    procedure CreerCritere(Sender : TObject);
  end ;

Implementation

procedure TOF_PGMULAUGMEXCLUS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGMULAUGMEXCLUS.OnArgument (S : String ) ;
var    {$IFNDEF EAGLCLIENT}
        Liste : THDBGrid ;
        {$ELSE}
        Liste : THGrid ;
        {$ENDIF}
        BOuv:TToolBarButton97;
        LeCritere : String;
begin
  Inherited ;
     Q_Mul := THQuery(Ecran.FindComponent('Q')) ;
        {$IFNDEF EAGLCLIENT}
        Liste := THDBGrid(GetControl('FListe'));
        {$ELSE}
        Liste := THGrid(GetControl('FListe'));
        {$ENDIF}
        If Liste <> Nil Then Liste.OnDblClick := GrilleDblClick ;
        BOuv := TToolbarButton97 (getcontrol ('BInsert'));
        if Bouv <> NIL then  BOuv.OnClick := CreerCritere;
        LeCritere := ReadTokenPipe(S,';');
        If LeCritere > '' then
        begin
             SetControlText('PAE_CRITEREAUGM',LeCritere);
             SetControlEnabled('PAE_CRITEREAUGM',False);
        end;
end ;

procedure TOF_PGMULAUGMEXCLUS.GrilleDblClick (Sender : TObject) ;
var St : String ;
begin
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(TFmul(Ecran).FListe.Row-1) ;
        {$ENDIF}
        St := Q_Mul.FindField('PAE_ORDRE').AsString+';'+Q_Mul.FindField('PAE_CRITEREAUGM').AsString;
        AGLLanceFiche('PAY','AUGMEXCLUS','',St,GetControlText('PAE_CRITEREAUGM'));
end ;

procedure TOF_PGMULAUGMEXCLUS.CreerCritere (Sender : TObject);
begin
        AGLLanceFiche('PAY','AUGMEXCLUS','','',GetControlText('PAE_CRITEREAUGM')+';ACTION=CREATION');
        TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

Initialization
  registerclasses ( [ TOF_PGMULAUGMEXCLUS ] ) ;
end.

