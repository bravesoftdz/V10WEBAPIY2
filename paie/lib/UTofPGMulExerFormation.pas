{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 15/05/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : EXERFORMATION ()
Mots clefs ... : TOF;PGEXERFORMATION
*****************************************************************}
Unit UTofPGMulExerFormation ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,DBGrids,Mul,FE_Main,
{$ELSE}
     eMul,MaineAGL,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,HQry ;

Type
  TOF_PGEXERFORMATION = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    {$IFNDEF EAGLCLIENT}
    Liste : THDBGrid ;
    {$ELSE}
    Liste : THGrid ;
    {$ENDIF}
    procedure GrilleDblClick(Sender : TObject);
  end ;

Implementation

procedure TOF_PGEXERFORMATION.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGEXERFORMATION.OnArgument (S : String ) ;
begin
  Inherited ;
        {$IFNDEF EAGLCLIENT}
        Liste := THDBGrid(GetControl('FListe'));
        {$ELSE}
        Liste := THGrid(GetControl('FListe'));
        {$ENDIF}
        If Liste <> Nil Then Liste.OnDblClick := GrilleDblClick ;
end ;

procedure TOF_PGEXERFORMATION.GrilleDblClick(Sender : TObject);
var St : String;
     Q_Mul : THQuery ;

begin
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(Liste.Row-1) ;
        {$ENDIF}
        Q_Mul := THQuery(Ecran.FindComponent('Q')) ;
        St := Q_Mul.FindField('PFE_MILLESIME').AsString;
        AGLLanceFiche('PAY','INVESTFORMATION','','',St);
        TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

Initialization
  registerclasses ( [ TOF_PGEXERFORMATION ] ) ;
end.

