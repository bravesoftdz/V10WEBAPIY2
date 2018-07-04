Unit UTOF_TBSTATPAIE;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     dbtables,
     mul,
     Fe_Main,
{$else}
     eMul,
     uTob,
     MaineAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     stat,
     DPTableauBordLibrairie;

Type
  TOF_TBSTATPAIE = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnUpdate                 ; override ;    
   private 
    ChaineSql : String;
    procedure AfterShow;
  end ;

//---------------------------------
//--- Déclaration des procédures
//---------------------------------
procedure Aff_TBStatPaie (ChSql : String);

Implementation

//--------------------------------------------------------
//--- Nom   : Aff_DPMulPaie
//--- Objet : Affichage de la fiche Multicritére Paie
//--------------------------------------------------------
procedure Aff_TBStatPaie (ChSQl : String);
begin
 AGLLanceFiche('DP','TBSTATPAIE','','',ChSql);
end;

//---------------------------------------------------
//--- Nom   : OnArgument
//--- Objet : Initialisation de la fiche TBSTATPAIE
//---------------------------------------------------
procedure TOF_TBSTATPAIE.OnArgument (S : String ) ;
begin
 Inherited;
 ChaineSql:=READTOKENST(S);
 ChaineSql:=StringReplace (ChaineSql,'''','"',[rfReplaceAll, rfIgnoreCase]);
 ChaineSql:=StringReplace (ChaineSql,V_PGI.DbName+'.dbo.','',[rfReplaceAll, rfIgnoreCase]);

 // while (RemplacerChaine (ChaineSql,'''','"')) do;
 // while (RemplacerChaine (ChaineSQl,V_PGI.DbName+'.dbo.','')) do;

 TFStat(Ecran).OnAfterFormShow := AfterShow;
 TFStat(Ecran).FiltreDisabled:=true ;
end ;

//----------------------------------------------------
//--- Nom   : AfterShow
//--- Objet : Initialisation de la fiche TBSTATPAIE
//----------------------------------------------------
procedure TOF_TBSTATPAIE.AfterShow;
begin
 Inherited;
 TFStat(Ecran).BChercheClick(Nil);
end;

//------------------------------------------
//--- Nom   : OnUpdate
//--- Objet : Maj de la fiche TBSTATPAIE
//------------------------------------------
procedure TOF_TBSTATPAIE.OnUpdate ;
begin
 Inherited;
 TFStat(Ecran).StSql:=ChaineSql;
end ;

Initialization
  registerclasses ( [ TOF_TBSTATPAIE ] ) ;
end.

