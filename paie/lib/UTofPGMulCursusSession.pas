{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 14/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULCURSUSSESSION ()
Mots clefs ... : TOF;PGMULCURSUSSESSION
*****************************************************************
---- JL 17/10/2006 Modification contrôle des exercices de formations -----
}
Unit UTofPGMulCursusSession ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,DBCtrls,Mul,Fe_Main,DBGrids,
{$ELSE}
     MaineAGL,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,ParamDat,HTB97,Hqry,UTob,PGOutilsFormation ;

Type
  TOF_PGMULCURSUSSESSION = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    procedure GrilleDblClick(Sender : TObject);
    procedure DateElipsisclick(Sender : TObject);
    procedure CreatCursus(Sender : TObject);
  end ;

Implementation

procedure TOF_PGMULCURSUSSESSION.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGMULCURSUSSESSION.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGMULCURSUSSESSION.OnArgument (S : String ) ;
var DateDeb,dateFin : TDateTime;
    {$IFNDEF EAGLCLIENT}
    Grille : THDBGrid;
    {$ELSE}
    Grille : THGrid;
    {$ENDIF}
    Edit : THEdit;
    BOuv : TToolBarButton97;
begin
  Inherited ;
        DateDeb := Date;
        dateFin := Date;
        RendMillesimeRealise(DateDeb,DateFin);
        SetControlText('PCU_DATEDEBUT',DateToStr(DateDeb));
        SetControlText('PCU_DATEDEBUT_',DateToStr(DateFin));
        SetControlText('PCU_DATEFIN',DateToStr(DateDeb));
        SetControlText('PCU_DATEFIN_',DateToStr(DateFin));
        {$IFNDEF EAGLCLIENT}
        Grille := THDBGrid (GetControl ('Fliste'));
        {$ELSE}
        Grille := THGrid (GetControl ('Fliste'));
        {$ENDIF}
        If Grille <> Nil then Grille.OnDblClick := GrilleDblClick;
        Edit := THEdit(GetControl('PCU_DATEDEBUT'));
        if Edit <> Nil then Edit.OnElipsisClick := DateElipsisclick;
        Edit := THEdit(GetControl('PCU_DATEDEBUT_'));
        if Edit <> Nil then Edit.OnElipsisClick := DateElipsisclick;
        Edit := THEdit(GetControl('PCU_DATEFIN'));
        if Edit <> Nil then Edit.OnElipsisClick := DateElipsisclick;
        Edit := THEdit(GetControl('PCU_DATEFIN_'));
        if Edit <> Nil then Edit.OnElipsisClick := DateElipsisclick;
        BOuv  :=  TToolbarButton97 (getcontrol ('BInsert'));
        if Bouv  <>  NIL then  BOuv.OnClick  :=  CreatCursus;
end ;

procedure TOF_PGMULCURSUSSESSION.GrilleDblClick(Sender : TObject);
var Q_Mul : THQuery;
    St : String;
    Cursus,Rang : String;
    Bt : TToolBarButton97;
begin
        Q_Mul := THQuery(Ecran.FindComponent('Q'));
        If Q_Mul = Nil Then Exit;
        Cursus := Q_MUL.FindField('PCU_CURSUS').AsString;
        Rang := IntToStr(Q_MUL.FindField('PCU_RANGCURSUS').AsInteger);
        St := Cursus + ';' + Rang;
        AGLLanceFiche('PAY','CURSUSSESSION','','','MODIFICATION;'+St);
        Bt  :=  TToolbarButton97 (GetControl('BCherche'));
        if Bt  <>  NIL then Bt.click;
end;

procedure TOF_PGMULCURSUSSESSION.DateElipsisclick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGMULCURSUSSESSION.CreatCursus(Sender : TObject);
var Bt : TToolBarButton97;
begin
        AGLLanceFiche('PAY','CURSUSSESSION','','','CREATION');
        Bt  :=  TToolbarButton97 (GetControl('BCherche'));
        if Bt  <>  NIL then Bt.click;
end;

Initialization
  registerclasses ( [ TOF_PGMULCURSUSSESSION ] ) ;
end.

