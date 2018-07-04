{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGENVOIFORM_MUL ()
Mots clefs ... : TOF;PGENVOIFORM_MUL
*****************************************************************
---- JL 17/10/2006 Modification contrôle des exercices de formations -----
}
Unit UTofPGMulEnvoiFormFact ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,DBGrids,HDB,DBCtrls,Mul,
{$ELSE}
      mainEAGL,UtilEagl,EMul,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,HTB97,Hqry,ParamDat,Utob ,PGOutilsFormation;

Type
  TOF_PGENVOIFORMFACT_MUL = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    Q_Mul : THQuery;
    Arg : String;
    procedure AfficheEnvoi(Sender : TObject);
    procedure DateElipsisclick(Sender :  TObject);
  end ;

Implementation

procedure TOF_PGENVOIFORMFACT_MUL.OnLoad ;
var Where : String;
begin
  Inherited ;
        Where := '';
        If StrtoInt(GetControlText('NUMENVOI'))>0 then Where := 'PVF_NUMENVOI='+GetControlText('NUMENVOI')+'';
        SetControlText('XX_WHERE',Where);
        If GetControlText('VETATFACT') <> 'TOUS' then
        begin
                If Where <> '' then Where := Where + ' AND ';
                if GetControlText('VETATFACT') = 'FRO' then Where := Where + 'PVF_RETOUROPCA="X"';
                if GetControlText('VETATFACT') = 'TF' then Where := Where + 'PVF_TRANSFACT="X"';
                if GetControlText('VETATFACT') = 'RF' then Where := Where + 'PVF_RETOURFAC="X"';
        end;
end ;

procedure TOF_PGENVOIFORMFACT_MUL.OnArgument (S : String ) ;
var {$IFNDEF EAGLCLIENT}
    Grille : THDBGrid;
    {$ELSE}
    Grille : THGrid;
    {$ENDIF}
    Date1,Date2 : THEdit;
    DateDebut,DateFin : TDateTime;
    Q : TQuery;
begin
  Inherited ;
        Arg := ReadTokenPipe(S,';');
        Q_Mul := THQuery(Ecran.FindComponent('Q'));
        {$IFNDEF EAGLCLIENT}
        Grille := THDBGrid (GetControl ('Fliste'));
        {$ELSE}
        Grille := THGrid (GetControl ('Fliste'));
        {$ENDIF}
        if Grille <> NIL then Grille.OnDblClick := AfficheEnvoi;
        Date1 := THEdit(GetControl('PVF_DATEENVOI'));
        If Date1 <> Nil then Date1.OnElipsisClick := DateElipsisClick;
        Date2 := THEdit(GetControl('PVF_DATEENVOI_'));
        If Date2 <> Nil then Date2.OnElipsisClick := DateElipsisClick;
        RendMillesimeRealise(DateDebut,DateFin);
        SetcontrolText('PVF_DATEENVOI',DateToStr(dateDebut));
        SetcontrolText('PVF_DATEENVOI_',DateToStr(datefin));
        SetControltext('VETATFACT','TOUS');
        SetControltext('PVF_STATUTENVOI','010');
end ;

procedure TOF_PGENVOIFORMFACT_MUL.AfficheEnvoi(Sender : TObject);
var St : String;
begin
        st := IntToStr(Q_MUL.FindField('PVF_NUMENVOI').AsInteger);
        AGLLanceFiche('PAY','ENVOIFORMATION','',St,Arg+';ACTION=MODIFICATION');
        TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

procedure TOF_PGENVOIFORMFACT_MUL.DateElipsisclick(Sender : TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

Initialization
  registerclasses ( [ TOF_PGENVOIFORMFACT_MUL ] ) ;
end.

