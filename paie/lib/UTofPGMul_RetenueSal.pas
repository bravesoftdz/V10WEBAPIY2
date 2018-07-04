{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 01/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULRETENUESAL ()
Mots clefs ... : TOF;PGMULRETENUESAL
*****************************************************************
PT1 16/01/2006 JL V_650 FQ 12818 Saisie atuo du matricule complet sur 10 caractères
PT2 30/01/2007 V_80 FC Mise en place filtrage des habilitations/poupulations
PT3 27/04/2007 V_720 JL Gestion accès depuis fiche salarié
PT4 16/07/2007 V_800 JL FQ 14541 Gestion elipsis salarié
PT5 21/12/2007 FC V_81 Concept accessibilité fiche salarié
}
Unit UTofPGMul_RetenueSal ;

Interface

Uses StdCtrls,Controls,Classes,                     
{$IFNDEF EAGLCLIENT}
      HDB,Mul,db,FE_Main,
{$ELSE}
     UtileAGL,MaineAgl,emul,
{$ENDIF}
     sysutils,HCtrls,HEnt1,UTOF,Hqry,HTB97,ParamDat,EntPaie,PGOutils2,P5Def,LookUp ;

Type
  TOF_PGMULRETENUESAL = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    Action:String;//PT5
    procedure GrilleDblClick (Sedner : Tobject);
    procedure CreerRetenue (Sender : TObject);
    procedure DateElipsisclick(Sender : TObject);
    procedure ExitEdit(Sender: TObject);    // PT1
    procedure SalarieElipsisClick(Sender : TObject);
  end ;

Implementation

procedure TOF_PGMULRETENUESAL.OnLoad ;
var StWhere, StConf : String;
  Where2: THEdit; //PT2
  Etab:String;  //PT5
begin
  Inherited ;
  //DEB PT5
  Etab:='';
  if GetControlText('ETABLISSEMENT') <> '' then
    Etab:='PSA_ETABLISSEMENT = "' + GetControlText('ETABLISSEMENT') + '"';
  //FIN PT5
  StWhere := '';
  StConf := SQLConf('SALARIES');
  If GetControlText('ETATRETENUE') = 'ACTIF' then StWhere := 'PRE_ACTIF ="X"';
  If GetControlText('ETATRETENUE') = 'NONACTIF' then StWhere := 'PRE_ACTIF ="-"';
  If StConf <> '' then
  begin
          If StWhere <> '' then
            if Etab<>'' then
              StWhere := StWhere + ' AND PRE_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE '+StConf+' AND ' +Etab+')'
            else
              StWhere := StWhere + ' AND PRE_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE '+StConf+')'
          else
            if Etab<>'' then
              StWhere := 'PRE_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE '+StConf+' AND ' + Etab+')'
            else
              StWhere := 'PRE_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE '+StConf+')';
  end;
  if (StWhere='') and (Etab<>'') then
    StWhere := 'PRE_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE ' + Etab+')';
  SetControlText('XX_WHERE',StWhere);

  // DEB PT2
  if Assigned(MonHabilitation) and (MonHabilitation.LeSQL<>'') then
  begin
    Where2 := THEdit(GetControl('XX_WHERE2'));
    if Where2 <> nil then
      SetControlText('XX_WHERE2', 'PRE_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE '+MonHabilitation.LeSQL+')');
  end;
  // FIN PT2

  //DEB PT5
  if Action='ACTION=CONSULTATION' then
    SetControlEnabled('BInsert',false);
  //FIN PT5
end ;

procedure TOF_PGMULRETENUESAL.OnArgument (S : String ) ;
var {$IFNDEF EAGLCLIENT}
    Liste:THDBGrid;
    {$ELSE}
    Liste:THGrid;
    {$ENDIF}
    BOuv:TToolBarButton97;
    Edit,Defaut : THEdit;
    Salarie : String;
begin
  Inherited ;
        Salarie := ReadTokenPipe(S,';');
        If Salarie <> '' then             //PT3
        begin
          SetControlEnabled('PRE_SALARIE',False);
          SetControlText('PRE_SALARIE',Salarie);
        end;
        //DEB PT5
        Action:='';
        if S <> '' then
          Action:=ReadTokenPipe(S,';');
        //FIN PT5
        {$IFNDEF EAGLCLIENT}
        Liste:=THDBGrid(GetControl('FLISTE'));
        {$ELSE}
        Liste:=THGrid(GetControl('FLISTE'));
        {$ENDIF}
        if Liste <> NIL then Liste.OnDblClick := GrilleDblClick;
        BOuv := TToolbarButton97 (getcontrol ('BInsert'));
        if Bouv <> NIL then  BOuv.OnClick := CreerRetenue;
        Edit := THEdit(GetControl('PRE_DATEDEBUT'));
        If Edit <> Nil then Edit.OnElipsisClick := DateElipsisClick;
        Edit := THEdit(GetControl('PRE_DATEDEBUT_'));
        If Edit <> Nil then Edit.OnElipsisClick := DateElipsisClick;
        Edit := THEdit(GetControl('PRE_DATEFIN'));
        If Edit <> Nil then Edit.OnElipsisClick := DateElipsisClick;
        Edit := THEdit(GetControl('PRE_DATEFIN_'));
        If Edit <> Nil then Edit.OnElipsisClick := DateElipsisClick;
        SetControlText('PRE_DATEDEBUT',DateToStr(IDate1900));
        SetControlText('PRE_DATEDEBUT_',DateToStr(IDate1900));
        SetControlText('PRE_DATEFIN',DateToStr(IDate1900));
        SetControlText('PRE_DATEFIN_',DateToStr(IDate1900));
        SetControlCaption('LIBSAL','');
        SetControlCaption('LIBETAB','');//PT2
        // d PT1
        Defaut:=ThEdit(getcontrol('PRE_SALARIE'));
        if Defaut<>nil then
        begin
          Defaut.OnElipsisClick := SalarieElipsisClick;  //PT4
          Defaut.OnExit:=ExitEdit;
        end;
        // f PT1
end ;

procedure TOF_PGMULRETENUESAL.GrilleDblClick (Sedner : Tobject);
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
        St := Q_Mul.FindField ('PRE_SALARIE').AsString + ';' + IntToStr(Q_Mul.FindField ('PRE_ORDRE').AsInteger) +
        ';' + DateToStr(Q_Mul.FindField ('PRE_DATEDEBUT').AsDateTime);
        if Action='ACTION=CONSULTATION' then
          AGLLanceFiche('PAY','RETENUESALAIRE','',St,'ACTION=CONSULTATION')
        else
          AGLLanceFiche('PAY','RETENUESALAIRE','',St,'ACTION=MODIFICATION');
        TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

procedure TOF_PGMULRETENUESAL.CreerRetenue (Sender : TObject);
begin
        AGLLanceFiche('PAY','RETENUESALAIRE','','','ACTION=CREATION;'+GetControlText('PRE_SALARIE'));     //PT3
        TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

procedure TOF_PGMULRETENUESAL.DateElipsisclick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGMULRETENUESAL.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and
       (length(Edit.text)<11) and
       (isnumeric(edit.text)) then
      edit.text:=AffectDefautCode(edit,10);
end;  { fin ExitEdit}

procedure TOF_PGMULRETENUESAL.SalarieElipsisClick(Sender : TObject); //PT4
begin
  LookupList(THEdit(Sender), 'Liste des salariés','SALARIES','PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', '', 'PSA_SALARIE', TRUE, -1);
end;


Initialization
  registerclasses ( [ TOF_PGMULRETENUESAL ] ) ;
end.

