{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 15/04/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULGESTCARRIERE ()
Mots clefs ... : TOF;PGMULGESTCARRIERE
*****************************************************************}
Unit UTOFPGMulGestionCarriere ;

Interface

Uses StdCtrls,Classes,
{$IFNDEF EAGLCLIENT}
     db,FE_Main,HDB,
{$else}
     eMul,uTob,MaineAGL,
{$ENDIF}
     sysutils,HCtrls,HEnt1,UTOF,HQry,EntPaie,P5Def,PGOutils2 ;

Type
  TOF_PGMULGESTCARRIERE = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    Private
    procedure GrilleDblClick (Sender : TObject) ;
    procedure ExitEdit(Sender: TObject);
    procedure OnClickSalarieSortie(Sender: TObject);
  end ;

  TOF_PGMULCHANGEPOSTE = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    Private
    procedure GrilleDblClick (Sender : TObject) ;
    procedure ExitEdit(Sender: TObject);
    procedure OnClickSalarieSortie(Sender: TObject);
  end ;

Implementation

procedure TOF_PGMULGESTCARRIERE.OnLoad ;
var DateArret : TDateTime;
    StDateArret : String;
begin
  Inherited ;
        if (GetControlText('CKSORTIE')='X') and (IsValidDate(GetControlText('DATEARRET')))then
        Begin
                DateArret:=StrtoDate(GetControlText('DATEARRET'));
                StDateArret:=' AND (PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) ';
                StDateArret:=StDateArret + ' AND PSA_DATEENTREE <="'+UsDateTime(DateArret)+'"';
                SetControlText('XX_WHERE',StDateArret);
        End
        else SetControlText('XX_WHERE','');
end ;

procedure TOF_PGMULGESTCARRIERE.OnArgument (S : String ) ;
var    {$IFNDEF EAGLCLIENT}
        Liste : THDBGrid ;
        {$ELSE}
        Liste : THGrid ;
        {$ENDIF}
        Check : TCheckBox;
        Num : Integer;
        Defaut : THEdit;
begin
  Inherited ;
        {$IFNDEF EAGLCLIENT}
        Liste := THDBGrid(GetControl('FListe'));
        {$ELSE}
        Liste := THGrid(GetControl('FListe'));
        {$ENDIF}
        If Liste <> Nil Then Liste.OnDblClick := GrilleDblClick ;
        SetControlvisible('DATEARRET',True);
        SetControlvisible('TDATEARRET',True);
        SetControlEnabled('DATEARRET',False);
        SetControlEnabled('TDATEARRET',False);
        Check:=TCheckBox(GetControl('CKSORTIE'));
        if Check=nil then
        Begin
                SetControlVisible('DATEARRET',False);
                SetControlVisible('TDATEARRET',False);
        End
        else Check.OnClick:=OnClickSalarieSortie;
        For Num := 1 to VH_Paie.PGNbreStatOrg do
        begin
                if Num >4 then Break;
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
        end;
        VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ; //PT3
        Defaut:=ThEdit(getcontrol('PSA_SALARIE'));
        If Defaut<>nil then Defaut.OnExit:=ExitEdit;
end ;

procedure TOF_PGMULGESTCARRIERE.GrilleDblClick (Sender : TOBject);
var St : String ;
    Q_Mul : THQuery ;
begin
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(TFmul(Ecran).FListe.Row-1) ;
        {$ENDIF}
        Q_Mul := THQuery(Ecran.FindComponent('Q')) ;
        St := Q_Mul.FindField('PSA_SALARIE').AsString;
        AGLLanceFiche('PAY','GESTCARRIERE','','',St);
end ;

procedure TOF_PGMULGESTCARRIERE.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

procedure TOF_PGMULGESTCARRIERE.OnClickSalarieSortie(Sender: TObject);
begin
SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
end;

{TOF_PGMULCHANGEPOSTE}

procedure TOF_PGMULCHANGEPOSTE.OnLoad ;
var DateArret : TDateTime;
    StDateArret : String;
begin
  Inherited ;
        if (GetControlText('CKSORTIE')='X') and (IsValidDate(GetControlText('DATEARRET')))then
        Begin
                DateArret:=StrtoDate(GetControlText('DATEARRET'));
                StDateArret:=' AND (PSI_DATESORTIE>="'+UsDateTime(DateArret)+'" OR PSI_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSI_DATESORTIE IS NULL) ';
                StDateArret:=StDateArret + ' AND PSI_DATEENTREE <="'+UsDateTime(DateArret)+'"';
                SetControlText('XX_WHERE',StDateArret);
        End
        else SetControlText('XX_WHERE','');
end ;

procedure TOF_PGMULCHANGEPOSTE.OnArgument (S : String ) ;
var    {$IFNDEF EAGLCLIENT}
        Liste : THDBGrid ;
        {$ELSE}
        Liste : THGrid ;
        {$ENDIF}
        Check : TCheckBox;
        Num : Integer;
        Defaut : THEdit;
begin
  Inherited ;
        {$IFNDEF EAGLCLIENT}
        Liste := THDBGrid(GetControl('FListe'));
        {$ELSE}
        Liste := THGrid(GetControl('FListe'));
        {$ENDIF}
        If Liste <> Nil Then Liste.OnDblClick := GrilleDblClick ;
        SetControlvisible('DATEARRET',True);
        SetControlvisible('TDATEARRET',True);
        SetControlEnabled('DATEARRET',False);
        SetControlEnabled('TDATEARRET',False);
        Check:=TCheckBox(GetControl('CKSORTIE'));
        if Check=nil then
        Begin
                SetControlVisible('DATEARRET',False);
                SetControlVisible('TDATEARRET',False);
        End
        else Check.OnClick:=OnClickSalarieSortie;
        For Num := 1 to VH_Paie.PGNbreStatOrg do
        begin
                if Num >4 then Break;
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSI_TRAVAILN'+IntToStr(Num)),GetControl ('TPSI_TRAVAILN'+IntToStr(Num)));
        end;
        VisibiliteStat (GetControl ('PSI_CODESTAT'),GetControl ('TPSI_CODESTAT')) ; //PT3
        Defaut:=ThEdit(getcontrol('PSI_SALARIE'));
        If Defaut<>nil then Defaut.OnExit:=ExitEdit;
end ;

procedure TOF_PGMULCHANGEPOSTE.GrilleDblClick (Sender : TOBject);
var St : String ;
    Q_Mul : THQuery ;
begin
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(TFmul(Ecran).FListe.Row-1) ;
        {$ENDIF}
        Q_Mul := THQuery(Ecran.FindComponent('Q')) ;
        St := Q_Mul.FindField('PSI_INTERIMAIRE').AsString;
        AGLLanceFiche('PAY','CHANGEPOSTE','','',St);
end ;

procedure TOF_PGMULCHANGEPOSTE.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

procedure TOF_PGMULCHANGEPOSTE.OnClickSalarieSortie(Sender: TObject);
begin
SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
end;


Initialization
  registerclasses ( [ TOF_PGMULGESTCARRIERE,TOF_PGMULCHANGEPOSTE ] ) ;
end.

