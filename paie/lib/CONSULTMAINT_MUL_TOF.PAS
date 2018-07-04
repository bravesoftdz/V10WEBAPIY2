{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 28/07/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CONSULTMAINT_MUL ()
Mots clefs ... : TOF;CONSULTMAINT_MUL
*****************************************************************}
{
PT1     12/01/2004 MF V_60 Formatage du matricule
PT2     06/03/2006 MF V_65 Ajout champ PSA_LIBELLE (nom commençant par) et
                           fourchette sur PCN_SALARIE remplacé par MultiValComboBox
}
Unit CONSULTMAINT_MUL_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
//unused     db, 
//unused     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
    mul,
{$else}
    eMul,
//unused     uTob,
{$ENDIF}
     Entpaie,
//unused     forms,
     sysutils, 
//unused     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     HQry,
     ParamDat,
     PgOutils2,
     UTOF ; 

Type
  TOF_CONSULTMAINT_MUL = Class (TOF)
    private
    AIntegrer                 : TCheckBox;
    WW                        : THEdit;
    Q_Mul                     : THQuery;
    WdateDeb, WDateFin        : TDateTime;
    DateDeb,DateFin : THEdit;

    procedure ActiveWhere(Okok : Boolean);
    procedure DateElipsisclick(Sender: TObject);
    procedure Change(Sender: TObject);
    procedure ExitEdit(Sender: TObject);    // PT1

    public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

procedure TOF_CONSULTMAINT_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CONSULTMAINT_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CONSULTMAINT_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CONSULTMAINT_MUL.OnLoad ;
var
  Okok : Boolean;

begin
  Inherited ;
  Okok := TRUE;
  ActiveWhere (Okok);

end ;

procedure TOF_CONSULTMAINT_MUL.OnArgument (S : String ) ;
// PT1
var
  Defaut                        : THEdit;
// PT1
begin
  Inherited ;
  WW:=THEdit (GetControl ('XX_WHERE'));

  DateDeb:= THEdit(getcontrol('XX_VARIABLED'));
  DateFin:= THEdit(getcontrol('XX_VARIABLED_'));

  AIntegrer := TCheckBox(getcontrol('AINTEGRER'));

  if (DateDeb <> NIL) and (DateFin <> NIL) then
  begin
    DateDeb.OnElipsisClick := DateElipsisclick;
    DateDeb.OnExit := Change;
    DateFin.OnElipsisClick := DateElipsisclick;
    DateFin.OnExit := Change;
  end;
  WDateDeb := IDate1900;
  if (DateDeb <> NIL) and (DateDeb.text = '') then
    WDateDeb := StrToDate(DateDeb.text);
  WDatefin := IDate1900;
  if (DateFin <> NIL) and (DateFin.text = '') then
    WDateFin := StrToDate(DateFin.text);
// d PT1
  Defaut:=ThEdit(getcontrol('PMT_SALARIE'));
  if Defaut<>nil then Defaut.OnExit:=ExitEdit;
// f PT1
end ;

procedure TOF_CONSULTMAINT_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CONSULTMAINT_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CONSULTMAINT_MUL.OnCancel () ;
begin
  Inherited ;
end ;
procedure TOF_CONSULTMAINT_MUL.ActiveWhere(Okok : Boolean);
var
  st                                    : string;
// d PT2
   Salarie                                   : String;
   StCondition                               : String;
   Indst                                     : integer;
// f PT2

begin
  WW.Text := '';

  st := '';

  if (AIntegrer.checked = True) then
  begin
    st := st + 'PMT_DATEDEBUT = "'+UsDateTime(IDate1900)+'" AND '+
               'PMT_DATEFIN = "'+UsDateTime(IDate1900)+'"';
    DateDeb.text := '';
    DateFin.text := '';
  end
  else
    if (WDatedeb <> IDate1900) and (WDatefin <> IDate1900) then
     st := st + 'PMT_DATEDEBUT = "'+UsDateTime(WDatedeb)+'" AND '+
                'PMT_DATEFIN = "'+UsDateTime(WDatefin)+'"';
// d PT2
  Salarie := (THMultiValComboBox(GetControl('SALARIE')).value);
  if (salarie <> '' ) then
  begin
       st := st + 'AND (';
       Indst := 0;
       while(salarie <> '') do
       begin
            if (Indst <> 0) then
               st := st + ' OR ';
            StCondition := ReadTokenPipe(Salarie,';');
            st := st + 'PMT_SALARIE = "'+ StCondition+'"';
            Indst := 1;
       end;
       st := st +')';
  end;

// f PT2

  if St <> '' then WW.Text := st;
  if Q_Mul <> NIL then
  begin
    TFMul(Ecran).SetDBListe('PGMAINTIEN');
  end;
end;
procedure TOF_CONSULTMAINT_MUL.DateElipsisclick(Sender: TObject);
var
  key : char;
begin
  key := '*';
  ParamDate (Ecran, Sender, Key);
end;  {fin DateElipsisclick}
procedure TOF_CONSULTMAINT_MUL.Change(Sender: TObject);
begin
  if not IsValidDate(GetControlText('XX_VARIABLED')) then
  {Pour générer message erreur si date erronnée}
  begin
    PGIBox('La date de début est erronée.',Ecran.caption);
    SetControlText('XX_VARIABLED',DatetoStr(Date));

  end;
  WDateDeb :=  StrToDate(GetControlText('XX_VARIABLED'));
  if not IsValidDate(GetControlText('XX_VARIABLED_'))then

  {Pour générer message erreur si date erronnée}
  begin
    PGIBox('La date de fin est erronée.',Ecran.caption);
    SetControlText('XX_VARIABLED_',DatetoStr(Date));
  end;
    WDateFin :=  StrToDate(GetControlText('XX_VARIABLED_'));

end;  {fin Change}

// d PT1
procedure TOF_CONSULTMAINT_MUL.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and
       (length(Edit.text)<11) and
       (isnumeric(edit.text)) then
      edit.text:=AffectDefautCode(edit,10);
end;  { fin ExitEdit}
// f PT1
Initialization
  registerclasses ( [ TOF_CONSULTMAINT_MUL ] ) ;
end.
