{***********UNITE*************************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 12/05/2004
Modifié le ... : 09/06/2004
Description .. : Source TOF de la FICHE : MUL_REGLTIJSS ()
Suite ........ : Suivi du régelement des IJSS
Mots clefs ... : TOF;MUL_REGLTIJSS; IJSS
*****************************************************************}
{
 PT1    : 06/03/2006 MF V_65 : Ajout champ PSA_LIBELLE (Nom commençant par)
                               et remplacement du choix de fourchette pour matricule
                               par un multivalcombobox

}
Unit MUL_REGLTIJSS_TOF ;

Interface

Uses
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
//unused     StdCtrls, 
     Controls,
     Classes,
     Hctrls,
{$IFNDEF EAGLCLIENT}
//unused     db,
//unused     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,
     HDB,
     mul,
{$else}
     eMul,
     MainEAgl,
//unused     uTob,
{$ENDIF}
     Entpaie,
//unused     forms,
     HTB97,
     sysutils,
//unused     ComCtrls,
     HEnt1,
     HMsgBox,
     HQry,
     ParamDat,
//unuesd     PgOutils,
     PgOutils2,
     UTOF ; 

Type
  TOF_MUL_REGLTIJSS = Class (TOF)
    public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

    private
    WW : THEdit;
    WdateDeb, WDateFin : TDateTime;
    Q_Mul:THQuery;
    
    procedure CalculerClick(Sender: TObject);
    procedure ActiveWhere (Okok : Boolean);
    procedure DateElipsisclick(Sender: TObject);
    procedure Change(Sender: TObject);
    procedure ExitEdit(Sender: TObject);

  end ;

Implementation

procedure TOF_MUL_REGLTIJSS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_MUL_REGLTIJSS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_MUL_REGLTIJSS.OnUpdate ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 09/06/2004
Modifié le ... :   /  /    
Description .. : procédure OnLoad
Suite ........ : procédure de chargement des données de la fiche Lance 
Suite ........ : ActiveWhere
Mots clefs ... : IJSS
*****************************************************************}
procedure TOF_MUL_REGLTIJSS.OnLoad ;
var
  Okok : Boolean;

begin
  Inherited ;
  Okok := TRUE;
  ActiveWhere (Okok);
end ; {fin OnLoad}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 09/06/2004
Modifié le ... : 09/06/2004
Description .. : procédure OnArgument
Suite ........ : chargement de la fiche
Suite ........ : Les dates servant de critère sont alimentées par défaut
Suite ........ :  de "date du jour - 1an" à  "date du jour"
Mots clefs ... : IJSS
*****************************************************************}
procedure TOF_MUL_REGLTIJSS.OnArgument (S : String ) ;
var
  Ouvrir                        : TToolbarButton97;
  DateDeb,DateFin               : THEdit;
{$IFNDEF EAGLCLIENT}
  FListe                        : THDBGrid;
{$ELSE}
  FListe                        : THGrid;
{$ENDIF}
//PT1  Defaut                        : THEdit;
  Defaut                        : THMultiValComboBox;

begin
  Inherited ;
  Ouvrir := TToolbarButton97(GetControl('BOUVRIR'));
  if Ouvrir <> NIL then
  begin
    Ouvrir.Visible := True;
    Ouvrir.OnClick := CalculerClick;
  end;

{$IFNDEF EAGLCLIENT}
FListe := THDBGrid(GetControl('FListe'));
{$ELSE}
FListe := THGrid(GetControl('FListe'));
{$ENDIF}
if FListe<>nil then Fliste.OnDblClick:=CalculerClick;

  WW:=THEdit (GetControl ('XX_WHERE'));

  DateDeb:= THEdit(getcontrol('XX_VARIABLED'));
  DateFin:= THEdit(getcontrol('XX_VARIABLED_'));

  if (DateDeb <> NIL) and (DateFin <> NIL) then
  begin
    DateDeb.OnElipsisClick := DateElipsisclick;
    DateDeb.OnExit := Change;
    DateFin.OnElipsisClick := DateElipsisclick;
    DateFin.OnExit := Change;
  end;
  if DateFin <> NIL then
  begin
    DateFin.text := DateToStr(Date);
    WDateFin := Date;
  end;
  if DateDeb <> NIL then
  begin
    WDateDeb := PlusMois(StrToDate(DateFin.text),-12);
    DateDeb.text := DateToStr(WDateDeb);
  end;

//PT1  Defaut:=ThEdit(getcontrol('PCN_SALARIE'));
  Defaut:=THMultiValComboBox(getcontrol('SALARIE'));
  if Defaut<>nil then Defaut.OnExit:=ExitEdit;
end ; { fin OnArgument}

procedure TOF_MUL_REGLTIJSS.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_MUL_REGLTIJSS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_MUL_REGLTIJSS.OnCancel () ;
begin
  Inherited ;
end ;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 09/06/2004
Modifié le ... : 09/06/2004
Description .. : procédure CalculerClick
Suite ........ : Charge les paramètre de lancement de la fiche REGLTIJSS
Suite ........ : et  la lance.
Mots clefs ... : IJSS
*****************************************************************}
procedure TOF_MUL_REGLTIJSS.CalculerClick(Sender: TObject);
var
  st :                          string;
begin
  st := '';
  if (GetField('PCN_SALARIE') <> '') and (GetField('PCN_SALARIE') <> NULL)   then
  begin
    st := GetField('PCN_SALARIE')+';';
    st := st + GetField('PCN_ETABLISSEMENT')+';';
    st := st + GetField('PSA_LIBELLE')+';';
    st := st + GetField('PSA_PRENOM')+';';
    st := st + DateToStr(WdateDeb)+';'+DateToStr(WDateFin)+';';
    st := st + GetField('PSA_NUMEROSS')+';';
    st := st + DateToStr(GetField('PSA_DATENAISSANCE'));


    AglLanceFiche ('PAY','REGLTIJSS', '', '' , st);
  end;
end; {fin CalculerClick}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 09/06/2004
Modifié le ... : 09/06/2004
Description .. : procédure ActiveWher
Suite ........ : Complète la requête.
Mots clefs ... : IJSS
*****************************************************************}
procedure TOF_MUL_REGLTIJSS.ActiveWhere(Okok : Boolean);
var
   St                                        : String;
// d PT1
   Salarie                                   : String;
   StCondition                               : String;
   Indst                                     : integer;
// f PT1
begin
  WW.Text := '';

  st := '';
// d PT1
  st := 'PCN_SALARIE  in (select PCN_SALARIE FROM ABSENCESALARIE WHERE ';
//st := st + '((PCN_DATEDEBUTABS >="'+UsDateTime(WDateDeb)+'" AND '+
  st := st + 'PCN_GESTIONIJSS = "X" AND ';
  st := st + '(((PCN_DATEDEBUTABS >="'+UsDateTime(WDateDeb)+'" AND '+
          'PCN_DATEDEBUTABS <="'+UsDateTime(WDateFin)+'") AND '+
         '(PCN_DATEFINABS >="'+UsDateTime(WDateDeb)+'" AND '+
//       'PCN_DATEFINABS <="'+UsDateTime(WDateFin)+'"))) ';
          'PCN_DATEFINABS <="'+UsDateTime(WDateFin)+'")) ';

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
            st := st + 'PCN_SALARIE = "'+ StCondition+'"';
            Indst := 1;
       end;
       st := st +')';
  end;
  st := st + '))';
// f PT1


  if St <> '' then WW.Text := st;
  if Q_Mul <> NIL then
  begin
    TFMul(Ecran).SetDBListe('PGMULABSIJSS');
  end;
end;  {fin ActiveWhere}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 09/06/2004
Modifié le ... :   /  /    
Description .. : procédure DateElipsisClick
Suite ........ : Affichage du calendrier qd double click sur champ date
Mots clefs ... : IJSS
*****************************************************************}
procedure TOF_MUL_REGLTIJSS.DateElipsisclick(Sender: TObject);
var
  key : char;
begin
  key := '*';
  ParamDate (Ecran, Sender, Key);
end;  {fin DateElipsisclick}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 09/06/2004
Modifié le ... :   /  /    
Description .. : Procédure Change
Suite ........ : contrôle de la validité des dates servant de critère
Mots clefs ... : IJSS
*****************************************************************}
procedure TOF_MUL_REGLTIJSS.Change(Sender: TObject);
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

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 23/06/2004
Modifié le ... : 23/06/2004
Description .. : procédure ExitEdit
Suite ........ : 
Suite ........ : formatage du matricule
Suite ........ :
Mots clefs ... : PAIE, IJSS
*****************************************************************}
procedure TOF_MUL_REGLTIJSS.ExitEdit(Sender: TObject);
var
   edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and
       (length(Edit.text)<11) and
       (isnumeric(edit.text)) then
      edit.text:=AffectDefautCode(edit,10);
end;  { fin ExitEdit}

Initialization
  registerclasses ( [ TOF_MUL_REGLTIJSS ] ) ;
end.
