{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 06/05/2004
Modifié le ... : 09/06/2004
Description .. : Source TOF de la FICHE : SUIVIIJSS_MUL ()
Suite ........ : Suivi des IJSS - Lance la fiche MvtAbsence
Mots clefs ... : TOF;SUIVIIJSS_MUL; IJSS 
*****************************************************************}
{
 PT1  : 23/02/2005 MF V_60 Suite erreur "utilisation incorrecte mot clé)))"
                           correction de la procédure ActiveWhere
PT2   19/06/2006 SB V_65 FQ 13231 Retrait des mvt absences annulées
}
Unit SUIVIIJSS_MUL_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
//unused     Fe_Main,
//unused     HDB,
     mul,
{$else}
     eMul,
//unused     MainEAgl,
     uTob,
{$ENDIF}
     Entpaie,
//unused     forms,
     ParamDat,
     sysutils,
//unused     ComCtrls,
     HTB97,
     HCtrls,
     HEnt1,
     HMsgBox,
     HQry,
//unused     PGOutils,
     PGOutils2,
     UTOF ; 

Type
  TOF_SUIVIIJSS_MUL = Class (TOF)
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
    IJSSNonSoldee                 : TCheckBox;
    procedure Change(Sender: TObject);
    procedure DateElipsisclick(Sender: TObject);
    procedure ActiveWhere (Okok : Boolean);
    procedure ExitEdit(Sender: TObject);

  end ;

Implementation

procedure TOF_SUIVIIJSS_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_SUIVIIJSS_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_SUIVIIJSS_MUL.OnUpdate ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 09/06/2004
Modifié le ... :   /  /    
Description .. : procédure de chargement des données de la fiche
Suite ........ : Lance ActiveWhere
Mots clefs ... : IJSS
*****************************************************************}
procedure TOF_SUIVIIJSS_MUL.OnLoad ;
var
  Okok : Boolean;

begin
  Inherited ;
  Okok := TRUE;
  ActiveWhere (Okok);
end ; {fin Oload}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - Mf
Créé le ...... : 09/06/2004
Modifié le ... : 09/06/2004
Description .. : Procédure OnArgument
Suite ........ : Chargement de la fiche.
Suite ........ : Les dates servant de critère sont alimentées par défaut
Suite ........ : de "date du jour - 1an" à  "date du jour"
Mots clefs ... : IJSS
*****************************************************************}
procedure TOF_SUIVIIJSS_MUL.OnArgument (S : String ) ;
var
  DateDeb,DateFin : THEdit;
  Defaut                        : THEdit;

begin
  Inherited ;
  TToolbarbutton97(getcontrol('BINSERT')).Visible := False;
  WW:=THEdit (GetControl ('XX_WHERE'));

  DateDeb:= THEdit(getcontrol('XX_VARIABLED'));
  DateFin:= THEdit(getcontrol('XX_VARIABLED_'));
  IJSSNonSoldee := TCheckBox(getcontrol('IJSSSOLDEE'));
  IJSSNonSoldee.checked := True;
  
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
  Defaut:=ThEdit(getcontrol('PCN_SALARIE'));
  if Defaut<>nil then Defaut.OnExit:=ExitEdit;

end ;   {fin OnArgument}

procedure TOF_SUIVIIJSS_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_SUIVIIJSS_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_SUIVIIJSS_MUL.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - Mf
Créé le ...... : 09/06/2004
Modifié le ... :   /  /
Description .. : procédure DateElispsisClick
Suite ........ : Affichage du calendrier qd double click sur champ date
Mots clefs ... : IJSS
*****************************************************************}
procedure TOF_SUIVIIJSS_MUL.DateElipsisclick(Sender: TObject);
var
  key : char;
begin
  key := '*';
  ParamDate (Ecran, Sender, Key);
end;  {fin DateElipsisclick}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - Mf
Créé le ...... : 09/06/2004
Modifié le ... : 09/06/2004
Description .. : procédure Change
Suite ........ : contrôle de la validité des dates servant de critère
Mots clefs ... : IJSS
*****************************************************************}
procedure TOF_SUIVIIJSS_MUL.Change(Sender: TObject);
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
Auteur  ...... : PAIE - Mf
Créé le ...... : 09/06/2004
Modifié le ... : 09/06/2004
Description .. :  procédure ActiveWhere
Suite ........ : Complète la requête.
Mots clefs ... : IJSS
*****************************************************************}
procedure TOF_SUIVIIJSS_MUL.ActiveWhere(Okok : Boolean);
var
   St                                          : STring;
   WTypeAbs                                    : string;
   i                                           : integer;
// d PT1
   Q                                           : Tquery;
   StTypeConge                                 : string;
// f PT1
begin
  WW.Text := '';

  WTypeAbs := GetControlText('TYPEABS');

// d PT1
  if (Length (WtypeAbs) = 0) or (WTypeAbs = '<<Tous>>') then
    st := st + ' AND (PMA_TYPEABS <>""'
  else
  begin
    st := st + ' AND (';
    i := 1;
    while (i < Length(WTypeAbs)) do
    begin
      st := st + 'PMA_TYPEABS = "'+ Copy(WTypeAbs,i,3)+'"';
      if (i+3 < Length(WTypeAbs)) then
        st := st + ' OR ';
      i := i + 4;
    end;
  end;
  st := st + ')';
  StTypeConge := '';
  Q:= OpenSql('SELECT PMA_MOTIFABSENCE FROM MOTIFABSENCE '+
              'WHERE ##PMA_PREDEFINI## AND PMA_GESTIONIJSS = "X"'+st,True);
  While  not Q.eof do
  begin
    if (StTypeConge = '') then
      StTypeConge := ' AND (PCN_TYPECONGE = '
    else
      StTypeConge := StTypeConge + ' OR PCN_TYPECONGE = ' ;

    StTypeConge := StTypeConge+'"'+ Q.FindField('PMA_MOTIFABSENCE').AsString+'"';
    Q.Next;
  end;
  if (StTypeConge <> '') then StTypeConge := StTypeConge + ')';
  ferme (Q);
// f PT1


  St := ' PCN_ETATPOSTPAIE <> "NAN" AND '; { PT2 }

  if (IJSSNonSoldee.checked = True) then
   st := St + '((PCN_IJSSSOLDEE <> "X") AND '
  else
   st := St + '(';

  st := st +'(PCN_DATEDEBUTABS >="'+UsDateTime(WDateDeb)+'" AND '+
          'PCN_DATEDEBUTABS <="'+UsDateTime(WDateFin)+'") AND '+
         '(PCN_DATEFINABS >="'+UsDateTime(WDateDeb)+'" AND '+
          'PCN_DATEFINABS <="'+UsDateTime(WDateFin)+'")';
// d PT1
  if (StTypeConge <> '') then
    st := st + StTypeConge
  else
    st := st + ' AND (PCN_TYPECONGE ="")';
// f PT1

  st := st + ')';

  if St <> '' then WW.Text := st;
  if Q_Mul <> NIL then
  begin
    TFMul(Ecran).SetDBListe('PGMULMVTIJSS');
  end;
end;  {fin ActiveWhere}

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
procedure TOF_SUIVIIJSS_MUL.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and
       (length(Edit.text)<11) and
       (isnumeric(edit.text)) then
      edit.text:=AffectDefautCode(edit,10);
end; { fin  ExitEdit}

Initialization
  registerclasses ( [ TOF_SUIVIIJSS_MUL ] ) ;
end.
